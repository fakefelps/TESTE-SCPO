# PROTEÇÃO ANTI-LOOP — deve ser a PRIMEIRA coisa no arquivo
import multiprocessing
multiprocessing.freeze_support()

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
import json
import os
import sys
import urllib.request
import urllib.parse
import http.cookiejar
from datetime import datetime
from dateutil.relativedelta import relativedelta
from io import BytesIO

# ─── Constantes fixas ────────────────────────────────────────────────────────
URL_BASE      = "https://scpo.mte.gov.br"
URL_LOGIN     = "https://scpo.mte.gov.br/Default.aspx"
URL_COMUNICAR = "https://scpo.mte.gov.br/DeclaracaoPreviaObra/Comunicar.aspx"
LOGIN_CPF     = "038.144.411-25"
EMAIL_FIXO    = "joaovitorcabral94@gmail.com"
TELEFONE_FIXO = "6299266-5923"
EMP_PRINCIPAL = "0"
EMP_TERCEIROS = "5"

# ─── Paleta padrão Morais ─────────────────────────────────────────────────────
COR_BG       = "#1e2a3a"
COR_LOG      = "#131c26"
COR_CAMPO    = "#2a3f55"
COR_BOTAO    = "#2e86de"
COR_BARRA    = "#4cd964"
COR_TEXTO    = "#ffffff"
COR_LABEL    = "#90adc4"
COR_LOG_TEXT = "#7ec8a0"

# ─── Config persistente ───────────────────────────────────────────────────────
CONFIG_PATH = os.path.join(
    os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__),
    "scpo_config.json"
)

def carregar_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {"senha": "SCPO123"}

def salvar_config(cfg: dict):
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f)

# ─── ViaCEP ───────────────────────────────────────────────────────────────────
def buscar_cep(cep: str) -> dict:
    cep_limpo = cep.replace("-", "").replace(".", "").strip()
    if len(cep_limpo) != 8:
        return {}
    try:
        url = f"https://viacep.com.br/ws/{cep_limpo}/json/"
        with urllib.request.urlopen(url, timeout=5) as r:
            dados = json.loads(r.read().decode())
        return {} if "erro" in dados else dados
    except Exception:
        return {}

# ─── Lógica de negócio ────────────────────────────────────────────────────────
def montar_nome_obra(rua: str, quadra: str, lote: str) -> str:
    return f"RESIDENCIAL {rua.upper()} QUADRA {quadra.upper()} LOTE {lote.upper()}"

def gerar_observacao(rua: str, quadra: str, lote: str,
                     esquina: bool, rua2: str, casas: list) -> str:
    partes = []
    for c in casas:
        label = f"CASA {c['numero']}"
        if esquina and c.get("rua"):
            label += f" SITUADA NA {c['rua'].upper()}"
        partes.append(label)
    casas_str = ", ".join(partes)
    if not esquina:
        return (f"OBRA RESIDENCIAL UNIFAMILIAR SITUADA NA {rua.upper()} "
                f"QUADRA {quadra.upper()} LOTE {lote.upper()} "
                f"COMPOSTA POR: {casas_str}")
    return (f"OBRA RESIDENCIAL UNIFAMILIAR SITUADA NA {rua.upper()} "
            f"E {rua2.upper()} QUADRA {quadra.upper()} LOTE {lote.upper()} "
            f"COMPOSTA POR: {casas_str}")

def data_termino_auto() -> str:
    return (datetime.today() + relativedelta(months=1)).strftime("%d/%m/%Y")

# ─── Parser HTML mínimo (extrai campos ASP.NET sem BeautifulSoup) ─────────────
def _extrair_campo(html: str, name: str) -> str:
    """Extrai value de <input name='X' value='Y'>"""
    import re
    padrao = rf'<input[^>]+name="{re.escape(name)}"[^>]+value="([^"]*)"'
    m = re.search(padrao, html, re.IGNORECASE)
    if m:
        return m.group(1)
    # tenta ordem invertida (value antes de name)
    padrao2 = rf'<input[^>]+value="([^"]*)"[^>]+name="{re.escape(name)}"'
    m2 = re.search(padrao2, html, re.IGNORECASE)
    return m2.group(1) if m2 else ""

def _extrair_campos_asp(html: str) -> dict:
    """Extrai __VIEWSTATE, __EVENTVALIDATION e outros campos ocultos ASP.NET."""
    campos = {}
    for campo in ["__VIEWSTATE", "__VIEWSTATEGENERATOR", "__EVENTVALIDATION",
                  "__EVENTTARGET", "__EVENTARGUMENT"]:
        campos[campo] = _extrair_campo(html, campo)
    return campos

# ─── Sessão HTTP (substitui Selenium por completo) ───────────────────────────
class SessaoSCPO:
    """
    Gerencia a sessão HTTP com o SCPO usando apenas urllib da stdlib.
    Sem Selenium, sem ChromeDriver, sem navegador.
    """
    HEADERS = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/124.0.0.0 Safari/537.36"),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
    }

    def __init__(self):
        self.jar = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(self.jar),
            urllib.request.HTTPSHandler(),
        )
        self.opener.addheaders = list(self.HEADERS.items())

    def get(self, url: str) -> str:
        """GET — retorna HTML decodificado."""
        req = urllib.request.Request(url, headers=self.HEADERS)
        with self.opener.open(req, timeout=20) as r:
            raw = r.read()
        # descomprime gzip se necessário
        encoding = ""
        for h in ["Content-Encoding"]:
            try: encoding = r.headers.get(h, "")
            except Exception: pass
        if encoding == "gzip":
            import gzip
            raw = gzip.decompress(raw)
        return raw.decode("utf-8", errors="replace")

    def post(self, url: str, campos: dict, referer: str = "") -> str:
        """POST com campos urlencoded — retorna HTML decodificado."""
        dados = urllib.parse.urlencode(campos).encode("utf-8")
        headers = dict(self.HEADERS)
        headers["Content-Type"] = "application/x-www-form-urlencoded"
        headers["Referer"] = referer or url
        req = urllib.request.Request(url, data=dados, headers=headers, method="POST")
        with self.opener.open(req, timeout=30) as r:
            raw = r.read()
        encoding = ""
        try: encoding = r.headers.get("Content-Encoding", "")
        except Exception: pass
        if encoding == "gzip":
            import gzip
            raw = gzip.decompress(raw)
        return raw.decode("utf-8", errors="replace")

    def get_imagem_captcha(self, url_img: str) -> bytes:
        """Baixa imagem do captcha como bytes."""
        req = urllib.request.Request(url_img, headers=self.HEADERS)
        with self.opener.open(req, timeout=10) as r:
            return r.read()


# ─── Automação principal (sem navegador) ─────────────────────────────────────
def executar_scpo(dados: dict, senha: str, step_cb, log_cb, done_cb):
    import traceback, re
    sessao = SessaoSCPO()
    try:
        # ── 1. Carregar página de login ───────────────────────────────────────
        log_cb("Abrindo pagina de login do SCPO...")
        step_cb(5, "Carregando pagina de login")
        html_login = sessao.get(URL_LOGIN)
        asp = _extrair_campos_asp(html_login)

        # ── 2. Extrair e exibir imagem do captcha ─────────────────────────────
        log_cb("Extraindo captcha...")
        step_cb(10, "Aguardando captcha")

        # URL da imagem do captcha está em <img src="CaptchaImage.axd?guid=...">
        m = re.search(r'src="(CaptchaImage\.axd\?[^"]+)"', html_login, re.IGNORECASE)
        if m:
            url_captcha = URL_BASE + "/" + m.group(1)
            img_bytes = sessao.get_imagem_captcha(url_captcha)
            # Envia imagem para UI exibir
            dados["captcha_img_bytes"] = img_bytes
            dados["evento_mostrar_captcha"].set()

        log_cb(">>> Veja o captcha no app, digite e clique 'Confirmar'")
        dados["evento_captcha"].wait()  # aguarda usuário digitar
        captcha_digitado = dados.get("captcha_valor", "")
        log_cb(f"Captcha recebido: {captcha_digitado}")

        # ── 3. POST de login ──────────────────────────────────────────────────
        step_cb(20, "Efetuando login")
        log_cb("Enviando login...")

        campos_login = {
            "__VIEWSTATE":          asp.get("__VIEWSTATE", ""),
            "__VIEWSTATEGENERATOR": asp.get("__VIEWSTATEGENERATOR", ""),
            "__EVENTVALIDATION":    asp.get("__EVENTVALIDATION", ""),
            "__EVENTTARGET":        "",
            "__EVENTARGUMENT":      "",
            "ct100$PlaceHolderConteudo$txtCPF":     LOGIN_CPF,
            "ct100$PlaceHolderConteudo$txtSenha":   senha,
            "ct100$PlaceHolderConteudo$txtCaptcha": captcha_digitado,
            "ct100$PlaceHolderConteudo$btnLogin":   "Entrar",
        }
        html_pos_login = sessao.post(URL_LOGIN, campos_login, referer=URL_LOGIN)

        # Verifica se login falhou
        if "txtCPF" in html_pos_login or "Entrar" in html_pos_login[:2000]:
            # Ainda na página de login — captcha ou senha errados
            if "captcha" in html_pos_login.lower() or "codigo" in html_pos_login.lower():
                raise Exception("Captcha incorreto. Tente novamente.")
            if "senha" in html_pos_login.lower() and "incorret" in html_pos_login.lower():
                raise Exception("Senha incorreta.")
            raise Exception("Login falhou. Verifique CPF, senha e captcha.")

        # Verifica pedido de troca de senha
        if "RedefinirSenha" in html_pos_login or "alterar" in html_pos_login.lower():
            log_cb("⚠ Site solicita alteracao de senha.")
            log_cb("Altere a senha no navegador e clique 'Senha alterada — Continuar'.")
            dados["evento_senha"].wait()
            nova_senha = dados.get("nova_senha", senha)
            salvar_config({"senha": nova_senha})
            senha = nova_senha
            log_cb("Nova senha salva. Prosseguindo...")
            html_pos_login = sessao.get(URL_LOGIN)

        log_cb("Login realizado com sucesso!")

        # ── 4. Navegar para Comunicar Obra ────────────────────────────────────
        step_cb(30, "Abrindo pagina Comunicar Obra")
        log_cb("Navegando para Comunicar Obra...")
        html_comunicar = sessao.get(URL_COMUNICAR)
        asp2 = _extrair_campos_asp(html_comunicar)

        # ── 5. POST — marcar sem CNPJ e preencher CPF ─────────────────────────
        step_cb(40, "Identificando empresa")
        log_cb("Marcando 'Obra nao tem CNPJ' e informando CPF...")

        # Simula clique no checkbox (postback ASP.NET)
        campos_empresa = {
            "__VIEWSTATE":          asp2.get("__VIEWSTATE", ""),
            "__VIEWSTATEGENERATOR": asp2.get("__VIEWSTATEGENERATOR", ""),
            "__EVENTVALIDATION":    asp2.get("__EVENTVALIDATION", ""),
            "__EVENTTARGET":        "ct100$PlaceHolderConteudo$chkObraSemCNPJ",
            "__EVENTARGUMENT":      "",
            "ct100$PlaceHolderConteudo$chkObraSemCNPJ": "on",
            "ct100$PlaceHolderConteudo$txtCPFProprietarioObra": LOGIN_CPF,
        }
        html_apos_checkbox = sessao.post(URL_COMUNICAR, campos_empresa, referer=URL_COMUNICAR)
        asp3 = _extrair_campos_asp(html_apos_checkbox)

        # ── 6. POST — clicar "Comunicar Obra" (abre formulário) ───────────────
        step_cb(45, "Abrindo formulario principal")
        log_cb("Clicando em Comunicar Obra...")

        campos_declarar = {
            "__VIEWSTATE":          asp3.get("__VIEWSTATE", ""),
            "__VIEWSTATEGENERATOR": asp3.get("__VIEWSTATEGENERATOR", ""),
            "__EVENTVALIDATION":    asp3.get("__EVENTVALIDATION", ""),
            "__EVENTTARGET":        "",
            "__EVENTARGUMENT":      "",
            "ct100$PlaceHolderConteudo$chkObraSemCNPJ": "on",
            "ct100$PlaceHolderConteudo$txtCPFProprietarioObra": LOGIN_CPF,
            "ct100$PlaceHolderConteudo$btnDeclararObra": "Comunicar Obra",
        }
        html_form = sessao.post(URL_COMUNICAR, campos_declarar, referer=URL_COMUNICAR)
        asp4 = _extrair_campos_asp(html_form)

        # ── 7. POST — formulário principal ────────────────────────────────────
        step_cb(60, "Preenchendo formulario")
        log_cb("Preenchendo dados da obra...")
        log_cb(f"  Nome: {dados['nome_obra']}")
        log_cb(f"  Obs: {dados['observacao'][:80]}...")

        # Converte data DD/MM/YYYY para o formato esperado pelo site
        dt_ini = dados["data_inicio"]   # DD/MM/YYYY
        dt_fim = dados["data_termino"]  # DD/MM/YYYY

        campos_form = {
            "__VIEWSTATE":          asp4.get("__VIEWSTATE", ""),
            "__VIEWSTATEGENERATOR": asp4.get("__VIEWSTATEGENERATOR", ""),
            "__EVENTVALIDATION":    asp4.get("__EVENTVALIDATION", ""),
            "__EVENTTARGET":        "",
            "__EVENTARGUMENT":      "",
            # Dados da obra
            "ct100$PlaceHolderConteudo$txtNomeObra":         dados["nome_obra"],
            "ct100$PlaceHolderConteudo$chkContratantePrincipal": "on",  # Sim
            "ct100$PlaceHolderConteudo$txtEmailObra":        EMAIL_FIXO,
            "ct100$PlaceHolderConteudo$txtTelefoneObra":     TELEFONE_FIXO,
            # Endereço
            "ct100$PlaceHolderConteudo$txtCEPObra":          dados["cep"],
            "ct100$PlaceHolderConteudo$txtComplementoObra":  f"QUADRA {dados['quadra']} LOTE {dados['lote']}",
            # Detalhamento
            "ct100$PlaceHolderConteudo$ddlClasseCNAE":       "4120-4",
            "ct100$PlaceHolderConteudo$ddlSubclasse":        "00",
            "ct100$PlaceHolderConteudo$ddlTipoConstrucao":   "1",   # Edifício
            "ct100$PlaceHolderConteudo$rdTipoObra":          "2",   # Privada
            "ct100$PlaceHolderConteudo$rdCaracteristicaObra":"1",   # Construção
            "ct100$PlaceHolderConteudo$txtDescricaoObra":    dados["observacao"],
            # FGTS — Não
            "ct100$PlaceHolderConteudo$rdObraFGTS":          "2",
            # Datas
            "ct100$PlaceHolderConteudo$txtDataInicioObra":   dt_ini,
            "ct100$PlaceHolderConteudo$txtDataTerminoObra":  dt_fim,
            # Empregados
            "ct100$PlaceHolderConteudo$txtNumEmpregadosPrincipal": EMP_PRINCIPAL,
            "ct100$PlaceHolderConteudo$txtNumEmpregadosTerceiros": EMP_TERCEIROS,
            # Botão submit
            "ct100$PlaceHolderConteudo$btnDeclararObra2":    "Comunicar Obra",
        }

        # Segunda rua (esquina)
        if dados.get("esquina") and dados.get("rua2"):
            campos_form["ct100$PlaceHolderConteudo$txtLogradouro2Obra"] = dados["rua2"]

        step_cb(80, "Submetendo formulario")
        log_cb("Submetendo...")
        html_resultado = sessao.post(URL_COMUNICAR, campos_form, referer=URL_COMUNICAR)

        # ── 8. Verificar resultado ────────────────────────────────────────────
        if "erro" in html_resultado.lower() and "obrigat" in html_resultado.lower():
            raise Exception("Formulario retornou erro de validacao. "
                            "Verifique os dados e tente novamente.")

        # Procura numero de protocolo/recibo na pagina de confirmacao
        m_prot = re.search(r'(\d{6,})', html_resultado)
        protocolo = m_prot.group(1) if m_prot else "N/D"
        log_cb(f"✔ Obra comunicada! Protocolo: {protocolo}")

        step_cb(100, "Concluido!")
        done_cb(True, f"SCPO preenchido com sucesso!\nProtocolo: {protocolo}")

    except Exception as e:
        tb = traceback.format_exc()
        log_cb(f"✖ ERRO: {e}")
        for linha in tb.splitlines():
            log_cb(linha)
        try:
            log_dir = os.path.join(
                os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "SCPOApp")
            os.makedirs(log_dir, exist_ok=True)
            log_path = os.path.join(log_dir, f"erro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            with open(log_path, "w", encoding="utf-8") as lf:
                lf.write(f"Erro: {e}\n\n{tb}")
            log_cb(f"Relatorio salvo em: {log_path}")
        except Exception:
            pass
        done_cb(False, str(e))


# ─── Interface Tkinter ────────────────────────────────────────────────────────
class AppSCPO(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SCPO Automation — Berçan Projetos")
        self.configure(bg=COR_BG)
        self.resizable(False, False)
        self.config_dados    = carregar_config()
        self.var_esquina     = tk.BooleanVar(value=False)
        self.var_data_ini    = tk.StringVar()
        self.var_senha       = tk.StringVar(value=self.config_dados.get("senha", "SCPO123"))
        self._evento_captcha       = threading.Event()
        self._evento_mostrar_captcha = threading.Event()
        self._evento_senha         = threading.Event()
        self._nova_senha_tmp       = ""
        self._build_ui()
        # Inicia watcher que mostra captcha quando disponível
        threading.Thread(target=self._aguardar_captcha_img, daemon=True).start()

    # ── Watcher: exibe janela de captcha quando a imagem chegar ──────────────
    def _aguardar_captcha_img(self):
        while True:
            self._evento_mostrar_captcha.wait()
            self._evento_mostrar_captcha.clear()
            img_bytes = self._dados_run.get("captcha_img_bytes", b"")
            if img_bytes:
                self.after(0, self._mostrar_janela_captcha, img_bytes)

    def _mostrar_janela_captcha(self, img_bytes: bytes):
        """Abre popup com imagem do captcha e campo para digitar."""
        try:
            from PIL import Image, ImageTk
            img = Image.open(BytesIO(img_bytes))
            # Aumenta para facilitar leitura
            img = img.resize((img.width * 3, img.height * 3), Image.NEAREST)
            photo = ImageTk.PhotoImage(img)

            win = tk.Toplevel(self)
            win.title("Codigo de Segurança")
            win.configure(bg=COR_BG)
            win.grab_set()
            win.resizable(False, False)

            tk.Label(win, text="Digite o codigo abaixo:", bg=COR_BG,
                     fg=COR_LABEL, font=("Consolas", 10)).pack(pady=(12, 4))
            lbl_img = tk.Label(win, image=photo, bg=COR_BG)
            lbl_img.image = photo  # mantém referência
            lbl_img.pack(padx=16, pady=4)

            var_cap = tk.StringVar()
            ent = tk.Entry(win, textvariable=var_cap, bg=COR_CAMPO, fg=COR_TEXTO,
                           font=("Consolas", 14, "bold"), justify="center",
                           insertbackground=COR_TEXTO, relief="flat", width=12)
            ent.pack(pady=8)
            ent.focus()

            def confirmar():
                self._dados_run["captcha_valor"] = var_cap.get().strip()
                self._evento_captcha.set()
                win.destroy()

            tk.Button(win, text="Confirmar", bg=COR_BOTAO, fg=COR_TEXTO,
                      font=("Consolas", 11, "bold"), relief="flat",
                      cursor="hand2", command=confirmar).pack(pady=(0, 12))
            ent.bind("<Return>", lambda e: confirmar())

        except ImportError:
            # Pillow indisponível — fallback: simpledialog
            val = simpledialog.askstring("Codigo de Segurança",
                                          "Digite o codigo de segurança:", parent=self)
            self._dados_run["captcha_valor"] = val or ""
            self._evento_captcha.set()

    # ── Construção da UI ──────────────────────────────────────────────────────
    def _build_ui(self):
        PAD = 12
        tk.Label(self, text="SCPO AUTOMATION", bg=COR_BG, fg=COR_TEXTO,
                 font=("Consolas", 14, "bold")).pack(pady=(PAD, 2))
        tk.Label(self, text="Bercan Projetos — Morais Engenharia", bg=COR_BG,
                 fg=COR_LABEL, font=("Consolas", 9)).pack(pady=(0, PAD))

        frame = tk.Frame(self, bg=COR_BG, padx=PAD, pady=4)
        frame.pack(fill="x")

        def lbl(row, txt):
            tk.Label(frame, text=txt, bg=COR_BG, fg=COR_LABEL,
                     font=("Consolas", 9), anchor="w"
                     ).grid(row=row, column=0, sticky="w", pady=3, padx=(0, 8))

        def ent(row, var=None, width=32):
            e = tk.Entry(frame, textvariable=var, bg=COR_CAMPO, fg=COR_TEXTO,
                         insertbackground=COR_TEXTO, font=("Consolas", 10),
                         relief="flat", width=width)
            e.grid(row=row, column=1, sticky="w", pady=3)
            return e

        # CEP
        lbl(0, "CEP:")
        frame_cep = tk.Frame(frame, bg=COR_BG)
        frame_cep.grid(row=0, column=1, sticky="w")
        self.ent_cep = tk.Entry(frame_cep, bg=COR_CAMPO, fg=COR_TEXTO,
                                 insertbackground=COR_TEXTO, font=("Consolas", 10),
                                 relief="flat", width=12)
        self.ent_cep.pack(side="left")
        tk.Button(frame_cep, text="Buscar", bg=COR_CAMPO, fg=COR_LOG_TEXT,
                  font=("Consolas", 8), relief="flat", cursor="hand2",
                  command=self._buscar_cep).pack(side="left", padx=6)
        self._lbl_cep_status = tk.Label(frame_cep, text="", bg=COR_BG,
                                         fg=COR_LOG_TEXT, font=("Consolas", 8))
        self._lbl_cep_status.pack(side="left")

        lbl(1, "Rua Principal:")
        self.ent_rua = ent(1)

        lbl(2, "Quadra:");    self.ent_quadra    = ent(2, width=10)
        lbl(3, "Lote:");      self.ent_lote      = ent(3, width=10)
        lbl(4, "Nr Casas:");  self.ent_num_casas = ent(4, width=5)
        self.ent_num_casas.bind("<FocusOut>", self._atualizar_casas_esquina)

        tk.Checkbutton(frame, text="Lote de esquina (frente para 2 ruas)",
                       variable=self.var_esquina, command=self._toggle_esquina,
                       bg=COR_BG, fg=COR_LABEL, selectcolor=COR_CAMPO,
                       activebackground=COR_BG, font=("Consolas", 9)
                       ).grid(row=5, column=0, columnspan=2, sticky="w", pady=3)

        lbl(6, "Rua 2 (esquina):")
        self.ent_rua2 = tk.Entry(frame, bg=COR_CAMPO, fg=COR_TEXTO,
                                  insertbackground=COR_TEXTO, font=("Consolas", 10),
                                  relief="flat", width=32, state="disabled")
        self.ent_rua2.grid(row=6, column=1, sticky="w", pady=3)

        self.frame_casas = tk.Frame(frame, bg=COR_BG)
        self.frame_casas.grid(row=7, column=0, columnspan=2, sticky="w")
        self._entries_rua_casa = []

        lbl(8, "Data de Inicio:")
        self.ent_data_ini = tk.Entry(frame, textvariable=self.var_data_ini,
                                      bg=COR_CAMPO, fg=COR_TEXTO,
                                      insertbackground=COR_TEXTO, font=("Consolas", 10),
                                      relief="flat", width=12)
        self.ent_data_ini.grid(row=8, column=1, sticky="w", pady=3)
        tk.Label(frame, text="(DD/MM/YYYY)", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 8)).grid(row=8, column=1, sticky="e")

        lbl(9, "Data de Termino:")
        tk.Label(frame, text=data_termino_auto(), bg=COR_BG,
                 fg=COR_LOG_TEXT, font=("Consolas", 9)
                 ).grid(row=9, column=1, sticky="w")
        tk.Label(frame, text="(automatico: +1 mes)", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 8)).grid(row=9, column=1, sticky="e")

        lbl(10, "Pasta Download:")
        self.ent_pasta = tk.Entry(frame, bg=COR_CAMPO, fg=COR_TEXTO,
                                   insertbackground=COR_TEXTO, font=("Consolas", 10),
                                   relief="flat", width=32)
        self.ent_pasta.insert(0, os.path.expanduser("~\\Downloads"))
        self.ent_pasta.grid(row=10, column=1, sticky="w", pady=3)

        tk.Frame(frame, height=1, bg=COR_CAMPO
                 ).grid(row=11, column=0, columnspan=2, sticky="ew", pady=8)

        lbl(12, "Senha SCPO:")
        frame_senha = tk.Frame(frame, bg=COR_BG)
        frame_senha.grid(row=12, column=1, sticky="w")
        self.ent_senha = tk.Entry(frame_senha, textvariable=self.var_senha,
                                   bg=COR_CAMPO, fg=COR_TEXTO, show="*",
                                   insertbackground=COR_TEXTO, font=("Consolas", 10),
                                   relief="flat", width=20)
        self.ent_senha.pack(side="left")
        tk.Button(frame_senha, text="Mostrar", bg=COR_CAMPO, fg=COR_LABEL,
                  font=("Consolas", 8), relief="flat", cursor="hand2",
                  command=self._toggle_senha).pack(side="left", padx=4)

        # Barra de progresso
        self._var_prog = tk.DoubleVar(value=0)
        self._var_desc = tk.StringVar(value="Aguardando...")
        tk.Label(self, textvariable=self._var_desc, bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9)).pack(pady=(PAD, 2))
        self._barra = ttk.Progressbar(self, variable=self._var_prog,
                                       maximum=100, length=460,
                                       style="v.Horizontal.TProgressbar")
        self._barra.pack(padx=PAD, pady=(0, 4))
        style = ttk.Style()
        style.theme_use("default")
        style.configure("v.Horizontal.TProgressbar",
                         troughcolor=COR_LOG, background=COR_BARRA)

        # Botões
        frame_btn = tk.Frame(self, bg=COR_BG)
        frame_btn.pack(pady=6)
        self._btn_run = tk.Button(frame_btn, text="▶  Automatizar SCPO",
                                   bg=COR_BOTAO, fg=COR_TEXTO,
                                   font=("Consolas", 11, "bold"),
                                   relief="flat", cursor="hand2",
                                   padx=16, pady=6, command=self._iniciar)
        self._btn_run.pack(side="left", padx=6)

        self._btn_senha_ok = tk.Button(frame_btn,
                                        text="✔ Senha alterada — Continuar",
                                        bg="#e67e22", fg=COR_TEXTO,
                                        font=("Consolas", 10), relief="flat",
                                        cursor="hand2", padx=12, pady=6,
                                        state="disabled",
                                        command=self._liberar_senha)
        self._btn_senha_ok.pack(side="left", padx=6)

        # Log
        tk.Label(self, text="Log:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9), anchor="w").pack(fill="x", padx=PAD)
        self._txt_log = tk.Text(self, bg=COR_LOG, fg=COR_LOG_TEXT,
                                 font=("Consolas", 9), height=10,
                                 relief="flat", state="disabled")
        self._txt_log.pack(fill="x", padx=PAD, pady=(0, PAD))

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _buscar_cep(self):
        cep = self.ent_cep.get().strip()
        if not cep: return
        d = buscar_cep(cep)
        if not d:
            messagebox.showwarning("CEP", f"CEP {cep} nao encontrado.")
            return
        rua = d.get("logradouro", "")
        self.ent_rua.delete(0, "end")
        self.ent_rua.insert(0, rua.upper())
        bairro = d.get("bairro", "")
        cidade = d.get("localidade", "")
        uf     = d.get("uf", "")
        self._lbl_cep_status.config(text=f"✔ {bairro} — {cidade}/{uf}")

    def _toggle_esquina(self):
        self.ent_rua2.config(state="normal" if self.var_esquina.get() else "disabled")
        self._atualizar_casas_esquina()

    def _atualizar_casas_esquina(self, event=None):
        for w in self.frame_casas.winfo_children():
            w.destroy()
        self._entries_rua_casa = []
        if not self.var_esquina.get(): return
        try:
            n = int(self.ent_num_casas.get())
        except ValueError:
            return
        tk.Label(self.frame_casas, text="Rua por casa:",
                 bg=COR_BG, fg=COR_LABEL, font=("Consolas", 9)
                 ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(4, 2))
        for i in range(n):
            tk.Label(self.frame_casas, text=f"  Casa {i+1}:",
                     bg=COR_BG, fg=COR_LABEL, font=("Consolas", 9)
                     ).grid(row=i+1, column=0, sticky="w")
            e = tk.Entry(self.frame_casas, bg=COR_CAMPO, fg=COR_TEXTO,
                         insertbackground=COR_TEXTO, font=("Consolas", 10),
                         relief="flat", width=28)
            e.grid(row=i+1, column=1, sticky="w", padx=4, pady=2)
            self._entries_rua_casa.append(e)

    def _toggle_senha(self):
        self.ent_senha.config(show="" if self.ent_senha.cget("show") == "*" else "*")

    def _log(self, msg):
        self.after(0, self._log_direto, msg)

    def _log_direto(self, msg):
        self._txt_log.config(state="normal")
        self._txt_log.insert("end", f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
        self._txt_log.see("end")
        self._txt_log.config(state="disabled")

    def _step(self, pct, desc):
        self.after(0, lambda: (self._var_prog.set(pct), self._var_desc.set(desc)))

    def _done(self, ok, msg):
        self.after(0, self._done_ui, ok, msg)

    def _done_ui(self, ok, msg):
        self._btn_run.config(state="normal")
        self._btn_senha_ok.config(state="disabled")
        (messagebox.showinfo if ok else messagebox.showerror)(
            "Concluido" if ok else "Erro", msg)

    def _liberar_senha(self):
        nova = simpledialog.askstring("Nova Senha",
                                      "Digite a nova senha cadastrada:", parent=self)
        if nova:
            self._nova_senha_tmp = nova
            self.var_senha.set(nova)
            self._evento_senha.set()
            self._btn_senha_ok.config(state="disabled")
            self._log_direto("Nova senha registrada.")

    # ── Validação e início ────────────────────────────────────────────────────
    def _validar(self):
        for v, n in [
            (self.ent_cep.get().strip(),      "CEP"),
            (self.ent_rua.get().strip(),       "Rua Principal"),
            (self.ent_quadra.get().strip(),    "Quadra"),
            (self.ent_lote.get().strip(),      "Lote"),
            (self.ent_num_casas.get().strip(), "Nr de Casas"),
            (self.var_data_ini.get().strip(),  "Data de Inicio"),
        ]:
            if not v:
                return False, f"Campo obrigatorio vazio: {n}"
        try:
            assert int(self.ent_num_casas.get()) >= 1
        except Exception:
            return False, "Nr de Casas deve ser inteiro positivo."
        try:
            datetime.strptime(self.var_data_ini.get().strip(), "%d/%m/%Y")
        except ValueError:
            return False, "Data de Inicio invalida — use DD/MM/YYYY."
        if self.var_esquina.get():
            if not self.ent_rua2.get().strip():
                return False, "Informe a Rua 2 para lote de esquina."
            for i, e in enumerate(self._entries_rua_casa):
                if not e.get().strip():
                    return False, f"Informe a rua da Casa {i+1}."
        if not self.var_senha.get().strip():
            return False, "Informe a senha do SCPO."
        return True, ""

    def _iniciar(self):
        ok, err = self._validar()
        if not ok:
            messagebox.showwarning("Campo invalido", err)
            return
        salvar_config({"senha": self.var_senha.get().strip()})
        n = int(self.ent_num_casas.get())
        casas = [{"numero": i+1,
                  "rua": (self._entries_rua_casa[i].get().strip()
                          if self.var_esquina.get() and i < len(self._entries_rua_casa) else "")}
                 for i in range(n)]

        self._dados_run = {
            "nome_obra":      montar_nome_obra(self.ent_rua.get().strip(),
                                               self.ent_quadra.get().strip(),
                                               self.ent_lote.get().strip()),
            "rua":            self.ent_rua.get().strip(),
            "rua2":           self.ent_rua2.get().strip(),
            "quadra":         self.ent_quadra.get().strip(),
            "lote":           self.ent_lote.get().strip(),
            "cep":            self.ent_cep.get().strip(),
            "esquina":        self.var_esquina.get(),
            "casas":          casas,
            "data_inicio":    self.var_data_ini.get().strip(),
            "data_termino":   data_termino_auto(),
            "observacao":     gerar_observacao(
                                  self.ent_rua.get().strip(),
                                  self.ent_quadra.get().strip(),
                                  self.ent_lote.get().strip(),
                                  self.var_esquina.get(),
                                  self.ent_rua2.get().strip(),
                                  casas),
            "pasta_download": self.ent_pasta.get().strip(),
            "evento_captcha":        self._evento_captcha,
            "evento_mostrar_captcha": self._evento_mostrar_captcha,
            "evento_senha":          self._evento_senha,
            "nova_senha":            self._nova_senha_tmp,
            "captcha_img_bytes":     b"",
            "captcha_valor":         "",
        }

        self._evento_captcha.clear()
        self._evento_mostrar_captcha.clear()
        self._evento_senha.clear()

        self._btn_run.config(state="disabled")
        self._btn_senha_ok.config(state="normal")
        self._var_prog.set(0)
        self._var_desc.set("Iniciando...")
        self._txt_log.config(state="normal")
        self._txt_log.delete("1.0", "end")
        self._txt_log.config(state="disabled")

        threading.Thread(
            target=executar_scpo,
            args=(self._dados_run, self.var_senha.get().strip(),
                  self._step, self._log, self._done),
            daemon=True
        ).start()


# ─── Entry point ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    AppSCPO().mainloop()
