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
from datetime import datetime
from dateutil.relativedelta import relativedelta

# ─── Selenium (imports explícitos — obrigatório para PyInstaller não cortar) ──
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions  # import direto evita "No module named"
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ─── Constantes fixas ────────────────────────────────────────────────────────
URL_SCPO      = "https://scpo.mte.gov.br/"
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
    """Retorna dict com logradouro/bairro/localidade/uf ou {} em caso de erro."""
    cep_limpo = cep.replace("-", "").replace(".", "").strip()
    if len(cep_limpo) != 8:
        return {}
    try:
        url = f"https://viacep.com.br/ws/{cep_limpo}/json/"
        with urllib.request.urlopen(url, timeout=5) as resp:
            dados = json.loads(resp.read().decode())
        return {} if "erro" in dados else dados
    except Exception:
        return {}

# ─── Lógica de negócio ────────────────────────────────────────────────────────
def montar_nome_obra(rua: str, quadra: str, lote: str) -> str:
    """RESIDENCIAL RUA DAS FLORES QUADRA 5 LOTE 10 — sem nome proprio"""
    return (f"RESIDENCIAL {rua.upper()} QUADRA {quadra.upper()} LOTE {lote.upper()}")

def gerar_observacao(residencial: str, rua: str, quadra: str, lote: str,
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


# ─── Automação Selenium ───────────────────────────────────────────────────────
def executar_scpo(dados: dict, senha: str, step_cb, log_cb, done_cb):
    import time
    driver = None
    try:
        # 1. Chrome ────────────────────────────────────────────────────────────
        log_cb("Iniciando Chrome...")
        step_cb(5, "Abrindo navegador")
        options = ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_experimental_option("prefs", {
            "download.default_directory": dados["pasta_download"],
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
        })
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager(driver_version=None).install()),
            options=options
        )
        wait = WebDriverWait(driver, 20)

        # 2. Login ────────────────────────────────────────────────────────────
        log_cb(f"Abrindo {URL_SCPO}...")
        step_cb(10, "Abrindo SCPO")
        driver.get(URL_SCPO)
        wait.until(EC.presence_of_element_located((By.NAME, "login"))).send_keys(LOGIN_CPF)
        driver.find_element(By.NAME, "senha").send_keys(senha)
        log_cb("Login e senha preenchidos.")
        log_cb(">>> Digite o código de segurança no navegador e clique 'Código digitado — Continuar'")
        step_cb(15, "Aguardando código de segurança")
        dados["evento_captcha"].wait()
        log_cb("Código confirmado. Clicando em Entrar...")
        step_cb(20, "Efetuando login")
        driver.find_element(By.NAME, "entrar").click()
        time.sleep(2)

        # Verifica pedido de troca de senha
        if "alterar" in driver.page_source.lower() and "senha" in driver.page_source.lower():
            log_cb("⚠ Site solicita alteração de senha. Altere no navegador e clique 'Senha alterada'.")
            step_cb(22, "Aguardando alteração de senha")
            dados["evento_senha"].wait()
            salvar_config({"senha": dados.get("nova_senha", senha)})
            log_cb("Nova senha salva. Prosseguindo...")

        # 3. Comunicação > Comunicar Obra ─────────────────────────────────────
        step_cb(30, "Navegando para Comunicar Obra")
        log_cb("Clicando em Comunicação > Comunicar Obra...")
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Comunicação"))).click()
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Comunicar Obra"))).click()

        # 4. Identificação da empresa ─────────────────────────────────────────
        step_cb(35, "Identificando empresa")
        log_cb("Marcando 'Obra não tem CNPJ'...")
        chk = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@type='checkbox' and contains(@id,'semCnpj')]")))
        if not chk.is_selected():
            chk.click()
        f = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[contains(@id,'cpfProprietario') or contains(@name,'cpfProprietario')]")))
        f.clear(); f.send_keys(LOGIN_CPF)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//input[@type='submit' and contains(@value,'Comunicar')] | "
                       "//button[contains(text(),'Comunicar Obra')]"))).click()

        # 5. Formulário ───────────────────────────────────────────────────────
        step_cb(45, "Preenchendo formulário")
        log_cb("Preenchendo dados da obra...")

        # Nome da obra
        f = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[contains(@id,'nomeObra') or contains(@name,'nomeObra')]")))
        f.clear(); f.send_keys(dados["nome_obra"])
        log_cb(f"  Nome: {dados['nome_obra']}")

        # Email e telefone
        driver.find_element(By.XPATH,
            "//input[contains(@id,'email') or contains(@name,'email')]").send_keys(EMAIL_FIXO)
        driver.find_element(By.XPATH,
            "//input[contains(@id,'telefone') or contains(@name,'telefone')]").send_keys(TELEFONE_FIXO)

        # CEP + lupa
        log_cb("  Preenchendo CEP e acionando auto-fill do site...")
        f = driver.find_element(By.XPATH,
            "//input[contains(@id,'cep') or contains(@name,'cep')]")
        f.clear(); f.send_keys(dados["cep"])
        try:
            driver.find_element(By.XPATH,
                "//img[contains(@src,'lupa') or contains(@alt,'Pesquisar')] | "
                "//button[contains(@onclick,'pesquisarCep')] | "
                "//a[contains(@onclick,'pesquisarCep')]").click()
            time.sleep(3)
        except Exception:
            log_cb("  ⚠ Lupa do CEP não localizada — verificar seletor no DevTools.")

        # Complemento
        f = driver.find_element(By.XPATH,
            "//input[contains(@id,'complemento') or contains(@name,'complemento')]")
        f.clear(); f.send_keys(f"QUADRA {dados['quadra']} LOTE {dados['lote']}")

        # 2ª rua (esquina)
        if dados.get("esquina") and dados.get("rua2"):
            log_cb("  Lote de esquina — preenchendo 2ª rua...")
            try:
                f = driver.find_element(By.XPATH,
                    "//input[contains(@id,'logradouro2') or contains(@name,'logradouro2') "
                    "or contains(@id,'logradouro_2')]")
                f.clear(); f.send_keys(dados["rua2"])
            except Exception:
                log_cb("  ⚠ Campo logradouro 2 não localizado — verificar ID no DevTools.")

        # Classe CNAE 4120-4
        step_cb(60, "Selecionando classe e tipo")
        sel = Select(wait.until(EC.presence_of_element_located(
            (By.XPATH, "//select[contains(@id,'classeCnae') or contains(@name,'classeCnae')]"))))
        for o in sel.options:
            if "4120" in o.text:
                sel.select_by_visible_text(o.text); break

        # Subclasse 00
        time.sleep(1)
        sel = Select(wait.until(EC.presence_of_element_located(
            (By.XPATH, "//select[contains(@id,'subclasse') or contains(@name,'subclasse')]"))))
        for o in sel.options:
            if o.text.strip().startswith("00"):
                sel.select_by_visible_text(o.text); break

        # Tipo de construção — Edifício
        sel = Select(driver.find_element(By.XPATH,
            "//select[contains(@id,'tipoConstrucao') or contains(@name,'tipoConstrucao')]"))
        for o in sel.options:
            if "dif" in o.text.lower():
                sel.select_by_visible_text(o.text); break

        # Tipo de obra — Privada
        driver.find_element(By.XPATH,
            "//input[@type='radio' and (@value='Privada' or @value='privada' "
            "or contains(@id,'privada'))]").click()

        # Característica — Construção
        driver.find_element(By.XPATH,
            "//input[@type='radio' and (@value='Construção' or @value='construcao' "
            "or contains(@id,'construcao'))]").click()

        # Observação
        step_cb(70, "Observação e datas")
        f = driver.find_element(By.XPATH,
            "//textarea[contains(@id,'observacao') or contains(@name,'observacao') "
            "or contains(@id,'descricao')]")
        f.clear(); f.send_keys(dados["observacao"])
        log_cb("  Observação preenchida.")

        # FGTS — Não
        driver.find_element(By.XPATH,
            "//input[@type='radio' and (contains(@id,'fgtsNao') "
            "or (@name='fgts' and @value='N'))]").click()

        # Datas
        f = driver.find_element(By.XPATH,
            "//input[contains(@id,'dataInicio') or contains(@name,'dataInicio')]")
        f.clear(); f.send_keys(dados["data_inicio"])
        log_cb(f"  Início: {dados['data_inicio']}  |  Término: {dados['data_termino']}")
        f = driver.find_element(By.XPATH,
            "//input[contains(@id,'dataTermino') or contains(@name,'dataTermino')]")
        f.clear(); f.send_keys(dados["data_termino"])

        # Empregados
        f = driver.find_element(By.XPATH,
            "//input[contains(@id,'empPrincipal') or contains(@name,'empPrincipal') "
            "or contains(@name,'numEmpregadosPrincipal')]")
        f.clear(); f.send_keys(EMP_PRINCIPAL)
        f = driver.find_element(By.XPATH,
            "//input[contains(@id,'empTerceiros') or contains(@name,'empTerceiros') "
            "or contains(@name,'numEmpregadosTerceiros')]")
        f.clear(); f.send_keys(EMP_TERCEIROS)

        # 6. Submeter ─────────────────────────────────────────────────────────
        step_cb(85, "Submetendo")
        log_cb("Clicando em 'Comunicar Obra'...")
        driver.find_element(By.XPATH,
            "//input[@type='submit' and contains(@value,'Comunicar')] | "
            "//button[contains(text(),'Comunicar Obra')]").click()
        time.sleep(3)

        if "erro" in driver.page_source.lower() or "inválido" in driver.page_source.lower():
            raise Exception("Site retornou erro após submissão. Verifique os dados no navegador.")

        log_cb("✔ Formulário submetido com sucesso!")
        step_cb(95, "Aguardando PDF")

        # 7. Baixar PDF de confirmação ────────────────────────────────────────
        pdf_ok = False
        for _ in range(10):
            try:
                driver.find_element(By.XPATH,
                    "//a[contains(@href,'.pdf')] | "
                    "//input[contains(@value,'Imprimir')] | "
                    "//button[contains(text(),'Imprimir')]").click()
                time.sleep(3)
                pdf_ok = True
                log_cb("✔ PDF gerado/baixado.")
                break
            except Exception:
                time.sleep(2)
        if not pdf_ok:
            log_cb("⚠ Botão de PDF não localizado — verifique a página de confirmação no navegador.")

        step_cb(100, "Concluído")
        done_cb(True, "SCPO preenchido com sucesso!")

    except Exception as e:
        log_cb(f"✖ ERRO: {e}")
        done_cb(False, str(e))
    # Navegador permanece aberto para conferência


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
        self._evento_captcha = threading.Event()
        self._evento_senha   = threading.Event()
        self._nova_senha_tmp = ""
        self._build_ui()

    def _build_ui(self):
        PAD = 12
        tk.Label(self, text="SCPO AUTOMATION", bg=COR_BG, fg=COR_TEXTO,
                 font=("Consolas", 14, "bold")).pack(pady=(PAD, 2))
        tk.Label(self, text="Berçan Projetos — Morais Engenharia", bg=COR_BG,
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

        # CEP com botão Buscar
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
        self._lbl_rua_status = tk.Label(frame_cep, text="", bg=COR_BG,
                                         fg=COR_LOG_TEXT, font=("Consolas", 8))
        self._lbl_rua_status.pack(side="left")

        # Rua — preenchida automaticamente pelo CEP, mas editável
        lbl(1, "Rua Principal:")
        self.ent_rua = ent(1)

        lbl(2, "Quadra:");    self.ent_quadra    = ent(2, width=10)
        lbl(3, "Lote:");      self.ent_lote      = ent(3, width=10)
        lbl(4, "Nº Casas:");  self.ent_num_casas = ent(4, width=5)
        self.ent_num_casas.bind("<FocusOut>", self._atualizar_casas_esquina)

        # Esquina
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

        # Data início
        lbl(8, "Data de Início:")
        self.ent_data_ini = tk.Entry(frame, textvariable=self.var_data_ini,
                                      bg=COR_CAMPO, fg=COR_TEXTO,
                                      insertbackground=COR_TEXTO, font=("Consolas", 10),
                                      relief="flat", width=12)
        self.ent_data_ini.grid(row=8, column=1, sticky="w", pady=3)
        tk.Label(frame, text="(DD/MM/YYYY)", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 8)).grid(row=8, column=1, sticky="e")

        # Data término — automático
        lbl(9, "Data de Término:")
        tk.Label(frame, text=data_termino_auto(), bg=COR_BG, fg=COR_LOG_TEXT,
                 font=("Consolas", 9)).grid(row=9, column=1, sticky="w")
        tk.Label(frame, text="(automático: +1 mês da data atual)",
                 bg=COR_BG, fg=COR_LABEL, font=("Consolas", 8)
                 ).grid(row=9, column=1, sticky="e")

        # Pasta de download
        lbl(10, "Pasta Download:")
        self.ent_pasta = tk.Entry(frame, bg=COR_CAMPO, fg=COR_TEXTO,
                                   insertbackground=COR_TEXTO, font=("Consolas", 10),
                                   relief="flat", width=32)
        self.ent_pasta.insert(0, os.path.expanduser("~\\Downloads"))
        self.ent_pasta.grid(row=10, column=1, sticky="w", pady=3)

        # Separador + Senha
        tk.Frame(frame, height=1, bg=COR_CAMPO
                 ).grid(row=12, column=0, columnspan=2, sticky="ew", pady=8)
        lbl(13, "Senha SCPO:")
        frame_s = tk.Frame(frame, bg=COR_BG)
        frame_s.grid(row=13, column=1, sticky="w")
        self.ent_senha = tk.Entry(frame_s, textvariable=self.var_senha,
                                   bg=COR_CAMPO, fg=COR_TEXTO, show="*",
                                   insertbackground=COR_TEXTO, font=("Consolas", 10),
                                   relief="flat", width=20)
        self.ent_senha.pack(side="left")
        tk.Button(frame_s, text="Mostrar", bg=COR_CAMPO, fg=COR_LABEL,
                  font=("Consolas", 8), relief="flat", cursor="hand2",
                  command=self._toggle_senha).pack(side="left", padx=4)

        # Progresso
        self._var_prog = tk.DoubleVar(value=0)
        self._var_desc = tk.StringVar(value="Aguardando...")
        tk.Label(self, textvariable=self._var_desc, bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9)).pack(pady=(PAD, 2))
        self._barra = ttk.Progressbar(self, variable=self._var_prog,
                                       maximum=100, length=480,
                                       style="verde.Horizontal.TProgressbar")
        self._barra.pack(padx=PAD, pady=(0, 4))
        s = ttk.Style(); s.theme_use("default")
        s.configure("verde.Horizontal.TProgressbar",
                    troughcolor=COR_LOG, background=COR_BARRA)

        # Botões
        frame_btn = tk.Frame(self, bg=COR_BG)
        frame_btn.pack(pady=6)
        self._btn_run = tk.Button(frame_btn, text="▶  Automatizar SCPO",
                                   bg=COR_BOTAO, fg=COR_TEXTO,
                                   font=("Consolas", 11, "bold"), relief="flat",
                                   cursor="hand2", padx=16, pady=6,
                                   command=self._iniciar)
        self._btn_run.pack(side="left", padx=6)
        self._btn_captcha = tk.Button(frame_btn,
                                       text="✔ Código digitado — Continuar",
                                       bg="#27ae60", fg=COR_TEXTO,
                                       font=("Consolas", 10), relief="flat",
                                       cursor="hand2", padx=12, pady=6,
                                       state="disabled",
                                       command=self._liberar_captcha)
        self._btn_captcha.pack(side="left", padx=6)
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

    # ─── Helpers ─────────────────────────────────────────────────────────────
    def _buscar_cep(self):
        d = buscar_cep(self.ent_cep.get().strip())
        if not d:
            self._lbl_rua_status.config(text="CEP não encontrado", fg="#e74c3c")
            return
        self.ent_rua.delete(0, "end")
        self.ent_rua.insert(0, d.get("logradouro", ""))
        self._lbl_rua_status.config(
            text=f"✔ {d.get('bairro','')} — {d.get('localidade','')}/{d.get('uf','')}",
            fg=COR_LOG_TEXT)

    def _toggle_esquina(self):
        self.ent_rua2.config(state="normal" if self.var_esquina.get() else "disabled")
        self._atualizar_casas_esquina()

    def _atualizar_casas_esquina(self, event=None):
        for w in self.frame_casas.winfo_children():
            w.destroy()
        self._entries_rua_casa = []
        if not self.var_esquina.get():
            return
        try:
            n = int(self.ent_num_casas.get())
        except ValueError:
            return
        tk.Label(self.frame_casas, text="Rua por casa (esquina):",
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
        self._btn_captcha.config(state="disabled")
        self._btn_senha_ok.config(state="disabled")
        (messagebox.showinfo if ok else messagebox.showerror)(
            "Concluído" if ok else "Erro", msg)

    def _liberar_captcha(self):
        self._evento_captcha.set()
        self._btn_captcha.config(state="disabled")
        self._log_direto("Código de segurança confirmado.")

    def _liberar_senha(self):
        nova = simpledialog.askstring("Nova Senha",
                                      "Digite a nova senha cadastrada:", parent=self)
        if nova:
            self._nova_senha_tmp = nova
            self.var_senha.set(nova)
            self._evento_senha.set()
            self._btn_senha_ok.config(state="disabled")
            self._log_direto("Nova senha registrada.")

    # ─── Validação e início ───────────────────────────────────────────────────
    def _validar(self):
        for v, n in [
            (self.ent_cep.get().strip(),         "CEP"),
            (self.ent_rua.get().strip(),          "Rua Principal"),
            (self.ent_quadra.get().strip(),       "Quadra"),
            (self.ent_lote.get().strip(),         "Lote"),
            (self.ent_num_casas.get().strip(),    "Nº de Casas"),
            (self.var_data_ini.get().strip(),     "Data de Início"),
        ]:
            if not v:
                return False, f"Campo obrigatório vazio: {n}"
        try:
            nc = int(self.ent_num_casas.get())
            assert nc >= 1
        except Exception:
            return False, "Nº de Casas deve ser inteiro positivo."
        try:
            datetime.strptime(self.var_data_ini.get().strip(), "%d/%m/%Y")
        except ValueError:
            return False, "Data de Início inválida — use DD/MM/YYYY."
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
            messagebox.showwarning("Campo inválido", err)
            return
        salvar_config({"senha": self.var_senha.get().strip()})
        n = int(self.ent_num_casas.get())
        casas = [{"numero": i+1,
                  "rua": (self._entries_rua_casa[i].get().strip()
                          if self.var_esquina.get() and i < len(self._entries_rua_casa) else "")}
                 for i in range(n)]
        dados = {
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
            "observacao":     gerar_observacao("",
                                               self.ent_rua.get().strip(),
                                               self.ent_quadra.get().strip(),
                                               self.ent_lote.get().strip(),
                                               self.var_esquina.get(),
                                               self.ent_rua2.get().strip(),
                                               casas),
            "pasta_download": self.ent_pasta.get().strip(),
            "evento_captcha": self._evento_captcha,
            "evento_senha":   self._evento_senha,
            "nova_senha":     self._nova_senha_tmp,
        }
        self._evento_captcha.clear()
        self._evento_senha.clear()
        self._btn_run.config(state="disabled")
        self._btn_captcha.config(state="normal")
        self._var_prog.set(0)
        self._var_desc.set("Iniciando...")
        self._txt_log.config(state="normal")
        self._txt_log.delete("1.0", "end")
        self._txt_log.config(state="disabled")
        threading.Thread(
            target=executar_scpo,
            args=(dados, self.var_senha.get().strip(),
                  self._step, self._log, self._done),
            daemon=True
        ).start()


# ─── Entry point ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    AppSCPO().mainloop()
