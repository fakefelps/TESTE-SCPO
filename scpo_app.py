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

# ─── Selenium ────────────────────────────────────────────────────────────────
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

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

# ─── Detecta caminho do EdgeDriver instalado no Windows ──────────────────────
def _encontrar_msedgedriver() -> str:
    """
    O EdgeDriver (msedgedriver.exe) vem junto com o Edge no Windows.
    Localiza automaticamente sem precisar baixar nada.
    """
    import subprocess, re as _re

    # Caminhos padrão onde o msedgedriver pode estar
    caminhos_fixos = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedgedriver.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedgedriver.exe",
    ]
    for p in caminhos_fixos:
        if os.path.exists(p):
            return p

    # Tenta localizar via where
    try:
        resultado = subprocess.check_output(
            ["where", "msedgedriver"], stderr=subprocess.DEVNULL, timeout=5
        ).decode().strip().splitlines()
        if resultado:
            return resultado[0]
    except Exception:
        pass

    # Busca na pasta de instalação do Edge pelo número de versão
    base = r"C:\Program Files (x86)\Microsoft\Edge\Application"
    if os.path.isdir(base):
        for entry in os.listdir(base):
            driver = os.path.join(base, entry, "msedgedriver.exe")
            if os.path.exists(driver):
                return driver

    return ""  # não encontrado — usa Selenium Manager como fallback


# ─── Automação Selenium com Edge ─────────────────────────────────────────────
def executar_scpo(dados: dict, senha: str, step_cb, log_cb, done_cb):
    import time, traceback
    driver = None
    try:
        # ── 1. Iniciar Edge ───────────────────────────────────────────────────
        log_cb("Iniciando Microsoft Edge...")
        step_cb(5, "Abrindo Edge")

        options = EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_experimental_option("prefs", {
            "download.default_directory": dados["pasta_download"],
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
        })

        driver_path = _encontrar_msedgedriver()
        if driver_path:
            log_cb(f"EdgeDriver encontrado: {driver_path}")
            driver = webdriver.Edge(
                service=EdgeService(driver_path), options=options)
        else:
            # Selenium Manager detecta automaticamente
            log_cb("EdgeDriver nao encontrado localmente — usando Selenium Manager...")
            driver = webdriver.Edge(options=options)

        wait = WebDriverWait(driver, 20)

        # ── 2. Login ──────────────────────────────────────────────────────────
        log_cb(f"Abrindo {URL_SCPO}...")
        step_cb(10, "Abrindo SCPO")
        driver.get(URL_SCPO)

        # IDs confirmados via DevTools
        wait.until(EC.presence_of_element_located(
            (By.ID, "PlaceHolderConteudo_txtCPF"))).send_keys(LOGIN_CPF)
        driver.find_element(
            By.ID, "PlaceHolderConteudo_txtSenha").send_keys(senha)

        log_cb("Login e senha preenchidos.")
        log_cb(">>> Digite o codigo de seguranca no navegador e clique 'Codigo digitado'")
        step_cb(15, "Aguardando codigo de seguranca")

        # Habilita botão de captcha na UI
        dados["fn_habilitar_captcha"]()
        dados["evento_captcha"].wait()

        log_cb("Codigo confirmado. Clicando em Entrar...")
        step_cb(20, "Efetuando login")
        driver.find_element(By.ID, "PlaceHolderConteudo_btnLogin").click()
        time.sleep(2)

        # Verifica troca de senha
        if "RedefinirSenha" in driver.current_url or (
                "alterar" in driver.page_source.lower() and
                "senha" in driver.page_source.lower()):
            log_cb("Site solicita alteracao de senha.")
            log_cb("Altere no navegador e clique 'Senha alterada'.")
            step_cb(22, "Aguardando alteracao de senha")
            dados["fn_habilitar_senha"]()
            dados["evento_senha"].wait()
            salvar_config({"senha": dados.get("nova_senha", senha)})
            log_cb("Nova senha salva.")

        log_cb("Login OK!")

        # ── 3. Navegação para Comunicar Obra ──────────────────────────────────
        step_cb(30, "Navegando para Comunicar Obra")
        log_cb("Abrindo menu Comunicacao...")
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@onclick,'subMenu01')]"))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@href,'DeclaracaoPreviaObra/Comunicar')]"))).click()

        # ── 4. Tela intermediária ─────────────────────────────────────────────
        step_cb(35, "Identificando empresa")
        log_cb("Marcando 'Obra nao tem CNPJ'...")
        chk = wait.until(EC.presence_of_element_located(
            (By.ID, "PlaceHolderConteudo_chkObraSemCNPJ")))
        if not chk.is_selected():
            chk.click()
        time.sleep(1)

        cpf_field = wait.until(EC.element_to_be_clickable(
            (By.ID, "txtCPFProprietarioObra")))
        cpf_field.clear()
        cpf_field.send_keys(LOGIN_CPF)

        wait.until(EC.element_to_be_clickable(
            (By.ID, "PlaceHolderConteudo_btnDeclararObra"))).click()

        # ── 5. Formulário principal ───────────────────────────────────────────
        step_cb(45, "Preenchendo formulario")
        log_cb("Preenchendo dados da obra...")

        # Nome da obra
        f = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[contains(@id,'NomeObra') or contains(@id,'nomeObra') or contains(@name,'NomeObra')]")))
        f.clear(); f.send_keys(dados["nome_obra"])
        log_cb(f"  Nome: {dados['nome_obra']}")

        # Email
        try:
            f = driver.find_element(By.XPATH,
                "//input[contains(@id,'Email') or contains(@id,'email')]")
            f.clear(); f.send_keys(EMAIL_FIXO)
        except Exception:
            log_cb("  campo email nao localizado")

        # Telefone
        try:
            f = driver.find_element(By.XPATH,
                "//input[contains(@id,'Telefone') or contains(@id,'telefone')]")
            f.clear(); f.send_keys(TELEFONE_FIXO)
        except Exception:
            log_cb("  campo telefone nao localizado")

        # CEP + lupa
        log_cb("  Preenchendo CEP...")
        f = driver.find_element(By.XPATH,
            "//input[contains(@id,'CEP') or contains(@id,'Cep') or contains(@id,'cep')]")
        f.clear(); f.send_keys(dados["cep"])
        try:
            lupa = driver.find_element(By.XPATH,
                "//img[contains(@src,'lupa') or contains(@title,'Pesquisar CEP')] | "
                "//a[contains(@onclick,'CEP') or contains(@onclick,'Cep')]")
            lupa.click()
            time.sleep(3)
        except Exception:
            log_cb("  lupa CEP nao localizada — preenchimento manual pode ser necessario")

        # Complemento
        try:
            f = driver.find_element(By.XPATH,
                "//input[contains(@id,'Complemento') or contains(@id,'complemento')]")
            f.clear(); f.send_keys(f"QUADRA {dados['quadra']} LOTE {dados['lote']}")
        except Exception:
            log_cb("  campo complemento nao localizado")

        # 2ª rua (esquina)
        if dados.get("esquina") and dados.get("rua2"):
            log_cb("  Lote de esquina — preenchendo 2a rua...")
            try:
                f = driver.find_element(By.XPATH,
                    "//input[contains(@id,'Logradouro2') or contains(@id,'logradouro2')]")
                f.clear(); f.send_keys(dados["rua2"])
            except Exception:
                log_cb("  campo logradouro2 nao localizado")

        # Classe CNAE 4120-4
        step_cb(60, "Selecionando classe e tipo")
        try:
            sel = Select(wait.until(EC.presence_of_element_located(
                (By.XPATH, "//select[contains(@id,'Classe') or contains(@id,'classe') or contains(@id,'CNAE')]"))))
            for o in sel.options:
                if "4120" in o.text:
                    sel.select_by_visible_text(o.text); break
        except Exception:
            log_cb("  select classe nao localizado")

        # Subclasse 00
        time.sleep(1)
        try:
            sel = Select(wait.until(EC.presence_of_element_located(
                (By.XPATH, "//select[contains(@id,'Subclasse') or contains(@id,'subclasse')]"))))
            for o in sel.options:
                if o.text.strip().startswith("00"):
                    sel.select_by_visible_text(o.text); break
        except Exception:
            log_cb("  select subclasse nao localizado")

        # Tipo de construção
        try:
            sel = Select(driver.find_element(By.XPATH,
                "//select[contains(@id,'TipoConstrucao') or contains(@id,'tipoConstrucao')]"))
            for o in sel.options:
                if "dif" in o.text.lower():
                    sel.select_by_visible_text(o.text); break
        except Exception:
            log_cb("  select tipo construcao nao localizado")

        # Tipo obra — Privada
        try:
            driver.find_element(By.XPATH,
                "//input[@type='radio' and "
                "(@value='Privada' or @value='privada' or contains(@id,'rivada'))]").click()
        except Exception:
            log_cb("  radio privada nao localizado")

        # Característica — Construção
        try:
            driver.find_element(By.XPATH,
                "//input[@type='radio' and "
                "(@value='Construcao' or @value='Construção' or contains(@id,'onstrucao'))]").click()
        except Exception:
            log_cb("  radio construcao nao localizado")

        # Observação
        step_cb(70, "Observacao e datas")
        try:
            f = driver.find_element(By.XPATH,
                "//textarea[contains(@id,'Observ') or contains(@id,'observ') "
                "or contains(@id,'Descri') or contains(@id,'descri')]")
            f.clear(); f.send_keys(dados["observacao"])
            log_cb("  Observacao preenchida.")
        except Exception:
            log_cb("  textarea observacao nao localizada")

        # FGTS — Não
        try:
            driver.find_element(By.XPATH,
                "//input[@type='radio' and "
                "(contains(@id,'FGTSNao') or contains(@id,'fgtsNao') or "
                "(@name and contains(@name,'FGTS') and @value='N') or "
                "(@name and contains(@name,'FGTS') and @value='2'))]").click()
        except Exception:
            log_cb("  radio FGTS nao localizado")

        # Datas
        try:
            f = driver.find_element(By.XPATH,
                "//input[contains(@id,'DataInicio') or contains(@id,'dataInicio')]")
            f.clear(); f.send_keys(dados["data_inicio"])
        except Exception:
            log_cb("  campo data inicio nao localizado")

        try:
            f = driver.find_element(By.XPATH,
                "//input[contains(@id,'DataTermino') or contains(@id,'dataTermino')]")
            f.clear(); f.send_keys(dados["data_termino"])
        except Exception:
            log_cb("  campo data termino nao localizado")

        log_cb(f"  Inicio: {dados['data_inicio']} | Termino: {dados['data_termino']}")

        # Empregados
        try:
            f = driver.find_element(By.XPATH,
                "//input[contains(@id,'EmpPrincipal') or contains(@id,'empPrincipal') "
                "or contains(@id,'NumEmpregadosPrincipal')]")
            f.clear(); f.send_keys(EMP_PRINCIPAL)
        except Exception:
            log_cb("  campo emp principal nao localizado")

        try:
            f = driver.find_element(By.XPATH,
                "//input[contains(@id,'EmpTerceiros') or contains(@id,'empTerceiros') "
                "or contains(@id,'NumEmpregadosTerceiros')]")
            f.clear(); f.send_keys(EMP_TERCEIROS)
        except Exception:
            log_cb("  campo emp terceiros nao localizado")

        # ── 6. Submeter ───────────────────────────────────────────────────────
        step_cb(85, "Submetendo")
        log_cb("Clicando em Comunicar Obra...")
        try:
            driver.find_element(By.XPATH,
                "//input[@type='submit' and (contains(@value,'Comunicar') or contains(@value,'comunicar'))] | "
                "//button[contains(text(),'Comunicar')]").click()
        except Exception:
            log_cb("  botao submit nao localizado — tentando via JavaScript...")
            driver.execute_script(
                "document.querySelector('input[type=submit]').click()")

        time.sleep(3)
        step_cb(100, "Concluido!")
        log_cb("Formulario submetido! Verifique o navegador para confirmar.")
        done_cb(True, "SCPO preenchido!\nVerifique o navegador Edge para confirmar o protocolo.")

    except Exception as e:
        tb = traceback.format_exc()
        log_cb(f"ERRO: {e}")
        for linha in tb.splitlines():
            log_cb(linha)
        try:
            log_dir = os.path.join(
                os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "SCPOApp")
            os.makedirs(log_dir, exist_ok=True)
            log_path = os.path.join(
                log_dir, f"erro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            with open(log_path, "w", encoding="utf-8") as lf:
                lf.write(f"Erro: {e}\n\n{tb}")
            log_cb(f"Relatorio: {log_path}")
        except Exception:
            pass
        done_cb(False, str(e))
    # Edge permanece aberto para conferência


# ─── Interface Tkinter ────────────────────────────────────────────────────────
class AppSCPO(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SCPO Automation — Bercan Projetos")
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
        self._lbl_cep = tk.Label(frame_cep, text="", bg=COR_BG,
                                  fg=COR_LOG_TEXT, font=("Consolas", 8))
        self._lbl_cep.pack(side="left")

        lbl(1, "Rua Principal:"); self.ent_rua    = ent(1)
        lbl(2, "Quadra:");        self.ent_quadra = ent(2, width=10)
        lbl(3, "Lote:");          self.ent_lote   = ent(3, width=10)
        lbl(4, "Nr Casas:");      self.ent_ncasas = ent(4, width=5)
        self.ent_ncasas.bind("<FocusOut>", self._atualizar_casas)

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
        self.ent_data = tk.Entry(frame, textvariable=self.var_data_ini,
                                  bg=COR_CAMPO, fg=COR_TEXTO,
                                  insertbackground=COR_TEXTO, font=("Consolas", 10),
                                  relief="flat", width=12)
        self.ent_data.grid(row=8, column=1, sticky="w", pady=3)
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
        frame_s = tk.Frame(frame, bg=COR_BG)
        frame_s.grid(row=12, column=1, sticky="w")
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

        self._btn_captcha = tk.Button(frame_btn,
                                       text="Codigo digitado — Continuar",
                                       bg="#27ae60", fg=COR_TEXTO,
                                       font=("Consolas", 10), relief="flat",
                                       cursor="hand2", padx=12, pady=6,
                                       state="disabled",
                                       command=self._liberar_captcha)
        self._btn_captcha.pack(side="left", padx=6)

        self._btn_senha_ok = tk.Button(frame_btn,
                                        text="Senha alterada — Continuar",
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
        self.ent_rua.delete(0, "end")
        self.ent_rua.insert(0, d.get("logradouro", "").upper())
        self._lbl_cep.config(
            text=f"✔ {d.get('bairro','')} — {d.get('localidade','')}/{d.get('uf','')}")

    def _toggle_esquina(self):
        self.ent_rua2.config(state="normal" if self.var_esquina.get() else "disabled")
        self._atualizar_casas()

    def _atualizar_casas(self, event=None):
        for w in self.frame_casas.winfo_children():
            w.destroy()
        self._entries_rua_casa = []
        if not self.var_esquina.get(): return
        try:
            n = int(self.ent_ncasas.get())
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
        self.ent_senha.config(
            show="" if self.ent_senha.cget("show") == "*" else "*")

    def _log(self, msg):
        self.after(0, self._log_direto, msg)

    def _log_direto(self, msg):
        self._txt_log.config(state="normal")
        self._txt_log.insert("end",
            f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
        self._txt_log.see("end")
        self._txt_log.config(state="disabled")

    def _step(self, pct, desc):
        self.after(0, lambda: (
            self._var_prog.set(pct), self._var_desc.set(desc)))

    def _done(self, ok, msg):
        self.after(0, self._done_ui, ok, msg)

    def _done_ui(self, ok, msg):
        self._btn_run.config(state="normal")
        self._btn_captcha.config(state="disabled")
        self._btn_senha_ok.config(state="disabled")
        (messagebox.showinfo if ok else messagebox.showerror)(
            "Concluido" if ok else "Erro", msg)

    def _liberar_captcha(self):
        self._evento_captcha.set()
        self._btn_captcha.config(state="disabled")
        self._log_direto("Codigo de seguranca confirmado.")

    def _liberar_senha(self):
        nova = simpledialog.askstring(
            "Nova Senha", "Digite a nova senha cadastrada:", parent=self)
        if nova:
            self._nova_senha_tmp = nova
            self.var_senha.set(nova)
            self._evento_senha.set()
            self._btn_senha_ok.config(state="disabled")

    # ── Validação e início ────────────────────────────────────────────────────
    def _validar(self):
        for v, n in [
            (self.ent_cep.get().strip(),     "CEP"),
            (self.ent_rua.get().strip(),      "Rua Principal"),
            (self.ent_quadra.get().strip(),   "Quadra"),
            (self.ent_lote.get().strip(),     "Lote"),
            (self.ent_ncasas.get().strip(),   "Nr de Casas"),
            (self.var_data_ini.get().strip(), "Data de Inicio"),
        ]:
            if not v:
                return False, f"Campo obrigatorio vazio: {n}"
        try:
            assert int(self.ent_ncasas.get()) >= 1
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
        n = int(self.ent_ncasas.get())
        casas = [{"numero": i+1,
                  "rua": (self._entries_rua_casa[i].get().strip()
                          if self.var_esquina.get() and
                          i < len(self._entries_rua_casa) else "")}
                 for i in range(n)]

        self._evento_captcha.clear()
        self._evento_senha.clear()

        dados = {
            "nome_obra":    montar_nome_obra(self.ent_rua.get().strip(),
                                             self.ent_quadra.get().strip(),
                                             self.ent_lote.get().strip()),
            "rua":          self.ent_rua.get().strip(),
            "rua2":         self.ent_rua2.get().strip(),
            "quadra":       self.ent_quadra.get().strip(),
            "lote":         self.ent_lote.get().strip(),
            "cep":          self.ent_cep.get().strip(),
            "esquina":      self.var_esquina.get(),
            "casas":        casas,
            "data_inicio":  self.var_data_ini.get().strip(),
            "data_termino": data_termino_auto(),
            "observacao":   gerar_observacao(
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
            "fn_habilitar_captcha": lambda: self.after(
                0, lambda: self._btn_captcha.config(state="normal")),
            "fn_habilitar_senha": lambda: self.after(
                0, lambda: self._btn_senha_ok.config(state="normal")),
        }

        self._btn_run.config(state="disabled")
        self._btn_captcha.config(state="disabled")
        self._btn_senha_ok.config(state="disabled")
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
