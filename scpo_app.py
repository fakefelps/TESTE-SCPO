# PROTEÇÃO ANTI-LOOP — deve ser a PRIMEIRA coisa no arquivo
import multiprocessing
multiprocessing.freeze_support()

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import json
import os
import sys
import urllib.request
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

# ─── Selenium ───────────────────────────────────────────────────────────────
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ─── Constantes fixas ────────────────────────────────────────────────────────
URL_SCPO      = "https://scpo.mte.gov.br/"
LOGIN_CPF     = "038.144.411-25"
EMAIL_FIXO    = "joaovitorcabral94@gmail.com"
TELEFONE_FIXO = "6299266-5923"
EMP_PRINCIPAL = "0"
EMP_TERCEIROS = "5"

# ─── Paleta padrão Morais ─────────────────────────────────────────────────────
COR_BG        = "#1e2a3a"
COR_LOG       = "#131c26"
COR_CAMPO     = "#2a3f55"
COR_BOTAO     = "#2e86de"
COR_BARRA     = "#4cd964"
COR_TEXTO     = "#ffffff"
COR_LABEL     = "#90adc4"
COR_LOG_TEXT  = "#7ec8a0"

# ─── Arquivo de configuração (senha persistente) ──────────────────────────────
CONFIG_PATH = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), "scpo_config.json")

def carregar_config():
    """Carrega configurações salvas (senha atual)."""
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {"senha": "SCPO123"}

def salvar_config(config: dict):
    """Salva configurações em disco (persiste entre sessões)."""
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f)

# ─── Lógica de negócio (pura, sem dependência da UI) ─────────────────────────

def gerar_observacao(residencial: str, rua: str, quadra: str, lote: str,
                     esquina: bool, rua2: str, casas: list[dict]) -> str:
    """
    Gera o texto do campo 'Observação' conforme padrão SCPO.
    casas: lista de dicts {"numero": int, "rua": str}
    Exemplo sem esquina: [{"numero": 1, "rua": ""}, {"numero": 2, "rua": ""}]
    Exemplo com esquina: [{"numero": 1, "rua": "RUA X"}, {"numero": 2, "rua": "RUA Y"}]
    """
    casas_str_list = []
    for c in casas:
        label = f"CASA {c['numero']}"
        if esquina and c.get("rua"):
            label += f" SITUADA NA {c['rua'].upper()}"
        casas_str_list.append(label)
    casas_str = ", ".join(casas_str_list)

    if not esquina:
        obs = (
            f"OBRA RESIDENCIAL UNIFAMILIAR SITUADA NA {rua.upper()} "
            f"QUADRA {quadra.upper()} LOTE {lote.upper()} "
            f"COMPOSTA POR: {casas_str}"
        )
    else:
        obs = (
            f"OBRA RESIDENCIAL UNIFAMILIAR SITUADA NA {rua.upper()} "
            f"E {rua2.upper()} QUADRA {quadra.upper()} LOTE {lote.upper()} "
            f"COMPOSTA POR: {casas_str}"
        )
    return obs

def nome_obra(residencial: str, rua: str, quadra: str, lote: str) -> str:
    """Monta o nome da obra conforme padrão: RESIDENCIAL X RUA Y QUADRA Z LOTE W."""
    return f"RESIDENCIAL {residencial.upper()} {rua.upper()} QUADRA {quadra.upper()} LOTE {lote.upper()}"

def buscar_cep(cep: str) -> dict:
    """Consulta ViaCEP. Retorna dict com logradouro, bairro, localidade, uf ou {} se falhar."""
    cep_limpo = cep.replace("-", "").replace(".", "").strip()
    if len(cep_limpo) != 8:
        return {}
    try:
        url = f"https://viacep.com.br/ws/{cep_limpo}/json/"
        with urllib.request.urlopen(url, timeout=5) as resp:
            dados = json.loads(resp.read().decode())
        if "erro" in dados:
            return {}
        return dados
    except Exception:
        return {}

def data_termino_auto() -> str:
    """Retorna data de término = 1 mês após hoje, formato DD/MM/YYYY."""
    return (datetime.today() + relativedelta(months=1)).strftime("%d/%m/%Y")


def executar_scpo(dados: dict, senha: str, step_cb, log_cb, done_cb):
    """
    Função principal de automação Selenium.
    Roda em thread separada.
    
    dados: dicionário com todos os campos do formulário
    step_cb(pct, descricao): atualiza barra de progresso
    log_cb(msg): envia mensagem ao log da UI
    done_cb(ok, msg): finaliza execução (True=sucesso, False=erro)
    """
    driver = None
    try:
        # ── 1. Iniciar Chrome ────────────────────────────────────────────────
        log_cb("Iniciando Chrome...")
        step_cb(5, "Abrindo navegador")
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        # Diretório de download configurado para capturar o PDF
        prefs = {
            "download.default_directory": dados["pasta_download"],
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,  # Força download ao invés de abrir
        }
        options.add_experimental_option("prefs", prefs)
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 20)

        # ── 2. Login ─────────────────────────────────────────────────────────
        log_cb(f"Abrindo {URL_SCPO}...")
        step_cb(10, "Abrindo SCPO")
        driver.get(URL_SCPO)

        # Preenche login e senha
        wait.until(EC.presence_of_element_located((By.NAME, "login"))).send_keys(LOGIN_CPF)
        driver.find_element(By.NAME, "senha").send_keys(senha)
        log_cb("Login e senha preenchidos. Aguardando código de segurança...")
        step_cb(15, "Aguardando código de segurança")

        # ── 3. Pausar para código de segurança ───────────────────────────────
        # A UI irá liberar o evento após o usuário clicar "Continuar"
        dados["evento_captcha"].wait()  # Bloqueia thread até UI liberar
        log_cb("Código de segurança confirmado. Efetuando login...")
        step_cb(20, "Efetuando login")

        # Clica "Entrar"
        driver.find_element(By.NAME, "entrar").click()

        # Verifica se houve erro de senha ou se pediu alteração
        import time
        time.sleep(2)
        page = driver.page_source.lower()
        if "alterar" in page and "senha" in page:
            log_cb("⚠ Site solicitou alteração de senha. Aguardando usuário alterar no navegador...")
            step_cb(22, "Aguardando alteração de senha")
            dados["evento_senha"].wait()  # UI libera após usuário alterar
            log_cb("Senha alterada. Prosseguindo...")
            # Salva nova senha
            salvar_config({"senha": dados["nova_senha"]})

        # ── 4. Navegar até "Comunicar Obra" ──────────────────────────────────
        step_cb(30, "Navegando para Comunicar Obra")
        log_cb("Navegando para Comunicação > Comunicar Obra...")

        # Clicar no menu "Comunicação"
        menu_com = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Comunicação")))
        menu_com.click()

        # Clicar em "Comunicar Obra" no submenu
        submenu = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Comunicar Obra")))
        submenu.click()

        # ── 5. Tela intermediária — Identificação da Empresa ─────────────────
        step_cb(35, "Identificando empresa")
        log_cb("Marcando 'Obra não tem CNPJ'...")

        # Marcar checkbox "Obra não tem CNPJ"
        chk_sem_cnpj = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@type='checkbox' and contains(@id,'semCnpj')]")))
        if not chk_sem_cnpj.is_selected():
            chk_sem_cnpj.click()

        # Preencher CPF do proprietário
        campo_cpf = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[contains(@id,'cpfProprietario') or contains(@name,'cpfProprietario')]")))
        campo_cpf.clear()
        campo_cpf.send_keys(LOGIN_CPF)

        # Clicar "Comunicar Obra" (botão de avançar)
        btn_comunicar = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//input[@type='submit' and @value='Comunicar Obra'] | //button[contains(text(),'Comunicar Obra')]")))
        btn_comunicar.click()

        # ── 6. Formulário principal ───────────────────────────────────────────
        step_cb(45, "Preenchendo formulário")
        log_cb("Preenchendo dados da obra...")

        # Nome da Obra
        campo_nome = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[contains(@id,'nomeObra') or contains(@name,'nomeObra')]")))
        campo_nome.clear()
        campo_nome.send_keys(dados["nome_obra"])
        log_cb(f"  Nome: {dados['nome_obra']}")

        # Contratante Principal: Sim (radio)
        # O campo já deve estar preenchido com CPF do proprietário

        # Email
        campo_email = driver.find_element(By.XPATH,
            "//input[contains(@id,'email') or contains(@name,'email')]")
        campo_email.clear()
        campo_email.send_keys(EMAIL_FIXO)

        # Telefone
        campo_tel = driver.find_element(By.XPATH,
            "//input[contains(@id,'telefone') or contains(@name,'telefone')]")
        campo_tel.clear()
        campo_tel.send_keys(TELEFONE_FIXO)

        # CEP
        log_cb("  Preenchendo CEP e aguardando auto-fill...")
        campo_cep = driver.find_element(By.XPATH,
            "//input[contains(@id,'cep') or contains(@name,'cep')]")
        campo_cep.clear()
        campo_cep.send_keys(dados["cep"])

        # Clicar na lupa para auto-fill de endereço
        btn_lupa = driver.find_element(By.XPATH,
            "//img[contains(@src,'lupa') or contains(@alt,'Pesquisar')] | //button[contains(@onclick,'pesquisarCep')]")
        btn_lupa.click()
        import time; time.sleep(3)  # Aguarda auto-fill

        # Complemento (Quadra + Lote)
        campo_compl = driver.find_element(By.XPATH,
            "//input[contains(@id,'complemento') or contains(@name,'complemento')]")
        campo_compl.clear()
        campo_compl.send_keys(f"QUADRA {dados['quadra']} LOTE {dados['lote']}")

        # Segunda rua (se esquina)
        if dados.get("esquina") and dados.get("rua2"):
            log_cb("  Lote de esquina: preenchendo segunda rua...")
            try:
                campo_logr2 = driver.find_element(By.XPATH,
                    "//input[contains(@id,'logradouro2') or contains(@name,'logradouro2') or contains(@id,'logradouro_2')]")
                campo_logr2.clear()
                campo_logr2.send_keys(dados["rua2"])
            except Exception:
                log_cb("  ⚠ Campo logradouro 2 não localizado. Verificar ID real do campo.")

        step_cb(60, "Selecionando classe/tipo")
        log_cb("  Selecionando Classe CNAE, tipo de obra...")

        # Classe CNAE — 4120-4
        sel_classe = Select(wait.until(EC.presence_of_element_located(
            (By.XPATH, "//select[contains(@id,'classeCnae') or contains(@name,'classeCnae')]"))))
        for opt in sel_classe.options:
            if "4120" in opt.text:
                sel_classe.select_by_visible_text(opt.text)
                break

        # Subclasse — 00
        import time; time.sleep(1)
        sel_subclasse = Select(wait.until(EC.presence_of_element_located(
            (By.XPATH, "//select[contains(@id,'subclasse') or contains(@name,'subclasse')]"))))
        for opt in sel_subclasse.options:
            if opt.text.strip().startswith("00"):
                sel_subclasse.select_by_visible_text(opt.text)
                break

        # Tipo de Construção — Edifício
        sel_tipo_const = Select(driver.find_element(By.XPATH,
            "//select[contains(@id,'tipoConstrucao') or contains(@name,'tipoConstrucao')]"))
        for opt in sel_tipo_const.options:
            if "dif" in opt.text.lower():  # Edifício
                sel_tipo_const.select_by_visible_text(opt.text)
                break

        # Tipo de Obra — Privada (radio)
        radio_privada = driver.find_element(By.XPATH,
            "//input[@type='radio' and (@value='Privada' or @value='privada' or contains(@id,'privada'))]")
        radio_privada.click()

        # Característica — Construção (radio)
        radio_construcao = driver.find_element(By.XPATH,
            "//input[@type='radio' and (@value='Construção' or @value='construcao' or contains(@id,'construcao'))]")
        radio_construcao.click()

        step_cb(70, "Preenchendo observação e datas")

        # Observação
        campo_obs = driver.find_element(By.XPATH,
            "//textarea[contains(@id,'observacao') or contains(@name,'observacao') or contains(@id,'descricao')]")
        campo_obs.clear()
        campo_obs.send_keys(dados["observacao"])
        log_cb("  Observação preenchida.")

        # FGTS — Não (radio)
        radio_fgts_nao = driver.find_element(By.XPATH,
            "//input[@type='radio' and (contains(@id,'fgtsNao') or (@name='fgts' and @value='N'))]")
        radio_fgts_nao.click()

        # Data de início
        campo_inicio = driver.find_element(By.XPATH,
            "//input[contains(@id,'dataInicio') or contains(@name,'dataInicio')]")
        campo_inicio.clear()
        campo_inicio.send_keys(dados["data_inicio"])
        log_cb(f"  Data início: {dados['data_inicio']}")

        # Data de término (1 mês após hoje, automático)
        campo_termino = driver.find_element(By.XPATH,
            "//input[contains(@id,'dataTermino') or contains(@name,'dataTermino')]")
        campo_termino.clear()
        campo_termino.send_keys(dados["data_termino"])
        log_cb(f"  Data término: {dados['data_termino']}")

        # Empregados — empresa principal: 0
        campo_emp_princ = driver.find_element(By.XPATH,
            "//input[contains(@id,'empPrincipal') or contains(@name,'empPrincipal') or contains(@name,'numEmpregadosPrincipal')]")
        campo_emp_princ.clear()
        campo_emp_princ.send_keys(EMP_PRINCIPAL)

        # Empregados — terceiros: 5
        campo_emp_terc = driver.find_element(By.XPATH,
            "//input[contains(@id,'empTerceiros') or contains(@name,'empTerceiros') or contains(@name,'numEmpregadosTerceiros')]")
        campo_emp_terc.clear()
        campo_emp_terc.send_keys(EMP_TERCEIROS)

        step_cb(85, "Submetendo formulário")
        log_cb("Submetendo formulário...")

        # Clicar "Comunicar Obra" final
        btn_final = driver.find_element(By.XPATH,
            "//input[@type='submit' and @value='Comunicar Obra'] | //button[contains(text(),'Comunicar Obra')]")
        btn_final.click()

        # ── 7. Aguardar confirmação ───────────────────────────────────────────
        import time; time.sleep(3)
        page_final = driver.page_source
        if "erro" in page_final.lower() or "inválido" in page_final.lower():
            raise Exception("Formulário retornou erro. Verifique os dados preenchidos.")

        log_cb("✔ Formulário submetido com sucesso!")
        step_cb(95, "Aguardando geração do PDF")
        log_cb("Aguardando PDF ser gerado pelo site...")

        # Tenta localizar botão/link de impressão/PDF na página de confirmação
        import time
        pdf_baixado = False
        for tentativa in range(10):
            try:
                btn_pdf = driver.find_element(By.XPATH,
                    "//a[contains(@href,'.pdf')] | //input[contains(@value,'Imprimir')] | //button[contains(text(),'Imprimir')]")
                btn_pdf.click()
                time.sleep(3)
                pdf_baixado = True
                log_cb("✔ PDF gerado/baixado.")
                break
            except Exception:
                time.sleep(2)

        if not pdf_baixado:
            log_cb("⚠ Botão de PDF não localizado automaticamente. Verifique a página no navegador.")

        step_cb(100, "Concluído")
        done_cb(True, "SCPO preenchido com sucesso!")

    except Exception as e:
        log_cb(f"✖ ERRO: {e}")
        done_cb(False, str(e))
    finally:
        # Não fecha o navegador automaticamente — usuário pode precisar verificar
        pass


# ─── Interface Tkinter ───────────────────────────────────────────────────────

class AppSCPO(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SCPO Automation — Berçan Projetos")
        self.configure(bg=COR_BG)
        self.resizable(False, False)

        # Carrega configuração salva (senha)
        self.config_dados = carregar_config()

        # Variáveis de controle
        self.var_esquina    = tk.BooleanVar(value=False)
        self.var_data_ini   = tk.StringVar()
        self.var_senha      = tk.StringVar(value=self.config_dados.get("senha", "SCPO123"))

        # Eventos de sincronização entre UI e thread Selenium
        self._evento_captcha = threading.Event()
        self._evento_senha   = threading.Event()
        self._nova_senha_tmp = ""

        self._build_ui()

    # ─── Construção da interface ──────────────────────────────────────────────

    def _build_ui(self):
        PAD = 12

        # ── Título ──
        tk.Label(self, text="SCPO AUTOMATION", bg=COR_BG, fg=COR_TEXTO,
                 font=("Consolas", 14, "bold")).pack(pady=(PAD, 2))
        tk.Label(self, text="Berçan Projetos — Morais Engenharia", bg=COR_BG,
                 fg=COR_LABEL, font=("Consolas", 9)).pack(pady=(0, PAD))

        # ── Frame principal ──
        frame = tk.Frame(self, bg=COR_BG, padx=PAD, pady=4)
        frame.pack(fill="x")

        def label(row, texto):
            tk.Label(frame, text=texto, bg=COR_BG, fg=COR_LABEL,
                     font=("Consolas", 9), anchor="w").grid(
                row=row, column=0, sticky="w", pady=3, padx=(0, 8))

        def entry(row, var=None, width=32):
            e = tk.Entry(frame, textvariable=var, bg=COR_CAMPO, fg=COR_TEXTO,
                         insertbackground=COR_TEXTO, font=("Consolas", 10),
                         relief="flat", width=width)
            e.grid(row=row, column=1, sticky="w", pady=3)
            return e

        # ── Campos de dados da obra ──
        # "Residencial" = apenas o nome (ex: TUPINIQUIM). App monta "RESIDENCIAL X RUA Y QD Z LT W"
        label(0, "Nome Residencial:")
        self.ent_residencial = entry(0)
        tk.Label(frame, text="(só o nome, ex: TUPINIQUIM)", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 7)).grid(row=0, column=1, sticky="e")

        # CEP com botão Buscar — preenche Rua automaticamente
        label(1, "CEP:")
        frame_cep = tk.Frame(frame, bg=COR_BG)
        frame_cep.grid(row=1, column=1, sticky="w", pady=3)
        self.ent_cep = tk.Entry(frame_cep, bg=COR_CAMPO, fg=COR_TEXTO,
                                 insertbackground=COR_TEXTO, font=("Consolas", 10),
                                 relief="flat", width=12)
        self.ent_cep.pack(side="left")
        self.ent_cep.bind("<FocusOut>", lambda e: self._buscar_cep())
        tk.Button(frame_cep, text="Buscar", bg=COR_BOTAO, fg=COR_TEXTO,
                  font=("Consolas", 9), relief="flat", cursor="hand2",
                  command=self._buscar_cep).pack(side="left", padx=6)

        label(2, "Rua Principal:")
        self.ent_rua = tk.Entry(frame, bg=COR_CAMPO, fg=COR_LOG_TEXT,
                                 insertbackground=COR_TEXTO, font=("Consolas", 10),
                                 relief="flat", width=32)
        self.ent_rua.grid(row=2, column=1, sticky="w", pady=3)
        tk.Label(frame, text="(preenchida pelo CEP)", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 7)).grid(row=2, column=1, sticky="e")

        label(3, "Quadra:")
        self.ent_quadra = entry(3, width=10)

        label(4, "Lote:")
        self.ent_lote = entry(4, width=10)

        label(5, "Nº de Casas:")
        self.ent_num_casas = entry(5, width=5)
        self.ent_num_casas.bind("<FocusOut>", self._atualizar_casas_esquina)

        # ── Lote de esquina ──
        tk.Checkbutton(frame, text="Lote de esquina (frente para 2 ruas)",
                       variable=self.var_esquina, command=self._toggle_esquina,
                       bg=COR_BG, fg=COR_LABEL, selectcolor=COR_CAMPO,
                       activebackground=COR_BG, font=("Consolas", 9)
                       ).grid(row=6, column=0, columnspan=2, sticky="w", pady=3)

        label(7, "Rua 2 (esquina):")
        self.ent_rua2 = tk.Entry(frame, bg=COR_CAMPO, fg=COR_TEXTO,
                                  insertbackground=COR_TEXTO, font=("Consolas", 10),
                                  relief="flat", width=32, state="disabled")
        self.ent_rua2.grid(row=7, column=1, sticky="w", pady=3)

        # ── Frame dinâmico das casas (rua por casa, só visível em esquina) ──
        self.frame_casas = tk.Frame(frame, bg=COR_BG)
        self.frame_casas.grid(row=8, column=0, columnspan=2, sticky="w")
        self._entries_rua_casa = []  # Lista de Entry para rua de cada casa

        # ── Data ──
        label(9, "Data de Início:")
        self.ent_data_ini = tk.Entry(frame, textvariable=self.var_data_ini,
                                      bg=COR_CAMPO, fg=COR_TEXTO,
                                      insertbackground=COR_TEXTO, font=("Consolas", 10),
                                      relief="flat", width=12)
        self.ent_data_ini.grid(row=9, column=1, sticky="w", pady=3)
        tk.Label(frame, text="(DD/MM/YYYY)", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 8)).grid(row=9, column=1, sticky="e")

        # Término automático
        tk.Label(frame, text="Data de Término:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9), anchor="w").grid(
            row=10, column=0, sticky="w", pady=3)
        self._lbl_termino = tk.Label(frame, text=data_termino_auto(),
                                      bg=COR_BG, fg=COR_LOG_TEXT, font=("Consolas", 9))
        self._lbl_termino.grid(row=10, column=1, sticky="w")
        tk.Label(frame, text="(automático: +1 mês da data atual)",
                 bg=COR_BG, fg=COR_LABEL, font=("Consolas", 8)
                 ).grid(row=10, column=1, sticky="e")

        # ── Pasta de download ──
        label(11, "Pasta Download:")
        pasta_padrao = os.path.expanduser("~\\Downloads")
        self.ent_pasta = tk.Entry(frame, bg=COR_CAMPO, fg=COR_TEXTO,
                                   insertbackground=COR_TEXTO, font=("Consolas", 10),
                                   relief="flat", width=32)
        self.ent_pasta.insert(0, pasta_padrao)
        self.ent_pasta.grid(row=11, column=1, sticky="w", pady=3)

        # ── Senha ──
        sep = tk.Frame(frame, height=1, bg=COR_CAMPO)
        sep.grid(row=12, column=0, columnspan=2, sticky="ew", pady=8)

        tk.Label(frame, text="Senha SCPO:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9), anchor="w").grid(row=13, column=0, sticky="w")
        frame_senha = tk.Frame(frame, bg=COR_BG)
        frame_senha.grid(row=13, column=1, sticky="w")
        self.ent_senha = tk.Entry(frame_senha, textvariable=self.var_senha,
                                   bg=COR_CAMPO, fg=COR_TEXTO, show="*",
                                   insertbackground=COR_TEXTO, font=("Consolas", 10),
                                   relief="flat", width=20)
        self.ent_senha.pack(side="left")
        tk.Button(frame_senha, text="Mostrar", bg=COR_CAMPO, fg=COR_LABEL,
                  font=("Consolas", 8), relief="flat", cursor="hand2",
                  command=self._toggle_senha).pack(side="left", padx=4)

        # ── Barra de progresso ──
        self._var_prog  = tk.DoubleVar(value=0)
        self._var_desc  = tk.StringVar(value="Aguardando...")
        tk.Label(self, textvariable=self._var_desc, bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9)).pack(pady=(PAD, 2))
        self._barra = ttk.Progressbar(self, variable=self._var_prog,
                                       maximum=100, length=440,
                                       style="verde.Horizontal.TProgressbar")
        self._barra.pack(padx=PAD, pady=(0, 4))
        style = ttk.Style()
        style.theme_use("default")
        style.configure("verde.Horizontal.TProgressbar",
                         troughcolor=COR_LOG, background=COR_BARRA)

        # ── Botões ──
        frame_btn = tk.Frame(self, bg=COR_BG)
        frame_btn.pack(pady=6)

        self._btn_run = tk.Button(frame_btn, text="▶  Automatizar SCPO",
                                   bg=COR_BOTAO, fg=COR_TEXTO, font=("Consolas", 11, "bold"),
                                   relief="flat", cursor="hand2", padx=16, pady=6,
                                   command=self._iniciar)
        self._btn_run.pack(side="left", padx=6)

        self._btn_captcha = tk.Button(frame_btn, text="✔ Código digitado — Continuar",
                                       bg="#27ae60", fg=COR_TEXTO, font=("Consolas", 10),
                                       relief="flat", cursor="hand2", padx=12, pady=6,
                                       state="disabled", command=self._liberar_captcha)
        self._btn_captcha.pack(side="left", padx=6)

        self._btn_senha_ok = tk.Button(frame_btn, text="✔ Senha alterada — Continuar",
                                        bg="#e67e22", fg=COR_TEXTO, font=("Consolas", 10),
                                        relief="flat", cursor="hand2", padx=12, pady=6,
                                        state="disabled", command=self._liberar_senha)
        self._btn_senha_ok.pack(side="left", padx=6)

        # ── Log ──
        tk.Label(self, text="Log:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9), anchor="w").pack(fill="x", padx=PAD)
        self._txt_log = tk.Text(self, bg=COR_LOG, fg=COR_LOG_TEXT,
                                 font=("Consolas", 9), height=10,
                                 relief="flat", state="disabled")
        self._txt_log.pack(fill="x", padx=PAD, pady=(0, PAD))

    # ─── Helpers de UI ────────────────────────────────────────────────────────

    def _buscar_cep(self):
        """Consulta ViaCEP e preenche campo Rua automaticamente."""
        cep = self.ent_cep.get().strip()
        if not cep:
            return
        dados = buscar_cep(cep)
        if not dados:
            messagebox.showwarning("CEP não encontrado", f"CEP {cep} não localizado. Verifique e tente novamente.")
            return
        # Preenche Rua (logradouro)
        rua = dados.get("logradouro", "")
        self.ent_rua.config(state="normal")
        self.ent_rua.delete(0, "end")
        self.ent_rua.insert(0, rua.upper())

    def _toggle_esquina(self):
        """Habilita/desabilita campo de rua 2 e campos de rua por casa."""
        if self.var_esquina.get():
            self.ent_rua2.config(state="normal")
        else:
            self.ent_rua2.config(state="disabled")
        self._atualizar_casas_esquina()

    def _atualizar_casas_esquina(self, event=None):
        """Monta dynamicamente campos 'Rua da Casa N' quando em esquina."""
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
        atual = self.ent_senha.cget("show")
        self.ent_senha.config(show="" if atual == "*" else "*")

    def _log(self, msg: str):
        """Escreve mensagem no log (thread-safe via after)."""
        self.after(0, self._log_direto, msg)

    def _log_direto(self, msg: str):
        self._txt_log.config(state="normal")
        self._txt_log.insert("end", f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
        self._txt_log.see("end")
        self._txt_log.config(state="disabled")

    def _step(self, pct: float, desc: str):
        """Atualiza barra de progresso (thread-safe)."""
        self.after(0, lambda: (
            self._var_prog.set(pct),
            self._var_desc.set(desc)
        ))

    def _done(self, ok: bool, msg: str):
        """Finaliza execução (thread-safe)."""
        self.after(0, self._done_ui, ok, msg)

    def _done_ui(self, ok: bool, msg: str):
        self._btn_run.config(state="normal")
        self._btn_captcha.config(state="disabled")
        self._btn_senha_ok.config(state="disabled")
        if ok:
            messagebox.showinfo("Concluído", msg)
        else:
            messagebox.showerror("Erro", f"Falha na automação:\n{msg}")

    def _liberar_captcha(self):
        """Usuário confirmou que digitou o código de segurança no navegador."""
        self._evento_captcha.set()
        self._btn_captcha.config(state="disabled")
        self._log_direto("Usuário confirmou código de segurança.")

    def _liberar_senha(self):
        """Usuário confirmou que alterou a senha no site."""
        nova = tk.simpledialog.askstring(
            "Nova Senha", "Digite a nova senha que você acabou de cadastrar:",
            parent=self)
        if nova:
            self._nova_senha_tmp = nova
            self.var_senha.set(nova)
            self._evento_senha.set()
            self._btn_senha_ok.config(state="disabled")
            self._log_direto(f"Nova senha registrada.")

    # ─── Validação e Início ───────────────────────────────────────────────────

    def _validar(self) -> tuple[bool, str]:
        """Valida campos obrigatórios. Retorna (ok, msg_erro)."""
        campos = [
            (self.ent_residencial.get().strip(), "Residencial"),
            (self.ent_rua.get().strip(), "Rua Principal"),
            (self.ent_quadra.get().strip(), "Quadra"),
            (self.ent_lote.get().strip(), "Lote"),
            (self.ent_cep.get().strip(), "CEP"),
            (self.ent_num_casas.get().strip(), "Nº de Casas"),
            (self.var_data_ini.get().strip(), "Data de Início"),
        ]
        for valor, nome in campos:
            if not valor:
                return False, f"Campo obrigatório vazio: {nome}"

        try:
            n = int(self.ent_num_casas.get())
            if n < 1:
                raise ValueError
        except ValueError:
            return False, "Nº de Casas deve ser um número inteiro positivo."

        try:
            datetime.strptime(self.var_data_ini.get().strip(), "%d/%m/%Y")
        except ValueError:
            return False, "Data de Início inválida. Use o formato DD/MM/YYYY."

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

        # Salva senha atual na config
        salvar_config({"senha": self.var_senha.get().strip()})

        # Monta lista de casas
        n_casas = int(self.ent_num_casas.get())
        casas = []
        for i in range(n_casas):
            rua_casa = self._entries_rua_casa[i].get().strip() if (
                self.var_esquina.get() and i < len(self._entries_rua_casa)) else ""
            casas.append({"numero": i + 1, "rua": rua_casa})

        # Monta dicionário de dados
        obs = gerar_observacao(
            residencial=self.ent_residencial.get().strip(),
            rua=self.ent_rua.get().strip(),
            quadra=self.ent_quadra.get().strip(),
            lote=self.ent_lote.get().strip(),
            esquina=self.var_esquina.get(),
            rua2=self.ent_rua2.get().strip(),
            casas=casas,
        )

        dados = {
            "nome_obra":       nome_obra(
                                   self.ent_residencial.get().strip(),
                                   self.ent_rua.get().strip(),
                                   self.ent_quadra.get().strip(),
                                   self.ent_lote.get().strip()),
            "residencial":     self.ent_residencial.get().strip(),
            "rua":             self.ent_rua.get().strip(),
            "rua2":            self.ent_rua2.get().strip(),
            "quadra":          self.ent_quadra.get().strip(),
            "lote":            self.ent_lote.get().strip(),
            "cep":             self.ent_cep.get().strip(),
            "esquina":         self.var_esquina.get(),
            "casas":           casas,
            "data_inicio":     self.var_data_ini.get().strip(),
            "data_termino":    data_termino_auto(),
            "observacao":      obs,
            "pasta_download":  self.ent_pasta.get().strip(),
            "evento_captcha":  self._evento_captcha,
            "evento_senha":    self._evento_senha,
            "nova_senha":      self._nova_senha_tmp,
        }

        # Reseta eventos
        self._evento_captcha.clear()
        self._evento_senha.clear()

        # UI: desabilita botão, habilita captcha
        self._btn_run.config(state="disabled")
        self._btn_captcha.config(state="normal")
        self._var_prog.set(0)
        self._var_desc.set("Iniciando...")
        self._txt_log.config(state="normal")
        self._txt_log.delete("1.0", "end")
        self._txt_log.config(state="disabled")

        # Dispara thread
        threading.Thread(
            target=executar_scpo,
            args=(dados, self.var_senha.get().strip(),
                  self._step, self._log, self._done),
            daemon=True
        ).start()


# ─── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import tkinter.simpledialog  # importação necessária para askstring
    AppSCPO().mainloop()
