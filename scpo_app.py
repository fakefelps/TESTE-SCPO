# PROTEÇÃO ANTI-LOOP — deve ser a PRIMEIRA coisa no arquivo
import multiprocessing
multiprocessing.freeze_support()

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
import json
import os
import sys
from datetime import datetime

# ─── Selenium ────────────────────────────────────────────────────────────────
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ─── Constantes ──────────────────────────────────────────────────────────────
URL_SCPO  = "https://scpo.mte.gov.br/"
LOGIN_CPF = "038.144.411-25"

COR_BG       = "#1e2a3a"
COR_LOG      = "#131c26"
COR_CAMPO    = "#2a3f55"
COR_BOTAO    = "#2e86de"
COR_BARRA    = "#4cd964"
COR_TEXTO    = "#ffffff"
COR_LABEL    = "#90adc4"
COR_LOG_TEXT = "#7ec8a0"

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

# ─── Automação — só login ─────────────────────────────────────────────────────
def executar_login(dados: dict, senha: str, step_cb, log_cb, done_cb):
    import time, traceback
    driver = None
    try:
        # ── Iniciar Edge ──────────────────────────────────────────────────────
        log_cb("Iniciando Edge...")
        step_cb(10, "Abrindo Edge")

        options = EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        # Tenta usar msedgedriver local (já vem com o Edge)
        caminhos_driver = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedgedriver.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedgedriver.exe",
        ]
        # Busca em subpastas versionadas
        for base in [r"C:\Program Files (x86)\Microsoft\Edge\Application",
                     r"C:\Program Files\Microsoft\Edge\Application"]:
            if os.path.isdir(base):
                for entry in os.listdir(base):
                    p = os.path.join(base, entry, "msedgedriver.exe")
                    if os.path.exists(p):
                        caminhos_driver.append(p)

        driver_path = ""
        for p in caminhos_driver:
            if os.path.exists(p):
                driver_path = p
                break

        if driver_path:
            log_cb(f"EdgeDriver: {driver_path}")
            driver = webdriver.Edge(
                service=EdgeService(driver_path), options=options)
        else:
            log_cb("Driver local nao encontrado — usando Selenium Manager...")
            driver = webdriver.Edge(options=options)

        # Guarda referência para fechar depois se necessário
        dados["driver"] = driver
        wait = WebDriverWait(driver, 30)  # timeout generoso: 30s

        # ── Abrir site ────────────────────────────────────────────────────────
        log_cb(f"Abrindo {URL_SCPO}...")
        step_cb(20, "Carregando site")
        driver.get(URL_SCPO)

        # Aguarda campo de login aparecer (até 30s)
        log_cb("Aguardando pagina carregar...")
        campo_login = wait.until(EC.presence_of_element_located(
            (By.ID, "PlaceHolderConteudo_txtCPF")))
        log_cb("Pagina carregada! Campo de login encontrado.")
        step_cb(40, "Preenchendo login")

        # ── Preenche login e senha ────────────────────────────────────────────
        campo_login.clear()
        campo_login.send_keys(LOGIN_CPF)

        driver.find_element(
            By.ID, "PlaceHolderConteudo_txtSenha").send_keys(senha)

        log_cb("Login e senha preenchidos.")
        log_cb(">>> Digite o codigo de seguranca no navegador")
        log_cb(">>> Depois clique em 'Codigo digitado — Continuar' aqui no app")
        step_cb(50, "Aguardando codigo de seguranca")

        # Habilita botão e aguarda usuário
        dados["fn_habilitar_captcha"]()
        dados["evento_captcha"].wait()

        # ── Clicar Entrar ─────────────────────────────────────────────────────
        log_cb("Clicando em Entrar...")
        step_cb(70, "Efetuando login")
        driver.find_element(By.ID, "PlaceHolderConteudo_btnLogin").click()
        time.sleep(3)

        # ── Verificar resultado ───────────────────────────────────────────────
        url_atual = driver.current_url
        log_cb(f"URL apos login: {url_atual}")

        if "Default.aspx" in url_atual and "login" in driver.page_source.lower():
            raise Exception("Login falhou — verifique senha e codigo de seguranca.")

        log_cb("Login realizado com sucesso!")
        step_cb(100, "Login OK!")
        done_cb(True, "Login realizado com sucesso!\nNavegador aberto — proxima etapa: navegacao.")

    except Exception as e:
        tb = traceback.format_exc()
        log_cb(f"ERRO: {e}")
        for linha in tb.splitlines():
            log_cb(linha)
        done_cb(False, str(e))
    # Navegador permanece aberto


# ─── Interface ────────────────────────────────────────────────────────────────
class AppSCPO(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SCPO — Login Test")
        self.configure(bg=COR_BG)
        self.resizable(False, False)
        self.config_dados    = carregar_config()
        self.var_senha       = tk.StringVar(value=self.config_dados.get("senha", "SCPO123"))
        self._evento_captcha = threading.Event()
        self._build_ui()

    def _build_ui(self):
        PAD = 16
        tk.Label(self, text="SCPO — TESTE DE LOGIN", bg=COR_BG, fg=COR_TEXTO,
                 font=("Consolas", 13, "bold")).pack(pady=(PAD, 2))
        tk.Label(self, text="Bercan Projetos — Morais Engenharia", bg=COR_BG,
                 fg=COR_LABEL, font=("Consolas", 9)).pack(pady=(0, PAD))

        frame = tk.Frame(self, bg=COR_BG, padx=PAD, pady=4)
        frame.pack(fill="x")

        # Senha
        tk.Label(frame, text="Senha SCPO:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 10), anchor="w"
                 ).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 8))
        frame_s = tk.Frame(frame, bg=COR_BG)
        frame_s.grid(row=0, column=1, sticky="w")
        self.ent_senha = tk.Entry(frame_s, textvariable=self.var_senha,
                                   bg=COR_CAMPO, fg=COR_TEXTO, show="*",
                                   insertbackground=COR_TEXTO,
                                   font=("Consolas", 11), relief="flat", width=20)
        self.ent_senha.pack(side="left")
        tk.Button(frame_s, text="Mostrar", bg=COR_CAMPO, fg=COR_LABEL,
                  font=("Consolas", 8), relief="flat", cursor="hand2",
                  command=self._toggle_senha).pack(side="left", padx=4)

        # CPF (só exibição)
        tk.Label(frame, text="Login (CPF):", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 10), anchor="w"
                 ).grid(row=1, column=0, sticky="w", pady=6, padx=(0, 8))
        tk.Label(frame, text=LOGIN_CPF, bg=COR_BG, fg=COR_LOG_TEXT,
                 font=("Consolas", 10)
                 ).grid(row=1, column=1, sticky="w")

        # Barra de progresso
        self._var_prog = tk.DoubleVar(value=0)
        self._var_desc = tk.StringVar(value="Aguardando...")
        tk.Label(self, textvariable=self._var_desc, bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9)).pack(pady=(PAD, 2))
        self._barra = ttk.Progressbar(self, variable=self._var_prog,
                                       maximum=100, length=400,
                                       style="v.Horizontal.TProgressbar")
        self._barra.pack(padx=PAD, pady=(0, 6))
        s = ttk.Style(); s.theme_use("default")
        s.configure("v.Horizontal.TProgressbar",
                     troughcolor=COR_LOG, background=COR_BARRA)

        # Botões
        frame_btn = tk.Frame(self, bg=COR_BG)
        frame_btn.pack(pady=6)

        self._btn_run = tk.Button(frame_btn, text="▶  Iniciar Login",
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

        # Log
        tk.Label(self, text="Log:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9), anchor="w").pack(fill="x", padx=PAD)
        self._txt_log = tk.Text(self, bg=COR_LOG, fg=COR_LOG_TEXT,
                                 font=("Consolas", 9), height=12,
                                 relief="flat", state="disabled")
        self._txt_log.pack(fill="x", padx=PAD, pady=(0, PAD))

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
        (messagebox.showinfo if ok else messagebox.showerror)(
            "Resultado", msg)

    def _liberar_captcha(self):
        self._evento_captcha.set()
        self._btn_captcha.config(state="disabled")
        self._log_direto("Codigo confirmado.")

    def _iniciar(self):
        if not self.var_senha.get().strip():
            messagebox.showwarning("Aviso", "Informe a senha.")
            return
        salvar_config({"senha": self.var_senha.get().strip()})
        self._evento_captcha.clear()

        dados = {
            "evento_captcha":       self._evento_captcha,
            "fn_habilitar_captcha": lambda: self.after(
                0, lambda: self._btn_captcha.config(state="normal")),
            "driver": None,
        }

        self._btn_run.config(state="disabled")
        self._btn_captcha.config(state="disabled")
        self._var_prog.set(0)
        self._var_desc.set("Iniciando...")
        self._txt_log.config(state="normal")
        self._txt_log.delete("1.0", "end")
        self._txt_log.config(state="disabled")

        threading.Thread(
            target=executar_login,
            args=(dados, self.var_senha.get().strip(),
                  self._step, self._log, self._done),
            daemon=True
        ).start()


if __name__ == "__main__":
    AppSCPO().mainloop()
