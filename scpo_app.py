# PROTEÇÃO ANTI-LOOP — deve ser a PRIMEIRA coisa no arquivo
import multiprocessing
multiprocessing.freeze_support()

import os, sys, json, threading, io
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk

# ── Pasta permanente para o Chromium (sobrevive entre execuções) ──────────────
# Fica em C:\Users\<usuario>\AppData\Local\SCPOBrowser\
BROWSER_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "SCPOBrowser"

# ── Configura PLAYWRIGHT_BROWSERS_PATH para pasta permanente ─────────────────
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(BROWSER_DIR)

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ── Constantes ────────────────────────────────────────────────────────────────
URL_LOGIN = "https://scpo.mte.gov.br/"
LOGIN_CPF = "038.144.411-25"

SEL_CPF     = "#PlaceHolderConteudo_txtCPF"
SEL_SENHA   = "#PlaceHolderConteudo_txtSenha"
SEL_CAPTCHA = "#txtCaptcha"
SEL_IMG_CAP = "img[src*='CaptchaImage']"
SEL_BTN     = "#PlaceHolderConteudo_btnLogin"

COR_BG       = "#1e2a3a"
COR_LOG      = "#131c26"
COR_CAMPO    = "#2a3f55"
COR_BOTAO    = "#2e86de"
COR_BARRA    = "#4cd964"
COR_TEXTO    = "#ffffff"
COR_LABEL    = "#90adc4"
COR_LOG_TEXT = "#7ec8a0"

CONFIG_PATH = Path(
    sys.executable if getattr(sys, "frozen", False) else __file__
).parent / "scpo_config.json"

def carregar_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text()) if CONFIG_PATH.exists() else {}
    except Exception:
        return {}

def salvar_config(cfg: dict):
    CONFIG_PATH.write_text(json.dumps(cfg))

# ── Instala Chromium se necessário ────────────────────────────────────────────
def garantir_chromium(log_cb) -> bool:
    """
    Verifica se o Chromium já está instalado em BROWSER_DIR.
    Se não, baixa automaticamente (~170MB, só na primeira vez).
    Retorna True se OK, False se falhou.
    """
    import subprocess
    
    # Verifica se já existe alguma pasta chromium-* em BROWSER_DIR
    BROWSER_DIR.mkdir(parents=True, exist_ok=True)
    chromium_dirs = list(BROWSER_DIR.glob("chromium-*"))
    
    if chromium_dirs:
        chrome_exe = chromium_dirs[0] / "chrome-win64" / "chrome.exe"
        if chrome_exe.exists():
            log_cb(f"Chromium encontrado: {chromium_dirs[0].name}")
            return True
    
    # Precisa baixar
    log_cb("Chromium nao encontrado. Baixando (~170MB)...")
    log_cb("Isso so acontece na PRIMEIRA execucao. Aguarde...")
    
    try:
        # Usa o playwright CLI para instalar
        exe = sys.executable
        result = subprocess.run(
            [exe, "-m", "playwright", "install", "chromium"],
            env={**os.environ, "PLAYWRIGHT_BROWSERS_PATH": str(BROWSER_DIR)},
            capture_output=True, text=True, timeout=300
        )
        if result.returncode == 0:
            log_cb("Chromium instalado com sucesso!")
            return True
        else:
            log_cb(f"Erro ao instalar Chromium: {result.stderr[:200]}")
            return False
    except Exception as e:
        log_cb(f"Erro ao instalar Chromium: {e}")
        return False

# ── Automação ─────────────────────────────────────────────────────────────────
def executar_login(senha: str, step_cb, log_cb, done_cb,
                   fn_mostrar_captcha, evento_captcha, dados_captcha):
    import traceback
    try:
        # ── Garantir Chromium ─────────────────────────────────────────────────
        log_cb("Verificando Chromium...")
        step_cb(5, "Verificando navegador")
        if not garantir_chromium(log_cb):
            raise Exception(
                "Nao foi possivel instalar o Chromium.\n"
                "Verifique sua conexao com a internet e tente novamente.")

        step_cb(15, "Iniciando Playwright")
        with sync_playwright() as p:
            log_cb("Iniciando Chromium...")
            step_cb(20, "Abrindo navegador")

            browser = p.chromium.launch(
                headless=False,
                args=["--start-maximized", "--disable-notifications"],
            )
            context = browser.new_context(
                locale="pt-BR",
                ignore_https_errors=True,
            )
            page = context.new_page()

            # ── Abrir site ────────────────────────────────────────────────────
            log_cb(f"Abrindo {URL_LOGIN}...")
            step_cb(30, "Carregando site")
            page.goto(URL_LOGIN, wait_until="networkidle", timeout=60_000)
            log_cb("Pagina carregada!")

            # ── Preencher login e senha ───────────────────────────────────────
            step_cb(45, "Preenchendo login")
            page.fill(SEL_CPF, LOGIN_CPF)
            page.fill(SEL_SENHA, senha)
            log_cb("CPF e senha preenchidos.")

            # ── Capturar captcha ──────────────────────────────────────────────
            step_cb(55, "Aguardando captcha")
            log_cb("Capturando imagem do captcha...")
            img_bytes = b""
            try:
                img_elem = page.locator(SEL_IMG_CAP).first
                img_elem.wait_for(state="visible", timeout=10_000)
                img_bytes = img_elem.screenshot()
                log_cb(f"Captcha capturado ({len(img_bytes)} bytes).")
            except Exception as ex:
                log_cb(f"Captcha nao capturado: {ex}")
                log_cb("Digite o codigo visivel no navegador.")

            # Mostra popup com imagem
            dados_captcha["img_bytes"] = img_bytes
            fn_mostrar_captcha()
            evento_captcha.wait()
            codigo = dados_captcha.get("valor", "")
            log_cb(f"Codigo: {codigo}")

            # ── Preencher captcha e entrar ────────────────────────────────────
            step_cb(70, "Efetuando login")
            if codigo:
                page.fill(SEL_CAPTCHA, codigo)

            log_cb("Clicando Entrar...")
            try:
                page.click(SEL_BTN, timeout=5_000)
            except PWTimeout:
                page.evaluate(
                    "document.getElementById('PlaceHolderConteudo_btnLogin').click()")

            page.wait_for_load_state("networkidle", timeout=30_000)
            url = page.url
            log_cb(f"URL apos login: {url}")

            # ── Verificar resultado ───────────────────────────────────────────
            step_cb(100, "Concluido!")
            if page.locator(SEL_CPF).count() > 0:
                raise Exception("Login falhou — senha ou captcha incorreto.")

            log_cb("LOGIN REALIZADO COM SUCESSO!")
            done_cb(True, "Login OK!\nNavegador aberto.")

    except Exception as e:
        tb = traceback.format_exc()
        log_cb(f"ERRO: {e}")
        for linha in tb.splitlines():
            log_cb(linha)
        done_cb(False, str(e))


# ── Interface ─────────────────────────────────────────────────────────────────
class AppSCPO(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SCPO — Login Test (Playwright)")
        self.configure(bg=COR_BG)
        self.resizable(False, False)
        cfg = carregar_config()
        self.var_senha       = tk.StringVar(value=cfg.get("senha", "SCPO123"))
        self._evento_captcha = threading.Event()
        self._dados_captcha  = {}
        self._build_ui()

    def _build_ui(self):
        PAD = 16
        tk.Label(self, text="SCPO — TESTE DE LOGIN", bg=COR_BG, fg=COR_TEXTO,
                 font=("Consolas", 13, "bold")).pack(pady=(PAD, 2))
        tk.Label(self, text="Playwright • Bercan Projetos", bg=COR_BG,
                 fg=COR_LABEL, font=("Consolas", 9)).pack(pady=(0, PAD))

        frame = tk.Frame(self, bg=COR_BG, padx=PAD)
        frame.pack(fill="x")

        tk.Label(frame, text="Login (CPF):", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 10), anchor="w"
                 ).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 8))
        tk.Label(frame, text=LOGIN_CPF, bg=COR_BG, fg=COR_LOG_TEXT,
                 font=("Consolas", 10)).grid(row=0, column=1, sticky="w")

        tk.Label(frame, text="Senha SCPO:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 10), anchor="w"
                 ).grid(row=1, column=0, sticky="w", pady=6, padx=(0, 8))
        fs = tk.Frame(frame, bg=COR_BG)
        fs.grid(row=1, column=1, sticky="w")
        self.ent_senha = tk.Entry(fs, textvariable=self.var_senha,
                                   bg=COR_CAMPO, fg=COR_TEXTO, show="*",
                                   insertbackground=COR_TEXTO,
                                   font=("Consolas", 11), relief="flat", width=20)
        self.ent_senha.pack(side="left")
        tk.Button(fs, text="Mostrar", bg=COR_CAMPO, fg=COR_LABEL,
                  font=("Consolas", 8), relief="flat", cursor="hand2",
                  command=self._toggle_senha).pack(side="left", padx=4)

        self._var_prog = tk.DoubleVar(value=0)
        self._var_desc = tk.StringVar(value="Aguardando...")
        tk.Label(self, textvariable=self._var_desc, bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9)).pack(pady=(PAD, 2))
        self._barra = ttk.Progressbar(self, variable=self._var_prog,
                                       maximum=100, length=420,
                                       style="v.Horizontal.TProgressbar")
        self._barra.pack(padx=PAD, pady=(0, 6))
        s = ttk.Style(); s.theme_use("default")
        s.configure("v.Horizontal.TProgressbar",
                     troughcolor=COR_LOG, background=COR_BARRA)

        self._btn_run = tk.Button(self, text="▶  Iniciar Login",
                                   bg=COR_BOTAO, fg=COR_TEXTO,
                                   font=("Consolas", 11, "bold"),
                                   relief="flat", cursor="hand2",
                                   padx=20, pady=8, command=self._iniciar)
        self._btn_run.pack(pady=6)

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
        (messagebox.showinfo if ok else messagebox.showerror)("Resultado", msg)

    def _mostrar_captcha(self):
        self.after(0, self._abrir_popup_captcha)

    def _abrir_popup_captcha(self):
        img_bytes = self._dados_captcha.get("img_bytes", b"")
        win = tk.Toplevel(self)
        win.title("Codigo de Seguranca")
        win.configure(bg=COR_BG)
        win.grab_set()
        win.resizable(False, False)
        win.attributes("-topmost", True)

        if img_bytes:
            try:
                img = Image.open(io.BytesIO(img_bytes))
                img = img.resize((img.width * 3, img.height * 3), Image.NEAREST)
                photo = ImageTk.PhotoImage(img)
                lbl = tk.Label(win, image=photo, bg=COR_BG)
                lbl.image = photo
                lbl.pack(padx=16, pady=(16, 4))
            except Exception:
                pass
        else:
            tk.Label(win, text="Digite o codigo visivel no navegador.",
                     bg=COR_BG, fg=COR_LABEL,
                     font=("Consolas", 9)).pack(pady=12)

        tk.Label(win, text="Digite o codigo:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 10)).pack()
        var = tk.StringVar()
        ent = tk.Entry(win, textvariable=var, bg=COR_CAMPO, fg=COR_TEXTO,
                       font=("Consolas", 14, "bold"), justify="center",
                       insertbackground=COR_TEXTO, relief="flat", width=12)
        ent.pack(pady=8)
        ent.focus()

        def confirmar(_=None):
            self._dados_captcha["valor"] = var.get().strip()
            self._evento_captcha.set()
            win.destroy()

        tk.Button(win, text="Confirmar", bg=COR_BOTAO, fg=COR_TEXTO,
                  font=("Consolas", 11, "bold"), relief="flat",
                  cursor="hand2", padx=16, pady=6,
                  command=confirmar).pack(pady=(0, 16))
        ent.bind("<Return>", confirmar)

    def _iniciar(self):
        if not self.var_senha.get().strip():
            messagebox.showwarning("Aviso", "Informe a senha.")
            return
        salvar_config({"senha": self.var_senha.get().strip()})
        self._evento_captcha.clear()
        self._dados_captcha.clear()
        self._btn_run.config(state="disabled")
        self._var_prog.set(0)
        self._var_desc.set("Iniciando...")
        self._txt_log.config(state="normal")
        self._txt_log.delete("1.0", "end")
        self._txt_log.config(state="disabled")

        threading.Thread(
            target=executar_login,
            args=(self.var_senha.get().strip(),
                  self._step, self._log, self._done,
                  self._mostrar_captcha,
                  self._evento_captcha,
                  self._dados_captcha),
            daemon=True
        ).start()


if __name__ == "__main__":
    AppSCPO().mainloop()
