# PROTEÇÃO ANTI-LOOP — deve ser a PRIMEIRA coisa no arquivo
import multiprocessing
multiprocessing.freeze_support()

import os, sys, json, threading, io
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk

# ── Runtime hook Playwright — antes de qualquer import playwright ─────────────
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    _cand = Path(sys._MEIPASS) / "playwright" / "driver" / "package" / ".local-browsers"
    if _cand.exists():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(_cand)

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ── Constantes ────────────────────────────────────────────────────────────────
URL_LOGIN = "https://scpo.mte.gov.br/"
LOGIN_CPF = "038.144.411-25"

# Seletores confirmados via DevTools
SEL_CPF     = "#PlaceHolderConteudo_txtCPF"
SEL_SENHA   = "#PlaceHolderConteudo_txtSenha"
SEL_CAPTCHA = "#txtCaptcha"
SEL_IMG_CAP = "img[src*='CaptchaImage']"
SEL_BTN     = "#PlaceHolderConteudo_btnLogin"

# ── Paleta Morais ─────────────────────────────────────────────────────────────
COR_BG       = "#1e2a3a"
COR_LOG      = "#131c26"
COR_CAMPO    = "#2a3f55"
COR_BOTAO    = "#2e86de"
COR_BARRA    = "#4cd964"
COR_TEXTO    = "#ffffff"
COR_LABEL    = "#90adc4"
COR_LOG_TEXT = "#7ec8a0"

# ── Config persistente ────────────────────────────────────────────────────────
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

# ── Automação — só login ──────────────────────────────────────────────────────
def executar_login(senha: str, step_cb, log_cb, done_cb,
                   fn_mostrar_captcha, evento_captcha, dados_captcha):
    import traceback
    try:
        with sync_playwright() as p:
            log_cb("Iniciando Chromium (Playwright)...")
            step_cb(10, "Abrindo navegador")

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
            step_cb(20, "Carregando site")
            page.goto(URL_LOGIN, wait_until="networkidle", timeout=60_000)
            log_cb("Pagina carregada.")

            # ── Preencher CPF e senha ─────────────────────────────────────────
            step_cb(35, "Preenchendo login")
            page.fill(SEL_CPF, LOGIN_CPF)
            page.fill(SEL_SENHA, senha)
            log_cb("CPF e senha preenchidos.")

            # ── Capturar imagem do captcha ────────────────────────────────────
            step_cb(50, "Aguardando captcha")
            log_cb("Capturando imagem do captcha...")
            try:
                img_elem = page.locator(SEL_IMG_CAP).first
                img_elem.wait_for(state="visible", timeout=10_000)
                img_bytes = img_elem.screenshot()
                log_cb(f"Captcha capturado ({len(img_bytes)} bytes).")
            except Exception as ex:
                log_cb(f"Captcha nao encontrado ({ex}) — tente digitar no navegador.")
                img_bytes = b""

            # Envia imagem para UI exibir popup
            dados_captcha["img_bytes"] = img_bytes
            fn_mostrar_captcha()          # dispara popup na thread da UI
            evento_captcha.wait()         # aguarda usuário confirmar
            codigo = dados_captcha.get("valor", "")
            log_cb(f"Codigo recebido: {codigo}")

            # ── Preencher captcha e clicar Entrar ─────────────────────────────
            step_cb(70, "Efetuando login")
            if codigo:
                page.fill(SEL_CAPTCHA, codigo)

            log_cb("Clicando em Entrar...")
            try:
                page.click(SEL_BTN, timeout=5_000)
            except PWTimeout:
                page.evaluate("__doPostBack('ctl00$PlaceHolderConteudo$btnLogin','')")

            page.wait_for_load_state("networkidle", timeout=30_000)
            url_atual = page.url
            log_cb(f"URL apos login: {url_atual}")

            # ── Verificar resultado ───────────────────────────────────────────
            step_cb(100, "Concluido")
            if "Default.aspx" in url_atual and page.locator(SEL_CPF).count() > 0:
                raise Exception("Login falhou — verifique senha e captcha.")

            log_cb("Login realizado com sucesso!")
            done_cb(True, "Login OK!\nNavegador aberto — proxima etapa: navegacao.")
            # Navegador permanece aberto

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

        # CPF (só leitura)
        tk.Label(frame, text="Login (CPF):", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 10), anchor="w"
                 ).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 8))
        tk.Label(frame, text=LOGIN_CPF, bg=COR_BG, fg=COR_LOG_TEXT,
                 font=("Consolas", 10)).grid(row=0, column=1, sticky="w")

        # Senha
        tk.Label(frame, text="Senha SCPO:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 10), anchor="w"
                 ).grid(row=1, column=0, sticky="w", pady=6, padx=(0, 8))
        frame_s = tk.Frame(frame, bg=COR_BG)
        frame_s.grid(row=1, column=1, sticky="w")
        self.ent_senha = tk.Entry(frame_s, textvariable=self.var_senha,
                                   bg=COR_CAMPO, fg=COR_TEXTO, show="*",
                                   insertbackground=COR_TEXTO,
                                   font=("Consolas", 11), relief="flat", width=20)
        self.ent_senha.pack(side="left")
        tk.Button(frame_s, text="Mostrar", bg=COR_CAMPO, fg=COR_LABEL,
                  font=("Consolas", 8), relief="flat", cursor="hand2",
                  command=self._toggle_senha).pack(side="left", padx=4)

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

        # Botão
        self._btn_run = tk.Button(self, text="▶  Iniciar Login",
                                   bg=COR_BOTAO, fg=COR_TEXTO,
                                   font=("Consolas", 11, "bold"),
                                   relief="flat", cursor="hand2",
                                   padx=20, pady=8, command=self._iniciar)
        self._btn_run.pack(pady=6)

        # Log
        tk.Label(self, text="Log:", bg=COR_BG, fg=COR_LABEL,
                 font=("Consolas", 9), anchor="w").pack(fill="x", padx=PAD)
        self._txt_log = tk.Text(self, bg=COR_LOG, fg=COR_LOG_TEXT,
                                 font=("Consolas", 9), height=12,
                                 relief="flat", state="disabled")
        self._txt_log.pack(fill="x", padx=PAD, pady=(0, PAD))

    # ── Helpers ───────────────────────────────────────────────────────────────
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
        """Chamado via after() — abre popup com imagem do captcha."""
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
                tk.Label(win, text="[imagem nao disponivel]",
                         bg=COR_BG, fg=COR_LABEL,
                         font=("Consolas", 9)).pack(pady=8)
        else:
            tk.Label(win,
                     text="Captcha nao capturado.\nDigite o codigo visivel no navegador.",
                     bg=COR_BG, fg=COR_LABEL,
                     font=("Consolas", 9), justify="center").pack(pady=12)

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

    # ── Iniciar ───────────────────────────────────────────────────────────────
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
            args=(
                self.var_senha.get().strip(),
                self._step, self._log, self._done,
                self._mostrar_captcha,
                self._evento_captcha,
                self._dados_captcha,
            ),
            daemon=True
        ).start()


if __name__ == "__main__":
    AppSCPO().mainloop()
