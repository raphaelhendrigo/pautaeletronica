# login.py
from __future__ import annotations
import re
from playwright.sync_api import Page, TimeoutError as PWTimeout, Error as PWError

def _try_fill(page: Page, selectors: list[str], text: str) -> bool:
    for sel in selectors:
        loc = page.locator(sel)
        if loc.count() > 0:
            try:
                el = loc.first
                el.click(timeout=3000)
                el.fill("", timeout=3000)
                el.fill(text, timeout=3000)
                return True
            except PWError:
                pass
    return False

def _try_click(page: Page, selectors: list[str]) -> bool:
    for sel in selectors:
        loc = page.locator(sel)
        if loc.count() > 0:
            try:
                loc.first.click(timeout=3000)
                return True
            except PWError:
                pass
    return False

def efetuar_login(page: Page, base_url: str, usuario: str, senha: str) -> None:
    login_url = f"{base_url}/paginas/login.aspx"
    page.goto(login_url, wait_until="domcontentloaded")
    try:
        page.wait_for_selector("input[id*='txtUsuario'], input[name*='txtUsuario'], input[id*='User'], input[placeholder*='Usu'], input[type='text'], input[id*='txtSenha'], input[name*='txtSenha'], input[type='password']", state="visible", timeout=15000)
    except PWTimeout:
        page.wait_for_load_state("load", timeout=15000)

    # Possíveis campos/botões (ASP.NET/DevExpress)
    user_selectors = [
        "input[id*='txtUsuario']",
        "input[name*='txtUsuario']",
        "input[id*='User']",
        "input[placeholder*='Usu']",
        "input[type='text']",
    ]
    pass_selectors = [
        "input[id*='txtSenha']",
        "input[name*='txtSenha']",
        "input[type='password']",
    ]
    login_btns = [
        "button:has-text('Entrar')",
        "input[type='submit'][value*='Entrar']",
        "button:has-text('Acessar')",
        "input[type='submit'][value*='Acessar']",
        "button:has-text('Login')",
        "input[type='submit'][value*='Login']",
    ]

    if not _try_fill(page, user_selectors, usuario):
        raise RuntimeError("Campo de usuário não encontrado na tela de login.")
    if not _try_fill(page, pass_selectors, senha):
        raise RuntimeError("Campo de senha não encontrado na tela de login.")

    # Dispara o login; alguns portais fazem navegação completa, outros só postback parcial.
    triggered = False
    try:
        with page.expect_navigation(url=re.compile(r".*(?<!login\.aspx)$"),
                                    wait_until="load", timeout=10000):
            if not _try_click(page, login_btns):
                page.keyboard.press("Enter")
            triggered = True
    except PWTimeout:
        # Sem navegação visível (postback parcial). Segue o fluxo.
        pass

    # Aguarda estabilizar requisições
    try:
        page.wait_for_load_state("load", timeout=15000)
    except PWTimeout:
        pass

    # Se ainda está na página de login, tente detectar mensagem de erro com segurança
    def _has_login_error() -> bool:
        try:
            # classes comuns: validação, popups DevExpress, alerts
            err = page.locator(".validation-summary-errors, .dxpc-content .alert-danger, .alert-danger, .text-danger")
            return err.first.is_visible()
        except PWError:
            return False

    if "login.aspx" in page.url.lower():
        # Se existir mensagem de erro visível, falha de credencial
        if _has_login_error():
            raise RuntimeError("Falha no login: credenciais inválidas ou erro no portal.")

        # Fallback: às vezes autentica, mas não redireciona. Tente ir direto à página alvo.
        page.goto(f"{base_url}/paginas/resultado/consultarSessoesParaGabinete.aspx", wait_until="domcontentloaded")
        try:
            page.wait_for_load_state("load", timeout=15000)
        except PWTimeout:
            pass

        if "login.aspx" in page.url.lower():
            # Continua preso na tela de login
            raise RuntimeError("Login não efetuado (permaneceu na página de login).")

    # Sucesso: estamos autenticados (ou já conseguimos abrir a página alvo).
