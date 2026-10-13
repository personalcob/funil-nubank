#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Robô de download do QuickSight (sem API) -> pasta de rede.

Fluxo por export configurado em config.yaml:
  1. abre o dashboard (reaproveitando a sessão salva; faz login só se precisar)
  2. localiza o visual pelo título, abre o menu "..." e clica em Exportar p/ Excel
  3. salva numa pasta local de staging
  4. valida o arquivo (colunas obrigatórias, qtd. de linhas, data mais recente)
  5. renomeia a aba para "Planilha1" (é o que o OPENROWSET da seção 5 lê)
  6. copia para a rede de forma atômica (o SQL nunca lê arquivo pela metade)
  7. guarda uma cópia datada no histórico e registra tudo em log + CSV

Uso:
  python qs_download.py                    executa todos os exports
  python qs_download.py --only pagamentos_CC
  python qs_download.py --check            só testa login/dashboard/menu, sem baixar
  python qs_download.py --debug            navegador visível, lento, abre o Inspector no erro
  python qs_download.py --set-credentials  grava usuário/senha no Cofre do Windows
  python qs_download.py --login-manual     você loga à mão (SSO/MFA push) e a sessão é salva
  python qs_download.py --sim-date 2026-09-01  simula período (agosto + setembro, para demo)

Códigos de saída: 0 = tudo OK | 1 = algum export falhou | 2 = erro de configuração | 3 = erro inesperado
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import functools
import getpass
import json
import logging
import os
import re
import shutil
import sys
import time
import traceback
import urllib.request
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path, PureWindowsPath

import yaml
from openpyxl import load_workbook
from playwright.sync_api import Error as PWError
from playwright.sync_api import sync_playwright

try:
    import keyring
except ImportError:  # opcional: sem ele, usa variáveis de ambiente
    keyring = None
try:
    import pyotp
except ImportError:  # opcional: só necessário se houver MFA por TOTP
    pyotp = None

BASE_DIR = Path(__file__).resolve().parent
log = logging.getLogger("qs_robo")


# =============================================================================
# Erros
# =============================================================================
class RoboError(Exception):
    """Erro esperado do robô (mensagem já é legível para o log)."""


class ConfigError(RoboError):
    pass


class LoginError(RoboError):
    pass


class CredencialInvalidaError(LoginError):
    """A tela de login recusou as credenciais: NÃO tentar de novo (evita bloquear a conta)."""


class EtapaError(RoboError):
    pass


class ValidacaoError(RoboError):
    pass


# =============================================================================
# Configuração
# =============================================================================
SELETORES_OBRIGATORIOS = [
    "account_input", "account_submit", "username_input", "username_submit",
    "password_input", "password_submit", "mfa_input", "mfa_submit", "login_error",
    "visual_container", "visual_menu_button", "export_menu_texts",
    "export_excel_texts", "export_csv_texts", "export_confirm_texts", "loading_indicators",
]


def carregar_config(caminho: Path) -> dict:
    if not caminho.exists():
        raise ConfigError(f"Arquivo de configuração não encontrado: {caminho}")
    with open(caminho, encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}

    for secao in ("quicksight", "browser", "paths", "exports", "selectors"):
        if secao not in cfg:
            raise ConfigError(f"Seção '{secao}' ausente em {caminho.name}")

    faltando = [k for k in SELETORES_OBRIGATORIOS if k not in cfg["selectors"]]
    if faltando:
        raise ConfigError(f"Seletores ausentes em 'selectors': {faltando}")
    for k in SELETORES_OBRIGATORIOS:
        cfg["selectors"][k] = list(cfg["selectors"][k] or [])

    b = cfg["browser"]
    b.setdefault("headless", True)
    b.setdefault("slow_mo_ms", 0)
    b.setdefault("default_timeout_s", 60)
    b.setdefault("navigation_timeout_s", 90)
    b.setdefault("login_timeout_s", 120)
    b.setdefault("dashboard_load_timeout_s", 180)
    b.setdefault("loading_timeout_s", 90)
    b.setdefault("download_timeout_s", 300)

    cfg.setdefault("retry", {})
    cfg["retry"].setdefault("attempts", 3)
    cfg["retry"].setdefault("wait_s", 60)
    cfg.setdefault("alerts", {})
    cfg.setdefault("periodo", {})

    nomes = set()
    for exp in cfg["exports"]:
        for k in ("name", "dashboard_url", "visual_title", "destination"):
            if not exp.get(k):
                raise ConfigError(f"Export sem o campo obrigatório '{k}': {exp.get('name', exp)}")
        if exp["name"] in nomes:
            raise ConfigError(f"Nome de export repetido: {exp['name']}")
        nomes.add(exp["name"])
        exp.setdefault("export_format", "excel")
        if exp["export_format"] not in ("excel", "csv"):
            raise ConfigError(f"[{exp['name']}] export_format deve ser 'excel' ou 'csv'")
        for chave in ("steps", "export_steps"):
            if chave in exp and (not isinstance(exp[chave], list) or not all(isinstance(x, dict) for x in exp[chave])):
                raise ConfigError(f"[{exp['name']}] '{chave}' deve ser uma lista de passos (cada um com 'role'/'text'/'css')")
        if "export_steps" in exp and not exp["export_steps"]:
            raise ConfigError(f"[{exp['name']}] 'export_steps' está vazio: o último passo é o clique que baixa o arquivo")
        exp.setdefault("header_row", 1)
        exp.setdefault("min_rows", 1)
        exp.setdefault("required_columns", [])
    return cfg


def preparar_pastas(cfg: dict) -> dict:
    p = cfg["paths"]
    pastas = {
        "staging": Path(p.get("staging_dir") or BASE_DIR / "staging"),
        "logs": Path(p.get("logs_dir") or BASE_DIR / "logs"),
        "sessao": Path(p.get("session_file") or BASE_DIR / "sessao" / "storage_state.json"),
    }
    pastas["evidencias"] = pastas["logs"] / "evidencias"
    for k in ("staging", "logs", "evidencias"):
        pastas[k].mkdir(parents=True, exist_ok=True)
    pastas["sessao"].parent.mkdir(parents=True, exist_ok=True)
    return pastas


# =============================================================================
# Log e histórico
# =============================================================================
def configurar_log(pasta_logs: Path, debug: bool) -> None:
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S")
    arq = TimedRotatingFileHandler(pasta_logs / "qs_robo.log", when="midnight",
                                   backupCount=60, encoding="utf-8")
    arq.setFormatter(fmt)
    arq.setLevel(logging.DEBUG)
    tela = logging.StreamHandler(sys.stdout)
    tela.setFormatter(fmt)
    tela.setLevel(logging.DEBUG if debug else logging.INFO)
    log.setLevel(logging.DEBUG)
    log.handlers.clear()
    log.addHandler(arq)
    log.addHandler(tela)


CAMPOS_HISTORICO = ["inicio", "fim", "duracao_s", "export", "status", "tentativas",
                    "linhas", "tamanho_kb", "data_mais_recente", "meses", "data_corte", "dias_uteis",
                    "fim_mes_anterior", "destino", "erro"]


def registrar_historico(pasta_logs: Path, linha: dict) -> None:
    """Uma linha por export por execução — abre direto no Excel (separador ';')."""
    arq = pasta_logs / "historico_execucoes.csv"
    novo = not arq.exists()
    try:
        if not novo:
            with open(arq, encoding="utf-8-sig", newline="") as f:
                cabecalho = next(csv.reader(f, delimiter=";"), [])
            if cabecalho != CAMPOS_HISTORICO:   # formato antigo: guarda o arquivo velho e começa outro
                arq.rename(pasta_logs / f"historico_execucoes_{carimbo()}_antigo.csv")
                novo = True
        with open(arq, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=CAMPOS_HISTORICO, delimiter=";")
            if novo:
                w.writeheader()
            w.writerow({k: linha.get(k, "") for k in CAMPOS_HISTORICO})
    except OSError as e:  # histórico aberto no Excel, por exemplo
        log.warning("Não consegui gravar o histórico (%s): %s", arq, e)


def enviar_alerta(cfg: dict, texto: str) -> None:
    url = (cfg.get("alerts") or {}).get("webhook_url")
    if not url:
        return
    try:
        dados = json.dumps({"text": texto}).encode("utf-8")
        req = urllib.request.Request(url, data=dados, headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=15)
    except Exception as e:
        log.warning("Não consegui enviar o alerta para o webhook: %s", e)


def carimbo() -> str:
    return dt.datetime.now().strftime("%Y%m%d_%H%M%S")


# =============================================================================
# Trava contra execução simultânea
# =============================================================================
class TravaExecucao:
    def __init__(self, caminho: Path, idade_max_h: float = 3):
        self.caminho = caminho
        self.idade_max_s = idade_max_h * 3600

    def __enter__(self):
        if self.caminho.exists():
            idade = time.time() - self.caminho.stat().st_mtime
            if idade < self.idade_max_s:
                raise RoboError(f"Já existe uma execução em andamento (trava {self.caminho}, "
                                f"criada há {idade / 60:.0f} min). Se não houver, apague o arquivo.")
            log.warning("Trava antiga (%.0f min) encontrada e ignorada", idade / 60)
        self.caminho.write_text(f"pid={os.getpid()} inicio={dt.datetime.now().isoformat()}")
        return self

    def __exit__(self, *exc):
        try:
            self.caminho.unlink()
        except FileNotFoundError:
            pass


# =============================================================================
# Credenciais (Cofre do Windows via keyring, ou variáveis de ambiente)
# =============================================================================
def _servico(cfg: dict) -> str:
    return cfg["quicksight"].get("credential_service", "quicksight_robo")


@functools.lru_cache(maxsize=1)
def _credenciais_cache(servico: str) -> tuple:
    usuario = os.environ.get("QS_USERNAME")
    senha = os.environ.get("QS_PASSWORD")
    totp = os.environ.get("QS_TOTP_SECRET")
    if keyring is not None:
        usuario = usuario or keyring.get_password(servico, "username")
        if usuario:
            senha = senha or keyring.get_password(servico, usuario)
        totp = totp or keyring.get_password(servico, "totp_secret")
    if not usuario or not senha:
        raise ConfigError("Credenciais não encontradas. Rode 'python qs_download.py --set-credentials' "
                          "com o MESMO usuário Windows que executa a tarefa agendada, "
                          "ou defina QS_USERNAME / QS_PASSWORD.")
    return usuario, senha, totp


def obter_credenciais(cfg: dict) -> tuple:
    return _credenciais_cache(_servico(cfg))


def definir_credenciais(cfg: dict) -> int:
    if keyring is None:
        print("O pacote 'keyring' não está instalado (pip install keyring).")
        return 2
    servico = _servico(cfg)
    print(f"Gravando credenciais no Cofre de Credenciais do Windows (serviço '{servico}').")
    print("Atenção: rode isto logado com o mesmo usuário Windows da tarefa agendada.\n")
    usuario = input("Usuário do QuickSight: ").strip()
    senha = getpass.getpass("Senha (não aparece na tela): ")
    totp = getpass.getpass("Segredo TOTP do MFA (Enter se não usa MFA por app): ").strip()
    if not usuario or not senha:
        print("Usuário e senha são obrigatórios.")
        return 2
    keyring.set_password(servico, "username", usuario)
    keyring.set_password(servico, usuario, senha)
    if totp:
        keyring.set_password(servico, "totp_secret", totp)
    print("Credenciais gravadas.")
    return 0


def gerar_codigo_mfa(segredo: str | None) -> str:
    if not segredo:
        raise LoginError("O QuickSight pediu MFA, mas não há segredo TOTP configurado. "
                         "Use --set-credentials ou --login-manual (ver README).")
    if pyotp is None:
        raise LoginError("MFA por TOTP requer o pacote 'pyotp' (pip install pyotp).")
    return pyotp.TOTP(segredo).now()


# =============================================================================
# Helpers de navegador
# =============================================================================
def _aspas(texto: str) -> str:
    return texto.replace("\\", "\\\\").replace('"', '\\"')


def seletores_por_texto(textos: list[str], papel: str = "menuitem") -> list[str]:
    sels = []
    for t in textos:
        t2 = _aspas(t)
        sels += [f'text="{t2}"', f'role={papel}[name="{t2}"]']
    return sels


def visivel_agora(escopo, candidatos: list[str]):
    """Primeiro elemento visível entre os seletores candidatos (ou None)."""
    for sel in candidatos:
        try:
            loc = escopo.locator(sel)
            for i in range(min(loc.count(), 5)):
                item = loc.nth(i)
                if item.is_visible():
                    return item
        except PWError:
            continue
    return None


def esperar_visivel(escopo, candidatos: list[str], timeout_s: float, oque: str, obrigatorio=True):
    limite = time.monotonic() + timeout_s
    while time.monotonic() < limite:
        loc = visivel_agora(escopo, candidatos)
        if loc is not None:
            return loc
        time.sleep(0.5)
    if obrigatorio:
        raise EtapaError(f"Não encontrei '{oque}' na tela em {timeout_s:.0f}s. Seletores tentados: {candidatos}")
    return None


def eh_tela_login(page, sel: dict) -> bool:
    url = (page.url or "").lower()
    if any(k in url for k in ("signin.aws", "/signin", "/auth/", "login", "saml", "oauth")):
        return True
    return visivel_agora(page, sel["account_input"] + sel["username_input"] + sel["password_input"]) is not None


def esperar_carregamento(page, sel: dict, timeout_s: float) -> None:
    """Espera os indicadores de carregamento sumirem (melhor esforço, não falha)."""
    limite = time.monotonic() + timeout_s
    time.sleep(0.5)
    while time.monotonic() < limite:
        if visivel_agora(page, sel["loading_indicators"]) is None:
            time.sleep(1)  # confirma que não voltou
            if visivel_agora(page, sel["loading_indicators"]) is None:
                return
        time.sleep(1)
    log.warning("Indicadores de carregamento ainda visíveis após %.0fs; seguindo mesmo assim", timeout_s)


def capturar_evidencias(page, nome: str, tentativa: int, pastas: dict) -> None:
    """Print + HTML da tela no momento do erro (a página NÃO guarda a senha no HTML)."""
    if page is None:
        return
    base = pastas["evidencias"] / f"{carimbo()}_{nome}_tentativa{tentativa}"
    try:
        page.screenshot(path=str(base.with_suffix(".png")), full_page=True)
        base.with_suffix(".html").write_text(page.content(), encoding="utf-8")
        log.info("Evidências do erro salvas em %s.png/.html", base)
    except Exception as e:
        log.debug("Não consegui capturar evidências: %s", e)


# =============================================================================
# Login
# =============================================================================
def validar_sessao(navegador, cfg: dict, pastas: dict, debug: bool = False) -> bool:
    """
    Verifica se a sessão salva é válida tentando abrir o dashboard específico.
    Retorna True se válida, False se expirada.
    """
    sessao_path = pastas["sessao"]
    
    if not sessao_path.exists():
        log.info("⚠️  Nenhuma sessão salva. Será necessário fazer login.")
        return False
    
    try:
        log.info("Validando sessão salva...")
        # Criar pasta de downloads se não existir
        pasta_downloads = Path.home() / "Downloads"
        pasta_downloads.mkdir(parents=True, exist_ok=True)
        
        contexto = navegador.new_context(
            storage_state=str(sessao_path),
            accept_downloads=True
        )
        page = contexto.new_page()
        
        # Usar o mesmo dashboard URL do primeiro export (melhor teste)
        dashboard_url = cfg.get("exports", [{}])[0].get("dashboard_url", "https://quicksight.aws.amazon.com/")
        
        try:
            page.goto(dashboard_url, wait_until="domcontentloaded", timeout=15000)
        except Exception:
            # Se falhar em carregar, considera expirada
            contexto.close()
            log.warning("⚠️  Não conseguiu carregar dashboard (sessão expirada ou conexão lenta)")
            return False
        
        # Dar um tempo para a página redirecionar se a sessão expirou
        import time
        time.sleep(2)
        
        # Se a URL ainda contém "login" ou "auth", a sessão expirou
        url_atual = page.url.lower()
        if any(x in url_atual for x in ["login", "auth", "signin"]):
            contexto.close()
            log.warning("⚠️  Sessão expirada (redirecionou para autenticação)")
            return False
        
        # Se conseguiu chegar ao dashboard, a sessão é válida
        contexto.close()
        log.info("✅ Sessão válida")
        return True
        
    except Exception as e:
        log.warning("⚠️  Erro ao validar sessão: %s", str(e)[:150])
        try:
            contexto.close()
        except:
            pass
        return False


def fazer_login(page, cfg: dict) -> None:
    qs, sel = cfg["quicksight"], cfg["selectors"]
    modo = qs.get("login_mode", "quicksight")
    if modo == "none":
        raise LoginError("A sessão salva expirou e login_mode=none. Rode 'python qs_download.py --login-manual'.")

    usuario, senha, totp = obter_credenciais(cfg)
    # (nome da etapa, seletores do campo, seletores do botão, valor, pode comparar valor atual?)
    etapas = [
        ("conta", sel["account_input"], sel["account_submit"], lambda: qs.get("account_name"), True),
        ("usuário", sel["username_input"], sel["username_submit"], lambda: usuario, True),
        ("senha", sel["password_input"], sel["password_submit"], lambda: senha, False),
        ("MFA", sel["mfa_input"], sel["mfa_submit"], lambda: gerar_codigo_mfa(totp), False),
    ]
    log.info("Sessão não autenticada: iniciando login (modo %s)", modo)
    limite = time.monotonic() + cfg["browser"]["login_timeout_s"]

    # Máquina de estados: a cada volta vê quais campos estão na tela e preenche
    # os que faltam. Funciona tanto para a tela em etapas do QuickSight
    # (conta -> usuário -> senha) quanto para a tela única do IAM.
    while time.monotonic() < limite:
        erro = visivel_agora(page, sel["login_error"])
        if erro is not None:
            txt = (erro.inner_text() or "").strip()
            if txt:
                raise CredencialInvalidaError(f"A tela de login recusou o acesso: {txt[:200]}")
        if not eh_tela_login(page, sel):
            log.info("Login concluído")
            return

        ultimo = None
        for nome, campos, botoes, valor_fn, comparar in etapas:
            campo = visivel_agora(page, campos)
            if campo is None:
                continue
            try:
                atual = (campo.input_value() or "").strip()
            except PWError:
                atual = ""
            if atual and not comparar:
                continue  # senha/MFA já digitados, aguardando resposta
            valor = valor_fn()
            if not valor:
                raise LoginError(f"O login pediu '{nome}', mas esse valor não está configurado")
            if comparar and atual.lower() == str(valor).lower():
                continue
            campo.fill(str(valor))
            ultimo = (nome, campo, botoes)

        if ultimo:
            nome, campo, botoes = ultimo
            botao = visivel_agora(page, botoes)
            if botao is not None:
                botao.click()
            else:
                campo.press("Enter")
            log.info("Login: etapa '%s' enviada", nome)  # nunca loga valores
            time.sleep(2.5)
        else:
            time.sleep(1)
    raise LoginError(f"Login não concluído em {cfg['browser']['login_timeout_s']}s")


def salvar_sessao(contexto, pastas: dict) -> None:
    try:
        contexto.storage_state(path=str(pastas["sessao"]))
    except Exception as e:
        log.warning("Não consegui salvar a sessão: %s", e)


def apagar_sessao(pastas: dict) -> None:
    try:
        pastas["sessao"].unlink()
        log.info("Sessão salva descartada (será feito login completo na próxima tentativa)")
    except FileNotFoundError:
        pass


# =============================================================================
# Dashboard, visual e exportação
# =============================================================================
# =============================================================================
# Passos configurados no config.yaml (gravados com o "playwright codegen")
# =============================================================================
_MES_BASE_FORCADO: dt.date | None = None


def definir_mes_base(mes: dt.date | None) -> None:
    """Força o "mês atual" (usado quando o último pagamento do QuickSight cai no mês anterior a D-1)."""
    global _MES_BASE_FORCADO
    _MES_BASE_FORCADO = mes.replace(day=1) if mes else None


def periodo_meses(hoje: dt.date | None = None) -> dict:
    """
    "Mês atual" = mês de D-1 (ou o mês base forçado pela data do último pagamento) e
    "mês anterior" = o mês antes dele. Formato AAAA-MM-01 (como o seletor do QuickSight mostra).
    """
    if _MES_BASE_FORCADO is not None and hoje is None:
        atual = _MES_BASE_FORCADO
    else:
        hoje = hoje or dt.date.today()
        atual = (hoje - dt.timedelta(days=1)).replace(day=1)
    anterior = (atual - dt.timedelta(days=1)).replace(day=1)
    return {"mes_atual": atual.isoformat(), "mes_anterior": anterior.isoformat()}


# =============================================================================
# Data de corte pelo ÚLTIMO PAGAMENTO do QuickSight + dias úteis (mesma regra do SQL)
# =============================================================================
def _feriados(cfg: dict) -> frozenset:
    """Datas que não contam como dia útil (opcional: config 'periodo: feriados'). Use as mesmas do SQL."""
    return frozenset(_para_data(str(d)) for d in (cfg.get("periodo") or {}).get("feriados", []) if _para_data(str(d)))


def _eh_util(d: dt.date, feriados: frozenset) -> bool:
    return d.weekday() < 5 and d not in feriados


def dias_uteis_entre(ini: dt.date, fim: dt.date, feriados: frozenset = frozenset()) -> int:
    n, d = 0, ini
    while d <= fim:
        n += _eh_util(d, feriados)
        d += dt.timedelta(days=1)
    return n


def enesimo_dia_util(ini: dt.date, n: int, feriados: frozenset = frozenset()):
    """N-ésimo dia útil do mês de 'ini' (None se o mês tiver menos que N)."""
    d, cont = ini, 0
    while d.month == ini.month:
        if _eh_util(d, feriados):
            cont += 1
            if cont == n:
                return d
        d += dt.timedelta(days=1)
    return None


def _fim_do_mes(d: dt.date) -> dt.date:
    return (d.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)


def calcular_corte(arquivo_csv: Path, exp: dict, hoje: dt.date | None = None, feriados: frozenset = frozenset()):
    """
    Lê o CSV baixado e devolve a data de corte: a maior data de pagamento (só linhas que o funil usa:
    payment_type=direct e new_old=downpayment), limitada a D-1. Daí tira:
      - mes_base: mês dessa data (o "mês atual" do funil)
      - dias_uteis (N): dias úteis de 01/mês_base até a data de corte
      - período do mês anterior: 01 até o N-ésimo dia útil (ou fim do mês, se ele tiver menos que N)
    Devolve None se o export não tem coluna de data configurada.
    """
    col = (exp.get("cutoff_column") or exp.get("months_column") or "").lower()
    if not col:
        return None
    filtros = exp.get("cutoff_filters")
    if filtros is None:
        filtros = {"payment_type": "direct", "new_old": "downpayment"}
    filtros = {str(k).lower(): str(v).strip().lower() for k, v in filtros.items()}
    hoje = hoje or dt.date.today()
    limite = hoje - dt.timedelta(days=1)

    maximos, cache = {}, {}
    with open(arquivo_csv, encoding="utf-8-sig", errors="replace", newline="") as f:
        primeira = f.readline()
        f.seek(0)
        leitor = csv.reader(f, delimiter=max((",", ";", "\t"), key=primeira.count))
        cab = [c.strip().lower() for c in next(leitor, [])]
        if col not in cab:
            raise ValidacaoError(f"Coluna da data de pagamento '{col}' não existe no arquivo. Cabeçalho: {cab}")
        i_data = cab.index(col)
        idx_filtros = []
        for k, v in filtros.items():
            if k in cab:
                idx_filtros.append((cab.index(k), v))
            else:
                log.warning("[%s] Filtro do corte ignorado: coluna '%s' não existe no arquivo", exp["name"], k)
        for row in leitor:
            if len(row) <= i_data:
                continue
            if any(i < len(row) and row[i].strip().lower() != v for i, v in idx_filtros):
                continue
            bruto = row[i_data]
            if bruto not in cache:
                cache[bruto] = _para_data(bruto)
            d = cache[bruto]
            if d is None:
                continue
            k = (d.year, d.month)
            if k not in maximos or d > maximos[k]:
                maximos[k] = d
    if not maximos:
        log.warning("[%s] Nenhuma data de pagamento válida no arquivo (filtros %s); não deu para calcular o corte",
                    exp["name"], filtros)
        return None

    maior = max(maximos.values())
    corte = min(maior, limite)
    mes_base = corte.replace(day=1)
    n = dias_uteis_entre(mes_base, corte, feriados)
    ant_ini = (mes_base - dt.timedelta(days=1)).replace(day=1)
    ant_fim = enesimo_dia_util(ant_ini, n, feriados) if n >= 1 else None
    ant_limitado = ant_fim is None
    if ant_limitado:
        ant_fim = _fim_do_mes(ant_ini)
    ant_max = maximos.get((ant_ini.year, ant_ini.month))
    return {
        "max_pagamento": maior, "data_corte": corte, "mes_base": mes_base, "dias_uteis": n,
        "ant_ini": ant_ini, "ant_fim": ant_fim, "ant_limitado_ao_fim_do_mes": ant_limitado,
        "ant_max_pagamento": ant_max, "ant_completo": bool(ant_max and ant_max >= ant_fim),
    }


def _log_corte(exp: dict, c: dict) -> None:
    f = lambda d: d.strftime("%d/%m/%Y") if d else "(sem dados)"   # noqa: E731
    nome = exp["name"]
    log.info("[%s] ===== DATA DE CORTE (último pagamento no QuickSight): %s =====", nome, f(c["data_corte"]))
    log.info("[%s] Mês atual   : %s a %s | %d dias úteis", nome, f(c["mes_base"]), f(c["data_corte"]), c["dias_uteis"])
    log.info("[%s] Mês anterior: %s a %s | mesmos %d dias úteis%s", nome, f(c["ant_ini"]), f(c["ant_fim"]),
             c["dias_uteis"], " (o mês tem menos dias úteis; vai até o fim do mês)" if c["ant_limitado_ao_fim_do_mes"] else "")
    if c["dias_uteis"] < 1:
        log.warning("[%s] A data de corte cai antes do 1º dia útil do mês; o SQL vai pedir @DIAS_UTEIS_FORCADO "
                    "ou @DATA_CORTE_MODO = 'D1'", nome)
    if not c["ant_completo"]:
        log.warning("[%s] O mês anterior ainda não tem pagamentos até %s (último: %s). Mantendo os mesmos %d dias úteis.",
                    nome, f(c["ant_fim"]), f(c["ant_max_pagamento"]), c["dias_uteis"])
    if c["max_pagamento"] > c["data_corte"]:
        log.info("[%s] (Há pagamento em %s, depois de D-1; o corte fica limitado a D-1)", nome, f(c["max_pagamento"]))


def baixar_com_corte(page, contexto, exp: dict, cfg: dict, pastas: dict, hoje: dt.date | None = None, debug: bool = False):
    """
    Baixa o arquivo, descobre a data de corte pelo último pagamento e confere se o QuickSight foi
    baixado com os MESES certos. Se o último pagamento cai num mês anterior ao pedido (acontece nos
    primeiros dias do mês, quando os pagamentos atrasam), refaz o download com o mês certo.

    Download incremental: se só alguns meses foram baixados (os outros estavam fechados) e o arquivo parcial
    não permite calcular o corte, baixa TODOS os meses uma vez para conferir (o corte precisa ver o mês anterior).
    """
    feriados = _feriados(cfg)
    corte = None
    deslocou = False
    for rodada in (1, 2, 3):
        arquivo = exportar_visual(page, exp, cfg, pastas, debug)
        if arquivo.suffix.lower() != ".csv":
            return arquivo, None
        corte = calcular_corte(arquivo, exp, hoje, feriados)
        if corte is None:
            ped = _ESTADO.get("pedido") or {}
            parcial = bool(ped.get("requeridos")) and set(ped["requeridos"]) != set(ped.get("baixados") or [])
            calcula_corte = bool(exp.get("cutoff_column") or exp.get("months_column"))
            if calcula_corte and parcial and not _ESTADO["baixar_tudo"] and rodada < 3:
                log.info("[%s] Só baixei %s e nele não há pagamento válido para calcular o corte. "
                         "Baixando todos os meses uma vez para conferir...", exp["name"], ped.get("baixados"))
                _ESTADO["baixar_tudo"] = True
                abrir_dashboard(page, contexto, exp, cfg, pastas)
                continue
            return arquivo, None
        pedido = periodo_meses()["mes_atual"]
        if corte["mes_base"].isoformat() == pedido:
            break
        if deslocou or rodada == 3:
            log.warning("[%s] Mesmo depois de baixar de novo, o mês do último pagamento (%s) não é o pedido (%s)",
                        exp["name"], corte["mes_base"].isoformat(), pedido)
            break
        log.info("[%s] O último pagamento no QuickSight é de %s, mas o mês pedido foi %s. "
                 "Baixando de novo com %s e o mês anterior a ele...", exp["name"],
                 corte["data_corte"].strftime("%d/%m/%Y"), pedido[:7], corte["mes_base"].isoformat()[:7])
        definir_mes_base(corte["mes_base"])
        deslocou = True
        abrir_dashboard(page, contexto, exp, cfg, pastas)
    _log_corte(exp, corte)
    return arquivo, corte


def _fmt(texto):
    """
    Troca os marcadores de mês pelas datas do dia. Para {mes_anterior} = 2026-08-01:
      {mes_anterior}        -> 2026-08-01
      {mes_anterior_busca}  -> 2026-08     (para digitar no campo de busca)
      {mes_anterior_curto}  -> 2026-08-01  (checkbox com ano completo)
    """
    if not isinstance(texto, str):
        return texto
    # {mes_retrasado} = o mês antes do mês anterior (ex.: 2026-08 quando o anterior é 2026-09).
    # Fica FORA de periodo_meses() de propósito: aquela função define os "dois meses esperados" que
    # a validação do arquivo confere. Este marcador só serve para exports que pedem um mês a mais.
    if "{mes_retrasado" in texto:
        _ant = dt.date.fromisoformat(periodo_meses()["mes_anterior"])
        _retr = (_ant - dt.timedelta(days=1)).replace(day=1).isoformat()
        texto = (texto.replace("{mes_retrasado_busca}", _retr[:7])
                      .replace("{mes_retrasado_curto}", _retr)
                      .replace("{mes_retrasado}", _retr))
    for chave, valor in periodo_meses().items():
        texto = (texto.replace("{" + chave + "_busca}", valor[:7])
                      .replace("{" + chave + "_curto}", valor)  # ← use valor inteiro, não valor[4:]
                      .replace("{" + chave + "}", valor))
    return texto


def _nome(valor):
    """Texto simples = 'contém' (sem diferenciar maiúscula). Prefixo re: = expressão regular."""
    if valor is None:
        return None
    v = _fmt(str(valor))
    return re.compile(v[3:], re.IGNORECASE) if v.startswith("re:") else v


def localizar(page, spec: dict):
    exato = bool(spec.get("exact", False))
    if spec.get("role"):
        return page.get_by_role(spec["role"], name=_nome(spec.get("name")), exact=exato)
    if spec.get("text"):
        return page.get_by_text(_nome(spec["text"]), exact=exato)
    if spec.get("css"):
        return page.locator(_fmt(spec["css"]))
    raise ConfigError(f"Passo sem 'role', 'text' ou 'css': {spec}")


def _descrever(spec: dict) -> str:
    alvo = spec.get("role") or ("texto" if spec.get("text") else "css")
    nome = spec.get("name") or spec.get("text") or spec.get("css") or ""
    valor = f" = '{_fmt(str(spec['value']))}'" if spec.get("value") is not None else ""
    return f"{spec.get('action', 'click')} {alvo} '{_fmt(str(nome))}'{valor}"


def _rolar_ate_achar(page, passo: dict, exp: dict) -> bool:
    """
    Listas grandes (como a de meses) só desenham os itens que estão na tela.
    Rola a lista com a roda do mouse até o item aparecer: primeiro para baixo, depois para cima.
    """
    alvo = localizar(page, passo).first

    def visivel() -> bool:
        try:
            return alvo.is_visible()
        except PWError:
            return False

    if visivel():
        return True
    if passo.get("scroll_anchor"):
        ancora = localizar(page, passo["scroll_anchor"]).first
    elif passo.get("role"):
        ancora = page.get_by_role(passo["role"]).first   # qualquer item da mesma lista
    else:
        ancora = None
    if ancora is not None:
        try:
            ancora.hover(timeout=3000)   # deixa o mouse em cima da lista, para a roda rolar ELA
        except PWError:
            log.debug("[%s] Não consegui posicionar o mouse na lista", exp["name"])
    px = int(passo.get("scroll_step_px", 300))
    maximo = int(passo.get("scroll_max", 25))
    log.info("[%s] Item fora da tela; rolando a lista para achar: %s", exp["name"], _descrever(passo))
    for direcao, voltas in ((1, maximo), (-1, maximo * 2)):
        for _ in range(voltas):
            page.mouse.wheel(0, direcao * px)
            time.sleep(0.4)
            if visivel():
                log.info("[%s] Item encontrado depois de rolar", exp["name"])
                return True
    log.warning("[%s] Rolei a lista toda e não achei: %s", exp["name"], _descrever(passo))
    return False


def _achar_visivel(loc, espera_s: float):
    """Primeiro elemento VISÍVEL que casa com o passo (o HTML pode ter outros iguais escondidos)."""
    limite = time.monotonic() + espera_s
    while True:
        try:
            for i in range(min(loc.count(), 10)):
                item = loc.nth(i)
                if item.is_visible():
                    return item
        except PWError:
            pass
        if time.monotonic() >= limite:
            return None
        time.sleep(0.2)


def _logar_itens_da_tela(page, passo: dict, exp: dict) -> None:
    """
    Quando um passo não acha o elemento, mostra no log o que EXISTE na tela com o mesmo papel
    (checkbox, botão, opção...). Assim dá para ver o nome real do item sem abrir captura de tela.
    """
    papel = passo.get("role")
    if not papel or page is None:
        return
    try:
        nomes = page.evaluate(r"""(papel) => {
            const sel = papel === 'checkbox' ? '[role="checkbox"], input[type="checkbox"]' : '[role="' + papel + '"]';
            return Array.from(document.querySelectorAll(sel)).slice(0, 40).map(e => {
                let t = e.getAttribute('aria-label') || '';
                if (!t && e.id) { const l = document.querySelector('label[for="' + e.id + '"]'); if (l) t = l.innerText; }
                if (!t) { const o = e.closest('[role="option"]'); if (o) t = o.innerText; }
                if (!t) t = e.innerText || '';
                return t.replace(/\s+/g, ' ').trim().slice(0, 70);
            });
        }""", papel)
        log.warning("[%s] Itens '%s' que existem na tela (%d, mostrando até 40): %s", exp["name"], papel, len(nomes), nomes)
    except Exception as e:
        log.debug("[%s] Não consegui listar os itens da tela: %s", exp["name"], str(e)[:100])


def executar_passo(page, passo: dict, exp: dict, cfg: dict) -> bool:
    """Executa um passo. Devolve False só se o passo era opcional e o elemento não apareceu."""
    acao = passo.get("action", "click")
    opcional = bool(passo.get("optional", False))
    espera = float(passo.get("timeout_s", 3 if opcional else 30))

    if acao == "pause":
        time.sleep(float(passo.get("seconds", 1)))
        return True
    if acao == "press":
        page.keyboard.press(passo["key"])
        time.sleep(float(passo.get("settle_s", 0.4)))
        return True
    if acao == "limpar_filtros":
        limpar_filtros_anteriores(page, exp, cfg)
        return True
    if acao == "marcar_todos":
        # Deixa o filtro aberto com TODOS os valores marcados (sem restringir o export). Idempotente: se já estiver
        # tudo marcado, não faz nada (check() só clica quando o checkbox ainda não está marcado).
        try:
            opcao = page.get_by_role("option", name="Selecionar todos").first
            opcao.wait_for(state="visible", timeout=espera * 1000)
            caixa = opcao.get_by_role("checkbox")
            if caixa.count():
                caixa.first.check()
            else:
                opcao.click()
            time.sleep(0.5)
            return True
        except PWError:
            if opcional:
                log.info("[%s] Passo opcional ignorado (não apareceu): %s", exp["name"], _descrever(passo))
                return False
            _logar_itens_da_tela(page, {"role": "option"}, exp)
            erro = EtapaError(f"[{exp['name']}] Não encontrei o elemento do passo '{_descrever(passo)}' (a opção "
                              f"'Selecionar todos' do filtro aberto) em {espera:.0f}s. Confira se o filtro anterior abriu.")
            erro.passo = passo
            raise erro
    if acao not in ("click", "check", "fill", "hover", "wait"):
        raise ConfigError(f"Ação desconhecida '{acao}' (use click, check, fill, hover, wait, press, pause, limpar_filtros ou marcar_todos)")
    if acao == "fill" and passo.get("value") is None:
        raise ConfigError(f"O passo 'fill' precisa de 'value': {_descrever(passo)}")

    loc = localizar(page, passo)
    if passo.get("scroll_to_find"):
        _rolar_ate_achar(page, passo, exp)
    alvo = _achar_visivel(loc, espera)
    if alvo is None:
        if opcional:
            log.info("[%s] Passo opcional ignorado (não apareceu): %s", exp["name"], _descrever(passo))
            return False
        _logar_itens_da_tela(page, passo, exp)
        erro = EtapaError(f"[{exp['name']}] Não encontrei o elemento do passo '{_descrever(passo)}' em "
                          f"{espera:.0f}s. Confira o nome no config (compare com o código gravado).")
        erro.passo = passo   # o retry precisa saber QUAL passo falhou
        raise erro

    if acao == "click":
        alvo.click()
    elif acao == "check":
        alvo.check()
    elif acao == "fill":
        alvo.fill(_fmt(str(passo["value"])))
    elif acao == "hover":
        alvo.hover()
    time.sleep(float(passo.get("settle_s", 0.4)))
    if passo.get("wait_loading"):
        esperar_carregamento(page, cfg["selectors"], cfg["browser"]["loading_timeout_s"])
    return True


def _opcao_marcada(opcao):
    """Lê se uma opção de lista está marcada (aria-selected/aria-checked ou checkbox interno). None = não deu para saber."""
    for atributo in ("aria-selected", "aria-checked"):
        valor = opcao.get_attribute(atributo)
        if valor in ("true", "false"):
            return valor == "true"
    caixa = opcao.get_by_role("checkbox")
    if caixa.count():
        return caixa.first.is_checked()
    return None


def limpar_filtros_anteriores(page, exp: dict, cfg: dict, debug: bool = False) -> None:
    """
    Limpa filtros anteriores antes de começar a selecionar novos meses.
    Estratégia: clica 2x no option "Selecionar todos" para toggle para estado vazio.
    """
    try:
        log.debug("[%s] Tentando limpar filtros anteriores...", exp["name"])
        import time
        
        # Procurar option "Selecionar todos" (é um toggle)
        # Usa get_by_role que é mais robusto que query_selector
        try:
            # get_by_role não aceita timeout: a espera (5 s) fica no wait_for.
            opcao_selecionar_todos = page.get_by_role("option", name="Selecionar todos").first
            opcao_selecionar_todos.wait_for(state="visible", timeout=5000)

            # Toggle: clica até "Selecionar todos" ficar DESMARCADO (no máximo 2 cliques).
            # Se o estado não puder ser lido, mantém o comportamento antigo (2 cliques).
            for clique in (1, 2):
                opcao_selecionar_todos.click()
                time.sleep(0.5)
                marcado = _opcao_marcada(opcao_selecionar_todos)
                log.debug("[%s] Clique %d em 'Selecionar todos' -> marcado=%s", exp["name"], clique, marcado)
                if marcado is False:
                    break

            log.info("[%s] Filtros anteriores limpos ('Selecionar todos' desmarcado)", exp["name"])
        except Exception as inner_e:
            log.debug("[%s] Option 'Selecionar todos' não encontrado (timeout de 5s): %s", exp["name"], str(inner_e)[:100])
    
    except Exception as e:
        log.warning("[%s] Erro ao limpar filtros: %s", exp["name"], str(e)[:100])


def executar_passos(page, passos: list, exp: dict, cfg: dict, rotulo: str, meses_info: dict = None) -> None:
    """
    Executa passos, pulando os de meses que estão fechados (se meses_info fornecido).
    meses_info: {'mes_anterior': {'baixar': True/False}, 'mes_atual': {'baixar': True/False}}
    """
    for i, passo in enumerate(passos, 1):
        # Verificar se deve pular este passo (se for de um mês fechado).
        # O mês do passo vem da chave explícita "mes: anterior|atual" do config. Só quando ela não existe
        # (config antigo) usa o texto da descrição, e só as expressões "mês anterior"/"mês atual":
        # "Limpar filtros ANTERIORES" e "aviso de ATUALIZAÇÃO" nunca podem ser tratados como passo de mês.
        if meses_info:
            desc_lower = (passo.get("desc") or _descrever(passo)).lower()
            mes_passo = passo.get("mes")
            if mes_passo is None:
                mes_passo = ("retrasado" if "mês retrasado" in desc_lower else
                             "anterior" if "mês anterior" in desc_lower else
                             "atual" if "mês atual" in desc_lower else None)

            if mes_passo and not meses_info.get(f"mes_{mes_passo}", {}).get("baixar", True):
                log.info("[%s] %s %d/%d: PULADO (mês %s já fechado no arquivo) - %s",
                        exp["name"], rotulo, i, len(passos), mes_passo, passo.get("desc") or _descrever(passo))
                continue

        log.info("[%s] %s %d/%d: %s", exp["name"], rotulo, i, len(passos),
                 passo.get("desc") or _descrever(passo))
        executar_passo(page, passo, exp, cfg)


def verificar_por_passos(page, exp: dict, cfg: dict) -> None:
    """Modo --check: percorre tudo, menos o clique final que baixaria o arquivo."""
    executar_passos(page, exp.get("steps", []), exp, cfg, "Passo", meses_info=None)
    executar_passos(page, exp["export_steps"][:-1], exp, cfg, "Exportar", meses_info=None)
    ultimo = exp["export_steps"][-1]
    alvo = localizar(page, ultimo).first
    alvo.wait_for(state="visible", timeout=30000)
    log.info("[%s] Botão final de exportação encontrado (não clicado, modo check): %s",
             exp["name"], _descrever(ultimo))


def exportar_por_passos(page, exp: dict, cfg: dict, pastas: dict, debug: bool = False) -> Path:
    b = cfg["browser"]
    
    # Determinar quais meses precisam ser baixados
    meses_info = determinar_meses_para_baixar(exp, pastas, cfg)
    papeis = _papeis_de_mes(exp)
    for p in papeis:
        r = meses_info[f"mes_{p}"]
        log.info("[%s] Mês %-9s: %s -> %s", exp["name"], p, "BAIXAR " if r["baixar"] else "já fechado", r["motivo"])
    _ESTADO["pedido"] = {"requeridos": [_mes_do_papel(p) for p in papeis],
                         "baixados": [_mes_do_papel(p) for p in papeis if meses_info[f"mes_{p}"]["baixar"]]}
    
    executar_passos(page, exp.get("steps", []), exp, cfg, "Passo", meses_info)
    anteriores, ultimo = exp["export_steps"][:-1], exp["export_steps"][-1]
    executar_passos(page, anteriores, exp, cfg, "Exportar", meses_info=None)
    log.info("[%s] Exportando: %s", exp["name"], ultimo.get("desc") or _descrever(ultimo))
    with page.expect_download(timeout=b["download_timeout_s"] * 1000) as info:
        executar_passo(page, ultimo, exp, cfg)
    download = info.value
    sufixo = Path(download.suggested_filename or "").suffix or (".csv" if exp["export_format"] == "csv" else ".xlsx")
    local = pastas["staging"] / f"{exp['name']}_{carimbo()}{sufixo}"
    download.save_as(str(local))
    falha = download.failure()
    if falha:
        raise EtapaError(f"[{exp['name']}] O download falhou: {falha}")
    log.info("[%s] Download concluído: %s (%.1f KB)", exp["name"], local.name, local.stat().st_size / 1024)
    return local


class MesEncolheuError(ValidacaoError):
    """Um mês baixado veio com muito menos linhas do que o arquivo publicado já tem (download incompleto?)."""


def _texto_celula(v) -> str:
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    return str(v)


def _contar_meses_xlsx(arquivo: Path, exp: dict):
    """Linhas por mês ('AAAA-MM') de um Excel, lendo só a coluna de data. Devolve ({mes: linhas}, linhas_sem_data)."""
    col = (exp.get("freshness_column") or exp.get("months_column") or "date").lower()
    wb = load_workbook(arquivo, read_only=True)
    try:
        ws = wb.worksheets[0]
        linhas = ws.iter_rows(values_only=True)
        cab = [_texto_celula(c).strip().lower() for c in next(linhas, [])]
        if col not in cab:
            raise ValidacaoError(f"Coluna de data '{col}' não existe no arquivo baixado. Cabeçalho: {cab}")
        i = cab.index(col)
        cont, sem_data, cache = {}, 0, {}
        for linha in linhas:
            if linha is None or not any(c not in (None, "") for c in linha):
                continue
            v = linha[i] if i < len(linha) else None
            if v not in cache:
                d = _para_data(v)
                cache[v] = d.strftime("%Y-%m") if d else None
            m = cache[v]
            if m is None:
                sem_data += 1
            else:
                cont[m] = cont.get(m, 0) + 1
        return cont, sem_data
    finally:
        wb.close()


def consolidar_com_cache(arquivo_novo: Path, arquivo_cache, exp: dict, cfg: dict, pastas: dict) -> Path:
    """
    Monta o arquivo que vai ser publicado, com TODOS os meses que a consulta usa:
      - meses baixados   -> vêm do download e SUBSTITUEM tudo o que havia deles no arquivo antigo;
      - meses fechados   -> vêm do arquivo antigo (não foram baixados);
      - linhas do download de meses que NÃO foram pedidos -> ignoradas (o dashboard pode devolver mais meses do que
        o filtro pediu: no LL isso acontece, e somar essas linhas duplicaria os meses fechados);
      - meses fora da lista usada pela consulta (mais antigos) -> saem do arquivo, que não cresce com o tempo.
    Trava de segurança: um mês baixado não pode vir com muito menos linhas do que o arquivo publicado já tem
    (um mês só cresce). Se vier, não substitui e levanta MesEncolheuError. Para aceitar mesmo assim: --baixar-tudo.
    Tudo como TEXTO, aba e formato iguais ao da conversão do CSV (o SQL lê o mesmo arquivo de sempre).
    Se não der para juntar com segurança, levanta erro: é melhor não publicar do que publicar errado.
    """
    pedido = _ESTADO.get("pedido") or {}
    requeridos, baixados = pedido.get("requeridos") or [], pedido.get("baixados") or []
    if not requeridos:
        return arquivo_novo          # sem informação de meses (export sem passos): publica como veio
    tem_cache = bool(arquivo_cache) and Path(arquivo_cache).exists()
    precisa_juntar = not set(requeridos) <= set(baixados)
    if precisa_juntar and not tem_cache:
        raise ValidacaoError(f"Só baixei {baixados}, mas o arquivo anterior que traria os outros meses não existe mais")

    novos, sem_data_novo = _contar_meses_xlsx(arquivo_novo, exp)

    # Trava: o mês baixado não pode encolher em relação ao que já está publicado
    if tem_cache and not _ESTADO.get("baixar_tudo_usuario"):
        inv = _inventario_do_arquivo(Path(arquivo_cache), exp)
        antigos = inv[0] if inv else {}
        razao = float(exp.get("min_ratio_substituir", 0.9))
        for m in baixados:
            a, n = antigos.get(m, 0), novos.get(m, 0)
            if a > 0 and n < a * razao:
                raise MesEncolheuError(
                    f"O mês {m} veio do dashboard com {n:,} linhas e o arquivo publicado já tem {a:,} ({n / a:.1%}). "
                    f"Um mês só cresce: o download parece incompleto, então NÃO substituí e mantive o arquivo anterior. "
                    f"Confira o dashboard; se for isso mesmo, rode com --baixar-tudo.".replace(",", "."))

    fora = {m: c for m, c in novos.items() if m not in baixados}
    if not precisa_juntar and not fora and not sem_data_novo:
        return arquivo_novo          # o download já é exatamente o que vai ser publicado

    try:
        import xlsxwriter
    except ImportError:
        raise RoboError("Falta a biblioteca 'xlsxwriter'. Rode:  pip install xlsxwriter")
    col = (exp.get("freshness_column") or exp.get("months_column") or "date").lower()
    saida = pastas["staging"] / f"{exp['name']}_consolidado_{carimbo()}.xlsx"
    wb_n = load_workbook(arquivo_novo, read_only=True)
    wb_c = load_workbook(arquivo_cache, read_only=True) if precisa_juntar else None
    wb_o = xlsxwriter.Workbook(str(saida), {"strings_to_numbers": False, "strings_to_formulas": False,
                                            "strings_to_urls": False})
    cache_d, descartadas_fora, sem_data_ant = {}, 0, 0
    excl_ant = 0
    cont = {}
    try:
        linhas_n = wb_n.worksheets[0].iter_rows(values_only=True)
        cab_n = [_texto_celula(c).strip() for c in next(linhas_n, [])]
        n_l = [c.lower() for c in cab_n]
        if col not in n_l:
            raise ValidacaoError(f"Coluna de data '{col}' não existe no arquivo baixado. Cabeçalho: {cab_n}")
        i_data_n = n_l.index(col)
        ws_o = wb_o.add_worksheet(exp.get("rename_sheet_to") or "Planilha1")
        for j, nome in enumerate(cab_n):
            ws_o.write_string(0, j, nome)
        pos = 1

        def mes_de(v):
            if v not in cache_d:
                d = _para_data(v)
                cache_d[v] = d.strftime("%Y-%m") if d else None
            return cache_d[v]

        if precisa_juntar:                                              # 1) meses FECHADOS, do arquivo antigo
            linhas_c = wb_c.worksheets[0].iter_rows(values_only=True)
            cab_c = [_texto_celula(c).strip() for c in next(linhas_c, [])]
            c_l = [c.lower() for c in cab_c]
            if col not in c_l:
                raise ValidacaoError(f"Coluna de data '{col}' não existe no arquivo anterior. Cabeçalho: {cab_c}")
            sem_coluna = [c for c in cab_n if c.lower() not in c_l]
            if sem_coluna:
                raise ValidacaoError(f"O arquivo anterior não tem as colunas {sem_coluna}; não dá para juntar com segurança")
            mapa = [c_l.index(c.lower()) for c in cab_n]               # coluna do novo -> posição no antigo
            i_data_c = c_l.index(col)
            # exclude_rows vale também para o que já estava no arquivo (ex.: LL sem 'not-commissioned' nos meses antigos)
            regras_c = [(c_l.index(k), vals) for k, vals in
                        {str(k).strip().lower(): {str(x).strip().lower() for x in (v if isinstance(v, (list, tuple)) else [v])}
                         for k, v in (exp.get("exclude_rows") or {}).items()}.items() if k in c_l]
            excl_ant = 0
            for linha in linhas_c:
                if linha is None or not any(c not in (None, "") for c in linha):
                    continue
                if regras_c and any(i < len(linha) and _texto_celula(linha[i]).strip().lower() in vals for i, vals in regras_c):
                    excl_ant += 1
                    continue
                m = mes_de(linha[i_data_c]) if i_data_c < len(linha) else None
                if m is None:
                    sem_data_ant += 1
                    continue
                if m not in requeridos:
                    descartadas_fora += 1
                    continue
                if m in baixados:
                    continue                                             # será substituído pelo download
                for j, k in enumerate(mapa):
                    v = _texto_celula(linha[k]) if k < len(linha) else ""
                    if v != "":
                        ws_o.write_string(pos, j, v)
                cont[m] = cont.get(m, 0) + 1
                pos += 1
        ignoradas_novo = 0
        for linha in linhas_n:                                          # 2) meses BAIXADOS, do download
            if linha is None or not any(c not in (None, "") for c in linha):
                continue
            m = mes_de(linha[i_data_n]) if i_data_n < len(linha) else None
            if m is None or m not in baixados:
                ignoradas_novo += 1                                      # mês que não foi pedido (ou linha sem data)
                continue
            for j in range(len(cab_n)):
                v = _texto_celula(linha[j]) if j < len(linha) else ""
                if v != "":
                    ws_o.write_string(pos, j, v)
            cont[m] = cont.get(m, 0) + 1
            pos += 1
    finally:
        wb_o.close()
        wb_n.close()
        if wb_c is not None:
            wb_c.close()
    if ignoradas_novo:
        log.info("[%s] %d linha(s) do download foram ignoradas: são de meses que não pedi (%s) ou não têm data. "
                 "Os meses fechados continuam vindo do arquivo anterior.", exp["name"], ignoradas_novo,
                 {m: c for m, c in sorted(fora.items())} or "sem data")
    if excl_ant:
        log.info("[%s] %d linha(s) do arquivo anterior foram EXCLUÍDAS pela regra exclude_rows (%s)", exp["name"], excl_ant,
                 exp.get("exclude_rows"))
    if sem_data_ant:
        log.warning("[%s] %d linha(s) do arquivo anterior sem data legível ficaram de fora da junção", exp["name"], sem_data_ant)
    log.info("[%s] Arquivo consolidado: %s | baixados: %s | fechados (do arquivo anterior): %s | linhas por mês: %s | "
             "%d linha(s) de meses antigos descartadas", exp["name"], saida.name, baixados,
             [m for m in requeridos if m not in baixados], dict(sorted(cont.items())), descartadas_fora)
    faltam = [m for m in requeridos if cont.get(m, 0) == 0]
    if faltam:
        log.warning("[%s] O arquivo consolidado não tem linhas de %s (o mês ainda não tem pagamentos?)", exp["name"], faltam)
    return saida


class FiltroTipoError(ValidacaoError):
    """O download não trouxe os tipos de pagamento esperados: o filtro do dashboard não ficou em 'todos'."""


def conferir_tipos(arquivo_csv: Path, exp: dict) -> None:
    """
    O dashboard tem um filtro com a ORIGEM do pagamento (CC: payment_type; LL: deal_type). Ele precisa estar em
    'todos': a consulta SQL usa pagamentos direct e também os indiretos. Se o download vier com poucos tipos
    (ex.: só 'direct'), o filtro não foi limpo e o arquivo estaria incompleto: o robô para antes de publicar.
    Config do export:  type_column: payment_type   |   min_distinct_types: 2
    """
    col = (exp.get("type_column") or "").strip().lower()
    minimo = int(exp.get("min_distinct_types", 0) or 0)
    if not col or minimo < 1:
        return
    with open(arquivo_csv, encoding="utf-8-sig", errors="replace", newline="") as f:
        primeira = f.readline()
        f.seek(0)
        leitor = csv.reader(f, delimiter=max((",", ";", "\t"), key=primeira.count))
        cab = [c.strip().lower() for c in next(leitor, [])]
        if col not in cab:
            raise ValidacaoError(f"Coluna de tipo '{col}' não existe no arquivo baixado. Cabeçalho: {cab}")
        i = cab.index(col)
        cont = {}
        for linha in leitor:
            if len(linha) > i:
                v = linha[i].strip()
                cont[v] = cont.get(v, 0) + 1
    resumo = ", ".join(f"{k or '(vazio)'}={v:,}".replace(",", ".") for k, v in sorted(cont.items(), key=lambda kv: -kv[1]))
    log.info("[%s] Tipos em '%s' no download: %s", exp["name"], col, resumo)
    if len(cont) < minimo:
        raise FiltroTipoError(
            f"O download só tem {len(cont)} valor(es) em '{col}' ({sorted(cont)}); esperava pelo menos {minimo}. "
            f"O filtro '{col}' do dashboard não ficou em 'todos'. Nada foi publicado.")


def csv_para_xlsx(arquivo: Path, exp: dict) -> Path:
    """
    Converte o CSV baixado em .xlsx (aba Planilha1) para a consulta SQL não mudar.
    - Tudo vai como TEXTO, exatamente como veio no CSV: o SQL já faz REPLACE(',', '.') + CONVERT nos valores
      e o driver do Excel não "adivinha" tipos errados. Célula vazia fica em branco (NULL no SQL).
    - Gravado com XlsxWriter porque ele produz um arquivo no formato do próprio Excel (tabela de textos
      compartilhados + campo 'dimension'). O driver ACE do SQL Server NÃO abre o arquivo gerado pelo
      openpyxl em modo de escrita rápida (textos "inline", sem dimension).
    """
    try:
        import xlsxwriter
    except ImportError:
        raise RoboError("Falta a biblioteca 'xlsxwriter'. Rode:  pip install xlsxwriter")

    destino = arquivo.with_suffix(".xlsx")
    limite_linhas = 1_048_576   # máximo de linhas de uma aba do Excel
    # exclude_rows: {coluna: [valores]} = linhas que NÃO entram no arquivo (ex.: LL -> deal_type: [not-commissioned])
    regras = {str(k).strip().lower(): {str(x).strip().lower() for x in (v if isinstance(v, (list, tuple)) else [v])}
              for k, v in (exp.get("exclude_rows") or {}).items()}
    excluidas = {}
    with open(arquivo, encoding="utf-8-sig", errors="replace", newline="") as f:
        amostra = f.read(8192)
        f.seek(0)
        primeira = amostra.splitlines()[0] if amostra else ""
        separador = max((",", ";", "\t"), key=primeira.count)
        wb = xlsxwriter.Workbook(str(destino), {"strings_to_numbers": False, "strings_to_formulas": False,
                                                "strings_to_urls": False})
        ws = wb.add_worksheet(exp.get("rename_sheet_to") or "Planilha1")
        linhas = 0
        idx_regras = []
        for linha in csv.reader(f, delimiter=separador):
            if not any(c.strip() for c in linha):
                continue
            if linhas == 0 and regras:                      # cabeçalho: descobre onde estão as colunas das regras
                cab_l = [c.strip().lower() for c in linha]
                faltam = [c for c in regras if c not in cab_l]
                if faltam:
                    raise ValidacaoError(f"exclude_rows usa a coluna {faltam}, que não existe no arquivo. Cabeçalho: {cab_l}")
                idx_regras = [(cab_l.index(c), vals, c) for c, vals in regras.items()]
            elif idx_regras:
                motivo = next(((c, linha[i].strip()) for i, vals, c in idx_regras
                               if i < len(linha) and linha[i].strip().lower() in vals), None)
                if motivo:
                    excluidas[motivo] = excluidas.get(motivo, 0) + 1
                    continue
            if linhas >= limite_linhas:
                wb.close()
                destino.unlink(missing_ok=True)
                raise ValidacaoError(f"O CSV tem mais de {limite_linhas:,} linhas: não cabe numa aba do Excel".replace(",", "."))
            for col, valor in enumerate(linha):
                if valor != "":
                    ws.write_string(linhas, col, valor)
            linhas += 1
    wb.close()
    log.info("[%s] CSV convertido para Excel: %s (%d linhas incl. cabeçalho, separador %r)",
             exp["name"], destino.name, linhas, separador)
    if excluidas:
        log.info("[%s] Linhas EXCLUÍDAS por regra (exclude_rows): %s", exp["name"],
                 {f"{c}={v}": n for (c, v), n in sorted(excluidas.items())})
    return destino


def abrir_dashboard(page, contexto, exp: dict, cfg: dict, pastas: dict) -> None:
    sel, b = cfg["selectors"], cfg["browser"]
    nome = exp["name"]
    log.info("[%s] Abrindo dashboard", nome)
    page.goto(exp["dashboard_url"], wait_until="domcontentloaded", timeout=b["navigation_timeout_s"] * 1000)

    # Pronto = aba (se configurada) ou título do visual visível
    if exp.get("sheet_tab"):
        alvo = seletores_por_texto([exp["sheet_tab"]], papel="tab")
    else:
        alvo = [f'text="{_aspas(exp["visual_title"])}"']

    pronto_spec = exp.get("ready")

    def pronto() -> bool:
        if pronto_spec:
            try:
                return localizar(page, pronto_spec).first.is_visible()
            except PWError:
                return False
        return visivel_agora(page, alvo) is not None

    fez_login = False
    limite = time.monotonic() + b["dashboard_load_timeout_s"]
    while True:
        if time.monotonic() > limite:
            raise EtapaError(f"[{nome}] Dashboard não carregou em {b['dashboard_load_timeout_s']}s "
                             f"(não apareceu: {_descrever(pronto_spec) if pronto_spec else (exp.get('sheet_tab') or exp['visual_title'])})")
        if pronto():
            break
        if eh_tela_login(page, sel):
            if fez_login:
                raise LoginError("Continua na tela de login depois de autenticar")
            fazer_login(page, cfg)
            salvar_sessao(contexto, pastas)
            fez_login = True
            page.goto(exp["dashboard_url"], wait_until="domcontentloaded",
                      timeout=b["navigation_timeout_s"] * 1000)
            limite = time.monotonic() + b["dashboard_load_timeout_s"]
            continue
        time.sleep(1)

    if exp.get("sheet_tab") and not pronto_spec:
        log.info("[%s] Selecionando a aba '%s'", nome, exp["sheet_tab"])
        visivel_agora(page, alvo).click()
        esperar_visivel(page, [f'text="{_aspas(exp["visual_title"])}"'],
                        b["dashboard_load_timeout_s"], f"visual '{exp['visual_title']}'")

    esperar_carregamento(page, sel, b["loading_timeout_s"])
    log.info("[%s] Dashboard carregado", nome)


def abrir_menu_do_visual(page, exp: dict, sel: dict) -> None:
    """
    Acha o botão '...' do visual SEM depender de conhecer o HTML exato do QuickSight:
    passa o mouse no título do visual e sobe pelos elementos-pai, do mais interno
    para o mais externo, até achar um que contenha um botão de menu visível.
    O mais interno que contém o botão é, por definição, o próprio visual.
    """
    titulo_txt = exp["visual_title"]
    titulo = page.locator(f'text="{_aspas(titulo_txt)}"').first
    titulo.scroll_into_view_if_needed()
    titulo.hover()
    time.sleep(0.8)

    candidatos = []
    for csel in sel["visual_container"]:
        try:
            loc = page.locator(csel).filter(has=page.locator(f'text="{_aspas(titulo_txt)}"'))
            n = loc.count()
            candidatos += [loc.nth(i) for i in range(n - 1, -1, -1)]  # mais interno primeiro
        except PWError:
            continue
    ancestrais = titulo.locator("xpath=ancestor::*[.//button]")
    n = ancestrais.count()
    candidatos += [ancestrais.nth(i) for i in range(n - 1, max(n - 9, -1), -1)]

    for rodada in (1, 2):
        for cont in candidatos:
            try:
                if rodada == 2:
                    cont.hover()
                    time.sleep(0.5)
                botao = visivel_agora(cont, sel["visual_menu_button"])
                if botao is not None:
                    botao.click()
                    log.info("[%s] Menu do visual aberto", exp["name"])
                    return
            except PWError:
                continue
    raise EtapaError(f"[{exp['name']}] Não achei o botão de menu do visual '{titulo_txt}'. "
                     f"Ajuste 'visual_menu_button' no config (use --debug).")


def exportar_visual(page, exp: dict, cfg: dict, pastas: dict, debug: bool = False) -> Path:
    if exp.get("export_steps"):
        return exportar_por_passos(page, exp, cfg, pastas, debug)
    sel, b = cfg["selectors"], cfg["browser"]
    textos = sel["export_excel_texts"] if exp["export_format"] == "excel" else sel["export_csv_texts"]
    itens = seletores_por_texto(textos)
    submenus = [f'text="{_aspas(t)}"' for t in sel["export_menu_texts"]]

    item = None
    for tentativa in (1, 2):
        abrir_menu_do_visual(page, exp, sel)
        item = esperar_visivel(page, itens, 5, "opção de exportar", obrigatorio=False)
        if item is None and submenus:
            sub = visivel_agora(page, submenus)
            if sub is not None:
                sub.hover()
                time.sleep(0.8)
                item = visivel_agora(page, itens)
                if item is None:
                    sub.click()
                    item = esperar_visivel(page, itens, 5, "opção de exportar", obrigatorio=False)
        if item is not None:
            break
        page.keyboard.press("Escape")
        time.sleep(1)
    if item is None:
        raise EtapaError(f"[{exp['name']}] Opção de exportação não encontrada no menu. "
                         f"Textos tentados: {textos}")

    log.info("[%s] Exportando (%s)...", exp["name"], exp["export_format"])
    confirmar = seletores_por_texto(sel["export_confirm_texts"], papel="button")
    with page.expect_download(timeout=b["download_timeout_s"] * 1000) as info:
        item.click()
        if confirmar:
            botao = esperar_visivel(page, confirmar, 5, "confirmação", obrigatorio=False)
            if botao is not None:
                botao.click()
    download = info.value
    sufixo = Path(download.suggested_filename or "").suffix or (".xlsx" if exp["export_format"] == "excel" else ".csv")
    local = pastas["staging"] / f"{exp['name']}_{carimbo()}{sufixo}"
    download.save_as(str(local))
    falha = download.failure()
    if falha:
        raise EtapaError(f"[{exp['name']}] O download falhou: {falha}")
    log.info("[%s] Download concluído: %s (%.1f KB)", exp["name"], local.name, local.stat().st_size / 1024)
    return local


# =============================================================================
# Validação e publicação do arquivo
# =============================================================================
def _para_data(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:  # número serial do Excel
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            return dt.datetime.strptime(s[:19], fmt).date()
        except ValueError:
            continue
    return None


def _validar_linhas(linhas_iter, exp: dict, info: dict, obrigatorias: list) -> dict:
    """
    Validação em passada única (sem carregar o arquivo inteiro na memória).
    Recebe as linhas de um CSV ou de um Excel e confere colunas, quantidade de linhas,
    frescor dos dados e a presença dos dois meses.
    """
    col_data = (exp.get("freshness_column") or "").lower()
    col_meses = (exp.get("months_column") or "").lower()
    cab_l, idx_data, idx_meses = None, None, None
    linhas, maior, presentes, cache = 0, None, set(), {}

    def data_de(v):
        chave = v if isinstance(v, (str, int, float)) else str(v)
        if chave not in cache:
            cache[chave] = _para_data(v)
        return cache[chave]

    for n, row in enumerate(linhas_iter):
        if n == 0:
            cab = [str(c).strip() if c is not None else "" for c in row]
            cab_l = [c.lower() for c in cab]
            faltando = [c for c in obrigatorias if c not in cab_l]
            if faltando:
                raise ValidacaoError(f"Colunas obrigatórias ausentes: {faltando}. Cabeçalho encontrado: {cab}")
            if col_data:
                if col_data not in cab_l:
                    raise ValidacaoError(f"Coluna de frescor '{exp['freshness_column']}' não existe no arquivo")
                idx_data = cab_l.index(col_data)
            if col_meses:
                if col_meses not in cab_l:
                    raise ValidacaoError(f"Coluna de meses '{exp['months_column']}' não existe no arquivo")
                idx_meses = cab_l.index(col_meses)
            continue
        if not any(v not in (None, "") for v in row):
            continue
        linhas += 1
        if idx_data is not None and idx_data < len(row):
            d = data_de(row[idx_data])
            if d and (maior is None or d > maior):
                maior = d
        if idx_meses is not None and idx_meses < len(row):
            d = data_de(row[idx_meses])
            if d:
                presentes.add(d.strftime("%Y-%m"))

    if cab_l is None:
        raise ValidacaoError("Arquivo sem cabeçalho")
    if linhas < exp["min_rows"]:
        raise ValidacaoError(f"Só {linhas} linha(s) de dados (mínimo {exp['min_rows']})")
    info["linhas"] = linhas

    if col_data:
        info["data_mais_recente"] = maior.isoformat() if maior else ""
        idade_max = int(exp.get("freshness_max_age_days", 4))
        if maior is None or (dt.date.today() - maior).days > idade_max:
            msg = (f"Dados possivelmente desatualizados: data mais recente em '{exp['freshness_column']}' = {maior} "
                   f"(tolerância {idade_max} dias). O dataset do QuickSight foi atualizado?")
            if exp.get("freshness_action", "warn") == "fail":
                raise ValidacaoError(msg)
            log.warning("[%s] %s", exp["name"], msg)
    if col_meses:
        info["meses"] = ", ".join(sorted(presentes))
        faltam = [m for m in _meses_esperados() if m not in presentes]
        if faltam:
            msg = (f"O arquivo não tem dados de {faltam} (meses encontrados: {sorted(presentes) or 'nenhum'}). "
                   f"A consulta SQL precisa do mês atual e do anterior. O seletor de mês do dashboard pegou os dois?")
            if exp.get("months_action", "warn") == "fail":
                raise ValidacaoError(msg)
            log.warning("[%s] %s", exp["name"], msg)
    return info


def _validar_xlsx_rapido(arquivo: Path, exp: dict, info: dict, obrigatorias: list) -> dict:
    wb = load_workbook(arquivo, read_only=True)
    try:
        return _validar_linhas(wb.worksheets[0].iter_rows(values_only=True), exp, info, obrigatorias)
    finally:
        wb.close()


def validar_e_ajustar(arquivo: Path, exp: dict) -> dict:
    """
    Garante que o arquivo tem o formato que a consulta SQL espera.
    NÃO reescreve os dados (evita mudar tipos/decimais que o SQL trata com REPLACE):
    só renomeia a aba e, se preciso, remove linhas acima do cabeçalho.
    """
    tamanho = arquivo.stat().st_size
    if tamanho == 0:
        raise ValidacaoError("Arquivo baixado está vazio (0 bytes)")
    info = {"tamanho_kb": round(tamanho / 1024, 1), "linhas": "", "data_mais_recente": ""}
    obrigatorias = [c.strip().lower() for c in exp["required_columns"]]

    if arquivo.suffix.lower() == ".csv":
        with open(arquivo, encoding="utf-8-sig", errors="replace", newline="") as f:
            primeira = f.readline()
            f.seek(0)
            separador = max((",", ";", "\t"), key=primeira.count)
            return _validar_linhas(csv.reader(f, delimiter=separador), exp, info, obrigatorias)

    try:
        wb_teste = load_workbook(arquivo, read_only=True)
        titulo = wb_teste.worksheets[0].title
        wb_teste.close()
    except Exception as e:
        raise ValidacaoError(f"Arquivo não é um Excel válido: {e}")
    if int(exp["header_row"]) == 1 and titulo == (exp.get("rename_sheet_to") or titulo):
        return _validar_xlsx_rapido(arquivo, exp, info, obrigatorias)

    wb = load_workbook(arquivo)   # caminho lento: só quando é preciso mexer no arquivo (ajustar aba/linhas)
    ws = wb.worksheets[0]
    alterou = False

    linha_cab = int(exp["header_row"])
    if linha_cab > 1:
        ws.delete_rows(1, linha_cab - 1)
        alterou = True

    cab = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    cab_l = [c.lower() for c in cab]
    faltando = [c for c in obrigatorias if c not in cab_l]
    if faltando:
        raise ValidacaoError(f"Colunas obrigatórias ausentes: {faltando}. Cabeçalho encontrado: {cab}")

    linhas = sum(1 for r in ws.iter_rows(min_row=2, values_only=True)
                 if any(v not in (None, "") for v in r))
    if linhas < exp["min_rows"]:
        raise ValidacaoError(f"Só {linhas} linha(s) de dados (mínimo {exp['min_rows']})")
    info["linhas"] = linhas

    col_data = exp.get("freshness_column")
    if col_data:
        if col_data.lower() not in cab_l:
            raise ValidacaoError(f"Coluna de frescor '{col_data}' não existe no arquivo")
        idx = cab_l.index(col_data.lower()) + 1
        maior = None
        for (v,) in ws.iter_rows(min_row=2, min_col=idx, max_col=idx, values_only=True):
            d = _para_data(v)
            if d and (maior is None or d > maior):
                maior = d
        info["data_mais_recente"] = maior.isoformat() if maior else ""
        idade_max = int(exp.get("freshness_max_age_days", 4))
        if maior is None or (dt.date.today() - maior).days > idade_max:
            msg = (f"Dados possivelmente desatualizados: data mais recente em '{col_data}' = {maior} "
                   f"(tolerância {idade_max} dias). O dataset do QuickSight foi atualizado?")
            if exp.get("freshness_action", "warn") == "fail":
                raise ValidacaoError(msg)
            log.warning("[%s] %s", exp["name"], msg)

    col_meses = exp.get("months_column")
    if col_meses:
        if col_meses.lower() not in cab_l:
            raise ValidacaoError(f"Coluna de meses '{col_meses}' não existe no arquivo")
        idx = cab_l.index(col_meses.lower()) + 1
        presentes = set()
        for (v,) in ws.iter_rows(min_row=2, min_col=idx, max_col=idx, values_only=True):
            d = _para_data(v)
            if d:
                presentes.add(d.strftime("%Y-%m"))
        info["meses"] = ", ".join(sorted(presentes))
        esperados = _meses_esperados()
        faltam = [m for m in esperados if m not in presentes]
        if faltam:
            msg = (f"O arquivo não tem dados de {faltam} (meses encontrados: {sorted(presentes) or 'nenhum'}). "
                   f"A consulta SQL precisa do mês atual e do anterior. O seletor de mês do dashboard pegou os dois?")
            if exp.get("months_action", "warn") == "fail":
                raise ValidacaoError(msg)
            log.warning("[%s] %s", exp["name"], msg)

    nova_aba = exp.get("rename_sheet_to")
    if nova_aba and ws.title != nova_aba:
        if nova_aba in wb.sheetnames:
            wb[nova_aba].title = f"{nova_aba}_original"[:31]
        ws.title = nova_aba
        alterou = True

    if alterou:
        wb.save(arquivo)
    wb.close()
    return info


def _eh_rede(caminho: Path) -> bool:
    return str(caminho).startswith("\\\\")   # começa com \\servidor


def publicar(arquivo: Path, exp: dict, cfg: dict) -> None:
    destino = Path(exp["destination"])
    rede_liberada = cfg["paths"].get("publicar_na_rede") is True
    alternativo = cfg["paths"].get("destino_alternativo")
    pasta_alt = Path(alternativo) if alternativo else None
    if pasta_alt is not None:
        # Teste em produção: grava numa pasta ALTERNATIVA (escolhida de propósito) e nunca no caminho do config.
        original = destino
        pasta_original = str(PureWindowsPath(str(original)).parent).rstrip("\\/").lower()
        if str(pasta_alt).rstrip("\\/").lower() == pasta_original:
            raise ConfigError("'destino_alternativo' é a MESMA pasta do 'destination' de produção; "
                              "escolha outra pasta para o teste (ex.: ...\\TESTE_ROBO)")
        destino = pasta_alt / PureWindowsPath(str(original)).name
        if not pasta_alt.exists() and pasta_alt.parent.exists():
            pasta_alt.mkdir()
        log.warning("[%s] DESTINO ALTERNATIVO: gravando em %s. O arquivo de produção (%s) NÃO é tocado.",
                    exp["name"], destino, original)
    elif _eh_rede(destino) and not rede_liberada:
        # Trava de segurança: sem 'paths: publicar_na_rede: true' o robô NUNCA grava em pasta de rede.
        original = destino
        destino = BASE_DIR / "teste_saida" / PureWindowsPath(str(destino)).name
        destino.parent.mkdir(parents=True, exist_ok=True)
        log.warning("[%s] MODO TESTE: 'publicar_na_rede' está desligado; NÃO gravei em %s. Arquivo gravado em %s",
                    exp["name"], original, destino)
    if not destino.parent.exists():
        raise RoboError(f"Pasta de destino inacessível: {destino.parent} "
                        f"(rede fora do ar ou usuário da tarefa sem permissão?)")

    # 1) copia com nome temporário  2) confere tamanho  3) troca de uma vez só.
    # Se o SQL estiver lendo nesse momento, ele continua vendo o arquivo antigo inteiro.
    parcial = destino.with_name(f"~{destino.stem}.parcial{destino.suffix}")
    shutil.copy2(arquivo, parcial)
    if parcial.stat().st_size != arquivo.stat().st_size:
        parcial.unlink(missing_ok=True)
        raise RoboError("Cópia para a rede ficou incompleta")
    for i in range(1, 7):
        try:
            os.replace(parcial, destino)
            break
        except PermissionError:
            if i == 6:
                parcial.unlink(missing_ok=True)
                raise RoboError(f"Destino bloqueado (aberto no Excel ou sendo lido pelo SQL?): {destino}")
            log.warning("[%s] Destino em uso, nova tentativa em 20s (%d/5)", exp["name"], i)
            time.sleep(20)
    log.info("[%s] Publicado em %s", exp["name"], destino)

    historico = exp.get("archive_dir") or cfg["paths"].get("archive_dir")
    if historico or pasta_alt is not None:
        pasta = pasta_alt / "historico" if pasta_alt is not None else Path(historico)
        if pasta_alt is None and _eh_rede(pasta) and not rede_liberada:
            pasta = BASE_DIR / "teste_saida" / "historico"
        try:
            pasta.mkdir(parents=True, exist_ok=True)
            copia = pasta / f"{destino.stem}_{carimbo()}{destino.suffix}"
            shutil.copy2(arquivo, copia)
            log.info("[%s] Cópia de histórico: %s", exp["name"], copia.name)
            limpar_antigos(pasta, f"{destino.stem}_*{destino.suffix}",
                           int(cfg["paths"].get("archive_retention_days", 60)))
        except OSError as e:
            log.warning("[%s] Falha ao gravar histórico (arquivo principal já publicado): %s", exp["name"], e)


def extrair_data_maxima_pagamento(arquivo: Path, exp: dict, cfg: dict) -> dt.date | None:
    """
    Extrai a data máxima de pagamento do arquivo baixado.
    Procura na coluna especificada em 'freshness_column' (padrão: 'date').
    """
    try:
        log.debug("[%s] Abrindo arquivo para extrair data máxima: %s", exp["name"], arquivo)
        wb = load_workbook(arquivo, data_only=True)
        ws = wb.active
        
        # Procurar a coluna de data (geralmente 'date')
        coluna_data = exp.get("freshness_column", "date")
        log.debug("[%s] Procurando coluna: %s", exp["name"], coluna_data)
        
        # Encontrar o header
        header_row = exp.get("header_row", 1)
        headers = {}
        for col_idx, cell in enumerate(ws[header_row], 1):
            if cell.value:
                headers[str(cell.value).lower()] = col_idx
        
        log.debug("[%s] Headers encontrados: %s", exp["name"], list(headers.keys())[:10])
        
        # Encontrar a coluna de data
        col_idx = headers.get(coluna_data.lower())
        if not col_idx:
            log.warning("[%s] Coluna '%s' não encontrada. Headers disponíveis: %s", 
                       exp["name"], coluna_data, list(headers.keys())[:10])
            return None
        
        log.debug("[%s] Coluna de data encontrada na coluna %d", exp["name"], col_idx)
        
        # Extrair todas as datas da coluna
        # Aceita: datetime, date, ou string no formato YYYY-MM-DD ou YYYY-MM-DD HH:MM:SS
        datas = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value:
                try:
                    if isinstance(cell.value, dt.datetime):
                        datas.append(cell.value.date())
                    elif isinstance(cell.value, dt.date):
                        datas.append(cell.value)
                    elif isinstance(cell.value, str):
                        # Tentar parsear como string
                        valor_str = cell.value.strip()
                        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y'):
                            try:
                                datas.append(dt.datetime.strptime(valor_str, fmt).date())
                                break
                            except ValueError:
                                continue
                except Exception as e:
                    pass
        
        log.debug("[%s] Encontradas %d datas", exp["name"], len(datas))
        
        if datas:
            data_max = max(datas)
            log.info("[%s] Data máxima de pagamento extraída: %s", exp["name"], data_max)
            return data_max
        else:
            log.warning("[%s] Nenhuma data encontrada na coluna %d", exp["name"], col_idx)
    except Exception as e:
        log.warning("[%s] Erro ao extrair data máxima: %s", exp["name"], str(e)[:200])
        log.debug(traceback.format_exc())
    
    return None


def eh_ultimo_dia_mes(data: dt.date) -> bool:
    """Verifica se a data é o último dia do mês."""
    proximo_dia = data + dt.timedelta(days=1)
    return proximo_dia.month != data.month


def mes_esta_fechado(data_maxima: dt.date) -> bool:
    """Verifica se o mês está fechado (data máxima é o último dia do mês)."""
    return eh_ultimo_dia_mes(data_maxima)


# =============================================================================
# Meses FECHADOS x DEFASADOS (download incremental)
#   O arquivo publicado precisa ter SEMPRE todos os meses que a consulta usa. Mês fechado (já está no arquivo
#   publicado e a data máxima dele chegou ao último dia do mês) não é baixado de novo: sai do arquivo antigo.
#   Mês defasado ou ausente é baixado e SUBSTITUI o que havia daquele mês.
# =============================================================================
_ESTADO = {"baixar_tudo": False, "pedido": None, "inventarios": {}}


def _papeis_de_mes(exp: dict) -> list:
    """Quais meses o export usa: sempre anterior e atual; também o retrasado se algum passo pedir (mes: retrasado)."""
    papeis = {"anterior", "atual"}
    for p in exp.get("steps", []) or []:
        if p.get("mes") == "retrasado":
            papeis.add("retrasado")
    return [r for r in ("retrasado", "anterior", "atual") if r in papeis]


def _mes_do_papel(papel: str) -> str:
    """'AAAA-MM' do mês, já considerando o mês-base forçado."""
    if papel == "retrasado":
        return _fmt("{mes_retrasado_busca}")
    return periodo_meses()["mes_" + papel][:7]


def _inventario_do_arquivo(arquivo: Path, exp: dict):
    """
    Lê só a coluna de data do arquivo já publicado.
    Devolve ({'AAAA-MM': linhas}, data_máxima) ou None se não der para ler.
    """
    try:
        st = arquivo.stat()
        chave = (str(arquivo), st.st_size, st.st_mtime_ns)
        if chave in _ESTADO["inventarios"]:
            return _ESTADO["inventarios"][chave]
        col = (exp.get("freshness_column") or exp.get("months_column") or "date").lower()
        cab_linha = int(exp.get("header_row", 1))
        wb = load_workbook(arquivo, read_only=True)
        try:
            ws = wb.worksheets[0]
            cab = None
            for n, linha in enumerate(ws.iter_rows(min_row=cab_linha, max_row=cab_linha, values_only=True), 1):
                cab = [str(c).strip().lower() if c is not None else "" for c in linha]
            if not cab or col not in cab:
                log.warning("[%s] O arquivo anterior não tem a coluna de data '%s'; baixo todos os meses.", exp["name"], col)
                return None
            idx = cab.index(col) + 1
            meses, maximo, cache = {}, None, {}
            for (v,) in ws.iter_rows(min_row=cab_linha + 1, min_col=idx, max_col=idx, values_only=True):
                if v is None or v == "":
                    continue
                if v not in cache:
                    cache[v] = _para_data(v)
                d = cache[v]
                if d is None:
                    continue
                m = d.strftime("%Y-%m")
                meses[m] = meses.get(m, 0) + 1
                if maximo is None or d > maximo:
                    maximo = d
        finally:
            wb.close()
        res = (meses, maximo)
        _ESTADO["inventarios"] = {chave: res}      # guarda só o último (não acumula memória)
        return res
    except Exception as e:
        log.warning("[%s] Não consegui ler o arquivo anterior (%s); baixo todos os meses.", exp["name"], str(e)[:120])
        return None


def _meses_esperados() -> list:
    """Meses que a conferência do arquivo baixado espera: os que foram realmente pedidos ao dashboard."""
    ped = _ESTADO.get("pedido")
    if ped and ped.get("baixados"):
        return list(ped["baixados"])
    return [m[:7] for m in periodo_meses().values()]


def determinar_meses_para_baixar(exp: dict, pastas: dict, cfg: dict) -> dict:
    """
    Decide, mês a mês, se precisa baixar. Retorna {'mes_retrasado'|'mes_anterior'|'mes_atual': {'baixar': bool, 'motivo': str}}.
    Fechado = o arquivo publicado tem linhas do mês E a data máxima dele já chegou ao último dia do mês.
    Qualquer dúvida (sem arquivo, sem coluna de data, erro ao ler, --baixar-tudo) => baixa o mês.
    """
    papeis = _papeis_de_mes(exp)
    resultado = {f"mes_{p}": {"baixar": True, "motivo": ""} for p in papeis}

    def todos(motivo: str) -> dict:
        for p in papeis:
            resultado[f"mes_{p}"]["motivo"] = motivo
        return resultado

    if _ESTADO["baixar_tudo"]:
        return todos("Forçado (baixa todos os meses)")
    if exp.get("skip_closed_months", True) is False:
        return todos("skip_closed_months desligado no config")
    arquivo_cache, _info = verificar_cache_arquivo(exp, pastas, cfg)
    if not arquivo_cache:
        return todos("Sem arquivo anterior")
    inv = _inventario_do_arquivo(Path(arquivo_cache), exp)
    if inv is None or inv[1] is None:
        return todos("Arquivo anterior ilegível ou sem datas")
    meses, data_max = inv
    for p in papeis:
        m = _mes_do_papel(p)
        ultimo_dia = _fim_do_mes(dt.date.fromisoformat(m + "-01"))
        r = resultado[f"mes_{p}"]
        if meses.get(m, 0) > 0 and data_max >= ultimo_dia:
            r["baixar"] = False
            r["motivo"] = f"{m} FECHADO ({meses[m]} linhas; data máxima do arquivo {data_max})"
        elif meses.get(m, 0) == 0:
            r["motivo"] = f"{m} ainda não está no arquivo"
        else:
            r["motivo"] = f"{m} defasado (arquivo vai até {data_max}; o mês acaba em {ultimo_dia})"
    return resultado


def gravar_data_maxima(arquivo: Path, exp: dict, cfg: dict, data_max: dt.date) -> None:
    """
    Grava a data máxima de pagamento em um arquivo de controle (JSON).
    Arquivo: data_maxima.json na mesma pasta do arquivo de saída.
    """
    try:
        pasta = arquivo.parent
        arquivo_controle = pasta / "data_maxima.json"
        
        # Ler dados existentes ou criar novo
        dados = {}
        if arquivo_controle.exists():
            try:
                dados = json.loads(arquivo_controle.read_text(encoding="utf-8"))
            except:
                dados = {}
        
        # Adicionar/atualizar data máxima
        dados[exp["name"]] = {
            "data_maxima": data_max.isoformat(),
            "atualizado_em": dt.datetime.now().isoformat(),
            "arquivo": arquivo.name
        }
        
        # Gravar
        arquivo_controle.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")
        log.info("[%s] Data máxima gravada em: %s", exp["name"], arquivo_controle)
    except Exception as e:
        log.warning("[%s] Erro ao gravar data máxima: %s", exp["name"], str(e)[:100])


def limpar_antigos(pasta: Path, padrao: str, dias: int) -> None:
    if dias <= 0:
        return
    corte = time.time() - dias * 86400
    for f in pasta.glob(padrao):
        try:
            if f.stat().st_mtime < corte:
                f.unlink()
                log.debug("Removido por retenção: %s", f)
        except OSError:
            pass


# =============================================================================
# Cache Inteligente: Verifica se arquivo já baixado precisa ser atualizado
# =============================================================================
def verificar_cache_arquivo(exp: dict, pastas: dict, cfg: dict) -> tuple[Path | None, dict]:
    """
    Devolve o arquivo que o robô vai SUBSTITUIR (o 'destination' do export, ou o mesmo nome na pasta de teste).
    É dele que saem os meses já fechados. NÃO procura arquivos parecidos na pasta: a pasta oficial guarda arquivos
    antigos (ex.: Pagamentos_CC_3.xlsx) e usar um deles como "arquivo anterior" fazia o robô achar que faltavam meses.
    Retorna (arquivo, info): arquivo é None se ainda não existe.
    """
    info = {"arquivo": None, "motivo": ""}
    nome_fixo = PureWindowsPath(str(exp["destination"])).name
    alternativo = (cfg.get("paths") or {}).get("destino_alternativo")
    pasta = Path(alternativo) if alternativo else Path(exp["destination"]).parent
    arquivo = pasta / nome_fixo
    try:
        if arquivo.is_file() and arquivo.stat().st_size > 0:
            info["arquivo"] = str(arquivo)
            info["motivo"] = f"Arquivo encontrado: {arquivo.name}"
            log.info("[%s] Arquivo anterior (o que será atualizado): %s", exp["name"], arquivo)
            return arquivo, info
    except OSError as e:
        log.warning("[%s] Não consegui acessar %s: %s", exp["name"], arquivo, str(e)[:100])
    info["motivo"] = f"Ainda não existe {nome_fixo} em {pasta}"
    return None, info


# =============================================================================
# Orquestração
# =============================================================================
def iniciar_navegador(pw, cfg: dict, visivel: bool):
    b = cfg["browser"]
    kw = {"headless": False if visivel else b["headless"],
          "slow_mo": 150 if visivel else b["slow_mo_ms"]}
    if visivel:
        kw["args"] = ["--start-maximized"]
    if b.get("channel"):
        kw["channel"] = b["channel"]  # "msedge" / "chrome" = usa o navegador já instalado
    if b.get("proxy"):
        kw["proxy"] = {"server": b["proxy"]}
    return pw.chromium.launch(**kw)


def novo_contexto(navegador, cfg: dict, pastas: dict, visivel: bool = False):
    b = cfg["browser"]
    kw = {"accept_downloads": True,
          "locale": b.get("locale", "pt-BR"),
          "timezone_id": b.get("timezone", "America/Sao_Paulo"),
          "viewport": b.get("viewport") or {"width": 1920, "height": 1080}}
    if visivel:
        # janela visível: a página usa o tamanho real da janela (sem forçar 1920x1080)
        kw.pop("viewport")
        kw["no_viewport"] = True
    if pastas["sessao"].exists():
        try:
            return navegador.new_context(storage_state=str(pastas["sessao"]), **kw)
        except Exception as e:
            log.warning("Sessão salva inválida (%s); descartando", e)
            apagar_sessao(pastas)
    return navegador.new_context(**kw)


def executar_export(navegador, exp: dict, cfg: dict, pastas: dict, args) -> dict:
    inicio = dt.datetime.now()
    res = {"inicio": inicio.strftime("%Y-%m-%d %H:%M:%S"), "export": exp["name"],
           "destino": exp["destination"], "status": "FALHA"}
    tentativas = int(cfg["retry"]["attempts"])
    mes_base_fallback = _ESTADO.get("base_simulada")   # (--sim-date) ou, depois, o mês anterior quando o mês atual não existe no dashboard
    _ESTADO["baixar_tudo"] = bool(getattr(args, "baixar_tudo", False))
    _ESTADO["baixar_tudo_usuario"] = bool(getattr(args, "baixar_tudo", False))
    _ESTADO["pedido"] = None
    
    # Verificar cache antes de começar
    arquivo_cache, info_cache = verificar_cache_arquivo(exp, pastas, cfg)
    cache_disponivel = arquivo_cache and info_cache.get("arquivo")  # Tem arquivo mesmo que defasado
    
    if arquivo_cache and info_cache.get("pode_reutilizar"):
        log.info("[%s] ✅ Reutilizando arquivo anterior (%s)", exp["name"], info_cache["motivo"])
        res["status"] = "OK_CACHE"
        res["tentativas"] = 0
        res["arquivo_cache"] = str(arquivo_cache)
        res["informacao"] = info_cache
        return res

    # Se TODOS os meses que a consulta usa já estão fechados no arquivo publicado, não há nada a baixar
    # (nem abre o navegador). Mês fechado = tem linhas dele e a data máxima do arquivo chegou ao último dia do mês.
    definir_mes_base(None)
    if arquivo_cache and not _ESTADO["baixar_tudo"]:
        info_meses = determinar_meses_para_baixar(exp, pastas, cfg)
        if not any(v["baixar"] for v in info_meses.values()):
            log.info("[%s] ✅ Todos os meses que a consulta usa já estão fechados em %s. Nada a baixar.",
                     exp["name"], Path(arquivo_cache).name)
            res["status"] = "OK_CACHE"
            res["tentativas"] = 0
            res["arquivo_cache"] = str(arquivo_cache)
            res["informacao"] = {"motivo": "Todos os meses fechados - nenhum download necessário",
                                 "meses": {k[4:]: v["motivo"] for k, v in info_meses.items()}}
            return res

    for n in range(1, tentativas + 1):
        contexto = page = None
        try:
            log.info("[%s] Tentativa %d/%d", exp["name"], n, tentativas)
            definir_mes_base(mes_base_fallback)
            _ESTADO["baixar_tudo"] = bool(getattr(args, "baixar_tudo", False))
            _ESTADO["pedido"] = None
            # Com o mês-base deslocado (mês atual ainda inexistente), pode não sobrar nada a baixar: todos os meses
            # já estão fechados. Exportar assim seria exportar SEM mês selecionado: encerra aqui, sem abrir o dashboard.
            if arquivo_cache and not _ESTADO["baixar_tudo"] and mes_base_fallback is not None:
                info_meses = determinar_meses_para_baixar(exp, pastas, cfg)
                if not any(v["baixar"] for v in info_meses.values()):
                    log.info("[%s] ✅ Com os meses %s todos já estão fechados em %s. Nada a baixar.", exp["name"],
                             " e ".join(_mes_do_papel(p) for p in _papeis_de_mes(exp)), Path(arquivo_cache).name)
                    res["status"] = "OK_CACHE"
                    res["tentativas"] = n - 1
                    res["erro"] = ""
                    res["arquivo_cache"] = str(arquivo_cache)
                    res["informacao"] = {"motivo": "Mês atual ainda inexistente no dashboard; os meses anteriores já estão fechados",
                                         "meses": {k[4:]: v["motivo"] for k, v in info_meses.items()}}
                    break
            contexto = novo_contexto(navegador, cfg, pastas, visivel=bool(args.debug))
            page = contexto.new_page()
            page.set_default_timeout(cfg["browser"]["default_timeout_s"] * 1000)

            abrir_dashboard(page, contexto, exp, cfg, pastas)
            if args.check:
                if exp.get("export_steps"):
                    verificar_por_passos(page, exp, cfg)
                else:
                    abrir_menu_do_visual(page, exp, cfg["selectors"])
                    page.keyboard.press("Escape")
                res["status"] = "CHECK_OK"
            else:
                arquivo, corte = baixar_com_corte(page, contexto, exp, cfg, pastas, debug=args.debug)
                log.info("[%s] Validando o arquivo baixado...", exp["name"])
                res.update(validar_e_ajustar(arquivo, exp))
                if arquivo.suffix.lower() == ".csv":
                    conferir_tipos(arquivo, exp)
                if corte:
                    res.update({"data_corte": corte["data_corte"].isoformat(), "dias_uteis": corte["dias_uteis"],
                                "fim_mes_anterior": corte["ant_fim"].isoformat()})
                if arquivo.suffix.lower() == ".csv" and Path(exp["destination"]).suffix.lower() in (".xlsx", ".xlsm"):
                    log.info("[%s] Convertendo CSV para Excel (arquivos grandes levam cerca de 1 minuto)...", exp["name"])
                    arquivo = csv_para_xlsx(arquivo, exp)
                # Junta com o arquivo já publicado: meses fechados vêm dele, meses baixados o substituem
                arquivo = consolidar_com_cache(arquivo, arquivo_cache, exp, cfg, pastas)
                publicar(arquivo, exp, cfg)
                
                # Extrair e gravar data máxima de pagamento
                data_max = extrair_data_maxima_pagamento(arquivo, exp, cfg)
                if data_max:
                    gravar_data_maxima(arquivo, exp, cfg, data_max)
                    res["data_maxima_pagamento"] = data_max.isoformat()
                
                # (Sem mesclagem: o arquivo baixado já traz os dois meses e é ele que foi publicado.)
                
                res["status"] = "OK"
            salvar_sessao(contexto, pastas)
            res["tentativas"] = n
            res["erro"] = ""
            break
        except Exception as e:
            res["erro"] = f"{type(e).__name__}: {e}"[:500]
            res["tentativas"] = n
            log.error("[%s] Falha na tentativa %d: %s", exp["name"], n, res["erro"])
            log.debug(traceback.format_exc())
            
            # Pular captura de evidências para erro de "elemento não encontrado" (mês sem dados)
            # para não pesar no disco com logs repetidos
            é_elemento_não_encontrado = isinstance(e, EtapaError) and "Não encontrei o elemento" in str(e)
            # SÓ é "mês atual sem dados" se falhou o passo marcado com fallback_mes_anterior: true no config
            # (o de marcar o mês atual, DEPOIS de o mês anterior já estar marcado). Qualquer outro passo que
            # falhe (ex.: o seletor de mês) deixaria o dashboard sem filtro: exportar assim daria dado errado com status OK.
            passo_falho = getattr(e, "passo", None) or {}
            pode_fallback_mes_anterior = é_elemento_não_encontrado and bool(passo_falho.get("fallback_mes_anterior"))
            if not pode_fallback_mes_anterior:
                capturar_evidencias(page, exp["name"], n, pastas)
            
            if args.debug and page is not None and not pode_fallback_mes_anterior:
                log.info("Modo debug: abrindo o Playwright Inspector. Feche-o para continuar.")
                try:
                    page.pause()
                except Exception:
                    pass
            if isinstance(e, (ConfigError, CredencialInvalidaError, MesEncolheuError, FiltroTipoError)):
                log.error("[%s] Erro que não se resolve tentando de novo; abortando este export", exp["name"])
                break
            # Se for erro de elemento não encontrado (ex: mês sem dados no passo de marcar mês atual)
            # Tenta exportar apenas COM O MÊS ANTERIOR selecionado
            if pode_fallback_mes_anterior and mes_base_fallback is None:
                # O mês atual ainda não existe na lista do dashboard (sem dados). O arquivo publicado precisa
                # ter SEMPRE dois meses (o funil compara períodos iguais): refaz tudo usando o mês anterior
                # como "mês atual" (ex.: 2026-08 + 2026-09), pelo caminho normal (validar, converter, publicar).
                pedido = periodo_meses()
                mes_base_fallback = dt.date.fromisoformat(pedido["mes_anterior"])
                novo_ant = (mes_base_fallback - dt.timedelta(days=1)).replace(day=1)
                log.warning("[%s] O mês atual (%s) ainda não tem dados no dashboard. Refazendo com os dois meses "
                            "mais recentes: %s e %s.", exp["name"], pedido["mes_atual"][:7],
                            novo_ant.isoformat()[:7], mes_base_fallback.isoformat()[:7])
                continue   # sem espera: não foi falha de rede, só o mês que ainda não existe
            elif é_elemento_não_encontrado:
                nome_passo = (passo_falho.get("desc") or _descrever(passo_falho)) if passo_falho else "?"
                if pode_fallback_mes_anterior:
                    log.error("[%s] O passo do mês atual ('%s') falhou de novo; não repito a exportação parcial.", exp["name"], nome_passo)
                else:
                    log.error("[%s] O passo '%s' não foi encontrado e NÃO é o passo do mês atual: não exporto com filtro incompleto. "
                              "Corrija o nome desse passo no config (grave os cliques com o codegen).", exp["name"], nome_passo)
                break
            
            if isinstance(e, LoginError):
                apagar_sessao(pastas)
                # Se for LoginError na primeira tentativa, tenta fazer login automático
                if n == 1:
                    log.info("[%s] Sessão expirada. Tentando fazer login automático...", exp["name"])
                    try:
                        # Criar um contexto novo sem sessão para fazer login
                        ctx_login = novo_contexto(navegador, cfg, pastas, visivel=True)
                        pg_login = ctx_login.new_page()
                        pg_login.goto(exp.get("dashboard_url", "https://quicksight.aws.amazon.com/"), 
                                     wait_until="domcontentloaded", timeout=10000)
                        
                        # Esperar o usuário fazer login
                        log.info("[%s] Faça o login no navegador (inclusive SSO/MFA). Depois volte aqui e pressione Enter...", exp["name"])
                        input("")
                        
                        # Salvar a nova sessão
                        salvar_sessao(ctx_login, pastas)
                        ctx_login.close()
                        log.info("[%s] Sessão renovada. Tentando de novo...", exp["name"])
                    except Exception as err_login:
                        log.warning("[%s] Erro ao fazer login automático: %s", exp["name"], str(err_login)[:100])
            if n < tentativas:
                espera = int(cfg["retry"]["wait_s"]) * n  # espera crescente
                log.info("[%s] Aguardando %ds antes de tentar de novo", exp["name"], espera)
                time.sleep(espera)
        finally:
            if contexto is not None:
                try:
                    contexto.close()
                except Exception:
                    pass

    # Se ainda falhou (retry também não funcionou ou página foi fechada), usar como fallback
    if res["status"] == "FALHA" and cache_disponivel:
        try:
            arquivo_fallback = Path(info_cache.get("arquivo", ""))
            if arquivo_fallback.exists():
                log.warning("[%s] Todas as tentativas falharam. Usando cache como fallback...", exp["name"])
                res["status"] = "OK_CACHE_FALLBACK"
                res["tentativas"] = tentativas
                res["arquivo_cache"] = str(arquivo_fallback)
                res["informacao"] = f"Fallback - arquivo anterior: {arquivo_fallback.name}"
                res["erro"] = ""  # Considerar como sucesso parcial
                
                # Extrair e gravar data máxima também para fallback
                data_max = extrair_data_maxima_pagamento(arquivo_fallback, exp, cfg)
                if data_max:
                    gravar_data_maxima(arquivo_fallback, exp, cfg, data_max)
                    res["data_maxima_pagamento"] = data_max.isoformat()
                    
                    # Verificar se ambos os meses estão fechados (agora que temos a data)
                    try:
                        hoje = dt.date.today()
                        ultimo_dia_mes_atual = (hoje.replace(day=1) + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
                        
                        if mes_esta_fechado(data_max) and hoje > ultimo_dia_mes_atual:
                            log.info("[%s] ✅ Ambos os meses estão fechados (última data: %s). Próximas execuções pulam download.", 
                                    exp["name"], data_max)
                    except Exception as e:
                        log.debug("[%s] Erro ao verificar meses fechados: %s", exp["name"], str(e)[:100])
        except Exception as e:
            log.warning("[%s] Erro ao ativar fallback: %s", exp["name"], str(e)[:100])
    
    fim = dt.datetime.now()
    res["fim"] = fim.strftime("%Y-%m-%d %H:%M:%S")
    res["duracao_s"] = round((fim - inicio).total_seconds())
    registrar_historico(pastas["logs"], res)
    return res


def _url_autenticada(url: str) -> bool:
    """True se a URL é do QuickSight e NÃO é tela de login/autenticação."""
    u = (url or "").lower()
    return "quicksight.aws.amazon.com" in u and not any(x in u for x in ("/auth", "signin", "login", "okta"))


def _sessao_confirmada(contexto, cfg: dict) -> bool:
    """
    Confirma o login de forma definitiva: abre uma aba NOVA no mesmo contexto (que compartilha
    os cookies do login, em qualquer aba que ele tenha terminado) e vê se o QuickSight abre
    sem redirecionar para a autenticação. Funciona mesmo que o usuário tenha fechado as abas.
    """
    teste = None
    try:
        teste = contexto.new_page()
        teste.goto(cfg["quicksight"]["start_url"], wait_until="domcontentloaded", timeout=30000)
        time.sleep(4)  # dá tempo de um eventual redirecionamento acontecer
        url = teste.url
        log.info("Teste de sessão: %s", url[:120])
        return _url_autenticada(url)
    except Exception as e:
        log.warning("Não consegui testar a sessão: %s", str(e)[:150])
        return False
    finally:
        if teste is not None:
            try:
                teste.close()
            except Exception:
                pass


def login_manual(navegador, cfg: dict, pastas: dict) -> int:
    """Abre o QuickSight numa janela VISÍVEL, espera o login e só salva a sessão se ele foi confirmado."""
    if pastas["sessao"].exists():
        apagar_sessao(pastas)  # parte de uma sessão limpa (a antiga está expirada)
    contexto = novo_contexto(navegador, cfg, pastas, visivel=True)
    page = contexto.new_page()
    page.goto(cfg["quicksight"]["start_url"], wait_until="domcontentloaded")
    try:
        page.bring_to_front()
    except Exception:
        pass
    print("\nFaça o login na janela do navegador que abriu (inclusive SSO/MFA).")
    print("Espere o dashboard do QuickSight carregar. NÃO feche a janela: volte aqui e pressione Enter.")
    for tentativa in range(1, 4):
        input("Quando o QuickSight estiver aberto, pressione Enter... ")
        if _sessao_confirmada(contexto, cfg):
            salvar_sessao(contexto, pastas)
            try:
                contexto.close()
            except Exception:
                pass
            log.info("Sessão salva em %s", pastas["sessao"])
            return 0
        abertas = []
        try:
            abertas = [pg.url[:100] for pg in contexto.pages if not pg.is_closed()]
        except Exception:
            pass
        print(f"O QuickSight ainda pediu login ({tentativa}/3). Abas abertas: {abertas or 'nenhuma'}")
    try:
        contexto.close()
    except Exception:
        pass
    log.error("Login não confirmado; sessão NÃO foi salva.")
    return 1


def fechar_navegador(navegador) -> None:
    """Fecha sem deixar um erro de 'Connection closed' esconder o resultado da execução."""
    try:
        navegador.close()
    except Exception as e:
        log.debug("Ao fechar o navegador: %s", e)


def limpar_staging(pasta: Path, dias: int = 7) -> None:
    limpar_antigos(pasta, "*", dias)


# Status que NÃO são falha: OK (baixou e publicou), CHECK_OK (modo --check) e OK_CACHE (todos os meses já estão
# fechados no arquivo publicado: nada a baixar é o dia a dia normal). OK_CACHE_FALLBACK é falha: o download
# não funcionou e o arquivo antigo foi mantido (a consulta SQL ficaria com dado velho).
STATUS_SUCESSO = ("OK", "CHECK_OK", "OK_CACHE")


def separar_falhas(resultados: list) -> list:
    return [r for r in resultados if r["status"] not in STATUS_SUCESSO]


def execucao_sem_operador(args) -> bool:
    """True quando ninguém pode digitar nada (Agendador de Tarefas, serviço) ou quando foi pedido com --nao-interativo."""
    if getattr(args, "nao_interativo", False):
        return True
    try:
        return not sys.stdin.isatty()
    except Exception:
        return True


def tratar_sessao_expirada(args, cfg: dict):
    """
    Sessão expirada. Com operador na frente: devolve None (o fluxo abre a janela para o login).
    Sem operador: NÃO abre janela nem espera Enter; avisa e encerra com código 4.
    """
    if not execucao_sem_operador(args):
        return None
    msg = ("A sessão do QuickSight expirou e esta execução não tem ninguém na frente para refazer o login. "
           "Nada foi baixado nem publicado. Faça o login: abra o PowerShell na pasta do robô e rode  "
           "python qs_download.py --login-manual")
    log.error(msg)
    enviar_alerta(cfg, "[Robô QuickSight] " + msg)
    return 4


def main() -> int:
    ap = argparse.ArgumentParser(description="Robô de download do QuickSight para a pasta de rede")
    ap.add_argument("--config", default=str(BASE_DIR / "config.yaml"))
    ap.add_argument("--only", nargs="*", help="executa só os exports com estes nomes")
    ap.add_argument("--check", action="store_true", help="testa login/dashboard/menu sem baixar")
    ap.add_argument("--nao-interativo", action="store_true",
                    help="nunca abre janela nem espera Enter (use no Agendador de Tarefas): sessão expirada = alerta + código 4")
    ap.add_argument("--baixar-tudo", action="store_true", help="ignora meses fechados: baixa todos os meses e substitui o arquivo")
    ap.add_argument("--debug", action="store_true", help="navegador visível e Inspector no erro")
    ap.add_argument("--set-credentials", action="store_true", help="grava credenciais no Cofre do Windows")
    ap.add_argument("--login-manual", action="store_true", help="login manual e salva a sessão")
    ap.add_argument("--sim-date", type=str, default=None, help="simular período (ex: --sim-date 2026-09-01 para set/ago)")
    args = ap.parse_args()

    # Suporte a simulação de período (para demo/testes)
    if args.sim_date:
        try:
            data_sim = dt.datetime.strptime(args.sim_date, "%Y-%m-%d").date()
            _ESTADO["base_simulada"] = data_sim.replace(day=1)   # usado por executar_export como mês-base
            log.info("SIMULAÇÃO: usando mês base de %s (mês atual: %s, anterior: %s)", 
                     args.sim_date, data_sim.replace(day=1), (data_sim.replace(day=1) - dt.timedelta(days=1)).replace(day=1))
        except ValueError:
            print(f"ERRO: --sim-date deve estar no formato YYYY-MM-DD (ex: 2026-09-01)")
            return 2

    try:
        cfg = carregar_config(Path(args.config))
        pastas = preparar_pastas(cfg)
    except ConfigError as e:
        print(f"ERRO de configuração: {e}")
        return 2
    configurar_log(pastas["logs"], args.debug)

    if args.set_credentials:
        return definir_credenciais(cfg)

    exports = cfg["exports"]
    if args.only:
        exports = [e for e in exports if e["name"] in set(args.only)]
        if not exports:
            log.error("Nenhum export com os nomes %s", args.only)
            return 2

    log.info("=" * 70)
    log.info("Início da execução | exports: %s%s", [e["name"] for e in exports],
             " | MODO CHECK" if args.check else "")
    try:
        with TravaExecucao(pastas["staging"] / ".execucao.lock"):
            limpar_staging(pastas["staging"])
            with sync_playwright() as pw:
                navegador = iniciar_navegador(pw, cfg, visivel=args.debug or args.login_manual)
                try:
                    if args.login_manual:
                        return login_manual(navegador, cfg, pastas)

                    # Validar sessão antes de começar (evita descobrir no meio da execução)
                    if not args.check and not validar_sessao(navegador, cfg, pastas, debug=args.debug):
                        codigo_saida = tratar_sessao_expirada(args, cfg)
                        if codigo_saida is not None:
                            return codigo_saida
                        log.warning("Sessão expirada. Abrindo o navegador VISÍVEL para você fazer login...")
                        # O navegador atual está invisível (headless): fecha e abre outro com janela
                        fechar_navegador(navegador)
                        navegador = iniciar_navegador(pw, cfg, visivel=True)
                        if login_manual(navegador, cfg, pastas) != 0:
                            return 1
                        # Login feito: volta ao modo normal e continua o download nesta mesma execução
                        fechar_navegador(navegador)
                        navegador = iniciar_navegador(pw, cfg, visivel=args.debug)

                    resultados = [executar_export(navegador, e, cfg, pastas, args) for e in exports]
                finally:
                    fechar_navegador(navegador)
    except ConfigError as e:
        log.error("Erro de configuração: %s", e)
        enviar_alerta(cfg, f"[Robô QuickSight] Erro de configuração: {e}")
        return 2
    except RoboError as e:
        log.error("%s", e)
        enviar_alerta(cfg, f"[Robô QuickSight] {e}")
        return 1
    except Exception as e:
        log.critical("Erro inesperado: %s\n%s", e, traceback.format_exc())
        enviar_alerta(cfg, f"[Robô QuickSight] Erro inesperado: {e}")
        return 3

    falhas = separar_falhas(resultados)
    for r in resultados:
        log.info("Resumo | %-20s | %-8s | tentativas=%s | linhas=%s | %s",
                 r["export"], r["status"], r.get("tentativas"), r.get("linhas", ""), r.get("erro", ""))
    if falhas:
        enviar_alerta(cfg, "[Robô QuickSight] Falha em: " +
                      "; ".join(f"{r['export']} ({r.get('erro', '')[:150]})" for r in falhas))
        log.info("Fim da execução COM FALHAS")
        return 1
    log.info("Fim da execução: tudo OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())