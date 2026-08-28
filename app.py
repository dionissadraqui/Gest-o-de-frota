"""
LUFT LOGISTICS — Sistema de Controle de Motoristas
Versão Streamlit — arquivo único (app.py)

Coloque na mesma pasta:
  • luft.png       — imagem de fundo da splash
  • credentials.json — credenciais Google Service Account

Execute com:
  streamlit run app.py
"""
#  python -m streamlit run app.py


import json
import base64
from datetime import datetime
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

import gspread
from google.oauth2.service_account import Credentials
import google.auth.transport.requests

# ─── Configurações ────────────────────────────────────────────────────────────
SHEET_ID   = "1PP0cUAIpQqv7zB0JCYsGxGW9zHsrRw182iRrZ52q8Kg"
SHEET_NAME = "motoristas_luft"
SHEET_NAME_ORG = "organograma_luft"
# Credenciais: arquivo credentials.json na mesma pasta do app.py
BASE_DIR         = Path(__file__).parent
CREDENTIALS_PATH = BASE_DIR / "gestao-operacional-499623.json"
LOGO_PATH        = BASE_DIR / "luft.png"

MESES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

COLUNAS = [
    "cpf", "nome", "filial", "telefone", "email", "foto",
    "reciclagem", "simulador",
    "excesso", "multas", "acidentes",
    "obsAcidente", "obsMultas", "obsGerais", "obsReciclagem", "obsSimulador",
    "cnh", "validadeCnh", "admissao",
]
for _mes in MESES:
    for _s in range(1, 5):
        COLUNAS.append(f"dss_{_mes}_{_s}")
COLUNAS += [
    "exame_periodico", "exame_toxicologico", "pontuação_cnh",
    "vencimento_cnh_mopp", "entrega_de_uniforme",
    "telefone_corporativo", "numero_linha", "modelo", "imei",
    "reciclagem_data", "reciclagem_validade_meses",
    "simulador_data", "simulador_validade_meses",
    "exame_periodico_validade_meses", "exame_toxicologico_validade_meses",
    "gestime", "obsGestime", "gestime_data", "gestime_validade_meses",
    "afastado", "obsAfastado",
    "desligado", "obsDesligamento",
]

COLUNAS_ORG = ["id", "tipo", "setor_ordem", "setor_titulo", "setor_icone", "pessoa_ordem", "nome", "cargo"]

ORGANOGRAMA_PADRAO = {
    "supervisor": {"nome": "RAFAELA SILVA", "cargo": "SUPERVISORA"},
    "setores": [
        {"titulo": ["CONTROLE DE", "JORNADA"], "icone": "clock", "pessoas": [
            ["Cristina Calixto", "Assist. Adm."],
            ["Giselle Freitas", "Assist. Adm."],
            ["Julia Haro", "Assist. Adm."],
            ["Dionis Sadraqui", "Assist. Adm."],
            ["João Eduardo", "Jovem Aprendiz"],
        ]},
        {"titulo": ["GESTÃO DE", "MOTORISTAS"], "icone": "wheel", "pessoas": [
            ["Issac Fernandes", "Analista Adm."],
            ["Ana", "Assist. Adm."],
        ]},
        {"titulo": ["COMPROVANTES", "DE ENTREGA"], "icone": "clipboard", "pessoas": [
            ["Fabio de Almeida", "Assist. Adm."],
            ["Jose Souza", "Analista Adm."],
            ["Geovanna Vitoria", "Analista Adm."],
        ]},
        {"titulo": ["ACERTO"], "icone": "calculator", "pessoas": [
            ["Geisa", "Analista Adm."],
        ]},
        {"titulo": ["MOTORISTAS"], "icone": "wheel", "pessoas": [
            ["", ""],
        ]},
    ],
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ─── Google Sheets helpers ────────────────────────────────────────────────────
def get_sheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=len(COLUNAS) + 5)
        ws.append_row(COLUNAS)
    return ws

def ler_todos_motoristas():
    ws = get_sheet()
    records = ws.get_all_records(default_blank="")
    motoristas = []
    for i, row in enumerate(records):
        cpf_bruto  = str(row.get("cpf", "")).strip()
        nome_bruto = str(row.get("nome", "")).strip()
        if not cpf_bruto and not nome_bruto:
            continue
        cpf_efetivo = cpf_bruto or f"SEMCPF-{i+1:04d}"
        dss_anual = {}
        for mes in MESES:
            semanas = []
            for s in range(1, 5):
                val = row.get(f"dss_{mes}_{s}", "")
                semanas.append(bool(val) and val not in ("", "0", 0, False))
            dss_anual[mes] = semanas
        motoristas.append({
            "cpf":           cpf_efetivo,
            "nome":          nome_bruto,
            "filial":        str(row.get("filial", "")).strip(),
            "telefone":      str(row.get("telefone", "")).strip(),
            "email":         str(row.get("email", "")).strip(),
            "foto":          str(row.get("foto", "")).strip(),
            "reciclagem":    str(row.get("reciclagem", "PENDENTE")).strip() or "PENDENTE",
            "simulador":     str(row.get("simulador", "PENDENTE")).strip() or "PENDENTE",
            "excesso":       max(0, int(row.get("excesso", 0) or 0)),
            "multas":        max(0, int(row.get("multas", 0) or 0)),
            "acidentes":     max(0, int(row.get("acidentes", 0) or 0)),
            "obsAcidente":   str(row.get("obsAcidente", "")).strip(),
            "obsMultas":     str(row.get("obsMultas", "")).strip(),
            "obsGerais":     str(row.get("obsGerais", "")).strip(),
            "obsReciclagem": str(row.get("obsReciclagem", "")).strip(),
            "obsSimulador":  str(row.get("obsSimulador", "")).strip(),
            "cnh":           str(row.get("cnh", "")).strip(),
            "validadeCnh":   str(row.get("validadeCnh", "")).strip(),
            "admissao":      str(row.get("admissao", "")).strip(),
            "examePeriodico":     str(row.get("exame_periodico", "")).strip(),
            "exameToxicologico":  str(row.get("exame_toxicologico", "")).strip(),
            "pontuacaoCnh":       max(0, int(row.get("pontuação_cnh", 0) or 0)),
            "vencimentoCnhMopp":  str(row.get("vencimento_cnh_mopp", "")).strip(),
            "entregaUniforme":    str(row.get("entrega_de_uniforme", "PENDENTE")).strip() or "PENDENTE",
            "telefoneCorporativo": str(row.get("telefone_corporativo", "NÃO")).strip() or "NÃO",
            "numeroLinha":         str(row.get("numero_linha", "")).strip(),
            "modelo":              str(row.get("modelo", "")).strip(),
            "imei":                str(row.get("imei", "")).strip(),
            "reciclagemData":       str(row.get("reciclagem_data", "")).strip(),
            "reciclagemValidadeMeses": max(0, int(row.get("reciclagem_validade_meses", 0) or 0)),
            "simuladorData":        str(row.get("simulador_data", "")).strip(),
            "simuladorValidadeMeses": max(0, int(row.get("simulador_validade_meses", 0) or 0)),
            "examePeriodicoValidadeMeses": max(0, int(row.get("exame_periodico_validade_meses", 0) or 0)),
            "exameToxicologicoValidadeMeses": max(0, int(row.get("exame_toxicologico_validade_meses", 0) or 0)),
            "gestime":        str(row.get("gestime", "PENDENTE")).strip() or "PENDENTE",
            "obsGestime":     str(row.get("obsGestime", "")).strip(),
            "gestimeData":    str(row.get("gestime_data", "")).strip(),
            "gestimeValidadeMeses": max(0, int(row.get("gestime_validade_meses", 0) or 0)),
            "afastado":       str(row.get("afastado", "NÃO")).strip() or "NÃO",
            "obsAfastado":    str(row.get("obsAfastado", "")).strip(),
            "desligado":      str(row.get("desligado", "NÃO")).strip() or "NÃO",
            "obsDesligamento": str(row.get("obsDesligamento", "")).strip(),
            "dssAnual":      dss_anual,
        })
    return motoristas

def salvar_todos_motoristas(lista):
    ws = get_sheet()
    all_rows = []
    for m in lista:
        row_data = [
            m.get("cpf", ""), m.get("nome", ""), m.get("filial", ""),
            m.get("telefone", ""), m.get("email", ""), m.get("foto", ""),
            m.get("reciclagem", "PENDENTE"), m.get("simulador", "PENDENTE"),
            m.get("excesso", 0), m.get("multas", 0), m.get("acidentes", 0),
            m.get("obsAcidente", ""), m.get("obsMultas", ""), m.get("obsGerais", ""),
            m.get("obsReciclagem", ""), m.get("obsSimulador", ""),
            m.get("cnh", ""), m.get("validadeCnh", ""), m.get("admissao", ""),
        ]
        dss = m.get("dssAnual", {})
        for mes in MESES:
            semanas = dss.get(mes, [False] * 4)
            for s in range(4):
                row_data.append(1 if (len(semanas) > s and semanas[s]) else 0)
        row_data += [
            m.get("examePeriodico", ""), m.get("exameToxicologico", ""),
            m.get("pontuacaoCnh", 0), m.get("vencimentoCnhMopp", ""),
            m.get("entregaUniforme", "PENDENTE"),
            m.get("telefoneCorporativo", "NÃO"), m.get("numeroLinha", ""),
            m.get("modelo", ""), m.get("imei", ""),
            m.get("reciclagemData", ""), m.get("reciclagemValidadeMeses", 0),
            m.get("simuladorData", ""), m.get("simuladorValidadeMeses", 0),
            m.get("examePeriodicoValidadeMeses", 0), m.get("exameToxicologicoValidadeMeses", 0),
            m.get("gestime", "PENDENTE"), m.get("obsGestime", ""),
            m.get("gestimeData", ""), m.get("gestimeValidadeMeses", 0),
            m.get("afastado", "NÃO"), m.get("obsAfastado", ""),
            m.get("desligado", "NÃO"), m.get("obsDesligamento", ""),
        ]
        all_rows.append(row_data)
    existing = ws.get_all_values()
    if len(existing) > 1:
        ws.delete_rows(2, len(existing))
    if all_rows:
        ws.append_rows(all_rows, value_input_option="USER_ENTERED")


def get_sheet_org():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(SHEET_NAME_ORG)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME_ORG, rows=200, cols=len(COLUNAS_ORG) + 5)
        ws.append_row(COLUNAS_ORG)
    return ws

def ler_organograma():
    ws = get_sheet_org()
    records = ws.get_all_records(default_blank="")
    if not records:
        return ORGANOGRAMA_PADRAO
    supervisor = {"nome": "", "cargo": ""}
    setores_map = {}
    for row in records:
        tipo = str(row.get("tipo", "")).strip()
        if tipo == "supervisor":
            supervisor = {
                "nome": str(row.get("nome", "")).strip(),
                "cargo": str(row.get("cargo", "")).strip(),
            }
            continue
        so = int(row.get("setor_ordem", 0) or 0)
        if so not in setores_map:
            titulo_bruto = str(row.get("setor_titulo", "")).strip()
            titulo = titulo_bruto.split("|") if titulo_bruto else [""]
            setores_map[so] = {
                "titulo": titulo,
                "icone": str(row.get("setor_icone", "")).strip() or "clipboard",
                "pessoas": [],
            }
        po = int(row.get("pessoa_ordem", 0) or 0)
        setores_map[so]["pessoas"].append(
            (po, str(row.get("nome", "")).strip(), str(row.get("cargo", "")).strip())
        )
    setores = []
    for so in sorted(setores_map.keys()):
        s = setores_map[so]
        s["pessoas"] = [[n, c] for _, n, c in sorted(s["pessoas"], key=lambda t: t[0])]
        setores.append(s)
    if not setores:
        return ORGANOGRAMA_PADRAO
    return {"supervisor": supervisor, "setores": setores}


# ─── Gera access token OAuth2 para o JS usar ─────────────────────────────────
def get_access_token():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    creds.refresh(google.auth.transport.requests.Request())
    return creds.token

# ─── Logo em base64 ───────────────────────────────────────────────────────────
def logo_b64():
    if LOGO_PATH.exists():
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""


# ─── Streamlit page config ────────────────────────────────────────────────────
st.set_page_config(
    page_title="LUFT Controle de Motoristas",
    page_icon=BASE_DIR / "luft1.png",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Oculta elementos padrão do Streamlit (menu, footer, padding)
st.markdown("""
<style>
  #MainMenu, footer, header { visibility: hidden; }
  .block-container { padding: 0 !important; max-width: 100% !important; }
  [data-testid="stAppViewContainer"] { padding: 0 !important; }
  [data-testid="stVerticalBlock"] { gap: 0 !important; }
  iframe { height: 100vh !important; min-height: 100vh !important; }
  [data-testid="stIFrame"] { height: 100vh !important; }
</style>
""", unsafe_allow_html=True)

# ─── HTML completo da aplicação ───────────────────────────────────────────────
_LOGO_B64 = logo_b64()
_LOGO_CSS  = (
    f"background:url('data:image/png;base64,{_LOGO_B64}') center center/cover no-repeat;"
    if _LOGO_B64 else
    "background:linear-gradient(135deg,#0a1440 0%,#1a3a6b 100%);"
)
_ACCESS_TOKEN = get_access_token()

# Credenciais de login (usuário/senha) — vêm do secrets.toml,
# nunca ficam escritas no código-fonte do app.py
_LOGIN_USERS_RAW = st.secrets.get("auth", {}).get("usuarios", [])
CREDENCIAIS_LOGIN = [
    {"usuario": u["usuario"], "senha": u["senha"], "nome": u["nome"]}
    for u in _LOGIN_USERS_RAW
]

HTML = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>LUFT Logistics — Controle de Motoristas</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}

/* ── TELA SPLASH ── */
#splash-screen{{
  position:fixed;top:0;left:0;width:100vw;height:100vh;
  z-index:999999;
  display:flex;align-items:center;justify-content:center;
  overflow:hidden;
}}
#splash-bg{{
  position:absolute;top:0;left:0;width:100%;height:100%;
  {_LOGO_CSS}
  filter:brightness(1);
}}
.splash-card{{
  position:relative;z-index:2;
  background:rgba(10,20,60,0.82);
  border:1.5px solid rgba(59,125,216,0.45);
  border-radius:16px;
  padding:40px 48px;
  text-align:center;
  backdrop-filter:blur(12px);
  box-shadow:0 24px 64px rgba(0,0,0,0.6);
  min-width:360px;
  max-width:480px;
}}
.splash-logo-txt{{font-size:36px;font-weight:900;color:#ffffff;letter-spacing:-1px;margin-bottom:4px}}
.splash-logo-txt span{{color:#22cc88}}
.splash-sub{{font-size:11px;letter-spacing:3px;color:#8ab4d8;text-transform:uppercase;margin-bottom:32px}}
.splash-label{{font-size:12px;color:#8ab4d8;letter-spacing:1px;text-transform:uppercase;font-weight:700;margin-bottom:12px}}
.splash-drop-area{{border:2px dashed rgba(59,125,216,0.5);border-radius:10px;padding:28px 20px;cursor:pointer;transition:all .2s;background:rgba(59,125,216,0.07);margin-bottom:14px}}
.splash-drop-area:hover,.splash-drop-area.drag-over{{border-color:#3b7dd8;background:rgba(59,125,216,0.18)}}
.splash-drop-icon{{font-size:32px;color:#3b7dd8;margin-bottom:8px}}
.splash-drop-txt{{font-size:13px;color:#a0bcd8;line-height:1.5}}
.splash-drop-txt strong{{color:#ffffff}}
#splashFileInput{{display:none}}
.splash-btn-escolher{{display:inline-block;margin-top:10px;background:#3b7dd8;color:#fff;padding:8px 22px;border-radius:6px;font-size:12px;font-weight:700;letter-spacing:1px;cursor:pointer;border:none;text-transform:uppercase;transition:background .2s}}
.splash-btn-escolher:hover{{background:#2563b0}}
.splash-status{{font-size:12px;color:#22cc88;margin-top:10px;min-height:18px;font-weight:600;letter-spacing:.5px}}
.splash-status.erro{{color:#ff4444}}
.splash-progress{{display:none;width:100%;height:4px;background:rgba(255,255,255,0.1);border-radius:2px;margin-top:12px;overflow:hidden}}
.splash-progress-bar{{height:100%;width:0%;background:#3b7dd8;border-radius:2px;transition:width .3s;animation:progressAnim .8s ease-in-out infinite alternate}}
@keyframes progressAnim{{0%{{opacity:.6}}100%{{opacity:1}}}}

/* ── TEMA GLOBAL ── */
.db{{background:#f0f4fa;color:#1a2a44;font-family:'Segoe UI',sans-serif;padding:0;font-size:14px}}

/* ── TOP BAR ── */
.top-bar{{background:#ffffff;border-bottom:2px solid #dde6f4;padding:10px 16px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(20,50,120,0.07);position:relative;}}
.brand{{display:flex;align-items:center;gap:10px}}
.brand-logo{{background:#f0f4fa;border:1.5px solid #c4d0e4;border-radius:6px;padding:6px 14px;display:flex;align-items:center;gap:8px}}
.dot-anim{{
  width:10px;height:10px;border-radius:50%;background:#e53e3e;
  flex-shrink:0;position:relative;
  box-shadow:0 0 4px #e53e3e,0 0 8px #e53e3e;
  animation:dotBlink 1.4s ease-in-out infinite;
}}
.dot-anim::after{{
  content:'';position:absolute;
  top:50%;left:50%;
  width:10px;height:10px;
  border-radius:50%;
  background:transparent;
  border:2px solid #e53e3e;
  transform:translate(-50%,-50%) scale(1);
  animation:sonarRing 1.4s ease-out infinite;
}}
@keyframes dotBlink{{
  0%,100%{{opacity:1;box-shadow:0 0 4px #e53e3e,0 0 8px #e53e3e;}}
  50%{{opacity:0.4;box-shadow:0 0 2px #e53e3e;}}
}}
@keyframes sonarRing{{
  0%{{transform:translate(-50%,-50%) scale(1);opacity:0.9;border-color:#e53e3e;}}
  100%{{transform:translate(-50%,-50%) scale(3.5);opacity:0;border-color:rgba(229,62,62,0);}}
}}
.brand-name{{font-size:13px;font-weight:700;letter-spacing:2px;color:#1a3a6b;text-transform:uppercase}}
.brand-sub{{font-size:11px;color:#1a7a4a;letter-spacing:1px}}
.luft-name{{font-size:26px;font-weight:900;color:#1a3a6b;letter-spacing:-1px}}
.luft-name span{{color:#1a7a4a}}
.pct-box{{background:#f0f4fa;border:1.5px solid #c4d0e4;border-radius:6px;padding:6px 14px;text-align:right}}
.pct-lbl{{font-size:11px;color:#5a6e8a;letter-spacing:1px;text-transform:uppercase;font-weight:600}}
.pct-val{{font-size:30px;font-weight:900;color:#16a34a}}

/* ── CONTEÚDO ── */
.content{{padding:12px 14px}}

/* ── KPIs ── */
.kpi-row{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:16px}}
.kpi{{border-radius:10px;padding:18px 20px;border:2px solid}}
.kpi.red{{background:#fff5f5;border-color:#ff4444;box-shadow:0 0 10px rgba(255,68,68,0.35),inset 0 0 6px rgba(255,68,68,0.06)}}
.kpi.green{{background:#f0fef4;border-color:#22cc88;box-shadow:0 0 10px rgba(34,204,136,0.35),inset 0 0 6px rgba(34,204,136,0.06)}}
.kpi.amber{{background:#fffbeb;border-color:#ffaa00;box-shadow:0 0 10px rgba(255,170,0,0.35),inset 0 0 6px rgba(255,170,0,0.06)}}
.kpi.blue{{background:#f0f6ff;border-color:#3b7dd8;box-shadow:0 0 10px rgba(59,125,216,0.35),inset 0 0 6px rgba(59,125,216,0.06)}}
.kpi.teal{{background:#f0fbfd;border-color:#0e9cc0;box-shadow:0 0 10px rgba(14,156,192,0.35),inset 0 0 6px rgba(14,156,192,0.06)}}
.kpi.purple{{background:#f5f0ff;border-color:#7c3aed;box-shadow:0 0 10px rgba(124,58,237,0.35),inset 0 0 6px rgba(124,58,237,0.06)}}
.kpi.indigo{{background:#eef2ff;border-color:#4f46e5;box-shadow:0 0 10px rgba(79,70,229,0.35),inset 0 0 6px rgba(79,70,229,0.06)}}
.kpi.gray{{background:repeating-linear-gradient(45deg,#eef1f5,#eef1f5 10px,#e0e4ea 10px,#e0e4ea 20px);border-color:#94a3b8;box-shadow:0 0 10px rgba(100,116,139,0.30),inset 0 0 6px rgba(100,116,139,0.08)}}
.kpi-lbl{{font-size:14px;letter-spacing:1.5px;text-transform:uppercase;color:#5a6e8a;margin-bottom:6px;font-weight:700}}
.kpi-val{{font-size:52px;font-weight:900;line-height:1}}
.kpi.red .kpi-val{{color:#dc2626}}
.kpi.green .kpi-val{{color:#16a34a}}
.kpi.amber .kpi-val{{color:#d97706}}
.kpi.blue .kpi-val{{color:#1a4fa0}}
.kpi.teal .kpi-val{{color:#0a7a9a}}
.kpi.purple .kpi-val{{color:#6d28d9}}
.kpi.indigo .kpi-val{{color:#4338ca}}
.kpi.gray .kpi-val{{color:#475569}}
.kpi-sub{{font-size:14px;color:#3b7dd8;margin-top:8px;text-transform:uppercase;letter-spacing:1px;font-weight:600}}

/* ── PAINEL / SEÇÕES ── */
.sec-title{{font-size:14px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#1a4fa0;margin-bottom:8px;display:flex;align-items:center;gap:6px}}
.sec-title::before{{content:'';display:inline-block;width:3px;height:10px;background:#3b7dd8;border-radius:2px}}
.panel{{background:#ffffff;border:1.5px solid #dde6f4;border-radius:10px;padding:12px;margin-bottom:12px;box-shadow:0 2px 8px rgba(20,50,120,0.06)}}

/* ── GRID FILIAIS ── */
.filial-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
.fc{{background:#f8fafd;border:1.5px solid #c4d0e4;border-radius:10px;padding:14px;display:flex;flex-direction:column;justify-content:space-between}}
.fc-name{{font-size:20px;font-weight:800;color:#1a3a6b;margin-bottom:6px}}
.fc-count{{font-size:48px;font-weight:900;line-height:1;margin-bottom:10px;color:#1a4fa0}}
.situation-bars{{display:flex;flex-direction:column;gap:6px}}
.sbar{{display:flex;align-items:center;gap:8px}}
.sbar-lbl{{font-size:13px;color:#5a6e8a;width:66px;flex-shrink:0;text-transform:uppercase;letter-spacing:.5px;font-weight:700}}
.sbar-track{{flex:1;height:7px;background:#e0e8f0;border-radius:3px;overflow:hidden}}
.sbar-fill{{height:100%;border-radius:3px;transition:width .3s}}
.sbar-cnt{{font-size:15px;font-weight:700;width:28px;text-align:right;flex-shrink:0}}
.sbar.ok .sbar-fill{{background:#16a34a}}.sbar.ok .sbar-cnt{{color:#16a34a}}
.sbar.neg .sbar-fill{{background:#dc2626}}.sbar.neg .sbar-cnt{{color:#dc2626}}
.sbar.pend .sbar-fill{{background:#d97706}}.sbar.pend .sbar-cnt{{color:#d97706}}

/* ── GRÁFICOS ── */
.chart-wrap{{position:relative;width:100%}}
.leg{{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:8px}}
.leg-item{{display:flex;align-items:center;gap:4px;font-size:14px;color:#5a6e8a;font-weight:600}}
.leg-sq{{width:9px;height:9px;border-radius:1px;flex-shrink:0}}
.btn-zoom{{background:rgba(59,125,216,0.07);color:#3b7dd8;border:1px solid rgba(59,125,216,0.2);border-radius:4px;padding:8px;font-size:13px;font-weight:700;cursor:pointer;text-transform:uppercase;text-align:center;margin-top:8px;display:flex;align-items:center;justify-content:center;gap:4px}}
.btn-zoom:hover{{background:#3b7dd8;color:#fff}}

/* ── TOAST ── */
.toast{{position:fixed;bottom:24px;right:24px;background:#ffffff;border:1.5px solid #c4d0e4;color:#1a2a44;padding:12px 20px;border-radius:8px;font-size:13px;z-index:99999;display:none;gap:10px;align-items:center;box-shadow:0 8px 24px rgba(20,50,120,0.15)}}
.toast.ok{{border-color:#86efac;color:#16a34a}}
.toast.erro{{border-color:#fca5a5;color:#dc2626}}
.toast.show{{display:flex}}

/* ── SPINNER ── */
.spinner-overlay{{position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(240,244,250,0.75);z-index:88888;display:none;align-items:center;justify-content:center}}
.spinner-overlay.show{{display:flex}}
.spinner{{width:44px;height:44px;border:4px solid #dde6f4;border-top-color:#3b7dd8;border-radius:50%;animation:spin 0.8s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}

/* ── MODAL FILIAL ── */
.modal-overlay{{position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(20,40,100,0.45);backdrop-filter:blur(8px);display:none;justify-content:center;align-items:center;z-index:9999}}
.modal-box{{background:#ffffff;border:1.5px solid #dde6f4;width:100%;max-width:100%;height:100vh;border-radius:0;padding:20px;display:flex;flex-direction:column;gap:12px;box-shadow:0 16px 48px rgba(20,50,120,0.18)}}
.modal-header{{display:flex;justify-content:space-between;align-items:center;border-bottom:1.5px solid #dde6f4;padding-bottom:10px}}
.modal-title{{font-size:18px;font-weight:800;color:#1a3a6b;text-transform:uppercase;display:flex;align-items:center;gap:8px}}
.btn-close{{background:#7a1a1a;color:#ffffff;border:1px solid #5c1212;width:28px;height:28px;border-radius:50%;cursor:pointer;display:flex;align-items:center;justify-content:center}}
.btn-close:hover{{background:#ff2222;color:#fff;border-color:#ff4444;box-shadow:0 0 10px rgba(255,50,50,0.7),0 0 20px rgba(255,50,50,0.4)}}
.modal-split{{display:grid;grid-template-columns:280px 1fr;gap:14px;flex:1;overflow:hidden}}
.modal-sidebar{{display:flex;flex-direction:column;gap:8px;overflow-y:auto;padding-right:4px}}
.modal-kpi-card{{background:#f8fafd;border:1.5px solid #dde6f4;padding:14px;border-radius:8px;cursor:pointer;transition:border-color .15s,box-shadow .15s}}
.modal-kpi-card:hover{{border-color:#3b7dd8;box-shadow:0 2px 12px rgba(59,125,216,0.18)}}
.m-lbl{{font-size:14px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:1px}}
.m-val{{font-size:36px;font-weight:900;color:#1a3a6b}}
.modal-main{{background:#f8fafd;border:1.5px solid #dde6f4;border-radius:8px;display:flex;flex-direction:column;overflow:hidden}}
.table-container{{flex:1;overflow-y:auto}}

/* ── Lista mobile (dentro do modal de Filial) ── */
.filial-mobile-list{{display:none;flex-direction:column;gap:10px;overflow-y:auto;padding:10px 12px;flex:1}}
.filial-mobile-backbar{{display:none;align-items:center;gap:10px;padding:10px 12px 8px;border-bottom:1px solid #eef3fb;background:#fff;flex-shrink:0}}
.btn-voltar-mobile{{background:transparent;color:#3b7dd8;border:1.5px solid #3b7dd8;padding:6px 14px;font-size:11px;font-weight:800;letter-spacing:1px;text-transform:uppercase;border-radius:6px;cursor:pointer;display:flex;align-items:center;gap:6px;flex-shrink:0}}
.btn-voltar-mobile:hover{{background:#3b7dd8;color:#fff}}
.filial-mobile-titulo{{font-size:14px;font-weight:800;color:#1a3a6b;text-transform:uppercase;letter-spacing:.5px}}
.m-table{{width:100%;border-collapse:collapse;text-align:left;font-size:14px}}
.m-table th{{background:#eef3fb;color:#1a4fa0;font-size:12px;font-weight:800;text-transform:uppercase;padding:14px 16px;border-bottom:1.5px solid #dde6f4;position:sticky;top:0}}
.m-table td{{padding:14px 16px;border-bottom:1px solid #eef3fb;color:#2a3a55}}
.driver-row{{cursor:pointer}}
.driver-row:hover{{background:#eef3fb!important}}
.m-name{{font-weight:700;color:#1a3a6b;font-size:15px}}
.m-cpf{{font-family:monospace;font-size:15px;color:#5a6e8a}}
.m-badge{{display:inline-block;padding:4px 10px;border-radius:4px;font-size:12px;font-weight:700}}
.m-badge.ok{{background:rgba(22,163,74,0.1);color:#16a34a;border:1px solid rgba(22,163,74,0.25)}}
.m-badge.pend{{background:rgba(217,119,6,0.1);color:#d97706;border:1px solid rgba(217,119,6,0.25)}}
.m-count-badge{{font-weight:700;color:#1a3a6b;background:#eef3fb;padding:4px 10px;border-radius:4px;border:1px solid #c4d0e4;font-size:14px}}

/* ── FORMULÁRIO ── */
.admin-panel{{background:#ffffff;border:1.5px solid #c4d0e4;border-radius:10px;margin-bottom:12px;overflow:hidden;box-shadow:0 2px 8px rgba(20,50,120,0.06)}}
.admin-panel-header{{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;cursor:pointer;user-select:none}}
.admin-panel-header:hover{{background:#f4f8ff}}
.admin-panel-title{{display:flex;align-items:center;gap:10px;font-size:14px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#16a34a}}
.admin-panel-title::before{{content:'';display:inline-block;width:3px;height:10px;background:#16a34a;border-radius:2px}}
.btn-toggle-form{{background:rgba(22,163,74,0.08);color:#16a34a;border:1px solid rgba(22,163,74,0.3);border-radius:5px;padding:8px 16px;font-size:13px;font-weight:700;text-transform:uppercase;cursor:pointer;display:flex;align-items:center;gap:6px;transition:.2s}}
.btn-toggle-form:hover{{background:#16a34a;color:#fff}}
.btn-toggle-form .chevron{{transition:transform .3s}}
.btn-toggle-form.open .chevron{{transform:rotate(180deg)}}
.admin-panel-body{{max-height:0;overflow:hidden;transition:max-height .35s ease,padding .35s ease;padding:0 16px}}
.admin-panel-body.open{{max-height:200px;padding:0 16px 16px}}
.form-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-top:10px}}
.form-group{{display:flex;flex-direction:column;gap:4px}}
.form-group label{{font-size:15px;color:#5a6e8a;text-transform:uppercase;font-weight:700}}
.form-group input,.form-group select{{background:#f4f7fc;border:1px solid #c4d0e4;color:#1a2a44;padding:8px 12px;border-radius:4px;font-size:14px;outline:none}}
.form-group input:focus,.form-group select:focus{{border-color:#3b7dd8;background:#fff}}
.btn-add{{background:#16a34a;color:#fff;border:none;font-weight:700;text-transform:uppercase;cursor:pointer;padding:0 16px;border-radius:4px;height:32px;margin-top:17px;display:flex;align-items:center;justify-content:center;gap:6px;font-size:11px}}
.btn-add:hover{{background:#15803d}}
.btn-save-master{{background:transparent;color:#16a34a;border:1.5px solid #16a34a;padding:7px 20px;font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:1.5px;border-radius:6px;cursor:pointer;display:flex;align-items:center;gap:7px;transition:color .2s,border-color .2s,box-shadow .2s,background .2s}}
.btn-save-master:hover{{background:transparent;color:#dc2626;border-color:#dc2626;box-shadow:0 0 12px rgba(220,38,38,0.18)}}

/* ── FICHA INDIVIDUAL ── */
.driver-profile-grid{{display:grid;grid-template-columns:340px 1fr;gap:16px;padding:18px;background:#f0f4fa}}
.profile-details-right{{display:flex;flex-direction:column;gap:14px}}
.info-section-box{{background:#ffffff;border-radius:12px;padding:0;overflow:hidden;box-shadow:0 2px 8px rgba(20,50,120,0.08);border:1.5px solid #dde6f4;transition:box-shadow .2s}}
.info-section-box:hover{{box-shadow:0 4px 16px rgba(20,50,120,0.13)}}
.card-stripe{{height:4px;width:100%;border-radius:0;display:block}}
.card-body{{padding:14px 16px 16px}}
.info-block-title{{font-size:17px;font-weight:800;text-transform:uppercase;letter-spacing:.8px;border-bottom:1px solid #e8eef8;padding-bottom:6px;margin-bottom:12px;display:flex;align-items:center;gap:7px}}
.card-condutor{{border-color:#3b7dd8}}
.card-condutor .card-stripe{{background:linear-gradient(90deg,#1a4fa0,#3b7dd8)}}
.card-condutor .info-block-title{{color:#1a4fa0}}
.avatar-outer{{position:relative;width:92px;height:92px;margin:0 auto;z-index:20;}}
.avatar-wrapper{{position:relative;width:92px;height:92px;cursor:pointer;border-radius:50%;border:2.5px dashed #3b7dd8;overflow:hidden;display:flex;align-items:center;justify-content:center;background:#e8f0fe}}
.avatar-wrapper img{{width:100%;height:100%;object-fit:cover}}
.avatar-wrapper .upload-hint{{position:absolute;bottom:0;background:rgba(26,79,160,0.82);width:100%;font-size:8px;color:#fff;padding:2px 0;text-transform:uppercase;font-weight:700}}
.avatar-menu{{position:absolute;top:100%;left:50%;transform:translateX(-50%);margin-top:6px;background:#ffffff;border:1.5px solid #c4d0e4;border-radius:8px;box-shadow:0 8px 24px rgba(20,50,120,0.18);z-index:9999;display:none;flex-direction:column;min-width:170px;overflow:visible;}}
.avatar-menu button{{background:none;border:none;padding:10px 14px;text-align:left;font-size:12px;font-weight:700;color:#1a3a6b;cursor:pointer;display:flex;align-items:center;gap:8px;transition:background .15s;width:100%;}}
.avatar-menu button:hover{{background:#eef3fb}}
.avatar-menu button#avatarMenuExcluir{{color:#dc2626}}
.avatar-menu button#avatarMenuExcluir:hover{{background:#fff5f5}}
.profile-card-left{{background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(20,50,120,0.10);border:1.5px solid #3b7dd8;display:flex;flex-direction:column}}
.profile-card-left .card-stripe{{background:linear-gradient(90deg,#1a4fa0,#3b7dd8)}}
.profile-card-left-body{{padding:16px;text-align:center;display:flex;flex-direction:column;gap:10px;flex:1;justify-content:space-between}}
.card-contato{{border-color:#0e9cc0}}
.card-contato .card-stripe{{background:linear-gradient(90deg,#0a7a9a,#0eb8e0)}}
.card-contato .info-block-title{{color:#0a7a9a;background:linear-gradient(135deg,#e8f8fc,#f0fcff);margin:-14px -16px 12px;padding:10px 16px 8px;border-radius:0}}
.card-docs{{border-color:#5a5fe8}}
.card-docs .card-stripe{{background:linear-gradient(90deg,#3a3ec8,#6c72f5)}}
.card-docs .info-block-title{{color:#3a3ec8;background:linear-gradient(135deg,#eeeeff,#f4f4ff);margin:-14px -16px 12px;padding:10px 16px 8px;border-radius:0}}
.card-exames{{border-color:#7c3aed}}
.card-exames .card-stripe{{background:linear-gradient(90deg,#6d28d9,#a78bfa)}}
.card-exames .info-block-title{{color:#6d28d9;background:linear-gradient(135deg,#f3f0ff,#f8f6ff);margin:-14px -16px 12px;padding:10px 16px 8px;border-radius:0}}
.card-seguranca{{border-color:#d97706}}
.card-seguranca .card-stripe{{background:linear-gradient(90deg,#b45309,#f59e0b)}}
.card-seguranca .info-block-title{{color:#b45309;background:linear-gradient(135deg,#fef5e6,#fff8ed);margin:-14px -16px 12px;padding:10px 16px 8px;border-radius:0}}
.card-dss{{border-color:#16a34a}}
.card-dss .card-stripe{{background:linear-gradient(90deg,#15803d,#22c55e)}}
.card-dss .info-block-title{{color:#15803d;background:linear-gradient(135deg,#ecfdf5,#f0fef8);margin:-14px -16px 12px;padding:10px 16px 8px;border-radius:0}}
@keyframes kpi-pulse-red{{0%,100%{{box-shadow:0 0 0 3px rgba(229,62,62,.20),0 4px 20px rgba(229,62,62,.12)}}50%{{box-shadow:0 0 0 5px rgba(229,62,62,.36),0 6px 28px rgba(229,62,62,.22)}}}}
@keyframes kpi-pulse-orange{{0%,100%{{box-shadow:0 0 0 3px rgba(221,107,32,.20),0 4px 20px rgba(221,107,32,.12)}}50%{{box-shadow:0 0 0 5px rgba(221,107,32,.36),0 6px 28px rgba(221,107,32,.22)}}}}
@keyframes kpi-pulse-green{{0%,100%{{box-shadow:0 0 0 3px rgba(22,163,74,.20),0 4px 20px rgba(22,163,74,.12)}}50%{{box-shadow:0 0 0 5px rgba(22,163,74,.36),0 6px 28px rgba(22,163,74,.22)}}}}
.card-highlight-vel{{border-color:#e53e3e!important;animation:kpi-pulse-red 2s ease-in-out infinite}}
.card-highlight-vel .card-stripe{{background:linear-gradient(90deg,#b91c1c,#e53e3e)!important}}
.card-highlight-mul{{border-color:#dd6b20!important;animation:kpi-pulse-orange 2s ease-in-out infinite}}
.card-highlight-mul .card-stripe{{background:linear-gradient(90deg,#c2410c,#dd6b20)!important}}
.card-highlight-acid{{border-color:#e53e3e!important;animation:kpi-pulse-red 2s ease-in-out infinite}}
.card-highlight-acid .card-stripe{{background:linear-gradient(90deg,#b91c1c,#e53e3e)!important}}
.card-highlight-dss{{border-color:#16a34a!important;animation:kpi-pulse-green 2s ease-in-out infinite}}
.card-highlight-dss .card-stripe{{background:linear-gradient(90deg,#15803d,#22c55e)!important}}
.meta-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
.meta-item{{display:flex;flex-direction:column;gap:3px}}
.meta-item label{{font-size:15px;color:#5a6e8a;text-transform:uppercase;font-weight:700}}
.meta-item input,.meta-item select{{background:#f4f7fc;border:1px solid #c4d0e4;color:#1a2a44;padding:9px 10px;border-radius:5px;font-size:17px;outline:none;transition:border-color .15s,background .15s,box-shadow .15s}}
.meta-item input:focus,.meta-item select:focus{{border-color:#3b7dd8;background:#ffffff;box-shadow:0 0 0 2px rgba(59,125,216,.12)}}
.card-contato .meta-item input:focus,.card-contato .meta-item select:focus{{border-color:#0e9cc0;box-shadow:0 0 0 2px rgba(14,156,192,.12)}}
.card-docs .meta-item input:focus,.card-docs .meta-item select:focus{{border-color:#5a5fe8;box-shadow:0 0 0 2px rgba(90,95,232,.12)}}
.card-seguranca .meta-item input:focus,.card-seguranca .meta-item select:focus{{border-color:#d97706;box-shadow:0 0 0 2px rgba(217,119,6,.12)}}
.card-dss .meta-item input:focus,.card-dss .meta-item select:focus{{border-color:#16a34a;box-shadow:0 0 0 2px rgba(22,163,74,.12)}}
.obs-input{{background:#f9fafd!important;border:1px solid #d0daea!important;color:#3a4a62!important;font-size:14px!important;font-style:italic;padding:8px 10px!important}}
.campo-valido{{border-color:#16a34a!important;background:#f0fef4!important;box-shadow:0 0 0 1px rgba(22,163,74,.25)!important}}
.campo-alerta-venc{{border-color:#d97706!important;background:#fffbeb!important;box-shadow:0 0 0 1px rgba(217,119,6,.25)!important}}
.campo-vencido-venc{{border-color:#dc2626!important;background:#fff5f5!important;box-shadow:0 0 0 1px rgba(220,38,38,.25)!important}}
.dss-matrix-container{{display:grid;grid-template-columns:repeat(4,1fr);gap:7px}}
.month-dss-box{{background:#f0faf4;border:1px solid #bbddc8;border-radius:7px;padding:7px}}
.month-name-lbl{{font-size:14px;font-weight:800;color:#15803d;text-transform:uppercase;margin-bottom:4px;text-align:center;border-bottom:1px solid #c8e8d4;padding-bottom:3px}}
.weeks-flex{{display:flex;justify-content:space-between;gap:2px}}
.week-checkbox-label{{display:flex;flex-direction:column;align-items:center;gap:2px;font-size:14px;color:#2d6a4a;cursor:pointer;flex:1;font-weight:600}}
.week-checkbox-label input{{cursor:pointer;accent-color:#16a34a}}
.btn-delete-driver{{background:#fff0f0;color:#cc2222;border:1px solid #e8aaaa;padding:8px;border-radius:7px;font-size:10px;font-weight:700;text-transform:uppercase;cursor:pointer;margin-top:12px;display:flex;align-items:center;justify-content:center;gap:6px;transition:.18s;width:100%}}
.btn-delete-driver:hover{{background:#cc2222;color:#fff;border-color:#cc2222}}
.btn-desligar-driver{{background:#fff7ed;color:#c2410c;border:1px solid #fbbf7a;padding:8px;border-radius:7px;font-size:10px;font-weight:700;text-transform:uppercase;cursor:pointer;margin-top:12px;display:flex;align-items:center;justify-content:center;gap:6px;transition:.18s;width:100%}}
.btn-desligar-driver:hover{{background:#c2410c;color:#fff;border-color:#c2410c}}
.btn-reativar-driver{{background:#f0fef4;color:#16a34a;border:1px solid #86efac;padding:8px;border-radius:7px;font-size:10px;font-weight:700;text-transform:uppercase;cursor:pointer;margin-top:12px;display:flex;align-items:center;justify-content:center;gap:6px;transition:.18s;width:100%}}
.btn-reativar-driver:hover{{background:#16a34a;color:#fff;border-color:#16a34a}}
#btnConfirmarFicha:hover{{background:#2ea84a!important;border-color:#22883a!important}}
#btnFecharFicha:hover{{background:#b52222!important;border-color:#8a1818!important}}
#btnVoltarFicha:hover{{background:#b52222!important;border-color:#8a1818!important}}
#driverModal .form-group label{{color:#5a6e8a}}
#driverModal .form-group input,#driverModal .form-group select{{background:#f0f4fb;border:1px solid #bccce0;color:#1a2a44}}
.kpi{{cursor:pointer;transition:transform .15s,box-shadow .15s}}
.kpi:hover{{transform:translateY(-2px);box-shadow:0 6px 20px rgba(20,50,120,0.12)}}

/* ── MODAL KPI ── */
.kpi-modal-overlay{{position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(20,40,100,0.45);backdrop-filter:blur(10px);display:none;justify-content:stretch;align-items:stretch;z-index:11000;padding:12px;box-sizing:border-box}}
.kpi-modal-overlay.show{{display:flex}}
.kpi-modal-box{{background:#ffffff;border:1.5px solid #dde6f4;width:100%;height:100%;max-width:none;max-height:none;border-radius:12px;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 16px 48px rgba(20,50,120,0.18)}}
.kpi-modal-head{{display:flex;align-items:center;justify-content:space-between;padding:14px 20px;border-bottom:1.5px solid #dde6f4;flex-shrink:0;background:#f8fafd}}
.kpi-modal-head-left{{display:flex;align-items:center;gap:12px}}
.kpi-modal-icon{{width:38px;height:38px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}}
.kpi-modal-label{{font-size:16px;font-weight:800;color:#1a3a6b;text-transform:uppercase;letter-spacing:1.5px}}
.kpi-modal-count{{font-size:14px;color:#5a6e8a;margin-top:2px}}
.kpi-modal-close{{background:#7a1a1a;color:#ffffff;border:1px solid #5c1212;width:30px;height:30px;border-radius:50%;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0}}
.kpi-modal-close:hover{{background:#ff2222;color:#fff;border-color:#ff4444;box-shadow:0 0 10px rgba(255,50,50,0.7),0 0 20px rgba(255,50,50,0.4)}}
.kpi-modal-search{{padding:10px 20px;border-bottom:1px solid #eef3fb;flex-shrink:0;background:#fff}}
.kpi-modal-search input{{width:100%;background:#f4f7fc;border:1.5px solid #c4d0e4;color:#1a2a44;padding:7px 12px;border-radius:6px;font-size:12px;outline:none}}
.kpi-modal-search input:focus{{border-color:#3b7dd8;background:#fff}}
.kpi-mes-filtro{{display:none;padding:8px 20px 0;gap:6px;flex-wrap:wrap;flex-shrink:0;background:#fff}}
.kpi-mes-filtro.visible{{display:flex}}
.mes-btn{{padding:6px 14px;border-radius:20px;border:1.5px solid #c4d0e4;background:#f4f7fc;color:#5a6e8a;font-size:13px;font-weight:700;cursor:pointer;letter-spacing:.5px;transition:all .15s}}
.mes-btn.ativo{{background:#16a34a;border-color:#16a34a;color:#fff}}
.dmc-semanas{{display:flex;gap:4px;margin-top:3px}}
.dmc-sem{{display:flex;flex-direction:column;align-items:center;gap:2px;flex:1}}
.dmc-sem-lbl{{font-size:13px;color:#8899aa;font-weight:700;text-transform:uppercase}}
.dmc-sem-dot{{width:28px;height:28px;border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:900}}
.dmc-sem-dot.ok{{background:rgba(22,163,74,0.15);color:#16a34a;border:1px solid rgba(22,163,74,0.35)}}
.dmc-sem-dot.pend{{background:rgba(220,38,38,0.08);color:#dc2626;border:1px solid rgba(220,38,38,0.2)}}

.driver-mini-card.card-ok{{border-color:rgba(22,163,74,0.45);background:#f0fef4}}.kpi-cards-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:14px;padding:20px 24px;overflow-y:auto;flex:1;align-content:start;background:#f0f4fa}}
.categoria-glass-panel{{background:rgba(255,255,255,0.55);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);border:1.5px solid rgba(196,208,228,0.6);border-radius:16px;padding:16px;cursor:pointer;transition:transform .18s,box-shadow .18s,border-color .18s;box-shadow:0 4px 18px rgba(20,50,120,0.08);display:flex;flex-direction:column;gap:12px}}
.categoria-glass-panel:hover{{transform:translateY(-3px);box-shadow:0 10px 30px rgba(20,50,120,0.16);border-color:#3b7dd8}}
.cgp-header{{display:flex;align-items:center;gap:10px}}
.cgp-icon{{width:38px;height:38px;border-radius:10px;background:rgba(22,163,74,0.12);color:#16a34a;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}}
.cgp-titulo{{font-size:17px;font-weight:800;color:#1a3a6b;letter-spacing:.3px}}
.cgp-stats{{display:grid;grid-template-columns:1fr 1fr;gap:8px}}
.cgp-stat{{border-radius:10px;padding:10px 8px;text-align:center;cursor:pointer;transition:transform .15s,box-shadow .15s;border:1.5px solid}}
.cgp-stat:hover{{transform:translateY(-2px)}}
.cgp-stat.ok{{background:rgba(22,163,74,0.08);border-color:rgba(22,163,74,0.35)}}
.cgp-stat.pend{{background:rgba(217,119,6,0.08);border-color:rgba(217,119,6,0.35)}}
.cgp-stat-val{{font-size:26px;font-weight:900;line-height:1}}
.cgp-stat.ok .cgp-stat-val{{color:#16a34a}}
.cgp-stat.pend .cgp-stat-val{{color:#d97706}}
.cgp-stat-lbl{{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:4px}}
.cgp-stat.ok .cgp-stat-lbl{{color:#16a34a}}
.cgp-stat.pend .cgp-stat-lbl{{color:#d97706}}
.driver-mini-card{{background:#ffffff;border:1.5px solid #dde6f4;border-radius:10px;padding:14px;cursor:pointer;transition:border-color .15s,transform .15s,box-shadow .15s;display:flex;flex-direction:column;gap:10px;box-shadow:0 2px 6px rgba(20,50,120,0.06)}}
.driver-mini-card.card-pend{{border-color:rgba(217,119,6,0.45);background:#fffbeb}}
.driver-mini-card:hover{{border-color:#3b7dd8;transform:translateY(-2px);box-shadow:0 8px 24px rgba(20,50,120,0.14)}}
.dmc-top{{display:flex;align-items:center;gap:10px}}
.dmc-avatar{{width:48px;height:48px;border-radius:50%;background:#eef3fb;border:2px solid #c4d0e4;display:flex;align-items:center;justify-content:center;font-size:20px;color:#3b7dd8;flex-shrink:0;overflow:hidden}}
.dmc-avatar img{{width:100%;height:100%;object-fit:cover;border-radius:50%}}
.dmc-info{{flex:1;min-width:0}}
.dmc-nome{{font-size:18px;font-weight:800;color:#1a3a6b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;letter-spacing:.2px;margin-bottom:3px}}
.dmc-filial{{font-size:15px;color:#3b7dd8;text-transform:uppercase;font-weight:700;letter-spacing:.5px;margin-bottom:2px}}
.dmc-cpf{{font-size:15px;color:#5a6e8a;font-family:monospace;letter-spacing:.8px;font-weight:600}}
.dmc-badges{{display:flex;flex-wrap:wrap;gap:5px}}
.dmc-badge{{padding:6px 14px;border-radius:4px;font-size:15px;font-weight:800;letter-spacing:.3px;display:flex;align-items:center;gap:4px}}
.dmc-badge.ok{{background:rgba(22,163,74,0.1);color:#16a34a;border:1px solid rgba(22,163,74,0.3)}}
.dmc-badge.pend{{background:rgba(217,119,6,0.1);color:#d97706;border:1px solid rgba(217,119,6,0.3)}}
.dmc-infracao{{display:flex;align-items:center;gap:10px;background:#f4f7fc;border-radius:8px;padding:10px 12px;border:1.5px solid #dde6f4}}
.dmc-inf-icon{{font-size:20px;flex-shrink:0}}
.dmc-inf-body{{flex:1;min-width:0}}
.dmc-inf-label{{font-size:12px;color:#5a6e8a;text-transform:uppercase;letter-spacing:1px;font-weight:700}}
.dmc-inf-val{{font-size:28px;font-weight:900;line-height:1;margin-top:1px}}
.dmc-inf-val.vel{{color:#dc2626}}.dmc-inf-val.mul{{color:#d97706}}.dmc-inf-val.acid{{color:#be185d}}

/* ── Card Prontuário ── */
.dmc-top-center{{display:flex;flex-direction:column;align-items:center;text-align:center;gap:4px}}
.dmc-top-center .dmc-avatar{{width:56px;height:56px;font-size:22px}}
.dmc-top-center .dmc-nome{{font-size:15px;white-space:normal;overflow:visible;text-overflow:unset;margin-bottom:0}}
.dmc-top-center .dmc-filial{{margin-bottom:0}}
.dmc-top-center .dmc-cpf{{font-size:12px}}
.dmc-pront-grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px 10px;background:#f8fafd;border:1px solid #eef3fb;border-radius:8px;padding:10px 12px}}
.dmc-pront-item{{display:flex;flex-direction:column;gap:2px;min-width:0}}
.dmc-pront-item.full{{grid-column:1/-1}}
.dmc-pront-lbl{{font-size:11px;color:#8899aa;text-transform:uppercase;letter-spacing:.5px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex-shrink:1;min-width:0}}
.dmc-pront-val{{font-size:15px;font-weight:700;color:#1a3a6b;white-space:nowrap;flex-shrink:0}}
.dmc-status-row{{display:grid;grid-template-columns:1fr 1fr;gap:8px}}
.dmc-status-pill{{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;padding:8px 6px;border-radius:8px;border:1.5px solid}}
.dmc-status-pill.ok{{background:rgba(22,163,74,0.08);border-color:rgba(22,163,74,0.35)}}
.dmc-status-pill.pend{{background:rgba(217,119,6,0.08);border-color:rgba(217,119,6,0.35)}}
.dmc-status-pill.vencido{{background:rgba(220,38,38,0.08);border-color:rgba(220,38,38,0.35)}}
.dmc-status-pill-lbl{{font-size:11px;text-transform:uppercase;letter-spacing:.5px;font-weight:700;color:#5a6e8a}}
.dmc-status-pill-val{{font-size:15px;font-weight:800}}
.dmc-status-pill.ok .dmc-status-pill-val{{color:#16a34a}}
.dmc-status-pill.pend .dmc-status-pill-val{{color:#d97706}}
.dmc-status-pill.vencido .dmc-status-pill-val{{color:#dc2626}}
.kpi-empty{{grid-column:1/-1;text-align:center;padding:40px;color:#9aaabb;font-size:12px}}
.kpi-empty i{{font-size:32px;margin-bottom:10px;display:block;color:#c4d0e4}}
.empty-state{{text-align:center;padding:40px;color:#9aaabb}}
.empty-state i{{font-size:36px;margin-bottom:12px;color:#c4d0e4}}
.empty-state p{{font-size:15px}}

/* ── BOTÕES DO CABEÇALHO DO ORGANOGRAMA (Salvar / Adicionar) ── */
.org-save-btn{{background:#1a5c2a;color:#ffffff;border:1px solid #14481f;padding:0 22px;height:36px;border-radius:6px;font-weight:800;font-size:13px;cursor:pointer;display:flex;align-items:center;gap:8px;transition:background .2s,box-shadow .2s}}
.org-save-btn:hover{{background:#22883a;box-shadow:0 0 10px rgba(34,153,66,0.55),0 0 18px rgba(34,153,66,0.3)}}
.org-add-btn-wrap{{position:relative;display:inline-block}}
.org-add-btn{{background:#16a34a;color:#ffffff;border:1px solid #0f7a37;padding:0 22px;height:36px;border-radius:6px;font-weight:800;font-size:13px;cursor:pointer;display:flex;align-items:center;gap:8px;transition:background .2s,box-shadow .2s}}
.org-add-btn:hover, .org-add-btn.open{{background:#15803d;box-shadow:0 0 10px rgba(22,163,74,0.55),0 0 18px rgba(22,163,74,0.35)}}
.org-quick-menu-panel{{position:absolute;top:calc(100% + 8px);left:50%;transform:translateX(-50%) translateY(-6px);max-height:0;overflow:hidden;opacity:0;transition:max-height .3s ease,opacity .25s ease,transform .25s ease,padding .3s ease;background:#ffffff;border:1.5px solid #c4d0e4;border-radius:12px;box-shadow:0 12px 32px rgba(20,50,120,0.2);padding:0 14px;width:240px;pointer-events:none;z-index:500}}
.org-quick-menu-panel.open{{max-height:400px;opacity:1;transform:translateX(-50%) translateY(0);padding:14px;pointer-events:auto}}
.org-quick-menu-title{{font-size:12px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:#16a34a;margin-bottom:10px}}
.org-add-setor-btn{{width:100%;background:#f0fef4;color:#16a34a;border:1.5px solid rgba(22,163,74,0.35);border-radius:6px;padding:9px 10px;font-size:12px;font-weight:700;cursor:pointer;display:flex;align-items:center;gap:8px;margin-bottom:8px;text-align:left;transition:background .15s,color .15s}}
.org-add-setor-btn:hover{{background:#16a34a;color:#fff}}
.org-add-setor-btn:last-child{{margin-bottom:0}}

/* ── RESPONSIVO ── */
@media (max-width:1024px){{
  .kpi-row{{grid-template-columns:repeat(3,1fr)}}
  .filial-grid{{grid-template-columns:repeat(2,1fr)}}
  .driver-profile-grid{{grid-template-columns:1fr;padding:12px}}
  .profile-card-left{{max-width:100%}}
  .meta-grid{{grid-template-columns:repeat(2,1fr)}}
  .dss-matrix-container{{grid-template-columns:repeat(3,1fr)}}
}}

@media (max-width:768px){{
  .top-bar{{flex-wrap:wrap;gap:8px;padding:8px 10px}}
  .luft-name{{font-size:20px;display:none}}
  .pct-val{{font-size:22px}}
  .kpi-row{{grid-template-columns:repeat(2,1fr);gap:8px}}
  .kpi-val{{font-size:36px}}
  .kpi{{padding:12px 14px}}
  .content{{padding:8px 8px}}
  .panel{{padding:8px}}
  .filial-grid{{grid-template-columns:1fr}}
  .fc-count{{font-size:36px}}
  .modal-split{{grid-template-columns:1fr;grid-template-rows:auto 1fr}}
  .modal-sidebar{{flex-direction:row;flex-wrap:wrap;gap:6px}}
  .modal-kpi-card{{flex:1;min-width:120px;padding:8px}}
  .m-val{{font-size:22px}}
  .modal-sidebar.mobile-hidden{{display:none}}
  .table-container.mobile-hidden{{display:none}}
  .filial-mobile-list.show{{display:flex}}
  .filial-mobile-backbar.show{{display:flex}}
  .modal-box{{padding:10px;gap:8px}}
  .driver-profile-grid{{grid-template-columns:1fr;padding:8px;gap:10px}}
  .profile-details-right{{gap:8px}}
  .meta-grid{{grid-template-columns:1fr}}
  .dss-matrix-container{{grid-template-columns:repeat(2,1fr)}}
  .kpi-cards-grid{{grid-template-columns:1fr;padding:10px 12px;gap:10px}}
  .kpi-modal-overlay{{padding:4px}}
  .kpi-modal-head{{padding:10px 12px}}
  .kpi-modal-label{{font-size:13px}}
  .admin-panel-body.open{{max-height:420px}}
  .form-grid{{grid-template-columns:1fr 1fr;gap:8px}}
  .charts-row{{grid-template-columns:1fr!important}}
  .org-save-btn,.org-add-btn{{padding:0 14px;font-size:11px;gap:5px}}
  .org-quick-menu-panel{{width:min(78vw,240px)}}
  .org-quick-menu-panel.open{{max-height:360px}}
}}

@media (max-width:900px){{
  .charts-row{{grid-template-columns:1fr!important}}
}}

@media (max-width:480px){{
  .kpi-row{{grid-template-columns:repeat(2,1fr);gap:6px}}
  .kpi-lbl{{font-size:10px;letter-spacing:.5px}}
  .kpi-sub{{font-size:10px}}
  .luft-name{{font-size:16px;display:none}}
  .brand-logo{{padding:4px 8px}}
  .pct-box{{padding:4px 8px}}
  .pct-val{{font-size:18px}}
  .pct-lbl{{font-size:9px}}
  .modal-title{{font-size:13px}}
  .modal-sidebar{{flex-direction:column}}
  .modal-kpi-card{{min-width:unset}}
  .dss-matrix-container{{grid-template-columns:repeat(2,1fr);gap:4px}}
  .month-dss-box{{padding:5px}}
  .month-name-lbl{{font-size:10px}}
  .week-checkbox-label{{font-size:10px}}
  .form-grid{{grid-template-columns:1fr;gap:6px}}
  .m-table{{font-size:12px}}
  .m-table th,.m-table td{{padding:8px 8px}}
  .fc-name{{font-size:16px}}
  .fc-count{{font-size:28px}}
  .sbar-lbl{{font-size:11px;width:52px}}
}}
</style>
</head>
<body class="db">

<!-- SPLASH -->
<div id="splash-screen">
  <div id="splash-bg"></div>
  <div class="splash-card">
    <div class="splash-logo-txt">Gestão<span> Operacional</span></div>
    <div class="splash-sub">Sistema de Controle de Motoristas</div>

    <!-- TELA DE LOGIN -->
    <div id="loginBox">
      <div class="splash-label" style="margin-bottom:18px;"><i class="fa-solid fa-lock" style="margin-right:6px"></i>Acesso Restrito</div>
      <div style="display:flex;flex-direction:column;gap:10px;margin-bottom:14px;">
        <input type="text" id="loginUser" placeholder="Usuário"
          style="background:rgba(255,255,255,0.08);border:1.5px solid rgba(59,125,216,0.4);color:#fff;padding:10px 14px;border-radius:8px;font-size:14px;outline:none;letter-spacing:.5px;"
           >
        <input type="password" id="loginPass" placeholder="Senha"
          style="background:rgba(255,255,255,0.08);border:1.5px solid rgba(59,125,216,0.4);color:#fff;padding:10px 14px;border-radius:8px;font-size:14px;outline:none;letter-spacing:.5px;"
          >
      </div>
       <button id="btnEntrar"
        style="width:100%;background:#3b7dd8;color:#fff;border:none;padding:11px;border-radius:8px;font-size:13px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;cursor:pointer;">
        <i class="fa-solid fa-right-to-bracket" style="margin-right:6px"></i>Entrar
      </button>
      <div id="loginErro" style="color:#ff4444;font-size:12px;margin-top:10px;min-height:16px;font-weight:600;"></div>
    </div>

    <!-- TELA DE CARREGAMENTO (oculta até login ok) -->
    <div id="loadingBox" style="display:none;">
      <div class="splash-label"><i class="fa-solid fa-database" style="margin-right:6px"></i>Conectando ao Google Sheets</div>
      <div class="splash-drop-area" style="cursor:default;pointer-events:none;margin-top:14px;">
        <div class="splash-drop-icon"><i class="fa-brands fa-google" style="color:#34a853"></i></div>
        <div class="splash-drop-txt">
          <strong>Google Sheets</strong><br>
          Carregando base de dados...
        </div>
      </div>
      <div class="splash-progress" id="splashProgress"><div class="splash-progress-bar" id="splashProgressBar"></div></div>
      <div class="splash-status" id="splashStatus">Aguardando conexão...</div>
    </div>
  </div>
</div>

<!-- Spinner global -->
<div class="spinner-overlay" id="spinnerOverlay"><div class="spinner"></div></div>

<!-- Toast global -->
<div class="toast" id="toastMsg"><i class="fa-solid fa-circle-check"></i><span id="toastText"></span></div>

<div class="top-bar">
  <div class="brand" style="min-width:220px;">
    <div class="brand-logo"><div class="dot-anim"></div><div class="brand-name">Controle<br><span class="brand-sub">Motoristas</span></div></div>
    <div id="topbarNomeUsuario" style="font-size:26px;font-weight:900;color:#1a4fa0;letter-spacing:-0.5px;"></div>
  </div>
  <div style="position:absolute;left:50%;transform:translateX(-50%);">
    <div class="luft-name" style="font-size:32px;">LUFT<span> LOGISTICS</span></div>
  </div>
  <div class="pct-box" style="display:flex;align-items:center;gap:16px;min-width:220px;justify-content:flex-end;">
    <div>
      <div class="pct-lbl">Regularidade Geral DSS</div>
      <div class="pct-val" id="macroPctDss">—</div>
    </div>
    <button class="btn-save-master" onclick="atualizarDadosDoSheets()" title="Buscar dados atualizados da planilha">
      <i class="fa-solid fa-rotate"></i> Atualizar
    </button>
  </div>
</div>

<div class="content">

  <div class="admin-panel">
    <div class="admin-panel-header" onclick="toggleFormulario()">
      <div class="admin-panel-title"><i class="fa-solid fa-user-plus"></i> Inclusão de Condutores</div>
      <button class="btn-toggle-form" id="btnToggleForm">
        <i class="fa-solid fa-plus"></i> Novo Condutor
        <i class="fa-solid fa-chevron-down chevron"></i>
      </button>
    </div>
    <div class="admin-panel-body" id="formBody">
      <div class="form-grid">
        <div class="form-group"><label>CPF do Motorista</label><input type="text" id="addCpf" placeholder="000.000.000-00"></div>
        <div class="form-group"><label>Nome Completo</label><input type="text" id="addNome" placeholder="Nome do profissional"></div>
        <div class="form-group"><label>Filial Base</label><input type="text" id="addFilial" placeholder="Ex: BARUERI"></div>
        <div class="form-group">
          <label>Curso Reciclagem</label>
          <select id="addRec"><option value="PENDENTE">PENDENTE</option><option value="OK">OK</option></select>
        </div>
        <div class="form-group">
          <label>Sessão Simulador</label>
          <select id="addSim"><option value="PENDENTE">PENDENTE</option><option value="OK">OK</option></select>
        </div>
        <button class="btn-add" onclick="adicionarNovoMotorista()"><i class="fa-solid fa-plus"></i> Inserir Condutor</button>
      </div>
    </div>
  </div>

  <div class="kpi-row">
    <div class="kpi purple" onclick="abrirKpiModal('prontuario')" title="Ver prontuário de exames dos motoristas">
      <div class="kpi-lbl">Prontuário</div>
      <div style="display:flex;align-items:flex-end;justify-content:space-between;">
        <div class="kpi-val" id="kpiProntuario">—</div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;padding-bottom:4px;gap:1px;">
          <div style="font-size:10px;color:#7c3aed;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8;">exames ok</div>
          <div id="kpiProntuarioOk" style="font-size:17px;font-weight:900;color:#7c3aed;font-family:'Courier New',monospace;letter-spacing:1px;text-shadow:0 0 6px rgba(124,58,237,0.5);">—</div>
        </div>
      </div>
      <div class="kpi-sub">Exames &amp; Complementares</div>
    </div>

    <div class="kpi amber" onclick="abrirVencimentoMenu('alerta')" title="Ver categorias com vencimento em até 30 dias" style="background:#fefce8;border-color:#eab308;box-shadow:0 0 10px rgba(234,179,8,0.35),inset 0 0 6px rgba(234,179,8,0.06);">
      <div class="kpi-lbl" style="color:#a16207;">Alertas de Vencimento</div>
      <div class="kpi-val" style="color:#eab308;font-size:34px;">ALERTA</div>
      <div class="kpi-sub" style="color:#ca8a04;">Vencendo em até 30 dias</div>
    </div>

    <div class="kpi amber" onclick="abrirVencimentoMenu('vencido')" title="Ver categorias já vencidas">
      <div class="kpi-lbl">Itens Vencidos</div>
      <div class="kpi-val" style="color:#dc2626;font-size:34px;">VENCIDOS</div>
      <div class="kpi-sub">Prazo já expirado</div>
    </div>

    <div class="kpi green" onclick="abrirCursosMenu()" title="Ver DSS e Reciclagem por categoria">
      <div class="kpi-lbl">Treinamentos</div>
      <div style="display:flex;align-items:flex-end;justify-content:space-between;">
        <div class="kpi-val" id="kpiCursosTotal">—</div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;padding-bottom:4px;gap:1px;">
          <div style="font-size:10px;color:#16a34a;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8;">regularidade DSS</div>
          <div id="kpiCursosSubPct" style="font-size:17px;font-weight:900;color:#16a34a;font-family:'Courier New',monospace;letter-spacing:1px;text-shadow:0 0 6px rgba(34,204,136,0.5);">—</div>
        </div>
      </div>
      <div class="kpi-sub">DSS · Reciclagem</div>
    </div>

    <div class="kpi red" onclick="abrirKpiModal('afastados')" title="Ver motoristas afastados">
      <div class="kpi-lbl">Motoristas Afastados</div>
      <div style="display:flex;align-items:flex-end;justify-content:space-between;">
        <div class="kpi-val" id="kpiAfastados">—</div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;padding-bottom:4px;gap:1px;">
          <div style="font-size:10px;color:#dc2626;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8;">do total</div>
          <div id="kpiAfastadosPct" style="font-size:17px;font-weight:900;color:#dc2626;font-family:'Courier New',monospace;letter-spacing:1px;text-shadow:0 0 6px rgba(255,68,68,0.5);">—</div>
        </div>
      </div>
      <div class="kpi-sub">Afastado = SIM</div>
    </div>

    <div class="kpi gray" onclick="abrirKpiModal('desligados')" title="Ver motoristas desligados (não contam nos demais indicadores)">
      <div class="kpi-lbl" style="color:#475569;">Motoristas Desligados</div>
      <div style="display:flex;align-items:flex-end;justify-content:space-between;">
        <div class="kpi-val" id="kpiDesligados">—</div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;padding-bottom:4px;gap:1px;">
          <div style="font-size:10px;color:#475569;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8;">do total geral</div>
          <div id="kpiDesligadosPct" style="font-size:17px;font-weight:900;color:#475569;font-family:'Courier New',monospace;letter-spacing:1px;">—</div>
        </div>
      </div>
      <div class="kpi-sub" style="color:#475569;">Desligado = SIM · Fora dos demais indicadores</div>
    </div>

    <div class="kpi teal" onclick="abrirKpiModal('telCorp')" title="Ver motoristas com celular corporativo">
      <div class="kpi-lbl">Celulares Corporativos</div>
      <div style="display:flex;align-items:flex-end;justify-content:space-between;">
        <div class="kpi-val" id="kpiTelCorp">—</div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;padding-bottom:4px;gap:1px;">
          <div style="font-size:10px;color:#0e9cc0;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8;">com linha SIM</div>
          <div id="kpiTelCorpPct" style="font-size:17px;font-weight:900;color:#0e9cc0;font-family:'Courier New',monospace;letter-spacing:1px;text-shadow:0 0 6px rgba(14,156,192,0.5);">—</div>
        </div>
      </div>
      <div class="kpi-sub">Telefone Corporativo = SIM</div>
    </div>

    <div class="kpi red" onclick="abrirKpiModal('excesso')" title="Ver motoristas com excesso de velocidade">
      <div class="kpi-lbl">Excesso Velocidade</div>
      <div style="display:flex;align-items:flex-end;justify-content:space-between;">
        <div class="kpi-val" id="kpiExcesso">—</div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;padding-bottom:4px;gap:1px;">
          <div style="font-size:10px;color:#dc2626;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8;">motoristas</div>
          <div id="kpiExcessoMot" style="font-size:17px;font-weight:900;color:#dc2626;font-family:'Courier New',monospace;letter-spacing:1px;text-shadow:0 0 6px rgba(255,68,68,0.5);">—</div>
        </div>
      </div>
      <div class="kpi-sub">Total Ocorrências</div>
    </div>

    <div class="kpi red" onclick="abrirKpiModal('multas')" title="Ver motoristas com multas">
      <div class="kpi-lbl">Total Multas</div>
      <div style="display:flex;align-items:flex-end;justify-content:space-between;">
        <div class="kpi-val" id="kpiMultas">—</div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;padding-bottom:4px;gap:1px;">
          <div style="font-size:10px;color:#dc2626;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8;">motoristas</div>
          <div id="kpiMultasMot" style="font-size:17px;font-weight:900;color:#dc2626;font-family:'Courier New',monospace;letter-spacing:1px;text-shadow:0 0 6px rgba(255,68,68,0.5);">—</div>
        </div>
      </div>
      <div class="kpi-sub">Total Ocorrências</div>
    </div>

    <div class="kpi red" onclick="abrirKpiModal('acidentes')" title="Ver motoristas com acidentes">
      <div class="kpi-lbl">Total Acidentes</div>
      <div style="display:flex;align-items:flex-end;justify-content:space-between;">
        <div class="kpi-val" id="kpiAcidentes">—</div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;padding-bottom:4px;gap:1px;">
          <div style="font-size:10px;color:#dc2626;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8;">motoristas</div>
          <div id="kpiAcidentesMot" style="font-size:17px;font-weight:900;color:#dc2626;font-family:'Courier New',monospace;letter-spacing:1px;text-shadow:0 0 6px rgba(255,68,68,0.5);">—</div>
        </div>
      </div>
      <div class="kpi-sub">Total Ocorrências</div>
    </div>

    <div class="kpi blue" onclick="abrirKpiModal('total')" title="Ver todos os motoristas">
      <div class="kpi-lbl">Total Motoristas</div>
      <div style="display:flex;align-items:flex-end;justify-content:space-between;">
        <div class="kpi-val" id="kpiTotal">—</div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;padding-bottom:4px;gap:1px;">
          <div style="font-size:10px;color:#1a4fa0;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8;">cadastrados</div>
          <div id="kpiTotalAnual" style="font-size:17px;font-weight:900;color:#1a4fa0;font-family:'Courier New',monospace;letter-spacing:1px;text-shadow:0 0 6px rgba(59,125,216,0.5);">—</div>
        </div>
      </div>
      <div class="kpi-sub">Todas as filiais</div>
    </div>

    <div class="kpi indigo" onclick="abrirOrganogramaModal()" title="Ver e editar o organograma da equipe">
      <div class="kpi-lbl">Gestão Organograma</div>
      <div class="kpi-val" style="font-size:34px;"><i class="fa-solid fa-sitemap"></i></div>
      <div class="kpi-sub">Equipe &amp; Estrutura</div>
    </div>
  </div>

 <div style="display:grid;grid-template-columns:2fr 1fr;gap:12px;margin-bottom:12px;width:100%;min-width:0;" class="charts-row">
    <div class="panel" style="display:flex;flex-direction:column;margin-bottom:0;min-width:0;overflow:hidden;">
      <div class="sec-title">DSS por sessão — __ANO__ (registros realizados)</div>
      <div style="display:flex;flex-wrap:wrap;gap:10px;margin-bottom:10px;align-items:center;">
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#16a34a;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#16a34a;"></span>100% adesão</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#3b7dd8;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#3b7dd8;"></span>+50% adesão</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#d97706;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#d97706;"></span>Menos de 50%</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#dc2626;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#dc2626;"></span>Sem registro</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#1a3a6b;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#1a3a6b;border:1px solid #c4d0e4;"></span>Semana atual</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#9aaabb;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:rgba(180,200,230,0.5);"></span>Futuro</span>
      </div>
      <div class="chart-wrap" id="dssChartWrap" style="height:260px;transition:height 0.4s ease;"><canvas id="dssChart"></canvas></div>
    </div>
    <div class="panel" style="display:flex;flex-direction:column;margin-bottom:0;min-width:0;overflow:hidden;">
      <div class="sec-title">Motoristas por filial — total e pendências DSS</div>
      <div class="leg">
        <div class="leg-item"><span class="leg-sq" style="background:#22cc88"></span>Com DSS</div>
        <div class="leg-item"><span class="leg-sq" style="background:#ff4444"></span>Sem DSS</div>
      </div>
      <div class="chart-wrap" style="flex:1;min-height:160px;"><canvas id="filialChart"></canvas></div>
    </div>
     <div class="panel" style="display:flex;flex-direction:column;margin-bottom:0;min-width:0;overflow:hidden;">
      <div class="sec-title">Status geral anual — indicadores por mês</div>
      <div style="display:flex;flex-wrap:wrap;gap:10px;margin-bottom:10px;align-items:center;">
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#16a34a;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#16a34a;"></span>100% adesão</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#3b7dd8;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#3b7dd8;"></span>+50% adesão</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#d97706;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#d97706;"></span>Menos de 50%</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#dc2626;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#dc2626;"></span>Sem registro</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#1a3a6b;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:#1a3a6b;border:1px solid #c4d0e4;"></span>Mês atual</span>
        <span style="display:flex;align-items:center;gap:5px;font-size:10px;font-weight:700;color:#9aaabb;"><span style="display:inline-block;width:22px;height:8px;border-radius:2px;background:rgba(180,200,230,0.5);"></span>Futuro</span>
      </div>
      <div style="position:relative;width:100%;flex:1;min-height:160px;"><canvas id="statusAnualChart" role="img" aria-label="Gráfico de barras com DSS realizados, acidentes, multas e excessos de velocidade por mês"></canvas></div>
    </div>
    <div class="panel" style="display:flex;flex-direction:column;margin-bottom:0;min-width:0;overflow:hidden;">
      <div class="sec-title">DSS anual por filial — sessões realizadas</div>
      <div class="leg">
        <div class="leg-item"><span class="leg-sq" style="background:#16a34a"></span>100% adesão</div>
        <div class="leg-item"><span class="leg-sq" style="background:#3b7dd8"></span>+50%</div>
        <div class="leg-item"><span class="leg-sq" style="background:#d97706"></span>&lt;50%</div>
        <div class="leg-item"><span class="leg-sq" style="background:#dc2626"></span>Sem DSS</div>
      </div>
      <div class="chart-wrap" style="flex:1;min-height:160px;"><canvas id="filialAnualChart"></canvas></div>
    </div>
  </div>
    <div style="margin-bottom:8px;display:flex;flex-wrap:wrap;gap:8px">
      <div class="leg-item"><span class="leg-sq" style="background:#22cc88"></span>Reciclagem ok</div>
      <div class="leg-item"><span class="leg-sq" style="background:#4a9eff"></span>Simulador ok</div>
      <div class="leg-item"><span class="leg-sq" style="background:#ff4444"></span>Acidentes / Multas</div>
      <div class="leg-item"><span class="leg-sq" style="background:#ffaa00"></span>Exc. velocidade</div>
    </div>
    <div class="filial-grid" id="filialGrid">
      <div class="empty-state" style="grid-column:1/-1">
        <i class="fa-solid fa-building-user"></i>
        <p>Nenhum motorista cadastrado ainda.<br>Use o formulário acima para inserir o primeiro condutor.</p>
      </div>
    </div>
  </div>
</div>

<!-- Modal KPI -->
<div class="kpi-modal-overlay" id="kpiModal">
  <div class="kpi-modal-box">
    <div class="kpi-modal-head">
      <div class="kpi-modal-head-left">
        <div class="kpi-modal-icon" id="kpiModalIcon"></div>
        <div>
          <div class="kpi-modal-label" id="kpiModalLabel">—</div>
          <div class="kpi-modal-count" id="kpiModalCount">0 motoristas</div>
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:8px;">
        <button id="btnVoltarCursos" style="display:none;background:transparent;color:#3b7dd8;border:1.5px solid #3b7dd8;padding:5px 14px;font-size:10px;font-weight:800;letter-spacing:1px;text-transform:uppercase;border-radius:5px;cursor:pointer;align-items:center;gap:6px;"><i class="fa-solid fa-arrow-left"></i> Voltar</button>
        <button id="btnBaixarPdfPendentes" style="display:none;background:transparent;color:#22cc88;border:1.5px solid #22cc88;padding:5px 14px;font-size:10px;font-weight:800;letter-spacing:1px;text-transform:uppercase;border-radius:5px;cursor:pointer;align-items:center;gap:6px;" onmouseover="this.style.color='#ff4444';this.style.borderColor='#ff4444'" onmouseout="this.style.color='#22cc88';this.style.borderColor='#22cc88'"><i class="fa-solid fa-file-pdf"></i> Baixar PDF</button>
        <button class="kpi-modal-close" onclick="fecharKpiModal()"><i class="fa-solid fa-xmark"></i></button>
      </div>
    </div>
    <div class="kpi-modal-search">
      <input type="text" id="kpiSearchInput" placeholder="🔍  Buscar por nome, CPF ou filial…" oninput="filtrarCardsKpi()">
    </div>
    <div class="kpi-mes-filtro" id="kpiMesFiltro"></div>
    <div class="kpi-cards-grid" id="kpiCardsGrid"></div>
  </div>
</div>

<!-- Modal Filial -->
<div class="modal-overlay" id="filialModal">
  <div class="modal-box">
    <div class="modal-header">
      <div class="modal-title"><i class="fa-solid fa-location-dot" style="color:#4a9eff"></i>&nbsp;Filial: <span id="mUnidadeName">...</span></div>
      <button class="btn-close" onclick="fecharJanelaFilial()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <div class="modal-split">
      <div class="modal-sidebar" id="filialSidebar">
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('todos')"><div class="m-lbl">Total de Motoristas</div><div class="m-val" id="mTotalDrivers">0</div></div>
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('dss')"><div class="m-lbl">DSS Realizados (Ano)</div><div class="m-val" id="mWithDss" style="color:#22cc88">0</div></div>
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('reciclagem')"><div class="m-lbl">Reciclagem OK</div><div class="m-val" id="mRecOk" style="color:#16a34a">0</div></div>
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('simulador')"><div class="m-lbl">Simulador OK</div><div class="m-val" id="mSimOk" style="color:#3b7dd8">0</div></div>
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('excesso')"><div class="m-lbl">Excesso Velocidade</div><div class="m-val" id="mExcVel" style="color:#dc2626">0</div></div>
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('multas')"><div class="m-lbl">Total Multas</div><div class="m-val" id="mMultas" style="color:#d97706">0</div></div>
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('acidentes')"><div class="m-lbl">Total Acidentes</div><div class="m-val" id="mAcidentes" style="color:#dc2626">0</div></div>
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('examePeriodico')"><div class="m-lbl">Exame Periódico OK</div><div class="m-val" id="mExamePerOk" style="color:#7c3aed">0</div></div>
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('exameToxicologico')"><div class="m-lbl">Exame Toxicológico OK</div><div class="m-val" id="mExameToxOk" style="color:#a78bfa">0</div></div>
        <div class="modal-kpi-card" onclick="filtrarFilialPorIndicador('telefoneCorporativo')"><div class="m-lbl">Telefone Corporativo OK</div><div class="m-val" id="mTelCorpOk" style="color:#0e9cc0">0</div></div>
      </div>
      <div class="modal-main">
        <div class="filial-mobile-backbar" id="filialMobileBackbar">
          <button class="btn-voltar-mobile" onclick="voltarSidebarFilialMobile()"><i class="fa-solid fa-arrow-left"></i> Voltar</button>
          <span class="filial-mobile-titulo" id="filialMobileTitulo"></span>
        </div>
        <div style="padding:10px 14px;border-bottom:1px solid #dde6f4;background:#fff;flex-shrink:0;">
          <input type="text" id="filialSearchInput" placeholder="🔍  Buscar por nome ou CPF…" oninput="filtrarTabelaFilial()" style="width:100%;background:#f4f7fc;border:1.5px solid #c4d0e4;color:#1a2a44;padding:7px 12px;border-radius:6px;font-size:13px;outline:none;">
        </div>
        <div class="table-container" id="filialTableContainer">
          <table class="m-table">
            <thead>
              <tr>
                <th>CPF / Motorista (Clique para abrir a ficha)</th>
                <th style="text-align:center">DSS Ano</th>
                <th>Reciclagem</th><th>Simulador</th>
                <th style="text-align:center">Excesso Vel.</th>
                <th style="text-align:center">Multas</th>
                <th style="text-align:center">Acidentes</th>
                <th>Exame Periódico</th><th>Exame Toxicológico</th><th>Tel. Corporativo</th>
              </tr>
            <tbody id="mDriversTableBody"></tbody>
          </table>
        </div>
        <div class="filial-mobile-list" id="filialMobileList"></div>
      </div>
    </div>
  </div>
</div>

<!-- Modal Ficha Individual -->
<div class="modal-overlay" id="driverModal" style="z-index:10000;background:rgba(0,0,0,0.55);backdrop-filter:blur(4px);padding:0;">
  <div class="modal-box" style="max-width:none;width:100%;height:100%;border-radius:0;border:none;background:#ffffff;display:flex;flex-direction:column;">
    <div class="modal-header" style="border-color:#d0d8e8;background:#f0f4fa;flex-shrink:0;">
      <div class="modal-title" style="font-size:13px;color:#1a3a6b;"><i class="fa-solid fa-id-card" style="color:#1a7a4a"></i> <span style="color:#1a3a6b;">FICHA INDIVIDUAL DO CONDUTOR — HISTÓRICO E COMPLIANCE</span></div>
      <div style="display:flex;align-items:center;gap:8px;">
        <button id="btnConfirmarFicha" onclick="confirmarEdicaoFicha()" style="background:#1a5c2a;color:#ffffff;border:1px solid #14481f;width:auto;padding:0 16px;height:28px;border-radius:4px;font-weight:700;font-size:11px;cursor:pointer;display:flex;align-items:center;gap:6px;">
          <i class="fa-solid fa-check"></i> Confirmar Alterações
        </button>
        <button id="btnVoltarFicha" onclick="voltarPaginaAnterior()" style="display:none;background:#7a1a1a;color:#ffffff;border:1px solid #5c1212;width:auto;padding:0 14px;height:28px;border-radius:4px;font-weight:700;font-size:11px;cursor:pointer;align-items:center;gap:6px;">
          <i class="fa-solid fa-arrow-left"></i> Voltar
        </button>
        <button id="btnFecharFicha" onclick="fecharJanelaDriver()" style="background:#7a1a1a;color:#ffffff;border:1px solid #5c1212;width:28px;height:28px;border-radius:50%;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:13px;"><i class="fa-solid fa-xmark"></i></button>
      </div>
    </div>
    <div class="driver-profile-grid" id="driverProfileContent" style="flex:1;overflow-y:auto;"></div>
  </div>
</div>

<!-- Modal Organograma -->
<div class="modal-overlay" id="organogramaModal" style="z-index:10000;background:rgba(0,0,0,0.55);backdrop-filter:blur(4px);padding:0;">
  <div class="modal-box" style="max-width:none;width:100%;height:100%;border-radius:0;border:none;background:#f0f4fa;display:flex;flex-direction:column;position:relative;">
    <div class="modal-header" style="border-color:#d0d8e8;background:#fff;flex-shrink:0;">
      <div class="modal-title" style="font-size:13px;color:#1a3a6b;"><i class="fa-solid fa-sitemap" style="color:#4338ca"></i> <span style="color:#1a3a6b;">GESTÃO ORGANOGRAMA — ESTRUTURA DA EQUIPE</span></div>
      <div style="display:flex;align-items:center;gap:8px;">
        <button class="org-save-btn" onclick="salvarOrganogramaAPI()">
          <i class="fa-solid fa-check"></i> Salvar Organograma
        </button>
        <div class="org-add-btn-wrap">
          <button type="button" class="org-add-btn" id="orgQuickMenuBtn" onclick="toggleOrgQuickMenu()" title="Adicionar colaborador">
            <i class="fa-solid fa-user-plus"></i> Adicionar Colaborador
          </button>
          <div class="org-quick-menu-panel" id="orgQuickMenuPanel">
            <div class="org-quick-menu-title"><i class="fa-solid fa-user-plus"></i> Adicionar em…</div>
            <div id="orgQuickMenuList"></div>
          </div>
        </div>
        <button onclick="gerarOrganogramaPdf()" style="background:transparent;color:#1a4fa0;border:1.5px solid #1a4fa0;width:auto;padding:0 16px;height:36px;border-radius:6px;font-weight:700;font-size:12px;cursor:pointer;display:flex;align-items:center;gap:6px;transition:background .2s,color .2s;" onmouseover="this.style.background='#1a4fa0';this.style.color='#fff'" onmouseout="this.style.background='transparent';this.style.color='#1a4fa0'">
          <i class="fa-solid fa-file-pdf"></i> Baixar PDF
        </button>
        <button class="btn-close" onclick="fecharOrganogramaModal()"><i class="fa-solid fa-xmark"></i></button>
      </div>
    </div>
    <div id="orgWrapOuter" style="position:relative;width:100%;flex:1;overflow:auto;background:#e9edf3;">
      <div id="orgWrapInner" style="position:absolute;top:0;left:0;transform-origin:top left;background:#fff;box-shadow:0 4px 24px rgba(20,50,120,0.12);"></div>
    </div>
  </div>
</div>

<input type="file" id="hiddenPhotoInput" accept="image/*" style="display:none;" onchange="processarFotoCarregada(this)">

<script>
const ACCESS_TOKEN = '{_ACCESS_TOKEN}';
const SHEET_ID_JS  = '{SHEET_ID}';
const SHEET_NAME_JS= '{SHEET_NAME}';
const SHEET_NAME_ORG_JS = '{SHEET_NAME_ORG}';
const SHEETS_BASE  = 'https://sheets.googleapis.com/v4/spreadsheets';
const DADOS_INICIAIS = {json.dumps(ler_todos_motoristas(), ensure_ascii=False)};
const DADOS_ORG_INICIAIS = {json.dumps(ler_organograma(), ensure_ascii=False)};
const CREDENCIAIS     = {json.dumps(CREDENCIAIS_LOGIN, ensure_ascii=False)};
const MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
               "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

const AVATAR_PADRAO = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDABALDA4MChAODQ4SERATGCgaGBYWGDEjJR0oOjM9PDkzODdASFxOQERXRTc4UG1RV19iZ2hnPk1xeXBkeFxlZ2P/2wBDARESEhgVGC8aGi9jQjhCY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2P/wAARCABPAFADASIAAhEBAxEB/8QAGwAAAQUBAQAAAAAAAAAAAAAABgADBAUHAQL/xAA2EAACAQMBBQUFBwUBAAAAAAABAgMABAURBhIhMVEiQWFxkTJSgaHREyNicrHB4QcUFRZCY//EABkBAQEBAQEBAAAAAAAAAAAAAAQDAQIFAP/EACMRAAICAgEEAwEBAAAAAAAAAAABAgMRIQQSMUJREyJSQSP/2gAMAwEAAhEDEQA/AD+lSqFk8lBjLVp524clUc2PQV8lnSMbxtkmWRIkLyMERRqWJ0AqkvNqrKAlYA07dV4L60KZLL3WVl3pW3YgezEp4D61FQUyvjrvIFbyX2iEr7WXLHsW8Sj8RJr1HtVda9uCMjw1FDwp1RV/hr9A5cmxf0LLTaW3lIWeNoj15iriOaOeMPEyuh5FTrQAqk6ADUmpqPd4iVDrul13jGeRHiKhOiPiVq5kvNZQb0qgY3Ix38O+nZYe2h5ip9EaaeGenGSkso8SOsaM7sFVRqSe4Vmebyz5bINLqRCvZiXoOvmaLts702uGMSnt3Dbnw5n6fGs+Q0iiPkHvl4j6U+lR0NPIaajz5EhamWdtNdzCKBCzfIedP4jDT5EiQ6xwd7kc/KjGzsoLGERQIFHee8+ZqNt6jpdylXFdjzLSImMxEViA76SS6e0RwHlVJtK2uTIH/KAfv+9F9BOfk3svP4aD5VGhuVmWW5UIwqUYr+kayvJLC6WaMnh7S+8OlHdvOlxCksZ1RhqDWcsdaKtkrsyWstux1MTar5H+a75MNdRxw7Gn0Mqv6gOTPZx9wVm/ShFTxou2/jP21lJ3FWX9KGsdjrrJXAhtYyx72PsqOpNfVNKCZexNzaORBnYKilmJ0AA1Jovwmy+gW4yI481i+v0qzwmz9tikDkCW5I4yEcvAdKualZe3qJ1XQluRxVCgBQABwAHdXqlSo4o5Wf5iTfy10f8A0I9OFaA3KsxupvtLqaT3nLeppXFW2wXM3FIRNXeyEp/ykqdzRE+hFDpaiDYtS2UmfuWEj1I+lJuf0YWhf6IIs3h4sxbxxyOU3JA28vPTvFeEuMPhIP7cTQwKvNQdWJ6nTjrVqwDAgjUHnWY7RYl8VkGUAmCQlom8OnmKBWur6tnpTfTtILLjbTGRtpEJ5fFU0HzIqP8A7zak8LObT8woFrtIVMSDukaDDtnj5CA8c8WveVBA9DVxaZGzvV1trhJNOYB4j4c6ycNTsNxJBIskLsjryZToRWOiL7Gq+S7mtTtuQSP7qk/KsqL6k0VY/acXeMube7IW5WFir8hJwPzoO362iLjnJze1PGB0tRtsVZmKwkunGhnbs/lH860JYjHy5W9S3jBC85H91a06CFLeFIYl3URQqjoKy+euk3j176h2oeRx8GStHt7hdUbkRzU9RUylRE8C2smX5nAXeKkJZTJb69mVRw+PQ1U1sbqrqVcAg8CCNQao7/ZHG3TF0Vrdjz+zPD0pMb/0HlT+TOKVGE2wkgP3F8pHR49P0Nci2EmJ++vo1H4Iyf3FU+WHsn8UvQIVY4jD3mWlAgQiMHRpWHZX6nwoystjsdbMHm37hh3OdF9BRBFGkSBI1VVHJVGgFTlf+SkafZCxGJgxNqIYBqTxdzzY1YVyu0ZvO2ISS0j/2Q==';

let motoristasDB         = DADOS_INICIAIS;
function listaAtiva(){{ return motoristasDB.filter(m => m.desligado !== 'SIM'); }}
function listaDesligados(){{ return motoristasDB.filter(m => m.desligado === 'SIM'); }}
let DADOS_ORG            = DADOS_ORG_INICIAIS;
let orgZoomFactor        = 0.85;
let dssChartInstance      = null;

// Garante que o setor "Motoristas" exista, mesmo que a planilha já tivesse dados salvos antes dele existir.
(function garantirSetorMotoristas(){{
  if(!DADOS_ORG || !Array.isArray(DADOS_ORG.setores)) return;
  const jaExiste = DADOS_ORG.setores.some(function(s){{
    return (s.titulo||[]).join(' ').trim().toUpperCase() === 'MOTORISTAS';
  }});
  if(!jaExiste){{
    DADOS_ORG.setores.push({{ titulo:['MOTORISTAS'], icone:'wheel', pessoas:[['','']] }});
  }}
}})();
let filialChartInstance   = null;
let filialAnualChartInst  = null;
let motoristaEmEdicaoCpf = null;
let fotoTemporariaBase64 = null;
let filialModalAtiva     = null;
let fichaOrigemModal     = null;
let houveEdicaoNaoSalva  = false;
let autoRefreshInterval  = null;
const AUTO_REFRESH_MS    = 30000; // intervalo padrão: 30 segundos

function mostrarSpinner(show){{ document.getElementById('spinnerOverlay').classList.toggle('show', show); }}

function toast(msg, tipo='ok'){{
  const el  = document.getElementById('toastMsg');
  const txt = document.getElementById('toastText');
  txt.textContent = msg;
  el.className = `toast ${{tipo}} show`;
  setTimeout(() => el.classList.remove('show'), 3500);
}}

function motoristasParaLinhas(lista){{
  return lista.map(m => {{
    const row = [
      m.cpf||'', m.nome||'', m.filial||'', m.telefone||'', m.email||'', m.foto||'',
      m.reciclagem||'PENDENTE', m.simulador||'PENDENTE',
      m.excesso||0, m.multas||0, m.acidentes||0,
      m.obsAcidente||'', m.obsMultas||'', m.obsGerais||'',
      m.obsReciclagem||'', m.obsSimulador||'',
      m.cnh||'', m.validadeCnh||'', m.admissao||''
    ];
    const meses = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                   "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
    meses.forEach(mes => {{
      const sems = m.dssAnual?.[mes] || [false,false,false,false];
      for(let s=0;s<4;s++) row.push(sems[s] ? 1 : 0);
    }});
    row.push(
      m.examePeriodico||'', m.exameToxicologico||'',
      m.pontuacaoCnh||0, m.vencimentoCnhMopp||'',
      m.entregaUniforme||'PENDENTE',
      m.telefoneCorporativo||'NÃO', m.numeroLinha||'',
      m.modelo||'', m.imei||'',
      m.reciclagemData||'', m.reciclagemValidadeMeses||0,
      m.simuladorData||'', m.simuladorValidadeMeses||0,
      m.examePeriodicoValidadeMeses||0, m.exameToxicologicoValidadeMeses||0,
      m.gestime||'PENDENTE', m.obsGestime||'',
      m.gestimeData||'', m.gestimeValidadeMeses||0,
      m.afastado||'NÃO', m.obsAfastado||'',
      m.desligado||'NÃO', m.obsDesligamento||''
    );
    return row;
  }});
}}

function _comprimirBase64(base64, maxPx, qualidade){{
  return new Promise(resolve => {{
    if(!base64 || !base64.startsWith('data:image')){{ resolve(base64 || ''); return; }}
    const img = new Image();
    img.onload = () => {{
      const canvas = document.createElement('canvas');
      const escala = Math.min(1, maxPx / Math.max(img.width, img.height));
      canvas.width  = Math.round(img.width  * escala);
      canvas.height = Math.round(img.height * escala);
      canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);

      // Tenta qualidade pedida, se passar de 35.000 chars reduz mais
      let resultado = canvas.toDataURL('image/jpeg', qualidade);
      if(resultado.length > 35000){{
        resultado = canvas.toDataURL('image/jpeg', 0.3);
      }}
      if(resultado.length > 35000){{
        // Última tentativa: reduz canvas para metade
        const c2 = document.createElement('canvas');
        c2.width  = Math.round(canvas.width  * 0.5);
        c2.height = Math.round(canvas.height * 0.5);
        c2.getContext('2d').drawImage(canvas, 0, 0, c2.width, c2.height);
        resultado = c2.toDataURL('image/jpeg', 0.25);
      }}
      // Se ainda assim passar, descarta a foto para não corromper o banco
      if(resultado.length > 35000){{
        resolve('');
        toast('Foto muito grande mesmo após compressão. Use uma imagem menor.', 'erro');
        return;
      }}
      resolve(resultado);
    }};
    img.onerror = () => resolve('');
    img.src = base64;
  }});
}}

function _sessaoExpirouRecarregar(){{
  toast('Sessão de acesso ao Google Sheets expirou. Recarregando a página...', 'erro');
  setTimeout(() => window.location.reload(), 2000);
}}

async function salvarTodosNaSheetsAPI(lista){{
  const auth = `Bearer ${{ACCESS_TOKEN}}`;
  const rangeBase = `${{SHEET_NAME_JS}}!A2:ZZ`;

  // Comprime fotos antes de montar o payload — evita estourar limite da API
  const listaSegura = await Promise.all(lista.map(async m => {{
    const fotoComprimida = m.foto ? await _comprimirBase64(m.foto, 80, 0.5) : '';
    return {{ ...m, foto: fotoComprimida }};
  }}));

  // Verifica tamanho antes de qualquer operação destrutiva
  const payload = JSON.stringify({{ values: motoristasParaLinhas(listaSegura) }});
  const tamanhoMB = new Blob([payload]).size / 1024 / 1024;
  if(tamanhoMB > 8){{
    return {{ ok: false, erro: `Foto muito grande (${{tamanhoMB.toFixed(1)}} MB). Reduza a imagem.` }};
  }}

  // 1. Limpa — só executa depois que o payload foi validado
  const clearResp = await fetch(
    `${{SHEETS_BASE}}/${{SHEET_ID_JS}}/values/${{encodeURIComponent(rangeBase)}}:clear`,
    {{ method: 'POST', headers: {{ 'Authorization': auth }} }}
  );
  if(!clearResp.ok){{
    const err = await clearResp.text();
    if(clearResp.status === 401 || err.includes('ACCESS_TOKEN_EXPIRED')){{
      _sessaoExpirouRecarregar();
      return {{ ok: false, erro: 'Sessão expirada. A página será recarregada automaticamente.' }};
    }}
    return {{ ok: false, erro: 'Erro ao limpar planilha: ' + err }};
  }}

  if(listaSegura.length === 0) return {{ ok: true }};

  // 2. Escreve as novas linhas
  const resp = await fetch(
    `${{SHEETS_BASE}}/${{SHEET_ID_JS}}/values/${{encodeURIComponent(rangeBase)}}?valueInputOption=RAW`,
    {{
      method: 'PUT',
      headers: {{ 'Authorization': auth, 'Content-Type': 'application/json' }},
      body: payload
    }}
  );
  if(!resp.ok){{
    const err = await resp.text();
    if(resp.status === 401 || err.includes('ACCESS_TOKEN_EXPIRED')){{
      _sessaoExpirouRecarregar();
      return {{ ok: false, erro: 'Sessão expirada. A página será recarregada automaticamente.' }};
    }}
    return {{ ok: false, erro: 'Erro ao salvar dados: ' + err }};
  }}
  return {{ ok: true }};
}}
// ── KPI Modal ──
const KPI_CONFIG = {{
  total:    {{ label:'Todos os Motoristas',         icon:'fa-users',             cor:'#7ab8ff', bg:'rgba(74,159,255,0.15)',  filtro: m => true, dssModal:true }},
  comDss:   {{ label:'Com DSS Ok',                  icon:'fa-circle-check',      cor:'#22cc88', bg:'rgba(34,204,136,0.15)', filtro: (m,mes) => dssOkNoMes(m, mes), dssModal:true }},
  semDss:   {{ label:'Pendentes DSS',               icon:'fa-clock',             cor:'#ffaa00', bg:'rgba(255,170,0,0.15)',  filtro: (m,mes) => !dssOkNoMes(m, mes), dssModal:true }},
  excesso:  {{ label:'Com Excesso de Velocidade',   icon:'fa-gauge-high',        cor:'#ff6666', bg:'rgba(255,68,68,0.15)',  filtro: m => Math.max(0, parseInt(m.excesso)   || 0) > 0 }},
  multas:   {{ label:'Com Multas Registradas',      icon:'fa-file-circle-xmark', cor:'#ff6666', bg:'rgba(255,68,68,0.15)',  filtro: m => Math.max(0, parseInt(m.multas)    || 0) > 0 }},
  acidentes:{{ label:'Com Acidentes Registrados',   icon:'fa-car-burst',         cor:'#ff6666', bg:'rgba(255,68,68,0.15)',  filtro: m => Math.max(0, parseInt(m.acidentes) || 0) > 0 }},
  reciclagemOk: {{ label:'Reciclagem OK',            icon:'fa-recycle',           cor:'#16a34a', bg:'rgba(22,163,74,0.15)',  filtro: m => reciclagemStatus(m) === 'OK' }},
  reciclagemPend: {{ label:'Reciclagem Pendente',    icon:'fa-clock',             cor:'#d97706', bg:'rgba(217,119,6,0.15)',  filtro: m => reciclagemStatus(m) === 'PENDENTE' }},
  simuladorOk: {{ label:'Simulador SEST SENAT OK',      icon:'fa-car-side',       cor:'#16a34a', bg:'rgba(22,163,74,0.15)',  filtro: m => simuladorStatus(m) === 'OK' }},
  simuladorPend: {{ label:'Simulador SEST SENAT Pendente', icon:'fa-clock',       cor:'#d97706', bg:'rgba(217,119,6,0.15)',  filtro: m => simuladorStatus(m) === 'PENDENTE' }},
  gestimeOk: {{ label:'Gestime OK',                     icon:'fa-clipboard-check', cor:'#16a34a', bg:'rgba(22,163,74,0.15)',  filtro: m => gestimeStatus(m) === 'OK' }},
  gestimePend: {{ label:'Gestime Pendente',             icon:'fa-clock',            cor:'#d97706', bg:'rgba(217,119,6,0.15)',  filtro: m => gestimeStatus(m) === 'PENDENTE' }},
  dssTodos:        {{ label:'DSS Mensal — Status Geral',       icon:'fa-calendar-check',  cor:'#16a34a', bg:'rgba(22,163,74,0.15)', filtro: m => true, dssModal:true }},
  reciclagemTodos: {{ label:'Reciclagem — Status Geral',       icon:'fa-recycle',         cor:'#16a34a', bg:'rgba(22,163,74,0.15)', filtro: m => true }},
  simuladorTodos:  {{ label:'Simulador SEST SENAT — Status Geral', icon:'fa-car-side',    cor:'#16a34a', bg:'rgba(22,163,74,0.15)', filtro: m => true }},
  gestimeTodos:    {{ label:'Gestime — Status Geral',          icon:'fa-clipboard-check', cor:'#16a34a', bg:'rgba(22,163,74,0.15)', filtro: m => true }},
  telCorp:  {{ label:'Celulares Corporativos',      icon:'fa-mobile-screen-button', cor:'#0eb8e0', bg:'rgba(14,156,192,0.15)', filtro: m => m.telefoneCorporativo === 'SIM' }},
  afastados:{{ label:'Motoristas Afastados',        icon:'fa-user-slash',        cor:'#dc2626', bg:'rgba(220,38,38,0.15)',  filtro: m => m.afastado === 'SIM' }},
  desligados:{{ label:'Motoristas Desligados',      icon:'fa-user-xmark',        cor:'#64748b', bg:'rgba(100,116,139,0.15)', filtro: m => m.desligado === 'SIM' }},
  prontuario:{{ label:'Prontuário — Exames & Complementares', icon:'fa-file-medical', cor:'#a78bfa', bg:'rgba(124,58,237,0.15)', filtro: m => true }},
}};

let kpiListaAtual = [];
let kpiMesAtual   = null;
let kpiTipoAtual  = null;
let kpiOrigemCursos = false;

// ── Vencimentos (Alerta 30 dias / Vencidos) ──
const VENCIMENTO_CATEGORIAS = {{
  reciclagem:        {{ label:'Reciclagem',            icon:'fa-recycle',           dataCampo:'reciclagemData',        mesesCampo:'reciclagemValidadeMeses' }},
  simulador:          {{ label:'Simulador SEST SENAT',  icon:'fa-car-side',          dataCampo:'simuladorData',         mesesCampo:'simuladorValidadeMeses' }},
  gestime:             {{ label:'Gestime',               icon:'fa-clipboard-check',   dataCampo:'gestimeData',           mesesCampo:'gestimeValidadeMeses' }},
  examePeriodico:      {{ label:'Exame Periódico',       icon:'fa-stethoscope',       dataCampo:'examePeriodico',        mesesCampo:'examePeriodicoValidadeMeses' }},
  exameToxicologico:   {{ label:'Exame Toxicológico',    icon:'fa-vial',              dataCampo:'exameToxicologico',     mesesCampo:'exameToxicologicoValidadeMeses' }},
  cnh:                 {{ label:'Validade CNH',          icon:'fa-id-card',           dataDireta:'validadeCnh' }},
  mopp:                {{ label:'Vencimento MOPP',       icon:'fa-id-card-clip',      dataDireta:'vencimentoCnhMopp' }},
}};

function diasParaVencerCategoria(m, catKey){{
  const cat = VENCIMENTO_CATEGORIAS[catKey];
  let venc;
  if(cat.dataDireta){{
    if(!m[cat.dataDireta]) return null;
    venc = new Date(m[cat.dataDireta] + 'T00:00:00');
  }} else {{
    venc = calcularVencimento(m[cat.dataCampo], m[cat.mesesCampo]);
  }}
  if(!venc || isNaN(venc)) return null;
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  return Math.floor((venc - hoje) / 86400000);
}}

function categoriaEmAlerta30(m, catKey){{ const d = diasParaVencerCategoria(m, catKey); return d !== null && d >= 0 && d <= 30; }}
function categoriaVencida(m, catKey){{ const d = diasParaVencerCategoria(m, catKey); return d !== null && d < 0; }}
function motoristaTemAlerta30(m){{ return Object.keys(VENCIMENTO_CATEGORIAS).some(k => categoriaEmAlerta30(m, k)); }}
function motoristaTemVencido(m){{ return Object.keys(VENCIMENTO_CATEGORIAS).some(k => categoriaVencida(m, k)); }}

let kpiVoltarFn = null; // função a chamar ao clicar em "Voltar"
let kpiCategoriaVencAtual = null;
let kpiTipoVencAtual = null;

function mesCorrente(){{ return MESES[new Date().getMonth()]; }}

function abrirVencimentoMenu(tipo){{ // tipo: 'alerta' | 'vencido'
  kpiOrigemCursos = false;
  kpiVoltarFn = null;
  const btnVoltarCursos = document.getElementById('btnVoltarCursos');
  if(btnVoltarCursos) btnVoltarCursos.style.display = 'none';

  kpiTipoAtual = 'vencimentoMenu';
  kpiMesAtual  = null;

  const isAlerta = tipo === 'alerta';
  const cor   = isAlerta ? '#eab308' : '#dc2626';
  const bg    = isAlerta ? 'rgba(234,179,8,0.15)' : 'rgba(220,38,38,0.15)';
  const icone = isAlerta ? 'fa-triangle-exclamation' : 'fa-circle-exclamation';
  const titulo= isAlerta ? 'Alertas de Vencimento (30 dias)' : 'Itens Vencidos';

  document.getElementById('kpiModalIcon').innerHTML = `<i class="fa-solid ${{icone}}" style="color:${{cor}}"></i>`;
  document.getElementById('kpiModalIcon').style.background = bg;
  document.getElementById('kpiModalLabel').textContent = titulo;
  document.getElementById('kpiModalCount').textContent = 'Selecione uma categoria';
  document.getElementById('kpiSearchInput').value = '';
  document.getElementById('kpiMesFiltro').classList.remove('visible');
  document.getElementById('kpiMesFiltro').innerHTML = '';
  const btnPdf = document.getElementById('btnBaixarPdfPendentes');
  if(btnPdf) btnPdf.style.display = 'none';

  const grid = document.getElementById('kpiCardsGrid');
  grid.innerHTML = Object.keys(VENCIMENTO_CATEGORIAS).map(catKey => {{
    const cat = VENCIMENTO_CATEGORIAS[catKey];
    const qtd = listaAtiva().filter(m => isAlerta ? categoriaEmAlerta30(m, catKey) : categoriaVencida(m, catKey)).length;
    return `<div class="categoria-glass-panel" onclick="abrirVencimentoCategoria('${{catKey}}','${{tipo}}')">
      <div class="cgp-header">
        <div class="cgp-icon" style="background:${{bg}};color:${{cor}}"><i class="fa-solid ${{cat.icon}}"></i></div>
        <div class="cgp-titulo">${{cat.label}}</div>
      </div>
      <div class="cgp-stats" style="grid-template-columns:1fr;">
        <div class="cgp-stat ${{isAlerta ? '' : ''}}" style="background:${{bg}};border-color:${{cor}}55;">
          <div class="cgp-stat-val" style="color:${{cor}}">${{qtd}}</div>
          <div class="cgp-stat-lbl" style="color:${{cor}}"><i class="fa-solid ${{icone}}"></i> ${{isAlerta ? 'a vencer' : 'vencidos'}}</div>
        </div>
      </div>
    </div>`;
  }}).join('');

  document.getElementById('kpiModal').classList.add('show');
}}

function abrirVencimentoCategoria(catKey, tipo){{
  const isAlerta = tipo === 'alerta';
  const cat = VENCIMENTO_CATEGORIAS[catKey];
  const cor   = isAlerta ? '#eab308' : '#dc2626';
  const bg    = isAlerta ? 'rgba(234,179,8,0.15)' : 'rgba(220,38,38,0.15)';
  const icone = isAlerta ? 'fa-triangle-exclamation' : 'fa-circle-exclamation';

  kpiTipoAtual = 'vencimentoCategoria';
  kpiCategoriaVencAtual = catKey;
  kpiTipoVencAtual = tipo;
  kpiMesAtual = null;

  kpiListaAtual = listaAtiva().filter(m => isAlerta ? categoriaEmAlerta30(m, catKey) : categoriaVencida(m, catKey));

  document.getElementById('kpiModalIcon').innerHTML = `<i class="fa-solid ${{cat.icon}}" style="color:${{cor}}"></i>`;
  document.getElementById('kpiModalIcon').style.background = bg;
  document.getElementById('kpiModalLabel').textContent = `${{cat.label}} — ${{isAlerta ? 'A vencer (30 dias)' : 'Vencidos'}}`;
  document.getElementById('kpiModalCount').textContent = `${{kpiListaAtual.length}} motorista${{kpiListaAtual.length !== 1 ? 's' : ''}}`;
  document.getElementById('kpiSearchInput').value = '';
  document.getElementById('kpiMesFiltro').classList.remove('visible');
  document.getElementById('kpiMesFiltro').innerHTML = '';
  const btnPdf = document.getElementById('btnBaixarPdfPendentes');
  if(btnPdf) btnPdf.style.display = 'none';

  kpiVoltarFn = () => abrirVencimentoMenu(tipo);
  const btnVoltarCursos = document.getElementById('btnVoltarCursos');
  if(btnVoltarCursos){{ btnVoltarCursos.style.display = 'flex'; btnVoltarCursos.onclick = () => kpiVoltarFn && kpiVoltarFn(); }}

  renderizarCardsKpi(kpiListaAtual);
  document.getElementById('kpiModal').classList.add('show');
}}

function abrirCategoriaCursos(tipo){{
  kpiOrigemCursos = true;
  abrirKpiModal(tipo);
}}

function abrirCursosMenu(){{
  kpiOrigemCursos = false;
  kpiVoltarFn = null;
  const btnVoltarCursos = document.getElementById('btnVoltarCursos');
  if(btnVoltarCursos) btnVoltarCursos.style.display = 'none';
  kpiTipoAtual = 'cursos';
  kpiMesAtual  = null;
  document.getElementById('kpiModalIcon').innerHTML = `<i class="fa-solid fa-graduation-cap" style="color:#16a34a"></i>`;
  document.getElementById('kpiModalIcon').style.background = 'rgba(22,163,74,0.15)';
  document.getElementById('kpiModalLabel').textContent = 'Treinamentos';
  document.getElementById('kpiModalCount').textContent = 'Selecione uma categoria';
  document.getElementById('kpiSearchInput').value = '';
  document.getElementById('kpiMesFiltro').classList.remove('visible');
  document.getElementById('kpiMesFiltro').innerHTML = '';
  const btnPdf = document.getElementById('btnBaixarPdfPendentes');
  if(btnPdf) btnPdf.style.display = 'none';

  const _mes    = mesCorrente();
  const _ativosCursos = listaAtiva();
  const totalM  = _ativosCursos.length;
  const nComDss = _ativosCursos.filter(m => dssOkNoMes(m, _mes)).length;
  const nSemDss = totalM - nComDss;
  const nRecOk  = _ativosCursos.filter(m => reciclagemStatus(m) === 'OK').length;
  const nRecPend= totalM - nRecOk;
  const nSimOk  = _ativosCursos.filter(m => simuladorStatus(m) === 'OK').length;
  const nSimPend= totalM - nSimOk;
  const nGestOk  = _ativosCursos.filter(m => gestimeStatus(m) === 'OK').length;
  const nGestPend= totalM - nGestOk;

  const grupos = [
    {{ titulo:'DSS Mensal',              icon:'fa-calendar-check',  tipoOk:'comDss',        tipoPend:'semDss',        tipoTodos:'dssTodos',        ok:nComDss, pend:nSemDss }},
    {{ titulo:'Reciclagem',              icon:'fa-recycle',         tipoOk:'reciclagemOk',  tipoPend:'reciclagemPend',tipoTodos:'reciclagemTodos', ok:nRecOk,  pend:nRecPend }},
    {{ titulo:'Simulador SEST SENAT',    icon:'fa-car-side',        tipoOk:'simuladorOk',   tipoPend:'simuladorPend', tipoTodos:'simuladorTodos',  ok:nSimOk,  pend:nSimPend }},
    {{ titulo:'Gestime',                 icon:'fa-clipboard-check', tipoOk:'gestimeOk',     tipoPend:'gestimePend',   tipoTodos:'gestimeTodos',    ok:nGestOk, pend:nGestPend }},
  ];

  document.getElementById('kpiCardsGrid').innerHTML = grupos.map(g => `
    <div class="categoria-glass-panel" onclick="abrirCategoriaCursos('${{g.tipoTodos}}')">
      <div class="cgp-header">
        <div class="cgp-icon"><i class="fa-solid ${{g.icon}}"></i></div>
        <div class="cgp-titulo">${{g.titulo}}</div>
      </div>
      <div class="cgp-stats">
        <div class="cgp-stat ok" onclick="event.stopPropagation(); abrirCategoriaCursos('${{g.tipoOk}}')">
          <div class="cgp-stat-val">${{g.ok}}</div>
          <div class="cgp-stat-lbl"><i class="fa-solid fa-check"></i> OK</div>
        </div>
        <div class="cgp-stat pend" onclick="event.stopPropagation(); abrirCategoriaCursos('${{g.tipoPend}}')">
          <div class="cgp-stat-val">${{g.pend}}</div>
          <div class="cgp-stat-lbl"><i class="fa-solid fa-clock"></i> Pendente</div>
        </div>
      </div>
    </div>
  `).join('');

  document.getElementById('kpiModal').classList.add('show');
}}

function abrirKpiModal(tipo, mes){{
  kpiTipoAtual = tipo;
  const cfg = KPI_CONFIG[tipo];
  if(!mes) mes = cfg.dssModal ? mesCorrente() : null;
  kpiMesAtual = mes;
  const btnVoltarCursos = document.getElementById('btnVoltarCursos');
  if(kpiOrigemCursos){{
    kpiVoltarFn = abrirCursosMenu;
    if(btnVoltarCursos){{ btnVoltarCursos.style.display = 'flex'; btnVoltarCursos.onclick = () => kpiVoltarFn && kpiVoltarFn(); }}
  }} else {{
    kpiVoltarFn = null;
    if(btnVoltarCursos) btnVoltarCursos.style.display = 'none';
  }}
  _aplicarFiltroKpi();
  document.getElementById('kpiModalIcon').innerHTML  = `<i class="fa-solid ${{cfg.icon}}" style="color:${{cfg.cor}}"></i>`;
  document.getElementById('kpiModalIcon').style.background = cfg.bg;
  document.getElementById('kpiModalLabel').textContent = cfg.label + (mes ? ` — ${{mes}}` : '');
  document.getElementById('kpiSearchInput').value = '';
  const mesFiltroEl = document.getElementById('kpiMesFiltro');
  if(cfg.dssModal){{
    mesFiltroEl.classList.add('visible');
    mesFiltroEl.innerHTML = MESES.map(m => `<button class="mes-btn${{m === kpiMesAtual ? ' ativo' : ''}}" onclick="trocarMesKpi('${{m}}')">${{m.substring(0,3).toUpperCase()}}</button>`).join('');
  }} else {{
    mesFiltroEl.classList.remove('visible');
    mesFiltroEl.innerHTML = '';
  }}
  document.getElementById('kpiModal').classList.add('show');
  const btnPdf = document.getElementById('btnBaixarPdfPendentes');
  if(btnPdf){{
    if(tipo === 'semDss' || tipo === 'comDss'){{
      btnPdf.style.display = 'flex';
      btnPdf.onclick = tipo === 'semDss' ? gerarRelatorioPdfPendentes : gerarRelatorioPdfRealizados;
      btnPdf.innerHTML = '<i class="fa-solid fa-file-pdf"></i> Baixar PDF';
    }} else {{ btnPdf.style.display = 'none'; }}
  }}
}}

function trocarMesKpi(mes){{
  kpiMesAtual = mes;
  const cfg = KPI_CONFIG[kpiTipoAtual];
  document.getElementById('kpiModalLabel').textContent = cfg.label + ` — ${{mes}}`;
  document.getElementById('kpiMesFiltro').querySelectorAll('.mes-btn').forEach(b => {{
    b.classList.toggle('ativo', b.textContent === mes.substring(0,3).toUpperCase());
  }});
  _aplicarFiltroKpi();
  document.getElementById('kpiSearchInput').value = '';
}}

function _aplicarFiltroKpi(){{
  const cfg = KPI_CONFIG[kpiTipoAtual];
  const base = kpiTipoAtual === 'desligados' ? listaDesligados() : listaAtiva();
  kpiListaAtual = base.filter(m => cfg.filtro(m, kpiMesAtual));
  document.getElementById('kpiModalCount').textContent = `${{kpiListaAtual.length}} motorista${{kpiListaAtual.length !== 1 ? 's' : ''}}`;
  renderizarCardsKpi(kpiListaAtual);
}}

function filtrarCardsKpi(){{
  const q = document.getElementById('kpiSearchInput').value.toLowerCase();
  const filtrados = q
    ? kpiListaAtual.filter(m => m.nome.toLowerCase().includes(q) || m.cpf.includes(q) || (m.filial||'').toLowerCase().includes(q))
    : kpiListaAtual;
  renderizarCardsKpi(filtrados);
}}

function renderizarCardsKpi(lista){{
  const grid = document.getElementById('kpiCardsGrid');
  if(lista.length === 0){{ grid.innerHTML = `<div class="kpi-empty"><i class="fa-solid fa-magnifying-glass"></i>Nenhum motorista encontrado.</div>`; return; }}

  if(kpiTipoAtual === 'vencimentoCategoria'){{
    const isAlerta = kpiTipoVencAtual === 'alerta';
    const cor = isAlerta ? '#eab308' : '#dc2626';
    const bg  = isAlerta ? '#fefce8' : '#fff5f5';
    grid.innerHTML = lista.map(m => {{
      const avatar = `<img src="${{m.foto || AVATAR_PADRAO}}" alt="">`;
      const d = diasParaVencerCategoria(m, kpiCategoriaVencAtual);
      const prazoTxt = isAlerta ? `Vence em ${{d}} dia${{d===1?'':'s'}}` : `Vencido há ${{Math.abs(d)}} dia${{Math.abs(d)===1?'':'s'}}`;
      return `<div class="driver-mini-card" style="border-color:${{cor}}55;background:${{bg}};" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge" style="background:${{cor}}22;color:${{cor}};border:1px solid ${{cor}}55;"><i class="fa-solid ${{isAlerta?'fa-triangle-exclamation':'fa-circle-exclamation'}}"></i> ${{prazoTxt}}</span></div>
      </div>`;
    }}).join('');
    return;
  }}

  const cfg = kpiTipoAtual ? KPI_CONFIG[kpiTipoAtual] : null;
  const isDssModal   = cfg && cfg.dssModal;
  const isExcesso    = kpiTipoAtual === 'excesso';
  const isMultas     = kpiTipoAtual === 'multas';
  const isAcidentes  = kpiTipoAtual === 'acidentes';
  const isInfracao   = isExcesso || isMultas || isAcidentes;
  const isTelCorp    = kpiTipoAtual === 'telCorp';
  const isAfastados  = kpiTipoAtual === 'afastados';
  const isDesligados = kpiTipoAtual === 'desligados';
  const isProntuario = kpiTipoAtual === 'prontuario';
  const isReciclagemOk = kpiTipoAtual === 'reciclagemOk';
  const isReciclagemPend = kpiTipoAtual === 'reciclagemPend';
  const isSimuladorOk = kpiTipoAtual === 'simuladorOk';
  const isSimuladorPend = kpiTipoAtual === 'simuladorPend';
  const isGestimeOk = kpiTipoAtual === 'gestimeOk';
  const isGestimePend = kpiTipoAtual === 'gestimePend';
  const isReciclagemTodos = kpiTipoAtual === 'reciclagemTodos';
  const isSimuladorTodos  = kpiTipoAtual === 'simuladorTodos';
  const isGestimeTodos    = kpiTipoAtual === 'gestimeTodos';
  grid.innerHTML = lista.map(m => {{
    const avatar = `<img src="${{m.foto || AVATAR_PADRAO}}" alt="">`;
    const nExc   = Math.max(0, parseInt(m.excesso)   || 0);
    const nMul   = Math.max(0, parseInt(m.multas)    || 0);
    const nAcid  = Math.max(0, parseInt(m.acidentes) || 0);
    if(isTelCorp){{
      return `<div class="driver-mini-card card-ok" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-infracao"><i class="fa-solid fa-mobile-screen-button dmc-inf-icon" style="color:#0e9cc0"></i>
          <div class="dmc-inf-body"><div class="dmc-inf-label">Número da Linha</div><div class="dmc-inf-val" style="color:#0e9cc0;font-size:17px;">${{m.numeroLinha || 'Não informado'}}</div></div>
        </div>
        <div class="dmc-badges">
          <span class="dmc-badge ok"><i class="fa-solid fa-mobile"></i> ${{m.modelo || 'Modelo não informado'}}</span>
          <span class="dmc-badge ok"><i class="fa-solid fa-barcode"></i> IMEI: ${{m.imei || 'Não informado'}}</span>
        </div>
      </div>`;
    }}
    if(isAfastados){{
      return `<div class="driver-mini-card card-pend" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge pend"><i class="fa-solid fa-user-slash"></i> Afastado</span></div>
        ${{m.obsAfastado ? `<div style="font-size:14px;color:#5a6e8a;font-style:italic;margin-top:4px;">${{m.obsAfastado}}</div>` : ''}}
      </div>`;
    }}
    if(isDesligados){{
      return `<div class="driver-mini-card" style="border-color:#94a3b8;background:repeating-linear-gradient(45deg,#f8fafc,#f8fafc 8px,#eef1f5 8px,#eef1f5 16px);" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge" style="background:#e2e8f0;color:#475569;border:1px solid #cbd5e1;"><i class="fa-solid fa-user-xmark"></i> Desligado</span></div>
        ${{m.obsDesligamento ? `<div style="font-size:14px;color:#5a6e8a;font-style:italic;margin-top:4px;">${{m.obsDesligamento}}</div>` : ''}}
      </div>`;
    }}
    if(isProntuario){{
      const exPerOk = exameOk(m.examePeriodico, m.examePeriodicoValidadeMeses);
      const exToxOk = exameOk(m.exameToxicologico, m.exameToxicologicoValidadeMeses);
      const svExPerK = statusVencimentoProntuario(m.examePeriodico, m.examePeriodicoValidadeMeses);
      const svExToxK = statusVencimentoProntuario(m.exameToxicologico, m.exameToxicologicoValidadeMeses);
      const recPill = statusPill3(m.reciclagemData, m.reciclagemValidadeMeses);
      const simPill = statusPill3(m.simuladorData, m.simuladorValidadeMeses);
      const uniOk   = m.entregaUniforme === 'OK';
      const afastadoSim = m.afastado === 'SIM';
      const svCnhK  = statusVencimentoData(m.validadeCnh);
      const svMoppK = statusVencimentoData(m.vencimentoCnhMopp);
      const fmtData = d => d ? new Date(d+'T00:00:00').toLocaleDateString('pt-BR') : 'PENDENTE';
      return `<div class="driver-mini-card ${{(exPerOk && exToxOk) ? 'card-ok' : 'card-pend'}}" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top-center">
          <div class="dmc-avatar">${{avatar}}</div>
          <div class="dmc-nome">${{m.nome}}</div>
          <div class="dmc-filial">${{m.filial||'—'}}</div>
          <div class="dmc-cpf">${{m.cpf}}</div>
        </div>
        <div class="dmc-pront-grid" style="background:none;border:none;padding:0;gap:8px;">
          <div class="dmc-pront-item full" style="background:rgba(255,255,255,0.55);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);border:1px solid rgba(196,208,228,0.6);border-radius:8px;padding:8px 10px;box-shadow:0 2px 6px rgba(20,50,120,0.05);gap:4px;">
            <div style="display:flex;justify-content:space-between;align-items:center;white-space:nowrap;"><span class="dmc-pront-lbl">Exame Periódico</span><span class="dmc-pront-val" style="color:${{exPerOk?'#16a34a':'#dc2626'}};white-space:nowrap;">${{fmtData(m.examePeriodico)}}</span></div>
            <div style="display:flex;justify-content:space-between;align-items:center;border-top:1px dashed rgba(196,208,228,0.6);padding-top:4px;white-space:nowrap;"><span class="dmc-pront-lbl">Venc. Periódico</span><span class="dmc-pront-val" style="color:${{svExPerK.cor}};white-space:nowrap;">${{svExPerK.venc ? svExPerK.venc.toLocaleDateString('pt-BR') : 'PENDENTE'}}</span></div>
          </div>
          <div class="dmc-pront-item full" style="background:rgba(255,255,255,0.55);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);border:1px solid rgba(196,208,228,0.6);border-radius:8px;padding:8px 10px;box-shadow:0 2px 6px rgba(20,50,120,0.05);gap:4px;">
            <div style="display:flex;justify-content:space-between;align-items:center;white-space:nowrap;"><span class="dmc-pront-lbl">Exame Toxicológico</span><span class="dmc-pront-val" style="color:${{exToxOk?'#16a34a':'#dc2626'}};white-space:nowrap;">${{fmtData(m.exameToxicologico)}}</span></div>
            <div style="display:flex;justify-content:space-between;align-items:center;border-top:1px dashed rgba(196,208,228,0.6);padding-top:4px;white-space:nowrap;"><span class="dmc-pront-lbl">Venc. Toxicológico</span><span class="dmc-pront-val" style="color:${{svExToxK.cor}};white-space:nowrap;">${{svExToxK.venc ? svExToxK.venc.toLocaleDateString('pt-BR') : 'PENDENTE'}}</span></div>
          </div>
          <div class="dmc-pront-item full" style="flex-direction:row;justify-content:space-between;align-items:center;background:rgba(255,255,255,0.55);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);border:1px solid rgba(196,208,228,0.6);border-radius:8px;padding:8px 10px;box-shadow:0 2px 6px rgba(20,50,120,0.05);"><span class="dmc-pront-lbl">Validade CNH</span><span class="dmc-pront-val" style="color:${{svCnhK.cor}}">${{fmtData(m.validadeCnh)}}</span></div>
          <div class="dmc-pront-item full" style="flex-direction:row;justify-content:space-between;align-items:center;background:rgba(255,255,255,0.55);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);border:1px solid rgba(196,208,228,0.6);border-radius:8px;padding:8px 10px;box-shadow:0 2px 6px rgba(20,50,120,0.05);"><span class="dmc-pront-lbl">Validade MOPP</span><span class="dmc-pront-val" style="color:${{svMoppK.cor}}">${{fmtData(m.vencimentoCnhMopp)}}</span></div>
          <div class="dmc-pront-item" style="background:rgba(255,255,255,0.55);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);border:1px solid rgba(196,208,228,0.6);border-radius:8px;padding:8px 10px;box-shadow:0 2px 6px rgba(20,50,120,0.05);"><span class="dmc-pront-lbl">Pontuação CNH</span><span class="dmc-pront-val">${{m.pontuacaoCnh||0}} pts</span></div>
          <div class="dmc-pront-item full" style="background:rgba(255,255,255,0.55);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);border:1px solid rgba(196,208,228,0.6);border-radius:8px;padding:8px 10px;box-shadow:0 2px 6px rgba(20,50,120,0.05);"><span class="dmc-pront-lbl">Entrega de Uniforme</span><span class="dmc-pront-val" style="color:${{uniOk?'#16a34a':'#dc2626'}}">${{m.entregaUniforme||'PENDENTE'}}</span></div>
          <div class="dmc-pront-item full" style="background:rgba(255,255,255,0.55);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);border:1px solid rgba(196,208,228,0.6);border-radius:8px;padding:8px 10px;box-shadow:0 2px 6px rgba(20,50,120,0.05);"><span class="dmc-pront-lbl">Afastado</span><span class="dmc-pront-val" style="color:${{afastadoSim?'#dc2626':'#16a34a'}}">${{m.afastado||'NÃO'}}${{afastadoSim && m.obsAfastado ? ' — ' + m.obsAfastado : ''}}</span></div>
        </div>
        <div class="dmc-status-row">
          <div class="dmc-status-pill ${{recPill.cls}}">
            <span class="dmc-status-pill-lbl"><i class="fa-solid fa-recycle"></i> Reciclagem</span>
            <span class="dmc-status-pill-val">${{recPill.txt}}</span>
          </div>
          <div class="dmc-status-pill ${{simPill.cls}}">
            <span class="dmc-status-pill-lbl"><i class="fa-solid fa-car-side"></i> Simulador</span>
            <span class="dmc-status-pill-val">${{simPill.txt}}</span>
          </div>
        </div>
      </div>`;
    }}
    if(isReciclagemOk){{
      const svR = statusVencimento(m.reciclagemData, m.reciclagemValidadeMeses);
      return `<div class="driver-mini-card card-ok" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge ok"><i class="fa-solid fa-recycle"></i> Reciclagem OK</span></div>
        <div style="font-size:13px;font-weight:700;color:${{svR.cor}}">${{svR.label}}</div>
      </div>`;
    }}
    if(isReciclagemPend){{
      const svR = statusVencimento(m.reciclagemData, m.reciclagemValidadeMeses);
      return `<div class="driver-mini-card card-pend" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge pend"><i class="fa-solid fa-clock"></i> Reciclagem Pendente</span></div>
        <div style="font-size:13px;font-weight:700;color:${{svR.cor}}">${{svR.label}}</div>
      </div>`;
    }}
    if(isSimuladorOk){{
      const svS = statusVencimento(m.simuladorData, m.simuladorValidadeMeses);
      return `<div class="driver-mini-card card-ok" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge ok"><i class="fa-solid fa-car-side"></i> Simulador OK</span></div>
        <div style="font-size:13px;font-weight:700;color:${{svS.cor}}">${{svS.label}}</div>
      </div>`;
    }}
    if(isSimuladorPend){{
      const svS = statusVencimento(m.simuladorData, m.simuladorValidadeMeses);
      return `<div class="driver-mini-card card-pend" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge pend"><i class="fa-solid fa-clock"></i> Simulador Pendente</span></div>
        <div style="font-size:13px;font-weight:700;color:${{svS.cor}}">${{svS.label}}</div>
      </div>`;
    }}
    if(isGestimeOk){{
      const svG = statusVencimento(m.gestimeData, m.gestimeValidadeMeses);
      return `<div class="driver-mini-card card-ok" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge ok"><i class="fa-solid fa-clipboard-check"></i> Gestime OK</span></div>
        <div style="font-size:13px;font-weight:700;color:${{svG.cor}}">${{svG.label}}</div>
      </div>`;
    }}
    if(isGestimePend){{
      const svG = statusVencimento(m.gestimeData, m.gestimeValidadeMeses);
      return `<div class="driver-mini-card card-pend" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge pend"><i class="fa-solid fa-clock"></i> Gestime Pendente</span></div>
        <div style="font-size:13px;font-weight:700;color:${{svG.cor}}">${{svG.label}}</div>
      </div>`;
    }}
    if(isReciclagemTodos){{
      const ok = reciclagemStatus(m) === 'OK';
      const svR = statusVencimento(m.reciclagemData, m.reciclagemValidadeMeses);
      return `<div class="driver-mini-card ${{ok?'card-ok':'card-pend'}}" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge ${{ok?'ok':'pend'}}"><i class="fa-solid ${{ok?'fa-recycle':'fa-clock'}}"></i> Reciclagem ${{ok?'OK':'Pendente'}}</span></div>
        <div style="font-size:13px;font-weight:700;color:${{svR.cor}}">${{svR.label}}</div>
      </div>`;
    }}
    if(isSimuladorTodos){{
      const ok = simuladorStatus(m) === 'OK';
      const svS = statusVencimento(m.simuladorData, m.simuladorValidadeMeses);
      return `<div class="driver-mini-card ${{ok?'card-ok':'card-pend'}}" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge ${{ok?'ok':'pend'}}"><i class="fa-solid ${{ok?'fa-car-side':'fa-clock'}}"></i> Simulador ${{ok?'OK':'Pendente'}}</span></div>
        <div style="font-size:13px;font-weight:700;color:${{svS.cor}}">${{svS.label}}</div>
      </div>`;
    }}
    if(isGestimeTodos){{
      const ok = gestimeStatus(m) === 'OK';
      const svG = statusVencimento(m.gestimeData, m.gestimeValidadeMeses);
      return `<div class="driver-mini-card ${{ok?'card-ok':'card-pend'}}" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-badges"><span class="dmc-badge ${{ok?'ok':'pend'}}"><i class="fa-solid ${{ok?'fa-clipboard-check':'fa-clock'}}"></i> Gestime ${{ok?'OK':'Pendente'}}</span></div>
        <div style="font-size:13px;font-weight:700;color:${{svG.cor}}">${{svG.label}}</div>
      </div>`;
    }}
    if(isInfracao){{
      let infVal, infCls, infIcon, infLabel;
      if(isExcesso)   {{ infVal=nExc;  infCls='vel';  infIcon='fa-gauge-high';          infLabel='Excessos de Velocidade'; }}
      if(isMultas)    {{ infVal=nMul;  infCls='mul';  infIcon='fa-file-circle-xmark';   infLabel='Multas Registradas'; }}
      if(isAcidentes) {{ infVal=nAcid; infCls='acid'; infIcon='fa-car-burst';            infLabel='Acidentes Registrados'; }}
      return `<div class="driver-mini-card" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
        <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
        <div class="dmc-cpf">${{m.cpf}}</div>
        <div class="dmc-infracao"><i class="fa-solid ${{infIcon}} dmc-inf-icon" style="color:${{infCls==='vel'?'#ff4444':infCls==='mul'?'#ff6622':'#ff4488'}}"></i>
          <div class="dmc-inf-body"><div class="dmc-inf-label">${{infLabel}}</div><div class="dmc-inf-val ${{infCls}}">${{infVal}}</div></div>
        </div></div>`;
    }}
    let semanasHtml = '';
    if(isDssModal && kpiMesAtual){{
      const sems = dssDoMes(m, kpiMesAtual);
      semanasHtml = `<div class="dmc-semanas">` +
        sems.map((ok,i) => `<div class="dmc-sem"><span class="dmc-sem-lbl">${{i+1}}ª S</span><span class="dmc-sem-dot ${{ok?'ok':'pend'}}">${{ok?'✓':'✗'}}</span></div>`).join('') +
        `</div>`;
    }}
    const nDssMes  = isDssModal && kpiMesAtual ? contarDssMes(m, kpiMesAtual) : contarDssSessoes(m);
    const dssOkMes = isDssModal && kpiMesAtual ? dssOkNoMes(m, kpiMesAtual)   : temDss(m);
    const mesLabel = kpiMesAtual ? kpiMesAtual.substring(0,3).toUpperCase() : '';
    const badges = [
      isDssModal
        ? (dssOkMes
            ? `<span class="dmc-badge ok"><i class="fa-solid fa-calendar-check"></i> ${{mesLabel}} ${{nDssMes}}/4 ✓</span>`
            : `<span class="dmc-badge pend"><i class="fa-solid fa-clock"></i> ${{mesLabel}} ${{nDssMes}}/4 pendente</span>`)
        : (temDss(m) ? `<span class="dmc-badge ok">DSS ok</span>` : `<span class="dmc-badge pend">Sem DSS</span>`)
    ].join('');
    const cardClass = dssOkMes ? 'card-ok' : 'card-pend';
    return `<div class="driver-mini-card ${{cardClass}}" onclick="irParaFichaViaKpi('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
      <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
      <div class="dmc-cpf">${{m.cpf}}</div>
      ${{semanasHtml}}<div class="dmc-badges">${{badges}}</div></div>`;
  }}).join('');
}}

function irParaFichaViaKpi(cpf){{ fichaOrigemModal='kpi'; fecharKpiModal(); setTimeout(()=>abrirFichaMotorista(cpf),120); }}
function fecharKpiModal(){{
  document.getElementById('kpiModal').classList.remove('show');
  kpiOrigemCursos = false;
  kpiVoltarFn = null;
  kpiCategoriaVencAtual = null;
  kpiTipoVencAtual = null;
  const btnVoltarCursos = document.getElementById('btnVoltarCursos');
  if(btnVoltarCursos) btnVoltarCursos.style.display = 'none';
}}
document.addEventListener('click', e => {{ if(e.target === document.getElementById('kpiModal')) fecharKpiModal(); }});

function toggleFormulario(){{
  const body = document.getElementById('formBody');
  const btn  = document.getElementById('btnToggleForm');
  const aberto = body.classList.toggle('open');
  btn.classList.toggle('open', aberto);
}}

async function carregarDados(){{
  mostrarSpinner(true);
  try{{
    const res = await apiFetch('/api/motoristas');
    if(res.ok){{ motoristasDB = res.motoristas; atualizarDashboardCompleto(); }}
    else {{ toast('Erro ao carregar dados: ' + res.erro, 'erro'); }}
  }} catch(e){{ toast('Falha de conexão com o servidor.', 'erro'); }}
  finally{{ mostrarSpinner(false); }}
}}

function gerarMatrizDssEmBranco(){{
  const d = {{}};
  MESES.forEach(m => {{ d[m] = [false,false,false,false]; }});
  return d;
}}

function temDss(m){{ return MESES.some(mes => m.dssAnual[mes] && m.dssAnual[mes].some(s => s)); }}
function contarDssSessoes(m){{ let t=0; MESES.forEach(mes=>{{ if(m.dssAnual[mes]) m.dssAnual[mes].forEach(s=>{{ if(s) t++; }}); }}); return t; }}
function dssDoMes(m, mes){{ return m.dssAnual && m.dssAnual[mes] ? m.dssAnual[mes] : [false,false,false,false]; }}
function contarDssMes(m, mes){{ return dssDoMes(m, mes).filter(Boolean).length; }}
function dssOkNoMes(m, mes){{ return contarDssMes(m, mes) >= 4; }}

function calcularVencimento(dataStr, meses){{
  if(!dataStr || !meses) return null;
  const d = new Date(dataStr+'T00:00:00');
  if(isNaN(d)) return null;
  d.setMonth(d.getMonth() + parseInt(meses));
  return d;
}}
function statusVencimento(dataStr, meses){{
  const venc = calcularVencimento(dataStr, meses);
  if(!venc) return {{ label:'Sem data/validade definida', cor:'#9aaabb', venc:null }};
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const diffDias = Math.floor((venc - hoje) / 86400000);
  if(diffDias < 0)   return {{ label:'VENCIDO em ' + venc.toLocaleDateString('pt-BR'), cor:'#dc2626', venc }};
  if(diffDias <= 30) return {{ label:'Vence em ' + venc.toLocaleDateString('pt-BR'), cor:'#d97706', venc }};
  return {{ label:'Válido até ' + venc.toLocaleDateString('pt-BR'), cor:'#16a34a', venc }};
}}

function statusVencimentoProntuario(dataStr, meses){{
  const venc = calcularVencimento(dataStr, meses);
  if(!venc) return {{ label:'Sem data/validade definida', cor:'#9aaabb', venc:null }};
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const diffDias = Math.floor((venc - hoje) / 86400000);
  if(diffDias < 0) return {{ label:'VENCIDO em ' + venc.toLocaleDateString('pt-BR'), cor:'#dc2626', venc }};
  const cor = diffDias <= 30 ? '#d97706' : '#16a34a';
  return {{ label: venc.toLocaleDateString('pt-BR'), cor, venc }};
}}

// Retorna true se a validade já passou (curso vencido)
function estaVencido(dataStr, meses){{
  const venc = calcularVencimento(dataStr, meses);
  if(!venc) return false; // sem data/validade definida -> não força vencimento
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  return venc < hoje;
}}

// Status EFETIVO do curso: se está "OK" mas a validade venceu,
// ele volta automaticamente para PENDENTE (a data NÃO é apagada em nenhum momento)
function statusEfetivo(campoStatus, dataStr, meses){{
  if(campoStatus !== 'OK') return 'PENDENTE';
  return estaVencido(dataStr, meses) ? 'PENDENTE' : 'OK';
}}

function reciclagemStatus(m){{ return statusEfetivo(m.reciclagem, m.reciclagemData, m.reciclagemValidadeMeses); }}
function simuladorStatus(m){{  return statusEfetivo(m.simulador,  m.simuladorData,  m.simuladorValidadeMeses); }}
function gestimeStatus(m){{    return statusEfetivo(m.gestime,    m.gestimeData,    m.gestimeValidadeMeses); }}

// Status em 3 estados (verde/laranja/vermelho) para os pills do Prontuário
function statusPill3(dataStr, meses){{
  const sv = statusVencimento(dataStr, meses);
  if(!sv.venc) return {{ cls:'pend', txt:'PENDENTE' }};
  if(sv.cor === '#dc2626') return {{ cls:'vencido', txt:'VENCIDO' }};
  if(sv.cor === '#d97706') return {{ cls:'pend', txt:'ALERTA' }};
  return {{ cls:'ok', txt:'OK' }};
}}
// Exame: considerado OK só se tem data preenchida E não está vencido.
// A data continua salva mesmo depois de vencer — só o status calculado muda.
function exameOk(dataStr, meses){{
  if(!dataStr) return false;
  return !estaVencido(dataStr, meses);
}}

// ── Validade CNH ──
// Diferente de reciclagem/exames: aqui a própria data preenchida É o vencimento
// (não soma meses). Se a data já passou, a CNH conta como VENCIDA.
function statusVencimentoData(dataStr){{
  if(!dataStr) return {{ label:'Sem data definida', cor:'#9aaabb', venc:null, vencida:false }};
  const venc = new Date(dataStr+'T00:00:00');
  if(isNaN(venc)) return {{ label:'Data inválida', cor:'#9aaabb', venc:null, vencida:false }};
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const diffDias = Math.floor((venc - hoje) / 86400000);
  if(diffDias < 0)   return {{ label:'VENCIDA em ' + venc.toLocaleDateString('pt-BR'), cor:'#dc2626', venc, vencida:true }};
  if(diffDias <= 30) return {{ label:'Vence em ' + venc.toLocaleDateString('pt-BR'), cor:'#d97706', venc, vencida:false }};
  return {{ label:'Válido até ' + venc.toLocaleDateString('pt-BR'), cor:'#16a34a', venc, vencida:false }};
}}
function cnhVencida(m){{
  return statusVencimentoData(m.validadeCnh).vencida;
}}
function rotuloVenc(sv){{
  if(!sv || !sv.venc) return 'PENDENTE';
  if(sv.cor === '#dc2626') return 'VENCIDO';
  if(sv.cor === '#d97706') return 'ALERTA';
  return 'OK';
}}

// ── Grupos de campos vinculados a vencimento (data + validade em meses) ──
const CAMPOS_VENCIMENTO_MESES = [
  ['editReciclagemData','editReciclagemValidadeMeses'],
  ['editSimuladorData','editSimuladorValidadeMeses'],
  ['editExamePeriodico','editExamePeriodicoValidadeMeses'],
  ['editExameToxicologico','editExameToxicologicoValidadeMeses'],
  ['editGestimeData','editGestimeValidadeMeses'],
];
// ── Campos cuja própria data já é o vencimento (sem soma de meses) ──
const CAMPOS_VENCIMENTO_DATA = ['editValidadeCnh','editVencimentoCnhMopp'];

function _aplicarCorPorStatus(elementos, cor){{
  elementos.forEach(el => {{
    if(!el) return;
    el.classList.remove('campo-valido','campo-alerta-venc','campo-vencido-venc');
    if(cor === '#16a34a') el.classList.add('campo-valido');
    else if(cor === '#d97706') el.classList.add('campo-alerta-venc');
    else if(cor === '#dc2626') el.classList.add('campo-vencido-venc');
  }});
}}

function aplicarEstadoVisualCampos(){{
  const container = document.getElementById('driverProfileContent');
  if(!container) return;

  const idsUsados = new Set();
  CAMPOS_VENCIMENTO_MESES.forEach(([dataId, anosId]) => {{
    idsUsados.add(dataId); idsUsados.add(anosId);
    const dataEl = document.getElementById(dataId);
    const anosEl = document.getElementById(anosId);
    if(!dataEl || !anosEl) return;
    const sv = statusVencimento(dataEl.value, anosEl.value);
    _aplicarCorPorStatus([dataEl, anosEl], sv.venc ? sv.cor : null);
  }});

  CAMPOS_VENCIMENTO_DATA.forEach(id => {{
    idsUsados.add(id);
    const el = document.getElementById(id);
    if(!el) return;
    const sv = statusVencimentoData(el.value);
    _aplicarCorPorStatus([el], sv.venc ? sv.cor : null);
  }});

  idsUsados.add('editObsAfastado');
  idsUsados.add('editAcidentes');
  idsUsados.add('editMultas');
  idsUsados.add('editExcesso');
  idsUsados.add('editObsAcidente');
  idsUsados.add('editObsMultas');
  idsUsados.add('editObsGerais');
  container.querySelectorAll('input[type="text"], input[type="email"], input[type="number"], input[type="date"]').forEach(el => {{
    if(idsUsados.has(el.id)) return;
    el.classList.remove('campo-valido','campo-alerta-venc','campo-vencido-venc');
    if((el.value||'').toString().trim() !== '') el.classList.add('campo-valido');
  }});

  // Acidentes / Multas / Exc. Velocidade: verde se 0, vermelho se 1 ou mais
  ['editAcidentes','editMultas','editExcesso'].forEach(id => {{
    const el = document.getElementById(id);
    if(!el) return;
    el.classList.remove('campo-valido','campo-alerta-venc','campo-vencido-venc');
    const val = parseInt(el.value) || 0;
    el.classList.add(val > 0 ? 'campo-vencido-venc' : 'campo-valido');
  }});

  // Obs. Acidente / Obs. Multas / Obs. Excesso: vermelho se preenchido
  ['editObsAcidente','editObsMultas','editObsGerais'].forEach(id => {{
    const el = document.getElementById(id);
    if(!el) return;
    el.classList.remove('campo-valido','campo-alerta-venc','campo-vencido-venc');
    if((el.value||'').toString().trim() !== '') el.classList.add('campo-vencido-venc');
  }});

  // Selects marcados como "OK" ficam verdes (Reciclagem, Simulador, Gestime, Entrega de Uniforme)
  ['editReciclagem','editSimulador','editGestime','editEntregaUniforme'].forEach(id => {{
    const el = document.getElementById(id);
    if(!el) return;
    el.classList.remove('campo-valido','campo-alerta-venc','campo-vencido-venc');
    if(el.value === 'OK') el.classList.add('campo-valido');
  }});

  // Telefone Corporativo = SIM fica verde
  const telCorpEl = document.getElementById('editTelefoneCorporativo');
  if(telCorpEl){{
    telCorpEl.classList.remove('campo-valido','campo-alerta-venc','campo-vencido-venc');
    if(telCorpEl.value === 'SIM') telCorpEl.classList.add('campo-valido');
  }}

  // Afastado: SIM fica vermelho, NÃO fica verde
  const afastadoEl = document.getElementById('editAfastado');
  if(afastadoEl){{
    afastadoEl.classList.remove('campo-valido','campo-alerta-venc','campo-vencido-venc');
    afastadoEl.classList.add(afastadoEl.value === 'SIM' ? 'campo-vencido-venc' : 'campo-valido');
  }}

  // Obs de Afastamento: fica vermelho quando preenchido
  const obsAfastadoEl = document.getElementById('editObsAfastado');
  if(obsAfastadoEl){{
    obsAfastadoEl.classList.remove('campo-valido','campo-alerta-venc','campo-vencido-venc');
    if((obsAfastadoEl.value||'').toString().trim() !== '') obsAfastadoEl.classList.add('campo-vencido-venc');
  }}

  // Label "*obrigatório" do Número da Linha fica verde quando o campo está preenchido
  const numLinhaEl = document.getElementById('editNumeroLinha');
  const numLinhaObrigLbl = document.getElementById('numeroLinhaObrigatorio');
  if(numLinhaEl && numLinhaObrigLbl){{
    const preenchido = (numLinhaEl.value||'').toString().trim() !== '';
    numLinhaObrigLbl.style.color = preenchido ? '#16a34a' : '#dc2626';
  }}
}}

document.addEventListener('input',  e => {{ if(e.target.closest && e.target.closest('#driverProfileContent')){{ aplicarEstadoVisualCampos(); houveEdicaoNaoSalva = true; }} }});
document.addEventListener('change', e => {{ if(e.target.closest && e.target.closest('#driverProfileContent')){{ aplicarEstadoVisualCampos(); houveEdicaoNaoSalva = true; }} }});

// Mantido por compatibilidade com o HTML já existente (oninput="checarCamposValidade(...)").
function checarCamposValidade(dataId, anosId){{
  aplicarEstadoVisualCampos();
}}

function obterChaveFilial(m){{
  return (m.filial||'').toUpperCase().trim() || 'AGUARDANDO FILIAL';
}}

function agruparPorFilial(lista){{
  lista = lista || listaAtiva();
  const mapa = {{}};
  const mesAtual  = MESES[new Date().getMonth()];
  const semAtual  = Math.min(3, Math.floor((new Date().getDate() - 1) / 7));
  lista.forEach(m => {{
    const f = obterChaveFilial(m);
    if(!mapa[f]) mapa[f] = {{ name:f, total:0, comDss:0, dssMax:0, recOk:0, simOk:0, acid:0, multas:0, excVel:0, acidMot:0, multasMot:0, excVelMot:0, examePerOk:0, exameToxOk:0, telCorpOk:0 }};
    mapa[f].total++;
    mapa[f].comDss += (m.dssAnual && m.dssAnual[mesAtual] && m.dssAnual[mesAtual][semAtual]) ? 1 : 0;
    mapa[f].dssMax += 1;
    if(reciclagemStatus(m) === 'OK') mapa[f].recOk++;
    if(simuladorStatus(m)  === 'OK') mapa[f].simOk++;
    mapa[f].acid   += Math.max(0, parseInt(m.acidentes || 0));
    mapa[f].multas += Math.max(0, parseInt(m.multas    || 0));
    mapa[f].excVel += Math.max(0, parseInt(m.excesso   || 0));
    if(Math.max(0, parseInt(m.acidentes || 0)) > 0) mapa[f].acidMot++;
    if(Math.max(0, parseInt(m.multas    || 0)) > 0) mapa[f].multasMot++;
    if(Math.max(0, parseInt(m.excesso   || 0)) > 0) mapa[f].excVelMot++;
    if(exameOk(m.examePeriodico, m.examePeriodicoValidadeMeses))       mapa[f].examePerOk++;
    if(exameOk(m.exameToxicologico, m.exameToxicologicoValidadeMeses)) mapa[f].exameToxOk++;
    if(m.telefoneCorporativo === 'SIM') mapa[f].telCorpOk++;
  }});
  return Object.values(mapa).sort((a,b) => b.total - a.total);
}}

function atualizarDashboardCompleto(){{
  const ativos     = listaAtiva();
  const desligados = listaDesligados();
  const filiais    = agruparPorFilial(ativos);
  const totalM     = ativos.length;
  const _mesDash   = MESES[new Date().getMonth()];
  const totalComDss= ativos.filter(m => dssOkNoMes(m, _mesDash)).length;
  const totalPend  = totalM - totalComDss;
  const totalExc   = ativos.reduce((acc,m) => acc + Math.max(0, parseInt(m.excesso)   || 0), 0);
  const totalMul   = ativos.reduce((acc,m) => acc + Math.max(0, parseInt(m.multas)    || 0), 0);
  const totalAcid  = ativos.reduce((acc,m) => acc + Math.max(0, parseInt(m.acidentes) || 0), 0);
  document.getElementById('kpiTotal').textContent    = totalM;
  document.getElementById('kpiExcesso').textContent  = totalExc;
  document.getElementById('kpiMultas').textContent   = totalMul;
  document.getElementById('kpiAcidentes').textContent= totalAcid;
  const pct = totalM > 0 ? ((totalComDss / totalM)*100).toFixed(1) + '%' : '—';
  document.getElementById('macroPctDss').textContent = pct;
  const cursosTotalEl = document.getElementById('kpiCursosTotal');
  const cursosSubEl   = document.getElementById('kpiCursosSubPct');
  if(cursosTotalEl) cursosTotalEl.textContent = totalM;
  if(cursosSubEl)   cursosSubEl.textContent   = pct;

  const totalDssAnual = ativos.reduce((acc, m) => {{
    MESES.forEach(mes => {{
      if(m.dssAnual && m.dssAnual[mes]) acc += m.dssAnual[mes].filter(Boolean).length;
    }});
    return acc;
  }}, 0);
  const totalSessoesAnual = totalM * MESES.length * 4;
  const motExcesso = ativos.filter(m => Math.max(0,parseInt(m.excesso)||0) > 0).length;
  const motMultas  = ativos.filter(m => Math.max(0,parseInt(m.multas)||0)  > 0).length;
  const motAcident = ativos.filter(m => Math.max(0,parseInt(m.acidentes)||0) > 0).length;
  const _s = (id,v) => {{ const e=document.getElementById(id); if(e) e.textContent=v; }};
  _s('kpiTotalAnual',   totalM);
  _s('kpiRecOkAnual',   totalDssAnual);
  _s('kpiPendAnual',    totalSessoesAnual - totalDssAnual);
  _s('kpiExcessoMot',   motExcesso);
  _s('kpiMultasMot',    motMultas);
  _s('kpiAcidentesMot', motAcident);

  const totalTelCorp      = ativos.filter(m => m.telefoneCorporativo === 'SIM').length;
  const totalProntuarioOk = ativos.filter(m =>
    exameOk(m.examePeriodico, m.examePeriodicoValidadeMeses) &&
    exameOk(m.exameToxicologico, m.exameToxicologicoValidadeMeses)
  ).length;
  _s('kpiTelCorp',     totalTelCorp);
  _s('kpiTelCorpPct',  totalM > 0 ? Math.round(totalTelCorp/totalM*100) + '%' : '—');
  _s('kpiProntuario',  totalM);
  _s('kpiProntuarioOk',totalProntuarioOk);

  const totalAfastados = ativos.filter(m => m.afastado === 'SIM').length;
  _s('kpiAfastados',    totalAfastados);
  _s('kpiAfastadosPct', totalM > 0 ? Math.round(totalAfastados/totalM*100) + '%' : '—');

  const totalRecOk  = ativos.filter(m => reciclagemStatus(m) === 'OK').length;
  const totalRecPend = ativos.filter(m => reciclagemStatus(m) === 'PENDENTE').length;
  _s('kpiReciclagemOk', totalRecOk);
  _s('kpiReciclagemOkPct', totalM > 0 ? Math.round(totalRecOk/totalM*100) + '%' : '—');
  _s('kpiReciclagemPend', totalRecPend);
  _s('kpiReciclagemPendPct', totalM > 0 ? Math.round(totalRecPend/totalM*100) + '%' : '—');

  const totalGeral = motoristasDB.length;
  _s('kpiDesligados',    desligados.length);
  _s('kpiDesligadosPct', totalGeral > 0 ? Math.round(desligados.length/totalGeral*100) + '%' : '—');

  renderizarGridFiliais(filiais);
  renderizarGraficos(filiais);
}}

function renderizarGridFiliais(filiais){{
  const filColors = ['#1a4fa0','#16a34a','#d97706','#dc2626','#7c3aed','#be185d','#0e7490','#9a3412','#166534','#1e40af','#9d174d','#0369a1'];
  const grid = document.getElementById('filialGrid');
  if(filiais.length === 0){{ grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1"><i class="fa-solid fa-building-user"></i><p>Nenhum motorista cadastrado ainda.<br>Use o formulário acima para inserir o primeiro condutor.</p></div>`; return; }}
  grid.innerHTML = '';
  filiais.forEach((f, i) => {{
    const color  = filColors[i % filColors.length];
    const pSem   = f.total > 0 ? Math.round((f.total - f.comDss)/ f.total *100) : 100;
    const recPct = f.total > 0 ? Math.round(f.recOk / f.total *100) : 0;
    const simPct = f.total > 0 ? Math.round(f.simOk / f.total *100) : 0;
    const acidPct= f.total > 0 ? Math.round(f.acidMot   / f.total *100) : 0;
    const multPct= f.total > 0 ? Math.round(f.multasMot / f.total *100) : 0;
    const velPct = f.total > 0 ? Math.round(f.excVelMot / f.total *100) : 0;
    const exPerPct  = f.total > 0 ? Math.round(f.examePerOk / f.total *100) : 0;
    const exToxPct  = f.total > 0 ? Math.round(f.exameToxOk / f.total *100) : 0;
    const telCorpPct= f.total > 0 ? Math.round(f.telCorpOk  / f.total *100) : 0;
    grid.innerHTML += `<div class="fc">
      <div class="fc-name">${{f.name}}</div>
      <div class="fc-count" style="color:${{color}}">${{f.total}}</div>
      <div class="situation-bars">
        <div class="sbar ok"><span class="sbar-lbl">Recicl</span><div class="sbar-track"><div class="sbar-fill" style="width:${{recPct}}%;background:#16a34a"></div></div><span class="sbar-cnt" style="color:#16a34a">${{f.recOk}}</span></div>
        <div class="sbar ok"><span class="sbar-lbl">Simul</span><div class="sbar-track"><div class="sbar-fill" style="width:${{simPct}}%;background:#3b7dd8"></div></div><span class="sbar-cnt" style="color:#3b7dd8">${{f.simOk}}</span></div>
        <div class="sbar neg"><span class="sbar-lbl">Acid</span><div class="sbar-track"><div class="sbar-fill" style="width:${{acidPct}}%;background:#dc2626"></div></div><span class="sbar-cnt">${{f.acid}}</span></div>
        <div class="sbar neg"><span class="sbar-lbl">Multas</span><div class="sbar-track"><div class="sbar-fill" style="width:${{multPct}}%;background:#dc2626"></div></div><span class="sbar-cnt">${{f.multas}}</span></div>
        <div class="sbar pend"><span class="sbar-lbl">Vel</span><div class="sbar-track"><div class="sbar-fill" style="width:${{velPct}}%;background:#d97706"></div></div><span class="sbar-cnt" style="color:#d97706">${{f.excVel}}</span></div>
        <div class="sbar ok"><span class="sbar-lbl">Ex.Per</span><div class="sbar-track"><div class="sbar-fill" style="width:${{exPerPct}}%;background:#7c3aed"></div></div><span class="sbar-cnt" style="color:#7c3aed">${{f.examePerOk}}</span></div>
        <div class="sbar ok"><span class="sbar-lbl">Ex.Tox</span><div class="sbar-track"><div class="sbar-fill" style="width:${{exToxPct}}%;background:#a78bfa"></div></div><span class="sbar-cnt" style="color:#a78bfa">${{f.exameToxOk}}</span></div>
        <div class="sbar ok"><span class="sbar-lbl">Tel.Corp</span><div class="sbar-track"><div class="sbar-fill" style="width:${{telCorpPct}}%;background:#0e9cc0"></div></div><span class="sbar-cnt" style="color:#0e9cc0">${{f.telCorpOk}}</span></div>
        <div class="sbar pend"><span class="sbar-lbl">DSS</span><div class="sbar-track" style="background:#dc2626;position:relative;overflow:hidden;"><div class="sbar-fill" style="width:${{f.dssMax>0?Math.round(f.comDss/f.dssMax*100):0}}%;background:#16a34a;position:absolute;left:0;top:0;height:100%;border-radius:3px;transition:width .3s;"></div></div><span class="sbar-cnt" style="color:${{f.comDss===f.dssMax&&f.dssMax>0?'#16a34a':'#dc2626'}};width:auto;min-width:36px;">${{f.comDss}}/${{f.dssMax}}</span></div>
      </div>
      <button class="btn-zoom" onclick="expandirFilial('${{f.name}}')"><i class="fa-solid fa-maximize"></i> Ver Condutores</button>
    </div>`;
  }});
}}

function renderizarGraficos(filiais){{
  if(dssChartInstance)      dssChartInstance.destroy();
  if(filialChartInstance)   filialChartInstance.destroy();
  if(filialAnualChartInst)  filialAnualChartInst.destroy();
  const now         = new Date();
  const mesAtualIdx = now.getMonth();
  const semAtualIdx = Math.min(3, Math.floor((now.getDate() - 1) / 7));
  const monthsShort = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];
  const dssLabels=[], dssData=[], dssBarColors=[], dssBorderColors=[];
  monthsShort.forEach((m, mi) => {{
    [0,1,2,3].forEach(wi => {{
      dssLabels.push(`${{wi+1}}ª ${{m}}`);
      const mesFull = MESES[mi];
      const count = listaAtiva().filter(mot => mot.dssAnual && mot.dssAnual[mesFull] && mot.dssAnual[mesFull][wi]).length;
      const isCurrent = mi === mesAtualIdx && wi === semAtualIdx;
      const isPast    = !isCurrent && (mi < mesAtualIdx || (mi === mesAtualIdx && wi <= semAtualIdx));
      dssData.push(count);
      if(isCurrent) {{ dssBarColors.push('#1a3a6b'); dssBorderColors.push('#1a3a6b'); }}
      else if(isPast) {{
        const total = motoristasDB.length;
        const pct   = total > 0 ? count / total : 0;
        if(pct >= 1.0)      {{ dssBarColors.push('#16a34a'); dssBorderColors.push('#16a34a'); }}
        else if(pct >= 0.5) {{ dssBarColors.push('#3b7dd8'); dssBorderColors.push('#3b7dd8'); }}
        else if(pct >  0)   {{ dssBarColors.push('#d97706'); dssBorderColors.push('#d97706'); }}
        else                {{ dssBarColors.push('#dc2626'); dssBorderColors.push('#dc2626'); }}
      }} else {{ dssBarColors.push('rgba(180,200,230,0.4)'); dssBorderColors.push('rgba(180,200,230,0.6)'); }}
    }});
  }});
  const maxVal = Math.max(...dssData, 1);
  const dynamicH = Math.min(500, Math.max(260, 160 + maxVal * 4));
  const wrap = document.getElementById('dssChartWrap');
  if(wrap) wrap.style.height = dynamicH + 'px';
  const yStep = maxVal <= 5 ? 1 : maxVal <= 20 ? 2 : maxVal <= 50 ? 5 : maxVal <= 100 ? 10 : Math.ceil(maxVal / 10);
  // Dataset de fundo: vermelho fino para semanas passadas sem registro
  const dssDataFundo = dssData.map((v, i) => {{
    const mi = Math.floor(i / 4);
    const wi = i % 4;
    const isCur = mi === mesAtualIdx && wi === semAtualIdx;
    const isPas = !isCur && (mi < mesAtualIdx || (mi === mesAtualIdx && wi <= semAtualIdx));
    return (isPas && v === 0) ? 1 : 0;
  }});
  const dssFundoCores = dssData.map((v, i) => {{
    const mi = Math.floor(i / 4);
    const wi = i % 4;
    const isCur = mi === mesAtualIdx && wi === semAtualIdx;
    const isPas = !isCur && (mi < mesAtualIdx || (mi === mesAtualIdx && wi <= semAtualIdx));
    return (isPas && v === 0) ? 'rgba(220,38,38,0.25)' : 'transparent';
  }});

  dssChartInstance = new Chart(document.getElementById('dssChart'), {{
    type:'bar',
    data:{{ labels:dssLabels, datasets:[
      {{ label:'Fundo Sem Registro', data:dssDataFundo, backgroundColor:dssFundoCores, borderColor:dssFundoCores, borderWidth:0, borderRadius:4, borderSkipped:false }},
      {{ label:'Sessões DSS', data:dssData.slice(), backgroundColor:dssBarColors, borderColor:dssBorderColors, borderWidth:1, borderRadius:4, borderSkipped:false }}
    ]}},
    options:{{
      responsive:true, maintainAspectRatio:false,
      animation:{{ duration:600, easing:'easeOutQuart' }},
      plugins:{{ legend:{{ display:false }},
        tooltip:{{
          callbacks:{{
            title: items => {{ const i=items[0].dataIndex; const mi=Math.floor(i/4); const wi=i%4; return `${{wi+1}}ª semana — ${{MESES[mi]}}`; }},
            label: ctx => {{
              const i=ctx.dataIndex; const mi=Math.floor(i/4); const wi=i%4;
              const isFut = mi > mesAtualIdx || (mi === mesAtualIdx && wi > semAtualIdx);
              if(isFut) return ' Ainda não realizado';
              const real=dssData[i]; const total=listaAtiva().length;
              const pct=total>0?Math.round(real/total*100):0;
              return [` ${{real}} de ${{total}} motorista${{total!==1?'s':''}}`, ` Adesão: ${{pct}}%`];
            }}
          }},
          backgroundColor:'#ffffff',borderColor:'#dde6f4',borderWidth:1,
          titleColor:'#1a3a6b',bodyColor:'#5a6e8a',padding:10,cornerRadius:6
        }}
      }},
      scales:{{
        x:{{ ticks:{{ color:'#5a6e8a',font:{{ size:9 }},maxRotation:45,autoSkip:false, callback(val,i){{ const mi=Math.floor(i/4);const wi=i%4;const isCur=mi===mesAtualIdx&&wi===semAtualIdx;return isCur?'▶ '+this.getLabelForValue(val):this.getLabelForValue(val); }}, color: (ctx) => {{ const mi=Math.floor(ctx.index/4);const wi=ctx.index%4;const isCur=mi===mesAtualIdx&&wi===semAtualIdx;return isCur?'#1a3a6b':'#5a6e8a'; }}, font: (ctx) => {{ const mi=Math.floor(ctx.index/4);const wi=ctx.index%4;const isCur=mi===mesAtualIdx&&wi===semAtualIdx;return {{size:isCur?11:9,weight:isCur?'900':'400'}}; }} }},grid:{{ color:'rgba(180,200,230,0.4)' }} }},
        y:{{ ticks:{{ color:'#5a6e8a',font:{{ size:10 }},stepSize:yStep,callback:v=>Number.isInteger(v)?v:'' }},grid:{{ color:'rgba(180,200,230,0.4)' }},min:0,max:Math.ceil(maxVal*1.15)||1 }}
      }}
    }}
  }});
 filialChartInstance = new Chart(document.getElementById('filialChart'), {{
    type:'bar',
    data:{{ labels:filiais.map(f=>f.name), datasets:[
      {{ label:'Com DSS',  data:filiais.map(f=>f.comDss),           backgroundColor:'#16a34a', borderRadius:3, borderSkipped:false }},
      {{ label:'Sem DSS',  data:filiais.map(f=>f.dssMax-f.comDss), backgroundColor:'rgba(220,38,38,0.12)', borderColor:'rgba(220,38,38,0.3)', borderWidth:1, borderRadius:3, borderSkipped:false }}
    ]}},
    options:{{ responsive:true, maintainAspectRatio:false, plugins:{{ legend:{{ display:false }},
      tooltip:{{
        callbacks:{{
          title: items => items[0].label,
          label: ctx => {{
            const i    = ctx.dataIndex;
            const real = filiais[i].comDss;
            const max  = filiais[i].dssMax;
            const pct  = max > 0 ? Math.round(real / max * 100) : 0;
            const semN = Math.min(3, Math.floor((new Date().getDate()-1)/7)) + 1;
            return [` ${{real}} de ${{max}} motoristas fizeram o DSS`, ` Semana ${{semN}} — Adesão: ${{pct}}%`];
          }}
        }},
        backgroundColor:'#ffffff', borderColor:'#dde6f4', borderWidth:1,
        titleColor:'#1a3a6b', bodyColor:'#5a6e8a', padding:10, cornerRadius:6
      }}
    }},
      scales:{{
        x:{{ stacked:true, ticks:{{ color:'#5a6e8a',font:{{ size:8 }},maxRotation:30,autoSkip:false }},grid:{{ color:'rgba(180,200,230,0.4)' }} }},
        y:{{ stacked:true, ticks:{{ color:'#5a6e8a',font:{{ size:9 }} }},grid:{{ color:'rgba(180,200,230,0.4)' }},min:0 }}
      }}
    }}
  }});

  if(window._statusAnualChartInstance) window._statusAnualChartInstance.destroy();
  const statusLabels   = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];
  const mesAtualIdxS   = new Date().getMonth();
  const totalMotS      = listaAtiva().length;
  const statusDssMes   = MESES.map(mes => {{
    let n = 0;
    listaAtiva().forEach(m => {{ if(dssOkNoMes(m, mes)) n++; }});
    return n;
  }});
  const statusBarCors  = MESES.map((mes, mi) => {{
    const pct     = totalMotS > 0 ? statusDssMes[mi] / totalMotS : 0;
    const isCur   = mi === mesAtualIdxS;
    const isPast  = mi < mesAtualIdxS;
    if(isCur)           return '#1a3a6b';
    if(!isPast)         return 'rgba(180,200,230,0.4)';
    if(pct >= 1.0)      return '#16a34a';
    if(pct >= 0.5)      return '#3b7dd8';
    if(pct >  0)        return '#d97706';
    return '#dc2626';
  }});
  const statusBordCors = MESES.map((mes, mi) => {{
    const pct     = totalMotS > 0 ? statusDssMes[mi] / totalMotS : 0;
    const isCur   = mi === mesAtualIdxS;
    const isPast  = mi < mesAtualIdxS;
    if(isCur)           return '#1a3a6b';
    if(!isPast)         return 'rgba(180,200,230,0.6)';
    if(pct >= 1.0)      return '#16a34a';
    if(pct >= 0.5)      return '#3b7dd8';
    if(pct >  0)        return '#d97706';
    return '#dc2626';
  }});
  const statusFundoCores = MESES.map((mes, mi) => {{
    const isCur  = mi === mesAtualIdxS;
    const isPast = mi < mesAtualIdxS;
    return (isPast && !isCur && statusDssMes[mi] === 0) ? 'rgba(220,38,38,0.25)' : 'transparent';
  }});
  const statusDataFundo = MESES.map((mes, mi) => {{
    const isCur  = mi === mesAtualIdxS;
    const isPast = mi < mesAtualIdxS;
    return (isPast && !isCur && statusDssMes[mi] === 0) ? 1 : 0;
  }});
 // ── Gráfico DSS Anual por Filial ──
  const filialAnualLabels = filiais.map(f => f.name);
  const filialAnualDss    = filiais.map(f => {{
    let total = 0;
    listaAtiva().filter(m => (m.filial||'').toUpperCase() === f.name).forEach(m => {{
      MESES.forEach(mes => {{ if(m.dssAnual && m.dssAnual[mes]) total += m.dssAnual[mes].filter(Boolean).length; }});
    }});
    return total;
  }});
  const filialAnualMax = filiais.map(f => f.total * MESES.length * 4);
  const filialAnualFalta = filiais.map((f, i) => Math.max(0, filialAnualMax[i] - filialAnualDss[i]));
  const filialAnualBarCors = filiais.map((f, i) => {{
    const pct = filialAnualMax[i] > 0 ? filialAnualDss[i] / filialAnualMax[i] : 0;
    if(pct >= 1.0)      return '#16a34a';
    if(pct >= 0.5)      return '#3b7dd8';
    if(pct >  0)        return '#d97706';
    return '#dc2626';
  }});
  filialAnualChartInst = new Chart(document.getElementById('filialAnualChart'), {{
    type: 'bar',
    data: {{ labels: filialAnualLabels, datasets: [
      {{ label: 'Sessões realizadas', data: filialAnualDss,   backgroundColor: filialAnualBarCors, borderRadius: 3, borderSkipped: false }},
      {{ label: 'Sessões em falta',   data: filialAnualFalta, backgroundColor: 'rgba(220,38,38,0.12)', borderColor: 'rgba(220,38,38,0.3)', borderWidth: 1, borderRadius: 3, borderSkipped: false }}
    ]}},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{ legend: {{ display: false }},
        tooltip: {{
          callbacks: {{
            title: items => items[0].label,
            label: ctx => {{
              const i     = ctx.dataIndex;
              const real  = filialAnualDss[i];
              const maxi  = filialAnualMax[i];
              const pct   = maxi > 0 ? Math.round(real / maxi * 100) : 0;
              return [` ${{real}} sessões realizadas de ${{maxi}}`, ` Adesão anual: ${{pct}}%`];
            }}
          }},
          backgroundColor:'#ffffff', borderColor:'#dde6f4', borderWidth:1,
          titleColor:'#1a3a6b', bodyColor:'#5a6e8a', padding:10, cornerRadius:6
        }}
      }},
      scales: {{
        x: {{ stacked: true, ticks: {{ color:'#5a6e8a', font:{{ size:8 }}, maxRotation:30, autoSkip:false }}, grid: {{ color:'rgba(180,200,230,0.4)' }} }},
        y: {{ stacked: true, ticks: {{ color:'#5a6e8a', font:{{ size:9 }}, callback: v => Number.isInteger(v) ? v : '' }}, grid: {{ color:'rgba(180,200,230,0.4)' }}, min: 0 }}
      }}
    }}
  }});

  window._statusAnualChartInstance = new Chart(document.getElementById('statusAnualChart'), {{
    type:'bar',
    data:{{ labels:statusLabels, datasets:[
      {{ label:'Fundo Sem Registro', data:statusDataFundo, backgroundColor:statusFundoCores, borderColor:statusFundoCores, borderWidth:0, borderRadius:4, borderSkipped:false }},
      {{ label:'DSS ok no mês', data:statusDssMes, backgroundColor:statusBarCors, borderColor:statusBordCors, borderWidth:1, borderRadius:4, borderSkipped:false }}
    ]}},
    options:{{
      responsive:true, maintainAspectRatio:false,
      animation:{{ duration:600, easing:'easeOutQuart' }},
      plugins:{{ legend:{{ display:false }},
        tooltip:{{
          callbacks:{{
            title: items => statusLabels[items[0].dataIndex] + ' ' + new Date().getFullYear(),
            label: ctx => {{
              const mi = ctx.dataIndex;
              const count = statusDssMes[mi];
              const total = totalMotS;
              const pct = total > 0 ? Math.round(count/total*100) : 0;
              if(mi > mesAtualIdxS) return ' Ainda não realizado';
              return [` ${{count}} de ${{total}} motorista${{total!==1?'s':''}}`, ` Adesão: ${{pct}}%`];
            }}
          }},
          backgroundColor:'#ffffff',borderColor:'#dde6f4',borderWidth:1,
          titleColor:'#1a3a6b',bodyColor:'#5a6e8a',padding:10,cornerRadius:6
        }}
      }},
      scales:{{
        x:{{ ticks:{{ color:'#5a6e8a',font:{{size:10}} }},grid:{{ color:'rgba(180,200,230,0.4)' }} }},
        y:{{ ticks:{{ color:'#5a6e8a',font:{{size:10}},stepSize:1,callback:v=>Number.isInteger(v)?v:'' }},grid:{{ color:'rgba(180,200,230,0.4)' }},min:0 }}
      }}
    }}
  }});
}}

async function adicionarNovoMotorista(){{
  let cpf      = document.getElementById('addCpf').value.trim();
  const nome   = document.getElementById('addNome').value.toUpperCase().trim();
  const filial = document.getElementById('addFilial').value.toUpperCase().trim();
  if(!nome || !filial){{ toast('Preencha ao menos Nome e Filial.', 'erro'); return; }}
  if(!cpf){{ cpf = `SEMCPF-${{Date.now()}}-${{Math.random().toString(36).slice(2,7)}}`; }}
  if(motoristasDB.some(m => m.cpf === cpf)){{ toast('Já existe um motorista com este CPF.', 'erro'); return; }}
  const novo = {{
    cpf, nome, filial, telefone:'', email:'', foto:'',
    reciclagem: document.getElementById('addRec').value,
    simulador:  document.getElementById('addSim').value,
    excesso:0, multas:0, acidentes:0,
    obsAcidente:'', obsMultas:'', obsGerais:'', obsReciclagem:'', obsSimulador:'',
    cnh:'', validadeCnh:'', admissao:'',
    examePeriodico:'', exameToxicologico:'',
    pontuacaoCnh:0, vencimentoCnhMopp:'', entregaUniforme:'PENDENTE',
    telefoneCorporativo:'NÃO', numeroLinha:'', modelo:'', imei:'',
    afastado:'NÃO', obsAfastado:'',
    desligado:'NÃO', obsDesligamento:'',
    dssAnual: gerarMatrizDssEmBranco()
  }};
  mostrarSpinner(true);
  try{{
    motoristasDB.push(novo);
    const res = await salvarTodosNaSheetsAPI(motoristasDB);
    if(res.ok){{
      document.getElementById('addCpf').value = '';
      document.getElementById('addNome').value = '';
      document.getElementById('addFilial').value = '';
      document.getElementById('formBody').classList.remove('open');
      document.getElementById('btnToggleForm').classList.remove('open');
      atualizarDashboardCompleto();
      toast('Condutor inserido e salvo no Google Sheets!');
    }} else {{
      motoristasDB.pop();
      toast(res.erro || 'Erro ao inserir.', 'erro');
    }}
  }} catch(e){{ toast('Falha de conexão: ' + e.message, 'erro'); }}
  finally{{ mostrarSpinner(false); }}
}}

function expandirFilial(nomeFilial){{
  filialModalAtiva = nomeFilial;
  document.getElementById('mUnidadeName').textContent = nomeFilial;
  const listagem = listaAtiva().filter(m => obterChaveFilial(m) === nomeFilial.toUpperCase());
  document.getElementById('mTotalDrivers').textContent = listagem.length;
  const totalDss = listagem.reduce((acc,m) => acc + contarDssSessoes(m), 0);
  document.getElementById('mWithDss').textContent = totalDss;
  document.getElementById('mRecOk').textContent = listagem.filter(m => reciclagemStatus(m) === 'OK').length;
  document.getElementById('mSimOk').textContent = listagem.filter(m => simuladorStatus(m) === 'OK').length;
  document.getElementById('mExcVel').textContent = listagem.reduce((acc,m) => acc + Math.max(0, parseInt(m.excesso)||0), 0);
  document.getElementById('mMultas').textContent = listagem.reduce((acc,m) => acc + Math.max(0, parseInt(m.multas)||0), 0);
  document.getElementById('mAcidentes').textContent = listagem.reduce((acc,m) => acc + Math.max(0, parseInt(m.acidentes)||0), 0);
  document.getElementById('mExamePerOk').textContent = listagem.filter(m => exameOk(m.examePeriodico, m.examePeriodicoValidadeMeses)).length;
  document.getElementById('mExameToxOk').textContent = listagem.filter(m => exameOk(m.exameToxicologico, m.exameToxicologicoValidadeMeses)).length;
  document.getElementById('mTelCorpOk').textContent = listagem.filter(m => m.telefoneCorporativo === 'SIM').length;
  const tbody = document.getElementById('mDriversTableBody');
  tbody.innerHTML = '';
  if(listagem.length === 0){{
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px;color:#666;">Nenhum motorista cadastrado nesta filial.</td></tr>';
  }} else {{
    listagem.forEach(m => {{
      console.log('DSS de', m.nome, JSON.stringify(m.dssAnual));
      let dssAno = 0;
      MESES.forEach(mes => {{
        if(m.dssAnual && m.dssAnual[mes]) {{
          m.dssAnual[mes].forEach(s => {{ if(s === true || s === 1) dssAno++; }});
        }}
      }});
      const dssMax = 48;
      const dssPct = Math.round(dssAno / dssMax * 100);
      const dssCor = dssPct >= 100 ? '#16a34a' : dssPct >= 50 ? '#3b7dd8' : dssPct > 0 ? '#d97706' : '#dc2626';
      tbody.innerHTML += `<tr class="driver-row" onclick="abrirFichaMotorista('${{m.cpf}}')">
        <td><div class="m-name">${{m.nome}}</div><div class="m-cpf">CPF: ${{m.cpf}}</div></td>
        <td style="text-align:center"><span class="m-count-badge" style="color:${{dssCor}};border-color:${{dssCor}};background:${{dssCor}}18;">${{dssAno}}/${{dssMax}}</span></td>
        <td><span class="m-badge ${{reciclagemStatus(m)==='OK'?'ok':'pend'}}">${{reciclagemStatus(m)}}</span></td>
        <td><span class="m-badge ${{simuladorStatus(m)==='OK'?'ok':'pend'}}">${{simuladorStatus(m)}}</span></td>
        <td style="text-align:center"><span class="m-count-badge">${{m.excesso}}</span></td>
        <td style="text-align:center"><span class="m-count-badge">${{m.multas}}</span></td>
        <td style="text-align:center"><span class="m-count-badge">${{m.acidentes}}</span></td>
        <td><span class="m-badge ${{exameOk(m.examePeriodico, m.examePeriodicoValidadeMeses)?'ok':'pend'}}">${{exameOk(m.examePeriodico, m.examePeriodicoValidadeMeses)?'OK':'PENDENTE'}}</span></td>
        <td><span class="m-badge ${{exameOk(m.exameToxicologico, m.exameToxicologicoValidadeMeses)?'ok':'pend'}}">${{exameOk(m.exameToxicologico, m.exameToxicologicoValidadeMeses)?'OK':'PENDENTE'}}</span></td>
        <td><span class="m-badge ${{m.telefoneCorporativo==='SIM'?'ok':'pend'}}">${{m.telefoneCorporativo||'NÃO'}}</span></td>
      </tr>`;
    }});
  }}
  voltarSidebarFilialMobile();
  document.getElementById('filialModal').style.display = 'flex';
}}

function fecharJanelaFilial(){{
  document.getElementById('filialModal').style.display = 'none';
  voltarSidebarFilialMobile();
}}
function filtrarTabelaFilial(){{
  const q = document.getElementById('filialSearchInput').value.toLowerCase();
  document.querySelectorAll('#mDriversTableBody tr').forEach(tr => {{
    tr.style.display = tr.textContent.toLowerCase().includes(q) ? '' : 'none';
  }});
}}
function isMobileView(){{ return window.innerWidth <= 768; }}

const FILIAL_INDICADOR_LABEL = {{
  todos:'Todos os Motoristas', dss:'DSS Realizados (Ano)', reciclagem:'Reciclagem OK',
  simulador:'Simulador OK', excesso:'Excesso de Velocidade', multas:'Total de Multas',
  acidentes:'Total de Acidentes', examePeriodico:'Exame Periódico OK',
  exameToxicologico:'Exame Toxicológico OK', telefoneCorporativo:'Telefone Corporativo OK'
}};

function filtrarFilialPorIndicador(tipo){{
  document.querySelectorAll('#mDriversTableBody tr').forEach(tr => {{
    tr.style.display = '';
  }});
  document.getElementById('filialSearchInput').value = '';
  const listagem = listaAtiva().filter(m => obterChaveFilial(m) === (filialModalAtiva||'').toUpperCase());
  const filtrados = listagem.filter(m => {{
   if(tipo==='todos')     return true;
    if(tipo==='dss')       return contarDssSessoes(m) > 0;
    if(tipo==='reciclagem') return reciclagemStatus(m) === 'OK';
    if(tipo==='simulador') return simuladorStatus(m) === 'OK';
    if(tipo==='excesso')   return Math.max(0,parseInt(m.excesso)||0)   > 0;
    if(tipo==='multas')    return Math.max(0,parseInt(m.multas)||0)    > 0;
    if(tipo==='acidentes') return Math.max(0,parseInt(m.acidentes)||0) > 0;
    if(tipo==='examePeriodico')      return exameOk(m.examePeriodico, m.examePeriodicoValidadeMeses);
    if(tipo==='exameToxicologico')   return exameOk(m.exameToxicologico, m.exameToxicologicoValidadeMeses);
    if(tipo==='telefoneCorporativo') return m.telefoneCorporativo === 'SIM';
    return true;
  }});
  const cpfsFiltrados = new Set(filtrados.map(m => m.cpf));
  document.querySelectorAll('#mDriversTableBody tr').forEach(tr => {{
    const cpfCell = tr.querySelector('.m-cpf');
    if(cpfCell){{
      const cpf = cpfCell.textContent.replace('CPF: ','').trim();
      tr.style.display = cpfsFiltrados.has(cpf) ? '' : 'none';
    }}
  }});

  // ── Versão mobile: mostra lista de cards com botão Voltar ──
  renderizarListaMobileFilial(filtrados, FILIAL_INDICADOR_LABEL[tipo] || 'Motoristas');
  if(isMobileView()){{
    document.getElementById('filialSidebar').classList.add('mobile-hidden');
    document.getElementById('filialTableContainer').classList.add('mobile-hidden');
    document.getElementById('filialMobileBackbar').classList.add('show');
    document.getElementById('filialMobileList').classList.add('show');
  }}
}}

function renderizarListaMobileFilial(lista, titulo){{
  const tituloEl = document.getElementById('filialMobileTitulo');
  if(tituloEl) tituloEl.textContent = `${{titulo}} (${{lista.length}})`;
  const cont = document.getElementById('filialMobileList');
  if(!cont) return;
  if(lista.length === 0){{
    cont.innerHTML = `<div class="empty-state"><i class="fa-solid fa-magnifying-glass"></i><p>Nenhum motorista encontrado.</p></div>`;
    return;
  }}
  cont.innerHTML = lista.map(m => {{
    const avatar = `<img src="${{m.foto || AVATAR_PADRAO}}" alt="">`;
    const recOk  = reciclagemStatus(m) === 'OK';
    const simOk  = simuladorStatus(m) === 'OK';
    return `<div class="driver-mini-card ${{(recOk && simOk) ? 'card-ok' : 'card-pend'}}" onclick="irParaFichaViaFilialMobile('${{m.cpf}}')" title="Abrir ficha de ${{m.nome}}">
      <div class="dmc-top"><div class="dmc-avatar">${{avatar}}</div><div class="dmc-info"><div class="dmc-nome">${{m.nome}}</div><div class="dmc-filial">${{m.filial||'—'}}</div></div></div>
      <div class="dmc-cpf">${{m.cpf}}</div>
      <div class="dmc-badges">
        <span class="dmc-badge ${{recOk?'ok':'pend'}}"><i class="fa-solid fa-recycle"></i> ${{recOk?'Reciclagem OK':'Reciclagem Pend.'}}</span>
        <span class="dmc-badge ${{simOk?'ok':'pend'}}"><i class="fa-solid fa-car-side"></i> ${{simOk?'Simulador OK':'Simulador Pend.'}}</span>
      </div>
    </div>`;
  }}).join('');
}}

function irParaFichaViaFilialMobile(cpf){{
  fichaOrigemModal = 'filialMobile';
  abrirFichaMotorista(cpf);
}}

function voltarSidebarFilialMobile(){{
  const sb  = document.getElementById('filialSidebar');
  const tb  = document.getElementById('filialTableContainer');
  const bar = document.getElementById('filialMobileBackbar');
  const lst = document.getElementById('filialMobileList');
  if(sb)  sb.classList.remove('mobile-hidden');
  if(tb)  tb.classList.remove('mobile-hidden');
  if(bar) bar.classList.remove('show');
  if(lst) lst.classList.remove('show');
}}
function filtrarTabelaFilial(){{
  const q = document.getElementById('filialSearchInput').value.toLowerCase();
  document.querySelectorAll('#mDriversTableBody tr').forEach(tr => {{
    tr.style.display = tr.textContent.toLowerCase().includes(q) ? '' : 'none';
  }});
}}

function abrirFichaMotorista(cpf){{
  const m = motoristasDB.find(x => x.cpf === cpf);
  if(!m) return;
  motoristaEmEdicaoCpf = cpf;
  fotoTemporariaBase64 = m.foto || null;
  houveEdicaoNaoSalva  = false;
  const avatarHtml = `<img src="${{m.foto || AVATAR_PADRAO}}" id="profilePreviewImg">`;
  let matrizHtml = '';
  MESES.forEach(mes => {{
    const semanas = m.dssAnual[mes] || [false,false,false,false];
    matrizHtml += `<div class="month-dss-box">
      <div class="month-name-lbl">${{mes}}</div>
      <div class="weeks-flex">
        ${{[0,1,2,3].map(i => `<label class="week-checkbox-label"><span>${{i+1}}ªS</span><input type="checkbox" id="dss-${{mes}}-${{i}}" ${{semanas[i]?'checked':''}}></label>`).join('')}}
      </div></div>`;
  }});
  const esc = s => (s||'').replace(/"/g,'&quot;');
  // Status efetivo: se venceu, volta para PENDENTE automaticamente (a data NÃO é apagada)
  const recEfetivo  = reciclagemStatus(m);
  const simEfetivo  = simuladorStatus(m);
  const gestEfetivo = gestimeStatus(m);
  const acidCor = (m.acidentes||0) > 0 ? '#dc2626' : '#16a34a';
  const acidBg  = (m.acidentes||0) > 0 ? '#fff5f5' : '#f0fef4';
  const acidBd  = (m.acidentes||0) > 0 ? '#fca5a5' : '#86efac';
  const multCor = (m.multas||0) > 0 ? '#dc2626' : '#16a34a';
  const multBg  = (m.multas||0) > 0 ? '#fff5f5' : '#f0fef4';
  const multBd  = (m.multas||0) > 0 ? '#fca5a5' : '#86efac';
  const velCor  = (m.excesso||0) > 0 ? '#dc2626' : '#16a34a';
  const velBg   = (m.excesso||0) > 0 ? '#fff5f5' : '#f0fef4';
  const velBd   = (m.excesso||0) > 0 ? '#fca5a5' : '#86efac';
  const svRec   = statusVencimento(m.reciclagemData, m.reciclagemValidadeMeses);
  const svSim   = statusVencimento(m.simuladorData, m.simuladorValidadeMeses);
  const svExPer = statusVencimento(m.examePeriodico, m.examePeriodicoValidadeMeses);
  const svExTox = statusVencimento(m.exameToxicologico, m.exameToxicologicoValidadeMeses);
  const svGest  = statusVencimento(m.gestimeData, m.gestimeValidadeMeses);
  const svCnh   = statusVencimentoData(m.validadeCnh);
  const svMopp  = statusVencimentoData(m.vencimentoCnhMopp);
  const bgPorCor = c => c==='#16a34a'?'#f0fef4':c==='#d97706'?'#fffbeb':c==='#dc2626'?'#fff5f5':'#f8fafd';
  const bdPorCor = c => c==='#16a34a'?'#86efac':c==='#d97706'?'#fde68a':c==='#dc2626'?'#fca5a5':'#dde6f4';
  const pontosCnh = m.pontuacaoCnh||0;
  const pontosCor = pontosCnh > 0 ? '#dc2626' : '#16a34a';
  const pontosBg  = pontosCnh > 0 ? '#fff5f5' : '#f0fef4';
  const pontosBd  = pontosCnh > 0 ? '#fca5a5' : '#86efac';
  const afastadoSimFlag = m.afastado === 'SIM';
  const afastadoCor = afastadoSimFlag ? '#dc2626' : '#16a34a';
  const afastadoBg  = afastadoSimFlag ? '#fff5f5' : '#f0fef4';
  const afastadoBd  = afastadoSimFlag ? '#fca5a5' : '#86efac';
  const uniformeOkFlag = m.entregaUniforme === 'OK';
  const uniformeCor = uniformeOkFlag ? '#16a34a' : '#dc2626';
  const uniformeBg  = uniformeOkFlag ? '#f0fef4' : '#fff5f5';
  const uniformeBd  = uniformeOkFlag ? '#86efac' : '#fca5a5';
  let highlightSeg='', highlightDss='', badgeSegHtml='', badgeDssHtml='';
  if(fichaOrigemModal === 'kpi'){{
    if(kpiTipoAtual==='excesso')  {{ highlightSeg=' card-highlight-vel';  badgeSegHtml='<span style="margin-left:auto;font-size:8px;font-weight:800;background:#fee2e2;color:#b91c1c;border:1px solid #fca5a5;padding:2px 8px;border-radius:20px;"><i class=\\"fa-solid fa-gauge-high\\" style=\\"margin-right:3px\\"></i>Excesso de Velocidade</span>'; }}
    if(kpiTipoAtual==='multas')   {{ highlightSeg=' card-highlight-mul';  badgeSegHtml='<span style="margin-left:auto;font-size:8px;font-weight:800;background:#fff3e0;color:#c2410c;border:1px solid #fbbf7a;padding:2px 8px;border-radius:20px;"><i class=\\"fa-solid fa-file-circle-xmark\\" style=\\"margin-right:3px\\"></i>Multas</span>'; }}
    if(kpiTipoAtual==='acidentes'){{ highlightSeg=' card-highlight-acid'; badgeSegHtml='<span style="margin-left:auto;font-size:8px;font-weight:800;background:#fee2e2;color:#b91c1c;border:1px solid #fca5a5;padding:2px 8px;border-radius:20px;"><i class=\\"fa-solid fa-car-burst\\" style=\\"margin-right:3px\\"></i>Acidentes</span>'; }}
    if(kpiTipoAtual==='comDss')   {{ highlightDss=' card-highlight-dss';  badgeDssHtml='<span style="margin-left:auto;font-size:8px;font-weight:800;background:#dcfce7;color:#15803d;border:1px solid #86efac;padding:2px 8px;border-radius:20px;"><i class=\\"fa-solid fa-circle-check\\" style=\\"margin-right:3px\\"></i>DSS Ok</span>'; }}
    if(kpiTipoAtual==='semDss')   {{ highlightDss=' card-highlight-dss';  badgeDssHtml='<span style="margin-left:auto;font-size:8px;font-weight:800;background:#fef9c3;color:#a16207;border:1px solid #fde68a;padding:2px 8px;border-radius:20px;"><i class=\\"fa-solid fa-clock\\" style=\\"margin-right:3px\\"></i>Pendente DSS</span>'; }}
  }}
  document.getElementById('driverProfileContent').innerHTML = `
    <div class="profile-card-left card-condutor">
      <span class="card-stripe"></span>
      <div class="profile-card-left-body">
        <div>
          <div class="avatar-outer" id="avatarOuterFicha">
            <div class="avatar-wrapper" id="avatarWrapperFicha" onclick="abrirMenuAvatar(event)">${{avatarHtml}}<div class="upload-hint">Alterar Foto</div></div>
            <div class="avatar-menu" id="avatarMenu">
              <button type="button" id="avatarMenuCarregar" onclick="event.stopPropagation(); dispararUploadFoto();"><i class="fa-solid fa-upload"></i> Carregar Imagem</button>
              <button type="button" id="avatarMenuSubstituir" style="display:none;" onclick="event.stopPropagation(); dispararUploadFoto();"><i class="fa-solid fa-rotate"></i> Substituir Imagem</button>
              <button type="button" id="avatarMenuExcluir" style="display:none;" onclick="event.stopPropagation(); excluirFotoAtual();"><i class="fa-solid fa-trash"></i> Excluir Imagem</button>
            </div>
          </div>
          <div class="form-group" style="width:100%;margin-top:15px;"><label>Nome do Condutor</label><input type="text" id="editNome" value="${{esc(m.nome)}}"></div>
          <div class="form-group" style="width:100%;margin-top:10px;"><label>Filial Base</label><input type="text" id="editFilial" value="${{esc(m.filial)}}"></div>
          <div style="display:flex;flex-direction:column;gap:8px;margin-top:16px;">
            <div style="background:${{afastadoBg}};border:1.5px solid ${{afastadoBd}};border-radius:8px;padding:10px 14px;display:flex;justify-content:space-between;align-items:center;">
              <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-user-slash" style="color:${{afastadoCor}};margin-right:6px"></i>Afastado</span>
              <span style="font-size:20px;font-weight:900;color:${{afastadoCor}};line-height:1;">${{m.afastado||'NÃO'}}</span>
            </div>
            <div style="background:${{bgPorCor(svCnh.cor)}};border:1.5px solid ${{bdPorCor(svCnh.cor)}};border-radius:8px;padding:10px 14px;display:flex;flex-direction:column;gap:2px;">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-id-card" style="color:${{svCnh.cor}};margin-right:6px"></i>Validade CNH</span>
                <span style="font-size:15px;font-weight:900;color:${{svCnh.cor}};letter-spacing:.5px;">${{rotuloVenc(svCnh)}}</span>
              </div>
              <span style="font-size:13px;font-weight:700;color:${{svCnh.cor}};">${{svCnh.venc ? svCnh.label : 'PENDENTE'}}</span>
            </div>
            <div style="background:${{bgPorCor(svMopp.cor)}};border:1.5px solid ${{bdPorCor(svMopp.cor)}};border-radius:8px;padding:10px 14px;display:flex;flex-direction:column;gap:2px;">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-id-card-clip" style="color:${{svMopp.cor}};margin-right:6px"></i>Validade MOPP</span>
                <span style="font-size:15px;font-weight:900;color:${{svMopp.cor}};letter-spacing:.5px;">${{rotuloVenc(svMopp)}}</span>
              </div>
              <span style="font-size:13px;font-weight:700;color:${{svMopp.cor}};">${{svMopp.venc ? svMopp.label : 'PENDENTE'}}</span>
            </div>
            <div style="background:${{bgPorCor(svExTox.cor)}};border:1.5px solid ${{bdPorCor(svExTox.cor)}};border-radius:8px;padding:10px 14px;display:flex;flex-direction:column;gap:2px;">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-vial" style="color:${{svExTox.cor}};margin-right:6px"></i>Exame Toxicológico</span>
                <span style="font-size:15px;font-weight:900;color:${{svExTox.cor}};letter-spacing:.5px;">${{rotuloVenc(svExTox)}}</span>
              </div>
              <span style="font-size:13px;font-weight:700;color:${{svExTox.cor}};">${{svExTox.venc ? svExTox.label : 'PENDENTE'}}</span>
            </div>
            <div style="background:${{bgPorCor(svExPer.cor)}};border:1.5px solid ${{bdPorCor(svExPer.cor)}};border-radius:8px;padding:10px 14px;display:flex;flex-direction:column;gap:2px;">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-stethoscope" style="color:${{svExPer.cor}};margin-right:6px"></i>Exame Periódico</span>
                <span style="font-size:15px;font-weight:900;color:${{svExPer.cor}};letter-spacing:.5px;">${{rotuloVenc(svExPer)}}</span>
              </div>
              <span style="font-size:13px;font-weight:700;color:${{svExPer.cor}};">${{svExPer.venc ? svExPer.label : 'PENDENTE'}}</span>
            </div>
            <div style="background:${{bgPorCor(svRec.cor)}};border:1.5px solid ${{bdPorCor(svRec.cor)}};border-radius:8px;padding:10px 14px;display:flex;flex-direction:column;gap:2px;">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-recycle" style="color:${{svRec.cor}};margin-right:6px"></i>Reciclagem</span>
                <span style="font-size:15px;font-weight:900;color:${{svRec.cor}};letter-spacing:.5px;">${{rotuloVenc(svRec)}}</span>
              </div>
              <span style="font-size:13px;font-weight:700;color:${{svRec.cor}};">${{svRec.venc ? svRec.label : 'PENDENTE'}}</span>
            </div>
            <div style="background:${{bgPorCor(svSim.cor)}};border:1.5px solid ${{bdPorCor(svSim.cor)}};border-radius:8px;padding:10px 14px;display:flex;flex-direction:column;gap:2px;">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-car-side" style="color:${{svSim.cor}};margin-right:6px"></i>Simulador</span>
                <span style="font-size:15px;font-weight:900;color:${{svSim.cor}};letter-spacing:.5px;">${{rotuloVenc(svSim)}}</span>
              </div>
              <span style="font-size:13px;font-weight:700;color:${{svSim.cor}};">${{svSim.venc ? svSim.label : 'PENDENTE'}}</span>
            </div>
            <div style="background:${{bgPorCor(svGest.cor)}};border:1.5px solid ${{bdPorCor(svGest.cor)}};border-radius:8px;padding:10px 14px;display:flex;flex-direction:column;gap:2px;">
              <div style="display:flex;justify-content:space-between;align-items:center;">
                <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-clipboard-check" style="color:${{svGest.cor}};margin-right:6px"></i>Gestime</span>
                <span style="font-size:15px;font-weight:900;color:${{svGest.cor}};letter-spacing:.5px;">${{rotuloVenc(svGest)}}</span>
              </div>
              <span style="font-size:13px;font-weight:700;color:${{svGest.cor}};">${{svGest.venc ? svGest.label : 'PENDENTE'}}</span>
            </div>
            <div style="background:${{uniformeBg}};border:1.5px solid ${{uniformeBd}};border-radius:8px;padding:10px 14px;display:flex;justify-content:space-between;align-items:center;">
              <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-shirt" style="color:${{uniformeCor}};margin-right:6px"></i>Entrega de Uniforme</span>
              <span style="font-size:18px;font-weight:900;color:${{uniformeCor}};line-height:1;">${{m.entregaUniforme||'PENDENTE'}}</span>
            </div>
            <div style="background:${{pontosBg}};border:1.5px solid ${{pontosBd}};border-radius:8px;padding:10px 14px;display:flex;justify-content:space-between;align-items:center;">
              <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-id-card-clip" style="color:${{pontosCor}};margin-right:6px"></i>Pontuação CNH</span>
              <span style="font-size:20px;font-weight:900;color:${{pontosCor}};line-height:1;">${{pontosCnh}} pts</span>
            </div>
            <div style="background:${{acidBg}};border:1.5px solid ${{acidBd}};border-radius:8px;padding:10px 14px;display:flex;justify-content:space-between;align-items:center;">
              <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-car-burst" style="color:${{acidCor}};margin-right:6px"></i>Acidentes</span>
              <span style="font-size:32px;font-weight:900;color:${{acidCor}};line-height:1;">${{m.acidentes||0}}</span>
            </div>
            <div style="background:${{multBg}};border:1.5px solid ${{multBd}};border-radius:8px;padding:10px 14px;display:flex;justify-content:space-between;align-items:center;">
              <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-file-circle-xmark" style="color:${{multCor}};margin-right:6px"></i>Multas</span>
              <span style="font-size:32px;font-weight:900;color:${{multCor}};line-height:1;">${{m.multas||0}}</span>
            </div>
            <div style="background:${{velBg}};border:1.5px solid ${{velBd}};border-radius:8px;padding:10px 14px;display:flex;justify-content:space-between;align-items:center;">
              <span style="font-size:12px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;"><i class="fa-solid fa-gauge-high" style="color:${{velCor}};margin-right:6px"></i>Exc. Velocidade</span>
              <span style="font-size:32px;font-weight:900;color:${{velCor}};line-height:1;">${{m.excesso||0}}</span>
            </div>
          </div>
        </div>
        <button onclick="gerarFichaPdf('${{m.cpf}}')" style="background:#fff0f8;color:#1a4fa0;border:1px solid #b0c8e8;padding:8px;border-radius:7px;font-size:10px;font-weight:700;text-transform:uppercase;cursor:pointer;margin-top:6px;display:flex;align-items:center;justify-content:center;gap:6px;width:100%;transition:.18s;" onmouseover="this.style.background='#1a4fa0';this.style.color='#fff'" onmouseout="this.style.background='#fff0f8';this.style.color='#1a4fa0'">
          <i class="fa-solid fa-file-pdf"></i> Baixar Ficha em PDF
        </button>
        <div style="margin-top:12px;padding-top:12px;border-top:1px dashed #d0d8e8;display:flex;flex-direction:column;gap:6px;">
          ${{m.desligado === 'SIM'
            ? `<button class="btn-reativar-driver" onclick="reativarMotoristaAtual('${{m.cpf}}')"><i class="fa-solid fa-rotate-left"></i> Reativar Motorista</button>`
            : `<button class="btn-desligar-driver" onclick="desligarMotoristaAtual('${{m.cpf}}')"><i class="fa-solid fa-user-xmark"></i> Desligar Motorista</button>`
          }}
          <input type="text" id="editObsDesligamento" class="obs-input" placeholder="Observação do desligamento" value="${{esc(m.obsDesligamento)}}">
        </div>
        <button class="btn-delete-driver" onclick="deletarMotoristaAtual('${{m.cpf}}','${{esc(m.nome)}}')">
          <i class="fa-solid fa-trash-can"></i> Excluir Condutor permanentemente
        </button>
      </div>
    </div>
    <div class="profile-details-right">
      <div class="info-section-box card-docs">
        <span class="card-stripe"></span>
        <div class="card-body">
          <div class="info-block-title"><i class="fa-solid fa-address-card"></i> Documentação</div>
          <div class="meta-grid">
            <div class="meta-item"><label>Nº Registro CNH</label><input type="text" id="editCnh" value="${{esc(m.cnh)}}"></div>
            <div class="meta-item">
              <label>Validade CNH</label>
              <input type="date" id="editValidadeCnh" value="${{esc(m.validadeCnh)}}">
              <div style="font-size:10px;font-weight:700;color:${{svCnh.cor}};margin-top:2px;">${{svCnh.label}}</div>
            </div>
            <div class="meta-item"><label>Data Admissão</label><input type="date" id="editAdmissao" value="${{esc(m.admissao)}}"></div>
          </div>
        </div>
      </div>
      <div class="info-section-box card-exames">
        <span class="card-stripe"></span>
        <div class="card-body">
          <div class="info-block-title"><i class="fa-solid fa-stethoscope"></i> Exames & Complementares</div>
          <div class="meta-grid">
            <div class="meta-item">
              <label>Afastado</label>
              <select id="editAfastado">
                <option value="NÃO" ${{m.afastado==='NÃO'?'selected':''}}>NÃO</option>
                <option value="SIM" ${{m.afastado==='SIM'?'selected':''}}>SIM</option>
              </select>
              <input type="text" id="editObsAfastado" class="obs-input" style="margin-top:4px;" value="${{esc(m.obsAfastado)}}" placeholder="Obs de Afastamento">
            </div>
            <div class="meta-item">
              <label>Validade MOPP</label>
              <input type="date" id="editVencimentoCnhMopp" value="${{esc(m.vencimentoCnhMopp)}}">
              <div style="font-size:10px;font-weight:700;color:${{svMopp.cor}};margin-top:2px;">${{svMopp.label}}</div>
            </div>
            <div class="meta-item">
              <label>Exame Toxicológico</label>
              <div style="display:flex;gap:4px;">
                <input type="date" id="editExameToxicologico" value="${{esc(m.exameToxicologico)}}" style="flex:1;" oninput="checarCamposValidade('editExameToxicologico','editExameToxicologicoValidadeMeses')">
                <input type="number" id="editExameToxicologicoValidadeMeses" value="${{m.exameToxicologicoValidadeMeses||0}}" min="0" style="width:52px;" title="Válido por (meses)" oninput="checarCamposValidade('editExameToxicologico','editExameToxicologicoValidadeMeses')">
              </div>
              <div style="font-size:10px;font-weight:700;color:${{svExTox.cor}};margin-top:2px;">${{svExTox.label}}</div>
            </div>
            <div class="meta-item">
              <label>Exame Periódico</label>
              <div style="display:flex;gap:4px;">
                <input type="date" id="editExamePeriodico" value="${{esc(m.examePeriodico)}}" style="flex:1;" oninput="checarCamposValidade('editExamePeriodico','editExamePeriodicoValidadeMeses')">
                <input type="number" id="editExamePeriodicoValidadeMeses" value="${{m.examePeriodicoValidadeMeses||0}}" min="0" style="width:52px;" title="Válido por (meses)" oninput="checarCamposValidade('editExamePeriodico','editExamePeriodicoValidadeMeses')">
              </div>
              <div style="font-size:10px;font-weight:700;color:${{svExPer.cor}};margin-top:2px;">${{svExPer.label}}</div>
            </div>
            <div class="meta-item"><label>Pontuação CNH</label><input type="number" id="editPontuacaoCnh" value="${{m.pontuacaoCnh||0}}"></div>
            <div class="meta-item">
              <label>Entrega de Uniforme</label>
              <select id="editEntregaUniforme">
                <option value="PENDENTE" ${{m.entregaUniforme==='PENDENTE'?'selected':''}}>PENDENTE</option>
                <option value="OK" ${{m.entregaUniforme==='OK'?'selected':''}}>OK</option>
              </select>
            </div>
          </div>
        </div>
      </div>
      <div class="info-section-box card-seguranca${{highlightSeg}}">
        <span class="card-stripe"></span>
        <div class="card-body">
          <div class="info-block-title"><i class="fa-solid fa-shield-halved"></i> Indicadores de Segurança & Observações${{badgeSegHtml}}</div>
          <div class="meta-grid" style="margin-bottom:12px;">
            <div class="meta-item">
              <label>Acidentes (Qtd)</label><input type="number" id="editAcidentes" value="${{m.acidentes||0}}">
              <input type="text" id="editObsAcidente" class="obs-input" style="margin-top:4px;" value="${{esc(m.obsAcidente)}}" placeholder="Obs de Acidente">
            </div>
            <div class="meta-item">
              <label>Multas (Qtd)</label><input type="number" id="editMultas" value="${{m.multas||0}}">
              <input type="text" id="editObsMultas" class="obs-input" style="margin-top:4px;" value="${{esc(m.obsMultas)}}" placeholder="Obs de Multas">
            </div>
            <div class="meta-item">
              <label>Excesso de Velocidade</label><input type="number" id="editExcesso" value="${{m.excesso||0}}">
              <input type="text" id="editObsGerais" class="obs-input" style="margin-top:4px;" value="${{esc(m.obsGerais)}}" placeholder="Obs de Excesso">
            </div>
          </div>
          <div class="meta-grid" style="border-top:1px dashed #e0d0b8;padding-top:10px;">
            <div class="meta-item">
              <label>Simulador SEST SENAT</label>
              <select id="editSimulador" style="margin-bottom:4px;">
                <option value="PENDENTE" ${{m.simulador==='PENDENTE'?'selected':''}}>PENDENTE</option>
                <option value="OK" ${{m.simulador==='OK'?'selected':''}}>OK</option>
              </select>
              <div style="display:flex;gap:4px;margin-bottom:4px;">
                <input type="date" id="editSimuladorData" value="${{esc(m.simuladorData)}}" style="flex:1;" title="Data de realização" oninput="checarCamposValidade('editSimuladorData','editSimuladorValidadeMeses')">
                <input type="number" id="editSimuladorValidadeMeses" value="${{m.simuladorValidadeMeses||0}}" min="0" style="width:52px;" title="Válido por (meses)" oninput="checarCamposValidade('editSimuladorData','editSimuladorValidadeMeses')">
              </div>
              <div style="font-size:10px;font-weight:700;color:${{svSim.cor}};margin-bottom:4px;">${{svSim.label}}</div>
              <input type="text" id="editObsSimulador" class="obs-input" value="${{esc(m.obsSimulador)}}" placeholder="Obs do Simulador">
            </div>
            <div class="meta-item">
              <label>Reciclagem</label>
              <select id="editReciclagem" style="margin-bottom:4px;">
                <option value="PENDENTE" ${{m.reciclagem==='PENDENTE'?'selected':''}}>PENDENTE</option>
                <option value="OK" ${{m.reciclagem==='OK'?'selected':''}}>OK</option>
              </select>
              <div style="display:flex;gap:4px;margin-bottom:4px;">
                <input type="date" id="editReciclagemData" value="${{esc(m.reciclagemData)}}" style="flex:1;" title="Data de realização" oninput="checarCamposValidade('editReciclagemData','editReciclagemValidadeMeses')">
                <input type="number" id="editReciclagemValidadeMeses" value="${{m.reciclagemValidadeMeses||0}}" min="0" style="width:52px;" title="Válido por (meses)" oninput="checarCamposValidade('editReciclagemData','editReciclagemValidadeMeses')">
              </div>
              <div style="font-size:10px;font-weight:700;color:${{svRec.cor}};margin-bottom:4px;">${{svRec.label}}</div>
              <input type="text" id="editObsReciclagem" class="obs-input" value="${{esc(m.obsReciclagem)}}" placeholder="Obs de Reciclagem">
            </div>
            <div class="meta-item">
              <label>Gestime</label>
              <select id="editGestime" style="margin-bottom:4px;">
                <option value="PENDENTE" ${{m.gestime==='PENDENTE'?'selected':''}}>PENDENTE</option>
                <option value="OK" ${{m.gestime==='OK'?'selected':''}}>OK</option>
              </select>
              <div style="display:flex;gap:4px;margin-bottom:4px;">
                <input type="date" id="editGestimeData" value="${{esc(m.gestimeData)}}" style="flex:1;" title="Data de realização" oninput="checarCamposValidade('editGestimeData','editGestimeValidadeMeses')">
                <input type="number" id="editGestimeValidadeMeses" value="${{m.gestimeValidadeMeses||0}}" min="0" style="width:52px;" title="Válido por (meses)" oninput="checarCamposValidade('editGestimeData','editGestimeValidadeMeses')">
              </div>
              <div style="font-size:10px;font-weight:700;color:${{svGest.cor}};margin-bottom:4px;">${{svGest.label}}</div>
              <input type="text" id="editObsGestime" class="obs-input" value="${{esc(m.obsGestime)}}" placeholder="Obs de Gestime">
            </div>
          </div>
        </div>
      </div>
      <div class="info-section-box card-contato">
        <span class="card-stripe"></span>
        <div class="card-body">
          <div class="info-block-title"><i class="fa-solid fa-phone"></i> Dados de Contato</div>
          <div class="meta-grid">
            <div class="meta-item"><label>Telefone / WhatsApp</label><input type="text" id="editTelefone" value="${{esc(m.telefone)}}" placeholder="(00) 00000-0000"></div>
            <div class="meta-item"><label>E-mail Corporativo</label><input type="email" id="editEmail" value="${{esc(m.email)}}" placeholder="nome@luft.com.br"></div>
            <div class="meta-item"><label>CPF do Motorista</label><input type="text" id="editCpf" value="${{esc(m.cpf)}}" placeholder="000.000.000-00"></div>
          </div>
          <div class="meta-grid" style="border-top:1px dashed #d0e4ec;padding-top:10px;margin-top:10px;">
            <div class="meta-item">
              <label>Possui Telefone Corporativo?</label>
              <select id="editTelefoneCorporativo" onchange="toggleNumeroLinha()">
                <option value="NÃO" ${{m.telefoneCorporativo==='NÃO'?'selected':''}}>NÃO</option>
                <option value="SIM" ${{m.telefoneCorporativo==='SIM'?'selected':''}}>SIM</option>
              </select>
            </div>
            <div class="meta-item">
              <label>Número da Linha <span id="numeroLinhaObrigatorio" style="color:#dc2626;${{m.telefoneCorporativo==='SIM'?'':'display:none;'}}">*obrigatório</span></label>
              <input type="text" id="editNumeroLinha" value="${{esc(m.numeroLinha)}}" placeholder="(00) 00000-0000" ${{m.telefoneCorporativo==='NÃO'?'disabled':''}} style="${{m.telefoneCorporativo==='NÃO'?'background:#eef1f5!important;color:#9aaabb!important;':''}}">
            </div>
            <div class="meta-item">
              <div style="display:flex;gap:4px;">
                <div style="display:flex;flex-direction:column;gap:3px;flex:1;min-width:0;">
                  <label>Modelo do Celular</label>
                  <input type="text" id="editModelo" value="${{esc(m.modelo)}}" placeholder="Ex: Samsung A54" style="width:100%;">
                </div>
                <div style="display:flex;flex-direction:column;gap:3px;flex:1;min-width:0;">
                  <label>IMEI</label>
                  <input type="text" id="editImei" value="${{esc(m.imei)}}" placeholder="IMEI" style="width:100%;" title="IMEI do celular">
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="info-section-box card-dss${{highlightDss}}">
        <span class="card-stripe"></span>
        <div class="card-body">
          <div class="info-block-title"><i class="fa-solid fa-calendar-check"></i> Matriz de Controle Semanal DSS (Ano Vigente)${{badgeDssHtml}}</div>
          <div class="dss-matrix-container">${{matrizHtml}}</div>
        </div>
      </div>
    </div>`;
  aplicarEstadoVisualCampos();
  document.getElementById('driverModal').style.display = 'flex';
  const btnVoltar = document.getElementById('btnVoltarFicha');
  if(fichaOrigemModal){{ btnVoltar.style.display = 'flex'; }} else {{ btnVoltar.style.display = 'none'; }}
}}

function toggleNumeroLinha(){{
  const sel = document.getElementById('editTelefoneCorporativo');
  const input = document.getElementById('editNumeroLinha');
  const aviso = document.getElementById('numeroLinhaObrigatorio');
  const ehSim = sel.value === 'SIM';
  input.disabled = !ehSim;
  input.style.background = ehSim ? '' : '#eef1f5';
  input.style.color = ehSim ? '' : '#9aaabb';
  aviso.style.display = ehSim ? 'inline' : 'none';
  if(!ehSim) input.value = '';
}}

function dispararUploadFoto(){{
  document.getElementById('avatarMenu').style.display = 'none';
  document.getElementById('hiddenPhotoInput').click();
}}

function abrirMenuAvatar(event){{
  event.stopPropagation();
  const menu    = document.getElementById('avatarMenu');
  const temFoto = !!fotoTemporariaBase64;
  document.getElementById('avatarMenuCarregar').style.display   = temFoto ? 'none' : 'flex';
  document.getElementById('avatarMenuSubstituir').style.display = temFoto ? 'flex' : 'none';
  document.getElementById('avatarMenuExcluir').style.display    = temFoto ? 'flex' : 'none';
  menu.style.display = (menu.style.display === 'flex') ? 'none' : 'flex';
}}

document.addEventListener('click', e => {{
  const menu = document.getElementById('avatarMenu');
  if(menu && menu.style.display === 'flex' && !menu.contains(e.target) && e.target.id !== 'avatarWrapperFicha'){{
    menu.style.display = 'none';
  }}
}});

function excluirFotoAtual(){{
  fotoTemporariaBase64 = '';
  const img = document.getElementById('profilePreviewImg');
  if(img){{ img.src = AVATAR_PADRAO; img.style.display = 'block'; }}
  document.getElementById('avatarMenu').style.display = 'none';
  toast('Foto removida. Avatar padrão restaurado. Clique em "Confirmar Alterações" para salvar.', 'ok');
}}

function processarFotoCarregada(input){{
  if(input.files && input.files[0]){{
    const reader = new FileReader();
    reader.onload = async e => {{
      // Comprime já na leitura — nunca guarda base64 gigante na memória
      fotoTemporariaBase64 = await _comprimirBase64(e.target.result, 80, 0.5);
      const img = document.getElementById('profilePreviewImg');
      img.src = fotoTemporariaBase64;
      img.style.display = 'block';
    }};
    reader.readAsDataURL(input.files[0]);
  }}
}}

async function confirmarEdicaoFicha(){{
  const idx = motoristasDB.findIndex(x => x.cpf === motoristaEmEdicaoCpf);
  if(idx === -1) return;

  let novoCpf = document.getElementById('editCpf').value.trim();
  if(!novoCpf){{
    novoCpf = `SEMCPF-${{Date.now()}}-${{Math.random().toString(36).slice(2,7)}}`;
    document.getElementById('editCpf').value = novoCpf;
  }}
  if(novoCpf !== motoristaEmEdicaoCpf && motoristasDB.some(m => m.cpf === novoCpf)){{
    toast('Já existe outro motorista cadastrado com este CPF.', 'erro');
    return;
  }}

  const telCorp = document.getElementById('editTelefoneCorporativo').value;
  const numLinha = document.getElementById('editNumeroLinha').value.trim();
  if(telCorp === 'SIM' && !numLinha){{
    toast('Informe o Número da Linha, pois o Telefone Corporativo está marcado como SIM.', 'erro');
    return;
  }}

  const recSelVal = document.getElementById('editReciclagem').value;
  const recData   = document.getElementById('editReciclagemData').value;
  const recMeses  = parseInt(document.getElementById('editReciclagemValidadeMeses').value) || 0;
  if(recSelVal === 'OK' && (!recData || recMeses <= 0)){{
    toast('Para marcar Reciclagem como OK, preencha a Data de Realização e a Validade (meses).', 'erro');
    return;
  }}
  const simSelVal = document.getElementById('editSimulador').value;
  const simData   = document.getElementById('editSimuladorData').value;
  const simMeses  = parseInt(document.getElementById('editSimuladorValidadeMeses').value) || 0;
  if(simSelVal === 'OK' && (!simData || simMeses <= 0)){{
    toast('Para marcar Simulador como OK, preencha a Data de Realização e a Validade (meses).', 'erro');
    return;
  }}
  const exPerData  = document.getElementById('editExamePeriodico').value;
  const exPerMeses = parseInt(document.getElementById('editExamePeriodicoValidadeMeses').value) || 0;
  if(exPerData && exPerMeses <= 0){{
    toast('Informe a Validade (meses) do Exame Periódico.', 'erro');
    return;
  }}
  const exToxData  = document.getElementById('editExameToxicologico').value;
  const exToxMeses = parseInt(document.getElementById('editExameToxicologicoValidadeMeses').value) || 0;
  if(exToxData && exToxMeses <= 0){{
    toast('Informe a Validade (meses) do Exame Toxicológico.', 'erro');
    return;
  }}
  const gestSelVal = document.getElementById('editGestime').value;
  const gestData   = document.getElementById('editGestimeData').value;
  const gestMeses  = parseInt(document.getElementById('editGestimeValidadeMeses').value) || 0;
  if(gestSelVal === 'OK' && (!gestData || gestMeses <= 0)){{
    toast('Para marcar Gestime como OK, preencha a Data de Realização e a Validade (meses).', 'erro');
    return;
  }}

  const dssAnual = {{}};
  MESES.forEach(mes => {{
    dssAnual[mes] = [0,1,2,3].map(i => {{
      const chk = document.getElementById(`dss-${{mes}}-${{i}}`);
      return chk ? chk.checked : false;
    }});
  }});
  const atualizado = {{
    ...motoristasDB[idx],
    cpf:           novoCpf,
    nome:          document.getElementById('editNome').value.toUpperCase(),
    filial:        document.getElementById('editFilial').value.toUpperCase(),
    telefone:      document.getElementById('editTelefone').value,
    imei:          document.getElementById('editImei').value,
    email:         document.getElementById('editEmail').value,
    cnh:           document.getElementById('editCnh').value,
    validadeCnh:   document.getElementById('editValidadeCnh').value,
    admissao:      document.getElementById('editAdmissao').value,
    examePeriodico:      document.getElementById('editExamePeriodico').value,
    exameToxicologico:   document.getElementById('editExameToxicologico').value,
    pontuacaoCnh:        parseInt(document.getElementById('editPontuacaoCnh').value)||0,
    vencimentoCnhMopp:   document.getElementById('editVencimentoCnhMopp').value,
    entregaUniforme:     document.getElementById('editEntregaUniforme').value,
    afastado:            document.getElementById('editAfastado').value,
    obsAfastado:         document.getElementById('editObsAfastado').value,
    telefoneCorporativo: telCorp,
    numeroLinha:         telCorp === 'SIM' ? numLinha : '',
    modelo:              document.getElementById('editModelo').value,
    reciclagem:    document.getElementById('editReciclagem').value,
    reciclagemData: document.getElementById('editReciclagemData').value,
    reciclagemValidadeMeses: parseInt(document.getElementById('editReciclagemValidadeMeses').value)||0,
    simulador:     document.getElementById('editSimulador').value,
    simuladorData: document.getElementById('editSimuladorData').value,
    simuladorValidadeMeses: parseInt(document.getElementById('editSimuladorValidadeMeses').value)||0,
    examePeriodicoValidadeMeses: parseInt(document.getElementById('editExamePeriodicoValidadeMeses').value)||0,
    exameToxicologicoValidadeMeses: parseInt(document.getElementById('editExameToxicologicoValidadeMeses').value)||0,
    gestime:        document.getElementById('editGestime').value,
    gestimeData:    document.getElementById('editGestimeData').value,
    gestimeValidadeMeses: parseInt(document.getElementById('editGestimeValidadeMeses').value)||0,
    obsGestime:     document.getElementById('editObsGestime').value,
    acidentes:     parseInt(document.getElementById('editAcidentes').value)||0,
    multas:        parseInt(document.getElementById('editMultas').value)||0,
    excesso:       parseInt(document.getElementById('editExcesso').value)||0,
    obsAcidente:   document.getElementById('editObsAcidente').value,
    obsMultas:     document.getElementById('editObsMultas').value,
    obsGerais:     document.getElementById('editObsGerais').value,
    obsReciclagem: document.getElementById('editObsReciclagem').value,
    obsSimulador:  document.getElementById('editObsSimulador').value,
    foto:          fotoTemporariaBase64 || '',
    dssAnual
  }};
  mostrarSpinner(true);
  try{{
    const anterior = motoristasDB[idx];
    motoristasDB[idx] = atualizado;
    const res = await salvarTodosNaSheetsAPI(motoristasDB);
    if(res.ok){{
      fecharJanelaDriver();
      atualizarDashboardCompleto();
      if(filialModalAtiva) expandirFilial(filialModalAtiva);
      toast('Ficha atualizada e salva no Google Sheets!');
    }} else {{
      motoristasDB[idx] = anterior;
      toast(res.erro || 'Erro ao salvar.', 'erro');
    }}
  }} catch(e){{ toast('Falha de conexão: ' + e.message, 'erro'); }}
  finally{{ mostrarSpinner(false); }}
}}

async function deletarMotoristaAtual(cpf, nome){{
  if(!confirm(`Remover permanentemente o condutor ${{nome}}?`)) return;
 mostrarSpinner(true);
  try{{
    const anterior = [...motoristasDB];
    motoristasDB = motoristasDB.filter(m => m.cpf !== cpf);
    const res = await salvarTodosNaSheetsAPI(motoristasDB);
    if(res.ok){{
      fecharJanelaDriver();
      atualizarDashboardCompleto();
      if(filialModalAtiva) expandirFilial(filialModalAtiva);
      toast('Condutor removido.');
    }} else {{
      motoristasDB = anterior;
      toast(res.erro || 'Erro ao remover.', 'erro');
    }}
  }} catch(e){{ toast('Falha de conexão: ' + e.message, 'erro'); }}
  finally{{ mostrarSpinner(false); }}
}}

async function desligarMotoristaAtual(cpf){{
  const idx = motoristasDB.findIndex(x => x.cpf === cpf);
  if(idx === -1) return;
  if(!confirm('Confirma o desligamento deste motorista? Ele deixará de contar em todos os indicadores, exceto no KPI de Desligados.')) return;
  const obsEl = document.getElementById('editObsDesligamento');
  const obs = obsEl ? obsEl.value : (motoristasDB[idx].obsDesligamento || '');
  mostrarSpinner(true);
  try{{
    const anterior = {{...motoristasDB[idx]}};
    motoristasDB[idx].desligado = 'SIM';
    motoristasDB[idx].obsDesligamento = obs;
    const res = await salvarTodosNaSheetsAPI(motoristasDB);
    if(res.ok){{
      fecharJanelaDriver();
      atualizarDashboardCompleto();
      if(filialModalAtiva) expandirFilial(filialModalAtiva);
      toast('Motorista desligado com sucesso.');
    }} else {{
      motoristasDB[idx] = anterior;
      toast(res.erro || 'Erro ao desligar motorista.', 'erro');
    }}
  }} catch(e){{ toast('Falha de conexão: ' + e.message, 'erro'); }}
  finally{{ mostrarSpinner(false); }}
}}

async function reativarMotoristaAtual(cpf){{
  const idx = motoristasDB.findIndex(x => x.cpf === cpf);
  if(idx === -1) return;
  if(!confirm('Confirma a reativação deste motorista?')) return;
  mostrarSpinner(true);
  try{{
    const anterior = {{...motoristasDB[idx]}};
    const obsEl = document.getElementById('editObsDesligamento');
    motoristasDB[idx].desligado = 'NÃO';
    motoristasDB[idx].obsDesligamento = obsEl ? obsEl.value : motoristasDB[idx].obsDesligamento;
    const res = await salvarTodosNaSheetsAPI(motoristasDB);
    if(res.ok){{
      fecharJanelaDriver();
      atualizarDashboardCompleto();
      if(filialModalAtiva) expandirFilial(filialModalAtiva);
      toast('Motorista reativado com sucesso.');
    }} else {{
      motoristasDB[idx] = anterior;
      toast(res.erro || 'Erro ao reativar motorista.', 'erro');
    }}
  }} catch(e){{ toast('Falha de conexão: ' + e.message, 'erro'); }}
  finally{{ mostrarSpinner(false); }}
}}

async function salvarTudoNoSheets(){{
  mostrarSpinner(true);
  try{{
    const res = await salvarTodosNaSheetsAPI(motoristasDB);
    if(res.ok) toast('Base salva com sucesso no Google Sheets!');
    else toast(res.erro || 'Erro ao salvar.', 'erro');
  }} catch(e){{ toast('Falha de conexão: ' + e.message, 'erro'); }}
  finally{{ mostrarSpinner(false); }}
}}

// ── Converte as linhas cruas da planilha de volta em objetos de motorista ──
function linhasParaMotoristas(linhas){{
  return linhas.filter(row => row && (row[0] || row[1])).map((row, idx) => {{
    let p = 0;
    const next = () => {{ const v = row[p]; p++; return (v === undefined || v === null) ? '' : v; }};
    const cpfBruto = String(next()).trim();
    const nome = String(next()).trim();
    const cpf = cpfBruto || `SEMCPF-${{String(idx+1).padStart(4,'0')}}`;
    const filial = String(next()).trim();
    const telefone = String(next()).trim();
    const email = String(next()).trim();
    const foto = String(next()).trim();
    const reciclagem = String(next()).trim() || 'PENDENTE';
    const simulador = String(next()).trim() || 'PENDENTE';
    const excesso = Math.max(0, parseInt(next())||0);
    const multas = Math.max(0, parseInt(next())||0);
    const acidentes = Math.max(0, parseInt(next())||0);
    const obsAcidente = String(next()).trim();
    const obsMultas = String(next()).trim();
    const obsGerais = String(next()).trim();
    const obsReciclagem = String(next()).trim();
    const obsSimulador = String(next()).trim();
    const cnh = String(next()).trim();
    const validadeCnh = String(next()).trim();
    const admissao = String(next()).trim();
    const dssAnual = {{}};
    MESES.forEach(mes => {{
      const semanas = [];
      for(let s=0; s<4; s++){{
        const val = next();
        semanas.push(val === '1' || val === 1 || val === true || val === 'TRUE');
      }}
      dssAnual[mes] = semanas;
    }});
    const examePeriodico = String(next()).trim();
    const exameToxicologico = String(next()).trim();
    const pontuacaoCnh = Math.max(0, parseInt(next())||0);
    const vencimentoCnhMopp = String(next()).trim();
    const entregaUniforme = String(next()).trim() || 'PENDENTE';
    const telefoneCorporativo = String(next()).trim() || 'NÃO';
    const numeroLinha = String(next()).trim();
    const modelo = String(next()).trim();
    const imei = String(next()).trim();
    const reciclagemData = String(next()).trim();
    const reciclagemValidadeMeses = Math.max(0, parseInt(next())||0);
    const simuladorData = String(next()).trim();
    const simuladorValidadeMeses = Math.max(0, parseInt(next())||0);
    const examePeriodicoValidadeMeses = Math.max(0, parseInt(next())||0);
    const exameToxicologicoValidadeMeses = Math.max(0, parseInt(next())||0);
    const gestime = String(next()).trim() || 'PENDENTE';
    const obsGestime = String(next()).trim();
    const gestimeData = String(next()).trim();
    const gestimeValidadeMeses = Math.max(0, parseInt(next())||0);
    const afastado = String(next()).trim() || 'NÃO';
    const obsAfastado = String(next()).trim();
    const desligado = String(next()).trim() || 'NÃO';
    const obsDesligamento = String(next()).trim();
    return {{
      cpf, nome, filial, telefone, email, foto,
      reciclagem, simulador, excesso, multas, acidentes,
      obsAcidente, obsMultas, obsGerais, obsReciclagem, obsSimulador,
      cnh, validadeCnh, admissao,
      examePeriodico, exameToxicologico,
      pontuacaoCnh, vencimentoCnhMopp, entregaUniforme,
      telefoneCorporativo, numeroLinha, modelo, imei,
      reciclagemData, reciclagemValidadeMeses,
      simuladorData, simuladorValidadeMeses,
      examePeriodicoValidadeMeses, exameToxicologicoValidadeMeses,
      gestime, obsGestime, gestimeData, gestimeValidadeMeses,
      afastado, obsAfastado,
      desligado, obsDesligamento,
      dssAnual
    }};
  }});
}}

// ── Botão "Atualizar": busca os dados direto do Sheets, sem salvar nada ──
async function atualizarDadosDoSheets(silencioso = false){{
  if(!silencioso) mostrarSpinner(true);
  try{{
    const auth  = `Bearer ${{ACCESS_TOKEN}}`;
    const range = `${{SHEET_NAME_JS}}!A2:ZZ`;
    const cacheBuster = Date.now(); // ── remove cache antigo da requisição ──
    const resp = await fetch(
      `${{SHEETS_BASE}}/${{SHEET_ID_JS}}/values/${{encodeURIComponent(range)}}?_=${{cacheBuster}}`,
      {{
        headers: {{
          'Authorization': auth,
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Pragma': 'no-cache'
        }},
        cache: 'no-store'
      }}
    );
    if(!resp.ok){{
      const err = await resp.text();
      if(resp.status === 401 || err.includes('ACCESS_TOKEN_EXPIRED')){{
        _sessaoExpirouRecarregar();
        return;
      }}
      if(!silencioso) toast('Erro ao atualizar dados: ' + err, 'erro');
      return;
    }}
    const data   = await resp.json();
    const linhas = data.values || [];
    motoristasDB = linhasParaMotoristas(linhas); // ── troca completa: descarta dados antigos em memória ──
    atualizarDashboardCompleto();

    const fichaAberta  = document.getElementById('driverModal').style.display === 'flex';
    const filialAberta = document.getElementById('filialModal').style.display === 'flex';
    const kpiAberto    = document.getElementById('kpiModal').classList.contains('show');

    // Redesenha a tela em que a pessoa já está, sem tirá-la do lugar
    if(fichaAberta && motoristaEmEdicaoCpf){{
      const aindaExiste = motoristasDB.some(m => m.cpf === motoristaEmEdicaoCpf);
      if(aindaExiste) abrirFichaMotorista(motoristaEmEdicaoCpf);
    }}
    if(filialAberta && filialModalAtiva){{
      expandirFilial(filialModalAtiva);
    }}
    if(kpiAberto && kpiTipoAtual){{
      if(kpiTipoAtual === 'vencimentoCategoria' && kpiCategoriaVencAtual && kpiTipoVencAtual){{
        abrirVencimentoCategoria(kpiCategoriaVencAtual, kpiTipoVencAtual);
      }} else if(KPI_CONFIG[kpiTipoAtual]){{
        abrirKpiModal(kpiTipoAtual, kpiMesAtual);
      }}
    }}

    if(!silencioso) toast('Dados atualizados com sucesso a partir do Google Sheets!');
  }} catch(e){{
    if(!silencioso) toast('Falha de conexão ao atualizar: ' + e.message, 'erro');
  }} finally{{
    if(!silencioso) mostrarSpinner(false);
  }}
}}

// ── AUTO-REFRESH PADRÃO ──────────────────────────────────────────
// Atualiza tudo periodicamente sem sair da tela.
// Pula o ciclo se: houver edição não salva na ficha, um salvamento em andamento,
// algum filtro/busca preenchido, ou a pessoa estiver digitando em algum campo.
function existeFiltroOuBuscaAtivo(){{
  // Campos de busca/filtro do sistema — se tiverem texto digitado, não atualiza
  const camposBusca = ['kpiSearchInput', 'filialSearchInput'];
  for(const id of camposBusca){{
    const el = document.getElementById(id);
    if(el && el.value && el.value.trim() !== '') return true;
  }}

  // Se a pessoa estiver com o foco em algum campo de digitação neste exato momento
  const ativo = document.activeElement;
  if(ativo){{
    const tag = ativo.tagName;
    const tipo = (ativo.type || '').toLowerCase();
    const editavel = (tag === 'INPUT' && !['button','checkbox','radio','submit'].includes(tipo))
                   || tag === 'TEXTAREA'
                   || tag === 'SELECT'
                   || ativo.isContentEditable;
    if(editavel) return true;
  }}

  return false;
}}

async function executarAutoRefresh(){{
  const fichaAberta   = document.getElementById('driverModal').style.display === 'flex';
  const salvandoAgora = document.getElementById('spinnerOverlay').classList.contains('show');

  if(salvandoAgora){{
    console.log('[auto-refresh] Pulado: uma operação de salvamento está em andamento.');
    return;
  }}
  if(fichaAberta && houveEdicaoNaoSalva){{
    console.log('[auto-refresh] Pulado: existem alterações não salvas na ficha aberta.');
    return;
  }}
  if(existeFiltroOuBuscaAtivo()){{
    console.log('[auto-refresh] Pulado: há um filtro/busca preenchido ou um campo em uso no momento.');
    return;
  }}
  await atualizarDadosDoSheets(true); // true = silencioso (sem spinner nem toast)
}}

function iniciarAutoRefresh(){{
  if(autoRefreshInterval) clearInterval(autoRefreshInterval);
  autoRefreshInterval = setInterval(executarAutoRefresh, AUTO_REFRESH_MS);
}}

function pararAutoRefresh(){{
  if(autoRefreshInterval){{ clearInterval(autoRefreshInterval); autoRefreshInterval = null; }}
}}

function voltarPaginaAnterior(){{
  const origem = fichaOrigemModal;
  fecharJanelaDriver();
  if(origem === 'kpi'){{ setTimeout(() => document.getElementById('kpiModal').classList.add('show'), 80); }}
  if(origem === 'filialMobile'){{ setTimeout(() => {{ document.getElementById('filialModal').style.display = 'flex'; }}, 80); }}
}}

function fecharJanelaDriver(){{
  document.getElementById('driverModal').style.display = 'none';
  motoristaEmEdicaoCpf = null;
  fotoTemporariaBase64 = null;
  fichaOrigemModal     = null;
  houveEdicaoNaoSalva  = false;
  document.getElementById('btnVoltarFicha').style.display = 'none';
}}

// ── Ficha Individual PDF ──
function gerarFichaPdf(cpf){{
  const m = motoristasDB.find(x => x.cpf === cpf);
  if(!m) return;
  const now    = new Date();
  const dtStr  = now.toLocaleDateString('pt-BR');
  const hrStr  = now.toLocaleTimeString('pt-BR', {{hour:'2-digit', minute:'2-digit'}});
  const esc    = s => (s||'—').replace(/</g,'&lt;');
  const fotoHtml = `<img src="${{m.foto || AVATAR_PADRAO}}" style="width:100px;height:100px;object-fit:cover;border-radius:6px;border:2px solid #c4d0e4;">`;
  const codigoDoc = `LUFT-${{(m.cpf||'').replace(/\\D/g,'').slice(-6) || '000000'}}-${{now.getFullYear()}}${{String(now.getMonth()+1).padStart(2,'0')}}${{String(now.getDate()).padStart(2,'0')}}`;
  const emissor = (typeof usuarioLogado !== 'undefined' && usuarioLogado) ? usuarioLogado.nome : 'Sistema';
  const afastadoSim = m.afastado === 'SIM';

  const svCnh   = statusVencimentoData(m.validadeCnh);
  const svMopp  = statusVencimentoData(m.vencimentoCnhMopp);
  const svExPer = statusVencimento(m.examePeriodico, m.examePeriodicoValidadeMeses);
  const svExTox = statusVencimento(m.exameToxicologico, m.exameToxicologicoValidadeMeses);
  const svRec   = statusVencimento(m.reciclagemData, m.reciclagemValidadeMeses);
  const svSim   = statusVencimento(m.simuladorData, m.simuladorValidadeMeses);
  const svGest  = statusVencimento(m.gestimeData, m.gestimeValidadeMeses);

  const badgeClasse = sv => !sv.venc ? 'badge-pend' : sv.cor === '#16a34a' ? 'badge-ok' : sv.cor === '#d97706' ? 'badge-alerta' : 'badge-red';
  const badgeTxt    = sv => rotuloVenc(sv);

  const dssLinhas = Object.entries(m.dssAnual||{{}}).map(([mes, sems])=>{{
    const boxes = sems.map((ok,i)=>`<td style="text-align:center;padding:4px;border:1px solid #dde6f4;background:${{ok?'#dcfce7':'#fff5f5'}};color:${{ok?'#16a34a':'#dc2626'}};font-weight:800;font-size:11px;">${{ok?'✓':'✗'}}</td>`).join('');
    return `<tr><td style="padding:4px 8px;border:1px solid #dde6f4;font-size:11px;font-weight:700;color:#1a3a6b;">${{mes}}</td>${{boxes}}</tr>`;
  }}).join('');

  const html = `<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<title>LUFT LOGISTICS — Ficha | ${{m.nome}}</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:'Segoe UI',Arial,sans-serif;background:#fff;color:#1a2a44;padding:28px 32px;font-size:12px}}
  .header{{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #1a3a6b;padding-bottom:12px;margin-bottom:10px}}
  .brand-title{{font-size:20px;font-weight:900;color:#1a3a6b}}
  .brand-title span{{color:#22cc88}}
  .brand-sub{{font-size:10px;color:#5a6e8a;letter-spacing:1px;text-transform:uppercase;margin-top:3px}}
  .doc-info{{text-align:right;font-size:10px;color:#5a6e8a;line-height:1.6}}
  .doc-info strong{{display:block;font-size:13px;color:#1a3a6b;font-weight:800}}
  .doc-code{{font-family:monospace;font-size:10px;color:#3b7dd8;font-weight:700}}
  .alerta-afastado{{background:#fff5f5;border:1.5px solid #fca5a5;border-radius:8px;padding:8px 14px;margin-bottom:14px;display:flex;align-items:center;gap:10px;font-size:12px;font-weight:800;color:#dc2626;text-transform:uppercase;letter-spacing:.5px}}
  .section{{border:1.5px solid #dde6f4;border-radius:8px;margin-bottom:12px;overflow:hidden;page-break-inside:avoid}}
  .section-head{{background:#1a3a6b;color:#fff;font-size:10px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;padding:6px 12px}}
  .section-body{{padding:12px}}
  .profile-row{{display:flex;gap:18px;align-items:flex-start}}
  .info-grid{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;flex:1}}
  .info-item label{{font-size:9px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;display:block;margin-bottom:2px}}
  .info-item span{{font-size:13px;font-weight:700;color:#1a2a44}}
  .badge{{display:inline-block;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:800;letter-spacing:.3px}}
  .badge-ok{{background:#dcfce7;color:#16a34a;border:1px solid #86efac}}
  .badge-pend{{background:#f1f5f9;color:#64748b;border:1px solid #cbd5e1}}
  .badge-alerta{{background:#fef9c3;color:#d97706;border:1px solid #fde68a}}
  .badge-red{{background:#fee2e2;color:#dc2626;border:1px solid #fca5a5}}
  .doc-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
  .doc-card{{border:1.5px solid #e2e8f0;border-radius:6px;padding:8px 10px;display:flex;flex-direction:column;gap:3px}}
  .doc-card label{{font-size:9px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px}}
  .doc-card .valor{{font-size:12px;font-weight:700;color:#1a2a44}}
  .kpi-row{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
  .kpi-box{{border-radius:6px;padding:10px 12px;text-align:center;border:1.5px solid}}
  .kpi-box label{{font-size:9px;text-transform:uppercase;font-weight:700;letter-spacing:.5px;display:block;margin-bottom:4px}}
  .kpi-box span{{font-size:26px;font-weight:900;line-height:1}}
  .kpi-acid{{border-color:#fca5a5;background:#fff5f5}}.kpi-acid span{{color:#dc2626}}
  .kpi-mul{{border-color:#fde68a;background:#fffbeb}}.kpi-mul span{{color:#d97706}}
  .kpi-vel{{border-color:#fed7aa;background:#fff7ed}}.kpi-vel span{{color:#ea580c}}
  table.dss{{width:100%;border-collapse:collapse}}
  table.dss th{{background:#eef3fb;color:#1a4fa0;font-size:10px;font-weight:800;text-transform:uppercase;padding:5px 8px;border:1px solid #dde6f4;text-align:center}}
  .assinatura-row{{display:flex;gap:32px;align-items:flex-end;margin-top:10px}}
  .footer{{border-top:1px solid #dde6f4;margin-top:10px;padding-top:6px;display:flex;justify-content:space-between;font-size:8px;color:#9aaabb}}
  @page{{margin:0}}
  @media print{{body{{padding:14mm 16mm}} .no-print{{display:none}} .section{{margin-bottom:8px}} .section-body{{padding:8px 10px}} .kpi-row{{gap:6px}} .kpi-box{{padding:6px 8px}} .kpi-box span{{font-size:20px}} table.dss td,table.dss th{{padding:3px 6px;font-size:10px}}}}
</style></head>
<body>

<div class="header">
  <div>
    <div class="brand-title">LUFT<span style="color:#22cc88"> LOGISTICS</span></div>
    <div class="brand-sub" style="color:#1a3a6b;font-weight:700;font-size:11px;margin-top:2px;letter-spacing:.5px;">Sistema de Controle de Motoristas</div>
    <div class="brand-sub">Ficha Individual do Condutor — Histórico & Compliance</div>
  </div>
  <div class="doc-info">
    <strong>Ficha do Condutor</strong>
    Emitido em: ${{dtStr}} às ${{hrStr}}<br>
    Emitido por: ${{esc(emissor)}}<br>
  </div>
</div>

${{afastadoSim ? `<div class="alerta-afastado"><span>⚠</span> Condutor Afastado${{m.obsAfastado ? ' — ' + esc(m.obsAfastado) : ''}}</div>` : ''}}

<!-- IDENTIFICAÇÃO -->
<div class="section">
  <div class="section-head">Identificação do Condutor</div>
  <div class="section-body">
    <div class="profile-row">
      <div style="flex-shrink:0">${{fotoHtml}}</div>
      <div style="flex:1;display:grid;grid-template-columns:1fr 1fr;gap:10px 24px;">
        <div class="info-item" style="grid-column:1/-1;border-bottom:1px solid #e8eef8;padding-bottom:8px;margin-bottom:4px;">
          <label>Nome Completo</label>
          <span style="font-size:17px;font-weight:900;color:#1a3a6b;">${{esc(m.nome)}}</span>
        </div>
        <div class="info-item"><label>CPF</label><span style="font-family:monospace">${{esc(m.cpf)}}</span></div>
        <div class="info-item"><label>Filial</label><span>${{esc(m.filial)}}</span></div>
        <div class="info-item"><label>Admissão</label><span>${{m.admissao ? new Date(m.admissao+'T00:00:00').toLocaleDateString('pt-BR') : '—'}}</span></div>
        <div class="info-item"><label>Telefone</label><span>${{esc(m.telefone)}}</span></div>
        <div class="info-item"><label>E-mail</label><span>${{esc(m.email)}}</span></div>
        <div class="info-item"><label>Situação</label><span class="badge ${{afastadoSim?'badge-red':'badge-ok'}}">${{afastadoSim?'AFASTADO':'ATIVO'}}</span></div>
      </div>
    </div>
  </div>
</div>

<!-- DOCUMENTAÇÃO -->
<div class="section">
  <div class="section-head">Documentação & Validades</div>
  <div class="section-body">
    <div class="doc-grid">
      <div class="doc-card">
        <label>CNH Nº</label><span class="valor" style="font-family:monospace">${{esc(m.cnh)}}</span>
      </div>
      <div class="doc-card">
        <label>Validade CNH</label>
        <span class="valor">${{m.validadeCnh ? new Date(m.validadeCnh+'T00:00:00').toLocaleDateString('pt-BR') : '—'}}</span>
        <span class="badge ${{badgeClasse(svCnh)}}">${{badgeTxt(svCnh)}}</span>
      </div>
      <div class="doc-card">
        <label>Vencimento MOPP</label>
        <span class="valor">${{m.vencimentoCnhMopp ? new Date(m.vencimentoCnhMopp+'T00:00:00').toLocaleDateString('pt-BR') : '—'}}</span>
        <span class="badge ${{badgeClasse(svMopp)}}">${{badgeTxt(svMopp)}}</span>
      </div>
      <div class="doc-card">
        <label>Pontuação CNH</label><span class="valor">${{m.pontuacaoCnh||0}} pontos</span>
      </div>
      <div class="doc-card">
        <label>Entrega de Uniforme</label>
        <span class="badge ${{m.entregaUniforme==='OK'?'badge-ok':'badge-pend'}}">${{m.entregaUniforme||'PENDENTE'}}</span>
      </div>
      <div class="doc-card">
        <label>Telefone Corporativo</label>
        <span class="badge ${{m.telefoneCorporativo==='SIM'?'badge-ok':'badge-pend'}}">${{m.telefoneCorporativo||'NÃO'}}</span>
      </div>
    </div>
  </div>
</div>

<!-- EXAMES OCUPACIONAIS -->
<div class="section">
  <div class="section-head">Exames Ocupacionais</div>
  <div class="section-body">
    <div class="doc-grid">
      <div class="doc-card">
        <label>Exame Periódico</label>
        <span class="valor">${{m.examePeriodico ? new Date(m.examePeriodico+'T00:00:00').toLocaleDateString('pt-BR') : '—'}}</span>
        <span class="badge ${{badgeClasse(svExPer)}}">${{badgeTxt(svExPer)}}</span>
      </div>
      <div class="doc-card">
        <label>Exame Toxicológico</label>
        <span class="valor">${{m.exameToxicologico ? new Date(m.exameToxicologico+'T00:00:00').toLocaleDateString('pt-BR') : '—'}}</span>
        <span class="badge ${{badgeClasse(svExTox)}}">${{badgeTxt(svExTox)}}</span>
      </div>
    </div>
  </div>
</div>

<!-- TREINAMENTOS -->
<div class="section">
  <div class="section-head">Treinamentos</div>
  <div class="section-body">
    <div class="doc-grid">
      <div class="doc-card">
        <label>Reciclagem</label>
        <span class="valor">${{m.reciclagemData ? new Date(m.reciclagemData+'T00:00:00').toLocaleDateString('pt-BR') : '—'}}</span>
        <span class="badge ${{badgeClasse(svRec)}}">${{badgeTxt(svRec)}}</span>
      </div>
      <div class="doc-card">
        <label>Simulador SEST SENAT</label>
        <span class="valor">${{m.simuladorData ? new Date(m.simuladorData+'T00:00:00').toLocaleDateString('pt-BR') : '—'}}</span>
        <span class="badge ${{badgeClasse(svSim)}}">${{badgeTxt(svSim)}}</span>
      </div>
      <div class="doc-card">
        <label>Gestime</label>
        <span class="valor">${{m.gestimeData ? new Date(m.gestimeData+'T00:00:00').toLocaleDateString('pt-BR') : '—'}}</span>
        <span class="badge ${{badgeClasse(svGest)}}">${{badgeTxt(svGest)}}</span>
      </div>
    </div>
  </div>
</div>

<!-- EQUIPAMENTOS -->
<div class="section">
  <div class="section-head">Equipamentos Corporativos</div>
  <div class="section-body">
    <div class="doc-grid">
      <div class="doc-card"><label>Número da Linha</label><span class="valor" style="font-family:monospace">${{esc(m.numeroLinha)}}</span></div>
      <div class="doc-card"><label>Modelo do Celular</label><span class="valor">${{esc(m.modelo)}}</span></div>
      <div class="doc-card"><label>IMEI</label><span class="valor" style="font-family:monospace">${{esc(m.imei)}}</span></div>
    </div>
  </div>
</div>

<!-- INDICADORES -->
<div class="section">
  <div class="section-head">Indicadores de Segurança</div>
  <div class="section-body">
    <div class="kpi-row">
      <div class="kpi-box kpi-acid"><label>Acidentes</label><span>${{m.acidentes||0}}</span></div>
      <div class="kpi-box kpi-mul"><label>Multas</label><span>${{m.multas||0}}</span></div>
      <div class="kpi-box kpi-vel"><label>Exc. Velocidade</label><span>${{m.excesso||0}}</span></div>
    </div>
    ${{(m.obsAcidente||m.obsMultas||m.obsGerais||m.obsReciclagem||m.obsSimulador||m.obsGestime) ? `
    <div style="margin-top:10px;display:grid;grid-template-columns:1fr 1fr;gap:8px;">
      ${{m.obsAcidente ? `<div style="background:#fff5f5;border:1px solid #fca5a5;border-radius:5px;padding:7px 10px;"><span style="font-size:9px;color:#dc2626;font-weight:700;text-transform:uppercase;">Obs. Acidentes</span><p style="font-size:11px;margin-top:3px;color:#1a2a44;">${{esc(m.obsAcidente)}}</p></div>` : ''}}
      ${{m.obsMultas  ? `<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:5px;padding:7px 10px;"><span style="font-size:9px;color:#d97706;font-weight:700;text-transform:uppercase;">Obs. Multas</span><p style="font-size:11px;margin-top:3px;color:#1a2a44;">${{esc(m.obsMultas)}}</p></div>` : ''}}
      ${{m.obsGerais  ? `<div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:5px;padding:7px 10px;"><span style="font-size:9px;color:#ea580c;font-weight:700;text-transform:uppercase;">Obs. Velocidade</span><p style="font-size:11px;margin-top:3px;color:#1a2a44;">${{esc(m.obsGerais)}}</p></div>` : ''}}
      ${{m.obsReciclagem ? `<div style="background:#f0fef4;border:1px solid #86efac;border-radius:5px;padding:7px 10px;"><span style="font-size:9px;color:#16a34a;font-weight:700;text-transform:uppercase;">Obs. Reciclagem</span><p style="font-size:11px;margin-top:3px;color:#1a2a44;">${{esc(m.obsReciclagem)}}</p></div>` : ''}}
      ${{m.obsSimulador ? `<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:5px;padding:7px 10px;"><span style="font-size:9px;color:#1a4fa0;font-weight:700;text-transform:uppercase;">Obs. Simulador</span><p style="font-size:11px;margin-top:3px;color:#1a2a44;">${{esc(m.obsSimulador)}}</p></div>` : ''}}
      ${{m.obsGestime ? `<div style="background:#f5f0ff;border:1px solid #ddd6fe;border-radius:5px;padding:7px 10px;"><span style="font-size:9px;color:#6d28d9;font-weight:700;text-transform:uppercase;">Obs. Gestime</span><p style="font-size:11px;margin-top:3px;color:#1a2a44;">${{esc(m.obsGestime)}}</p></div>` : ''}}
    </div>` : ''}}
  </div>
</div>

<!-- DSS -->
<div class="section">
  <div class="section-head">Controle Semanal DSS — Ano Vigente</div>
  <div class="section-body">
    <table class="dss">
      <thead><tr><th style="text-align:left;padding:5px 8px;">Mês</th><th>1ª Sem</th><th>2ª Sem</th><th>3ª Sem</th><th>4ª Sem</th></tr></thead>
      <tbody>${{dssLinhas}}</tbody>
    </table>
  </div>
</div>

<!-- ASSINATURA -->
<div class="section">
  <div class="section-head">Declaração e Assinatura</div>
  <div class="section-body">
    <p style="font-size:9px;color:#5a6e8a;margin-bottom:16px;line-height:1.6;font-style:italic;border-left:3px solid #1a3a6b;padding-left:8px;">
      Declaro que as informações contidas nesta ficha estão corretas e que estou ciente das normas de segurança e conformidade operacional da empresa.
    </p>
    <div class="assinatura-row">
      <div style="flex:2;display:flex;flex-direction:column;">
        <span style="font-size:8px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;margin-bottom:28px;display:block;">Assinatura do Condutor</span>
        <div style="border-bottom:1.5px solid #1a3a6b;margin-bottom:5px;"></div>
        <div style="font-size:9px;font-weight:700;color:#1a3a6b;">${{esc(m.nome)}}</div>
        <div style="font-size:8px;color:#5a6e8a;">CPF: ${{esc(m.cpf)}}</div>
      </div>
      <div style="flex:1;display:flex;flex-direction:column;">
        <span style="font-size:8px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;margin-bottom:28px;display:block;">Data</span>
        <div style="border-bottom:1.5px solid #1a3a6b;margin-bottom:5px;"></div>
      </div>
      <div style="flex:1;display:flex;flex-direction:column;">
        <span style="font-size:8px;color:#5a6e8a;text-transform:uppercase;font-weight:700;letter-spacing:.5px;margin-bottom:28px;display:block;">Local / Filial</span>
        <div style="border-bottom:1.5px solid #1a3a6b;margin-bottom:5px;"></div>
      </div>
    </div>
  </div>
</div>

<div class="footer">
  <span><strong style="color:#1a3a6b;">LUFT LOGISTICS</strong> — Documento interno e confidencial. Uso restrito à gestão operacional.</span>
  <span>Doc. ${{codigoDoc}} • Emitido por ${{esc(emissor)}} • ${{dtStr}} ${{hrStr}}</span>
</div>

</body></html>`;

  const blob = new Blob([html], {{type:'text/html;charset=utf-8'}});
  const url  = URL.createObjectURL(blob);
  const win  = window.open(url, '_blank');
  if(win) win.onload = () => {{ win.focus(); win.print(); }};
}}

// ── Relatório PDF Pendentes ──
function gerarRelatorioPdfPendentes(){{
  const mes   = kpiMesAtual || mesCorrente();
  const lista = listaAtiva().filter(m => !dssOkNoMes(m, mes));
  _gerarRelatorio(mes, lista, false);
}}
function gerarRelatorioPdfRealizados(){{
  const mes   = kpiMesAtual || mesCorrente();
  const lista = listaAtiva().filter(m => dssOkNoMes(m, mes));
  _gerarRelatorio(mes, lista, true);
}}
function _gerarRelatorio(mes, lista, realizado){{
  const total = listaAtiva().length;
  const now   = new Date();
  const dtStr = now.toLocaleDateString('pt-BR') + ' ' + now.toLocaleTimeString('pt-BR',{{hour:'2-digit',minute:'2-digit'}});
  const semanas = ['1ª Sem','2ª Sem','3ª Sem','4ª Sem'];
  const porFilial = {{}};
  lista.forEach(m => {{
    const f = m.filial || 'SEM FILIAL';
    if(!porFilial[f]) porFilial[f] = [];
    porFilial[f].push(m);
  }});
  const filialKeys = Object.keys(porFilial).sort();
  let linhas = '';
  let globalIdx = 0;
  const headerColor = realizado ? '#1a5c2a' : '#1a3a5c';
  filialKeys.forEach(filial => {{
    linhas += `<tr><td colspan="8" style="background:${{headerColor}};color:#fff;font-size:10px;font-weight:800;padding:7px 12px;letter-spacing:1px;text-transform:uppercase;">${{filial}} — ${{porFilial[filial].length}} ${{realizado?'realizado':'pendente'}}${{porFilial[filial].length!==1?'s':''}}</td></tr>`;
    porFilial[filial].forEach(m => {{
      const zebra = globalIdx++ % 2 === 0 ? '#ffffff' : (realizado ? '#f2fff6' : '#f4f8ff');
      const dssMes = m.dssAnual?.[mes] || [false,false,false,false];
      const caixas = semanas.map((s,i) => {{
        const feito = dssMes[i];
        return `<td style="text-align:center;padding:6px 4px;vertical-align:middle;border-bottom:1px solid #e0e8f0;background:${{zebra}};"><div style="width:16px;height:16px;border:2px solid ${{feito?'#22aa66':'#aaaaaa'}};border-radius:3px;margin:0 auto;display:flex;align-items:center;justify-content:center;background:${{feito?'#e8fff4':'#fff'}};">${{feito?'<span style=\\"color:#22aa66;font-size:12px;font-weight:900;line-height:1;\\">✕</span>':''}}</div></td>`;
      }}).join('');
      const recOk = reciclagemStatus(m)==='OK'; const simOk = simuladorStatus(m)==='OK';
      const recBox = `<div style="display:inline-flex;align-items:center;gap:4px;"><div style="width:14px;height:14px;border:2px solid ${{recOk?'#22aa66':'#aaa'}};border-radius:2px;display:flex;align-items:center;justify-content:center;background:${{recOk?'#e8fff4':'#fff'}}">${{recOk?'<span style=\\"color:#22aa66;font-size:11px;font-weight:900;\\">✕</span>':''}}</div><span style="font-size:9px;color:${{recOk?'#22aa66':'#cc4444'}};font-weight:700;">${{recOk?'OK':'Pend'}}</span></div>`;
      const simBox = `<div style="display:inline-flex;align-items:center;gap:4px;"><div style="width:14px;height:14px;border:2px solid ${{simOk?'#22aa66':'#aaa'}};border-radius:2px;display:flex;align-items:center;justify-content:center;background:${{simOk?'#e8fff4':'#fff'}}">${{simOk?'<span style=\\"color:#22aa66;font-size:11px;font-weight:900;\\">✕</span>':''}}</div><span style="font-size:9px;color:${{simOk?'#22aa66':'#cc4444'}};font-weight:700;">${{simOk?'OK':'Pend'}}</span></div>`;
      linhas += `<tr style="background:${{zebra}};"><td style="padding:6px 10px;font-size:10px;font-weight:700;color:#111;border-bottom:1px solid #e0e8f0;background:${{zebra}};">${{m.nome}}</td><td style="padding:6px 8px;font-size:9px;color:#555;font-family:monospace;border-bottom:1px solid #e0e8f0;background:${{zebra}};">${{m.cpf}}</td>${{caixas}}<td style="padding:6px 8px;border-bottom:1px solid #e0e8f0;background:${{zebra}};">${{recBox}}</td><td style="padding:6px 8px;border-bottom:1px solid #e0e8f0;background:${{zebra}};">${{simBox}}</td></tr>`;
    }});
  }});
  const titulo = realizado ? 'DSS Realizados' : 'Pendentes DSS';
  const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Relatório ${{titulo}} — ${{mes}} ${{now.getFullYear()}}</title>
  <style>*{{box-sizing:border-box;margin:0;padding:0}}body{{font-family:'Segoe UI',Arial,sans-serif;background:#fff;color:#111;padding:32px 28px}}
  .header{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:28px;border-bottom:3px solid #1a3a5c;padding-bottom:16px}}
  .brand-title{{font-size:22px;font-weight:900;color:#1a3a5c}}.brand-title span{{color:#e85c00}}
  .brand-sub{{font-size:11px;color:#666;letter-spacing:1px;text-transform:uppercase;margin-top:4px}}
  .report-tag{{font-size:13px;font-weight:800;color:${{realizado?'#22aa66':'#cc4444'}};text-transform:uppercase;letter-spacing:1px}}
  .report-mes{{font-size:18px;font-weight:900;color:#1a3a5c;margin:2px 0}}.report-dt{{font-size:10px;color:#888}}
  .summary-box{{display:flex;gap:20px;margin-bottom:24px}}
  .s-card{{background:#f4f8ff;border:1px solid #d0dff0;border-radius:8px;padding:12px 20px;text-align:center;flex:1}}
  .s-card .s-val{{font-size:28px;font-weight:900;color:#1a3a5c}}.s-card .s-lbl{{font-size:9px;color:#888;text-transform:uppercase;letter-spacing:1px;font-weight:700;margin-top:2px}}
  table{{width:100%;border-collapse:collapse;font-size:10px}}
  thead tr{{background:#1a3a5c}}thead th{{color:#fff;font-size:9px;font-weight:800;letter-spacing:1px;text-transform:uppercase;padding:8px 10px;text-align:left}}
  thead th.center{{text-align:center}}
  .footer{{margin-top:28px;border-top:1px solid #dde;padding-top:10px;font-size:9px;color:#aaa;display:flex;justify-content:space-between}}</style></head><body>
  <div class="header"><div><div class="brand-title">LUFT<span> LOGISTICS</span></div><div class="brand-sub">Controle de Motoristas — DSS</div></div>
  <div style="text-align:right"><div class="report-tag">${{realizado?'✓ DSS Realizados':'⚠ Pendentes DSS'}}</div><div class="report-mes">${{mes.toUpperCase()}} ${{now.getFullYear()}}</div><div class="report-dt">Gerado em ${{dtStr}}</div></div></div>
  <div class="summary-box">
    <div class="s-card"><div class="s-val">${{total}}</div><div class="s-lbl">Total de Motoristas</div></div>
    <div class="s-card"><div class="s-val" style="color:#22aa66">${{listaAtiva().filter(m=>dssOkNoMes(m,mes)).length}}</div><div class="s-lbl">DSS Realizados</div></div>
    <div class="s-card"><div class="s-val" style="color:${{realizado?'#22aa66':'#cc4444'}}">${{lista.length}}</div><div class="s-lbl">${{realizado?'Realizaram no Mês':'Pendentes no Mês'}}</div></div>
    <div class="s-card"><div class="s-val" style="color:#1a3a5c">${{filialKeys.length}}</div><div class="s-lbl">Filiais</div></div>
  </div>
  <table><thead><tr><th>Motorista</th><th>CPF</th><th class="center">1ª Sem</th><th class="center">2ª Sem</th><th class="center">3ª Sem</th><th class="center">4ª Sem</th><th class="center">Reciclagem</th><th class="center">Simulador</th></tr></thead><tbody>${{linhas}}</tbody></table>
  <div class="footer"><span>LUFT Logistics — Sistema de Controle de Motoristas</span><span>Relatório ${{titulo}} • ${{mes}} ${{now.getFullYear()}}</span></div>
  </body></html>`;
  const blob = new Blob([html],{{type:'text/html;charset=utf-8'}});
  const url  = URL.createObjectURL(blob);
  const win  = window.open(url,'_blank');
  if(win) win.onload = () => {{ win.focus(); win.print(); }};
}}

// ── ORGANOGRAMA ──────────────────────────────────────────────────
function escOrg(v){{
  return String(v==null?'':v).replace(/[&<>"']/g, function(c){{
    return {{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[c];
  }});
}}

function iconOrg(type,cx,cy){{
  const st = 'fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"';
  if(type === 'clock'){{
    return '<circle cx="'+cx+'" cy="'+cy+'" r="21" '+st+'/><path d="M'+cx+' '+(cy-15)+'V'+(cy+2)+'L'+(cx+11)+' '+(cy+9)+'" '+st+'/>';
  }}
  if(type === 'wheel'){{
    return '<circle cx="'+cx+'" cy="'+cy+'" r="21" '+st+'/><circle cx="'+cx+'" cy="'+cy+'" r="6" '+st+'/><path d="M'+(cx-20)+' '+(cy-5)+'Q'+cx+' '+(cy-19)+' '+(cx+20)+' '+(cy-5)+'" '+st+'/><path d="M'+(cx-20)+' '+(cy-5)+'H'+(cx-7)+'" '+st+'/><path d="M'+(cx+20)+' '+(cy-5)+'H'+(cx+7)+'" '+st+'/><path d="M'+cx+' '+(cy+6)+'V'+(cy+21)+'" '+st+'/>';
  }}
  if(type === 'clipboard'){{
    return '<rect x="'+(cx-14)+'" y="'+(cy-20)+'" width="28" height="40" rx="3" '+st+'/><rect x="'+(cx-7)+'" y="'+(cy-25)+'" width="14" height="9" rx="3" '+st+'/><path d="M'+(cx-8)+' '+(cy-5)+'l3 3 6-7" '+st+'/><path d="M'+(cx+5)+' '+(cy-5)+'h5" '+st+'/><path d="M'+(cx-8)+' '+(cy+9)+'l3 3 6-7" '+st+'/><path d="M'+(cx+5)+' '+(cy+9)+'h5" '+st+'/>';
  }}
  return '<rect x="'+(cx-15)+'" y="'+(cy-20)+'" width="30" height="40" rx="3" '+st+'/><rect x="'+(cx-10)+'" y="'+(cy-14)+'" width="20" height="8" rx="1" '+st+'/><circle cx="'+(cx-7)+'" cy="'+(cy+1)+'" r="1.7" fill="#fff"/><circle cx="'+cx+'" cy="'+(cy+1)+'" r="1.7" fill="#fff"/><circle cx="'+(cx+7)+'" cy="'+(cy+1)+'" r="1.7" fill="#fff"/><circle cx="'+(cx-7)+'" cy="'+(cy+9)+'" r="1.7" fill="#fff"/><circle cx="'+cx+'" cy="'+(cy+9)+'" r="1.7" fill="#fff"/><circle cx="'+(cx+7)+'" cy="'+(cy+9)+'" r="1.7" fill="#fff"/>';
}}

function personIconOrg(cx,cy){{
  return '<circle cx="'+cx+'" cy="'+(cy-10)+'" r="8" fill="none" stroke="#176bc2" stroke-width="1.8"/><path d="M'+(cx-15)+' '+(cy+16)+'Q'+(cx-15)+' '+(cy+1)+' '+cx+' '+(cy+1)+'Q'+(cx+15)+' '+(cy+1)+' '+(cx+15)+' '+(cy+16)+'" fill="none" stroke="#176bc2" stroke-width="1.8" stroke-linecap="round"/>';
}}

function normalizarPessoasOrganograma(){{
  // Remove entradas totalmente vazias (nome e cargo em branco).
  // Não cria mais um espaço vazio automático no fim de cada setor.
  (DADOS_ORG.setores||[]).forEach(function(setor){{
    if(!setor.pessoas) setor.pessoas = [];
    setor.pessoas = setor.pessoas.filter(function(p){{
      return String(p[0]||'').trim() !== '' || String(p[1]||'').trim() !== '';
    }});
  }});
}}

function adicionarPessoaOrganograma(setorIdx){{
  if(!DADOS_ORG.setores[setorIdx].pessoas) DADOS_ORG.setores[setorIdx].pessoas = [];
  DADOS_ORG.setores[setorIdx].pessoas.push(['','']);
  const novoIdx = DADOS_ORG.setores[setorIdx].pessoas.length - 1;
  fecharOrgQuickMenu();
  renderizarOrganograma({{tipo:'pessoaNome', setor:setorIdx, pessoa:novoIdx}}, 0);
}}

function toggleOrgQuickMenu(){{
  const btn   = document.getElementById('orgQuickMenuBtn');
  const panel = document.getElementById('orgQuickMenuPanel');
  const aberto = panel.classList.toggle('open');
  btn.classList.toggle('open', aberto);
  if(aberto) renderizarOrgQuickMenuLista();
}}

function fecharOrgQuickMenu(){{
  const btn   = document.getElementById('orgQuickMenuBtn');
  const panel = document.getElementById('orgQuickMenuPanel');
  if(panel) panel.classList.remove('open');
  if(btn) btn.classList.remove('open');
}}

function renderizarOrgQuickMenuLista(){{
  const cont = document.getElementById('orgQuickMenuList');
  if(!cont) return;
  cont.innerHTML = (DADOS_ORG.setores||[]).map(function(s, i){{
    const titulo = (s.titulo||[]).join(' ');
    return '<button type="button" class="org-add-setor-btn" onclick="adicionarPessoaOrganograma(' + i + ')"><i class="fa-solid fa-plus"></i> ' + escOrg(titulo) + '</button>';
  }}).join('');
}}

function renderizarOrganograma(focoRestaurar, cursorPos){{
  const W=1536;
  const Cline='#123d72', Ctext='#071c48', Cbright='#176bc2';
  const cardH=88, gap=12, firstY=474;

  const maxPessoas     = Math.max(1, ...(DADOS_ORG.setores||[]).map(s => (s.pessoas||[]).length));
  const lastCardBottom = firstY + (maxPessoas-1)*(cardH+gap) + cardH;
  const footerTop      = lastCardBottom + 26;
  const H              = footerTop + 36;

  let svg = '<svg viewBox="0 0 '+W+' '+H+'" xmlns="http://www.w3.org/2000/svg" style="position:absolute;top:0;left:0;width:100%;height:100%;">';
  svg += '<defs>';
  svg += '<linearGradient id="orgBackground" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#f6f8fb"/><stop offset=".48" stop-color="#ffffff"/><stop offset="1" stop-color="#f5f8fc"/></linearGradient>';
  svg += '<linearGradient id="orgSupervisor" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#112f62"/><stop offset="1" stop-color="#061a40"/></linearGradient>';
  svg += '<linearGradient id="orgHeader" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#155ba4"/><stop offset="1" stop-color="#06427f"/></linearGradient>';
  svg += '<linearGradient id="orgCircle" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#2d83d6"/><stop offset="1" stop-color="#1762b1"/></linearGradient>';
  svg += '<filter id="orgBoxShadow" x="-20%" y="-30%" width="140%" height="160%"><feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="#69798d" flood-opacity=".20"/></filter>';
  svg += '<filter id="orgCardShadow" x="-20%" y="-30%" width="140%" height="160%"><feDropShadow dx="0" dy="3" stdDeviation="5" flood-color="#778599" flood-opacity=".18"/></filter>';
  svg += '</defs>';
  svg += '<rect width="'+W+'" height="'+H+'" fill="url(#orgBackground)"/>';
  svg += '<text x="768" y="90" text-anchor="middle" font-size="55" font-weight="900" letter-spacing="-1" fill="#112d5b">GEST\u00c3O DE FROTA</text>';
  svg += '<rect x="700" y="112" width="137" height="5" rx="3" fill="#1b6dc2"/>';
  svg += '<g filter="url(#orgBoxShadow)"><rect x="542" y="144" width="451" height="119" rx="15" fill="url(#orgSupervisor)"/></g>';
  svg += '<circle cx="611" cy="203" r="41" fill="none" stroke="#fff" stroke-width="2"/>';
  svg += '<circle cx="611" cy="192" r="12" fill="none" stroke="#fff" stroke-width="2"/>';
  svg += '<path d="M591 224Q591 207 611 207Q631 207 631 224" fill="none" stroke="#fff" stroke-width="2"/>';

  const numSetores  = (DADOS_ORG.setores||[]).length || 1;
  const margemXOrg  = 39;
  const gapColOrg   = 53;
  const largColOrg  = (W - margemXOrg*2 - gapColOrg*(numSetores-1)) / numSetores;
  const scaleCol    = Math.min(1, largColOrg / 330);
  const xs = [], ws = [];
  for(let _i=0; _i<numSetores; _i++){{
    xs.push(margemXOrg + _i*(largColOrg+gapColOrg));
    ws.push(largColOrg);
  }}

  const centrosSetores = xs.map(function(x,i){{ return x+ws[i]/2; }});
  const barraX1     = centrosSetores[0];
  const barraX2     = centrosSetores[centrosSetores.length-1];
  const barraCentro = (barraX1+barraX2)/2;

  svg += '<path d="M'+barraX1+' 293H'+barraX2+'" fill="none" stroke="'+Cline+'" stroke-width="3"/>';
  svg += '<path d="M'+barraCentro+' 263V293" fill="none" stroke="'+Cline+'" stroke-width="3"/>';

  centrosSetores.forEach(function(cx){{
    svg += '<path d="M'+cx+' 293V326" stroke="'+Cline+'" stroke-width="3"/>';
  }});

  const overlays = [];
  overlays.push({{tipo:'supervisorNome', x:674, y:198, fs:29, fw:900, color:'#ffffff', w:300}});
  overlays.push({{tipo:'supervisorCargo', x:674, y:229, fs:20, fw:800, color:'#29a3fb', w:300}});

  const cardW     = Math.round(259 * scaleCol);
  const cardXOff  = Math.round(48  * scaleCol);
  const lineXOff  = Math.round(24  * scaleCol);
  const avatarOff = Math.round(38  * scaleCol);
  const textXOff  = Math.round(78  * scaleCol);
  const textW     = Math.round(171 * scaleCol);
  const nomeFs    = Math.max(11, Math.round(18 * scaleCol));
  const cargoFs   = Math.max(10, Math.round(17 * scaleCol));
  const iconXOff  = Math.round(61 * scaleCol);
  const iconR     = Math.round(42 * scaleCol);
  const titleSizeBase = Math.max(13, Math.round(24 * scaleCol));

  (DADOS_ORG.setores||[]).forEach(function(setor,i){{
    const x=xs[i], w=ws[i], headerY=326, iconX=x+iconXOff, iconY=381;
    const titleX = Math.round(x + (i===2 ? 116 : 120) * scaleCol);
    const titleSize = i===2 ? titleSizeBase-1 : titleSizeBase;
    svg += '<g><rect x="'+x+'" y="'+headerY+'" width="'+w+'" height="112" rx="14" fill="url(#orgHeader)" filter="url(#orgBoxShadow)"/><circle cx="'+iconX+'" cy="'+iconY+'" r="'+iconR+'" fill="url(#orgCircle)"/>' + iconOrg(setor.icone,iconX,iconY);
    if((setor.titulo||[]).length === 1){{
      svg += '<text x="'+titleX+'" y="'+(headerY+66)+'" font-size="'+titleSize+'" font-weight="900" fill="#fff">'+escOrg(setor.titulo[0])+'</text>';
    }} else {{
      svg += '<text x="'+titleX+'" y="'+(headerY+50)+'" font-size="'+titleSize+'" font-weight="900" fill="#fff">'+escOrg(setor.titulo[0])+'</text>';
      svg += '<text x="'+titleX+'" y="'+(headerY+82)+'" font-size="'+titleSize+'" font-weight="900" fill="#fff">'+escOrg(setor.titulo[1])+'</text>';
    }}
    svg += '</g>';

    const lineX=x+lineXOff, cardX=x+cardXOff;
    const pessoas = setor.pessoas||[];
    const lastCenter = firstY + (Math.max(pessoas.length,1)-1)*(cardH+gap) + cardH/2;
    svg += '<path d="M'+lineX+' 438V'+lastCenter+'" stroke="'+Cline+'" stroke-width="2" fill="none"/>';

    pessoas.forEach(function(p,j){{
      const y = firstY + j*(cardH+gap), cy = y+cardH/2;
      const nome  = String(p[0]||'').trim();
      const cargo = String(p[1]||'').trim();
      const vago  = !nome && !cargo;
      svg += '<path d="M'+lineX+' '+cy+'H'+cardX+'" stroke="'+(vago?'#a9bad4':Cline)+'" stroke-width="2" stroke-dasharray="'+(vago?'4,3':'0')+'"/>';
      svg += '<circle cx="'+lineX+'" cy="'+cy+'" r="4" fill="'+(vago?'#a9bad4':Cline)+'"/>';
      svg += '<rect x="'+cardX+'" y="'+y+'" width="'+cardW+'" height="'+cardH+'" rx="13" fill="'+(vago?'#f6f9fc':'#fff')+'" stroke="'+(vago?'#c7d4e8':'none')+'" stroke-width="'+(vago?'1.5':'0')+'" stroke-dasharray="'+(vago?'6,4':'0')+'" filter="'+(vago?'':'url(#orgCardShadow)')+'"/>';
      if(vago){{
        svg += '<circle cx="'+(cardX+avatarOff)+'" cy="'+cy+'" r="15" fill="none" stroke="#a9bad4" stroke-width="2" stroke-dasharray="3,3"/>';
        svg += '<path d="M'+(cardX+avatarOff)+' '+(cy-7)+'V'+(cy+7)+'M'+(cardX+avatarOff-7)+' '+cy+'H'+(cardX+avatarOff+7)+'" stroke="#a9bad4" stroke-width="2.4" stroke-linecap="round"/>';
      }} else {{
        svg += personIconOrg(cardX+avatarOff,cy);
      }}
      overlays.push({{tipo:'pessoaNome', setor:i, pessoa:j, x:cardX+textXOff, y:y+41, fs:nomeFs, fw:700, color: vago?'#9aa8bd':Ctext, w:textW, placeholder:'Nome do colaborador'}});
      overlays.push({{tipo:'pessoaCargo', setor:i, pessoa:j, x:cardX+textXOff, y:y+64, fs:cargoFs, fw:400, color: vago?'#b7c1d6':Cbright, w:textW, placeholder:'Cargo (ex: Assist. Adm.)'}});
    }});
  }});

  svg += '<rect x="0" y="'+footerTop+'" width="1222" height="23" fill="#07457f"/><path d="M1230 '+(footerTop-8)+'H1536V'+(footerTop+23)+'H1210Z" fill="#176bc0"/><path d="M1225 '+(footerTop-8)+'H1240L1220 '+(footerTop+23)+'H1205Z" fill="#f8fafc"/>';
  svg += '</svg>';

  const wrap = document.getElementById('orgWrapInner');
  wrap.innerHTML = svg;
  wrap.style.width  = W+'px';
  wrap.style.height = H+'px';
  wrap.dataset.orgW = W;
  wrap.dataset.orgH = H;

  overlays.forEach(function(o){{
    const inp = document.createElement('input');
    inp.type = 'text';
    if(o.placeholder) inp.placeholder = o.placeholder;
    if(o.tipo === 'supervisorNome')  inp.value = DADOS_ORG.supervisor.nome || '';
    if(o.tipo === 'supervisorCargo') inp.value = DADOS_ORG.supervisor.cargo || '';
    if(o.tipo === 'pessoaNome')      inp.value = DADOS_ORG.setores[o.setor].pessoas[o.pessoa][0] || '';
    if(o.tipo === 'pessoaCargo')     inp.value = DADOS_ORG.setores[o.setor].pessoas[o.pessoa][1] || '';
    inp.dataset.tipo = o.tipo;
    if(o.setor  !== undefined) inp.dataset.setor  = o.setor;
    if(o.pessoa !== undefined) inp.dataset.pessoa = o.pessoa;
    inp.style.position = 'absolute';
    inp.style.left = o.x + 'px';
    inp.style.top  = (o.y - o.fs) + 'px';
    inp.style.width = o.w + 'px';
    inp.style.fontSize = o.fs + 'px';
    inp.style.fontWeight = o.fw;
    inp.style.color = o.color;
    inp.style.background = 'transparent';
    inp.style.border = 'none';
    inp.style.outline = 'none';
    inp.style.padding = '0';
    inp.style.fontFamily = 'Arial,Helvetica,sans-serif';
    inp.onfocus = function(){{ inp.style.background = 'rgba(255,220,80,0.25)'; inp.style.borderRadius = '3px'; }};
    inp.onblur  = function(){{ inp.style.background = 'transparent'; }};
    inp.oninput = function(){{
      if(o.tipo === 'supervisorNome')  DADOS_ORG.supervisor.nome  = inp.value;
      if(o.tipo === 'supervisorCargo') DADOS_ORG.supervisor.cargo = inp.value;
      if(o.tipo === 'pessoaNome')      DADOS_ORG.setores[o.setor].pessoas[o.pessoa][0] = inp.value;
      if(o.tipo === 'pessoaCargo')     DADOS_ORG.setores[o.setor].pessoas[o.pessoa][1] = inp.value;
    }};
    wrap.appendChild(inp);
  }});

  if(focoRestaurar){{
    const sel = wrap.querySelector('input[data-tipo="'+focoRestaurar.tipo+'"][data-setor="'+focoRestaurar.setor+'"][data-pessoa="'+focoRestaurar.pessoa+'"]');
    if(sel){{ sel.focus(); if(cursorPos!=null){{ try{{ sel.setSelectionRange(cursorPos,cursorPos); }}catch(e){{}} }} }}
  }}

  ajustarEscalaOrganograma();
}}

function ajustarEscalaOrganograma(){{
  const outer = document.getElementById('orgWrapOuter');
  const inner = document.getElementById('orgWrapInner');
  if(!outer || !inner) return;
  const contentW = parseFloat(inner.dataset.orgW) || 1536;
  const contentH = parseFloat(inner.dataset.orgH) || 1024;
  const escalaBase = (outer.clientWidth / contentW) || 1; // usa toda a largura do modal, expandindo mais para os lados
  const escala   = escalaBase * orgZoomFactor;
  const offsetX  = Math.max(0, (outer.clientWidth  - contentW*escala) / 2);
  const offsetY  = Math.max(0, (outer.clientHeight - contentH*escala) / 2);
  inner.style.transform = 'translate(' + offsetX + 'px,' + offsetY + 'px) scale(' + escala + ')';
}}

function alterarZoomOrganograma(delta){{
  orgZoomFactor = Math.min(2.2, Math.max(0.4, +(orgZoomFactor + delta).toFixed(2)));
  ajustarEscalaOrganograma();
}}

function abrirOrganogramaModal(){{
  document.getElementById('organogramaModal').style.display = 'flex';
  fecharOrgQuickMenu();
  renderizarOrganograma();
  setTimeout(ajustarEscalaOrganograma, 30);
}}

document.addEventListener('click', function(e){{
  const menu = document.getElementById('orgQuickMenuPanel');
  const btn  = document.getElementById('orgQuickMenuBtn');
  if(!menu || !btn) return;
  if(menu.classList.contains('open') && !menu.contains(e.target) && !btn.contains(e.target)){{
    fecharOrgQuickMenu();
  }}
}});

function fecharOrganogramaModal(){{
  normalizarPessoasOrganograma();
  fecharOrgQuickMenu();
  document.getElementById('organogramaModal').style.display = 'none';
}}

window.addEventListener('resize', function(){{
  const m = document.getElementById('organogramaModal');
  if(m && m.style.display === 'flex') ajustarEscalaOrganograma();
}});

document.addEventListener('keydown', function(e){{
  const modal = document.getElementById('organogramaModal');
  if(!modal || modal.style.display !== 'flex') return;
  if(!(e.ctrlKey || e.metaKey)) return;
  if(e.key === '+' || e.key === '='){{
    e.preventDefault();
    alterarZoomOrganograma(0.1);
  }} else if(e.key === '-' || e.key === '_'){{
    e.preventDefault();
    alterarZoomOrganograma(-0.1);
  }} else if(e.key === '0'){{
    e.preventDefault();
    orgZoomFactor = 0.85;
    ajustarEscalaOrganograma();
  }}
}});

async function salvarOrganogramaAPI(){{
  mostrarSpinner(true);
  try{{
    normalizarPessoasOrganograma();
    const auth = 'Bearer ' + ACCESS_TOKEN;
    const rangeBase = SHEET_NAME_ORG_JS + '!A2:ZZ';
    const linhas = [];
    linhas.push(['1','supervisor','','','','', DADOS_ORG.supervisor.nome||'', DADOS_ORG.supervisor.cargo||'']);
    DADOS_ORG.setores.forEach(function(setor,so){{
      const tituloStr = (setor.titulo||[]).join('|');
      let po = 0;
      (setor.pessoas||[]).forEach(function(p){{
        const nome  = String(p[0]||'').trim();
        const cargo = String(p[1]||'').trim();
        if(!nome && !cargo) return; // não salva a linha vaga ainda não preenchida
        linhas.push([so+'-'+po,'pessoa',so,tituloStr,setor.icone||'',po,p[0]||'',p[1]||'']);
        po++;
      }});
    }});
    const clearResp = await fetch(SHEETS_BASE + '/' + SHEET_ID_JS + '/values/' + encodeURIComponent(rangeBase) + ':clear', {{ method:'POST', headers:{{'Authorization':auth}} }});
    if(!clearResp.ok){{ toast('Erro ao limpar a aba do organograma.', 'erro'); return; }}
    const resp = await fetch(SHEETS_BASE + '/' + SHEET_ID_JS + '/values/' + encodeURIComponent(rangeBase) + '?valueInputOption=RAW', {{
      method:'PUT',
      headers:{{'Authorization':auth,'Content-Type':'application/json'}},
      body: JSON.stringify({{values: linhas}})
    }});
    if(resp.ok) toast('Organograma salvo com sucesso!');
    else toast('Erro ao salvar organograma.', 'erro');
  }} catch(e){{ toast('Falha de conexão: ' + e.message, 'erro'); }}
  finally{{ mostrarSpinner(false); }}
}}

// ── PDF do Organograma ──
function gerarOrganogramaSvgTexto(incluirTituloTopo){{
  if(incluirTituloTopo === undefined) incluirTituloTopo = true;
  const W=1536;
  const Cline='#123d72', Ctext='#071c48', Cbright='#176bc2';
  const cardH=88, gap=12, firstY=474;
  const setoresComPessoas = (DADOS_ORG.setores||[]).map(function(s){{
    return {{ titulo:s.titulo, icone:s.icone, pessoas:(s.pessoas||[]).filter(function(p){{ return String(p[0]||'').trim() || String(p[1]||'').trim(); }}) }};
  }});
  const maxPessoas = Math.max(1, ...setoresComPessoas.map(s => s.pessoas.length));
  const lastCardBottom = firstY + (maxPessoas-1)*(cardH+gap) + cardH;
  const footerTop = lastCardBottom + 26;
  const H = footerTop + 36;
  let svg = '<svg viewBox="0 0 '+W+' '+H+'" xmlns="http://www.w3.org/2000/svg" width="'+W+'" height="'+H+'">';
  svg += '<defs>';
  svg += '<linearGradient id="orgBackgroundPdf" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#f6f8fb"/><stop offset=".48" stop-color="#ffffff"/><stop offset="1" stop-color="#f5f8fc"/></linearGradient>';
  svg += '<linearGradient id="orgSupervisorPdf" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#112f62"/><stop offset="1" stop-color="#061a40"/></linearGradient>';
  svg += '<linearGradient id="orgHeaderPdf" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#155ba4"/><stop offset="1" stop-color="#06427f"/></linearGradient>';
  svg += '<linearGradient id="orgCirclePdf" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#2d83d6"/><stop offset="1" stop-color="#1762b1"/></linearGradient>';
  svg += '</defs>';
  svg += '<rect width="'+W+'" height="'+H+'" fill="url(#orgBackgroundPdf)"/>';
  if(incluirTituloTopo){{
    svg += '<text x="768" y="90" text-anchor="middle" font-size="55" font-weight="900" letter-spacing="-1" fill="#112d5b">GEST\u00c3O DE FROTA</text>';
    svg += '<rect x="700" y="112" width="137" height="5" rx="3" fill="#1b6dc2"/>';
  }}
  svg += '<rect x="542" y="144" width="451" height="119" rx="15" fill="url(#orgSupervisorPdf)"/>';
  svg += '<circle cx="611" cy="203" r="41" fill="none" stroke="#fff" stroke-width="2"/>';
  svg += '<circle cx="611" cy="192" r="12" fill="none" stroke="#fff" stroke-width="2"/>';
  svg += '<path d="M591 224Q591 207 611 207Q631 207 631 224" fill="none" stroke="#fff" stroke-width="2"/>';
  svg += '<text x="674" y="198" font-size="29" font-weight="900" fill="#ffffff">'+escOrg(DADOS_ORG.supervisor.nome||'')+'</text>';
  svg += '<text x="674" y="229" font-size="20" font-weight="800" fill="#29a3fb">'+escOrg(DADOS_ORG.supervisor.cargo||'')+'</text>';

  const numSetoresPdf  = setoresComPessoas.length || 1;
  const margemXOrgPdf  = 39;
  const gapColOrgPdf   = 53;
  const largColOrgPdf  = (W - margemXOrgPdf*2 - gapColOrgPdf*(numSetoresPdf-1)) / numSetoresPdf;
  const scaleColPdf    = Math.min(1, largColOrgPdf / 330);
  const xs = [], ws = [];
  for(let _i=0; _i<numSetoresPdf; _i++){{
    xs.push(margemXOrgPdf + _i*(largColOrgPdf+gapColOrgPdf));
    ws.push(largColOrgPdf);
  }}

  const centrosSetoresPdf = xs.map(function(x,i){{ return x+ws[i]/2; }});
  const barraX1Pdf     = centrosSetoresPdf[0];
  const barraX2Pdf     = centrosSetoresPdf[centrosSetoresPdf.length-1];
  const barraCentroPdf = (barraX1Pdf+barraX2Pdf)/2;

  svg += '<path d="M'+barraX1Pdf+' 293H'+barraX2Pdf+'" fill="none" stroke="'+Cline+'" stroke-width="3"/>';
  svg += '<path d="M'+barraCentroPdf+' 263V293" fill="none" stroke="'+Cline+'" stroke-width="3"/>';

  centrosSetoresPdf.forEach(function(cx){{
    svg += '<path d="M'+cx+' 293V326" stroke="'+Cline+'" stroke-width="3"/>';
  }});

  const cardWPdf     = Math.round(259 * scaleColPdf);
  const cardXOffPdf  = Math.round(48  * scaleColPdf);
  const lineXOffPdf  = Math.round(24  * scaleColPdf);
  const avatarOffPdf = Math.round(38  * scaleColPdf);
  const textXOffPdf  = Math.round(78  * scaleColPdf);
  const nomeFsPdf    = Math.max(11, Math.round(18 * scaleColPdf));
  const cargoFsPdf   = Math.max(10, Math.round(17 * scaleColPdf));
  const iconXOffPdf  = Math.round(61 * scaleColPdf);
  const iconRPdf     = Math.round(42 * scaleColPdf);
  const titleSizeBasePdf = Math.max(13, Math.round(24 * scaleColPdf));
  setoresComPessoas.forEach(function(setor,i){{
    const x=xs[i], w=ws[i], headerY=326, iconX=x+iconXOffPdf, iconY=381;
    const titleX = Math.round(x + (i===2 ? 116 : 120) * scaleColPdf);
    const titleSize = i===2 ? titleSizeBasePdf-1 : titleSizeBasePdf;
    svg += '<rect x="'+x+'" y="'+headerY+'" width="'+w+'" height="112" rx="14" fill="url(#orgHeaderPdf)"/><circle cx="'+iconX+'" cy="'+iconY+'" r="'+iconRPdf+'" fill="url(#orgCirclePdf)"/>' + iconOrg(setor.icone,iconX,iconY);
    if((setor.titulo||[]).length === 1){{
      svg += '<text x="'+titleX+'" y="'+(headerY+66)+'" font-size="'+titleSize+'" font-weight="900" fill="#fff">'+escOrg(setor.titulo[0])+'</text>';
    }} else {{
      svg += '<text x="'+titleX+'" y="'+(headerY+50)+'" font-size="'+titleSize+'" font-weight="900" fill="#fff">'+escOrg(setor.titulo[0])+'</text>';
      svg += '<text x="'+titleX+'" y="'+(headerY+82)+'" font-size="'+titleSize+'" font-weight="900" fill="#fff">'+escOrg(setor.titulo[1])+'</text>';
    }}
    const lineX=x+lineXOffPdf, cardX=x+cardXOffPdf;
    const listaExibir = setor.pessoas.length ? setor.pessoas : [['','']];
    const lastCenter = firstY + (Math.max(listaExibir.length,1)-1)*(cardH+gap) + cardH/2;
    svg += '<path d="M'+lineX+' 438V'+lastCenter+'" stroke="'+Cline+'" stroke-width="2" fill="none"/>';
    listaExibir.forEach(function(p,j){{
      const y = firstY + j*(cardH+gap), cy = y+cardH/2;
      const nome  = String(p[0]||'').trim();
      const cargo = String(p[1]||'').trim();
      const vago  = !nome && !cargo;
      svg += '<path d="M'+lineX+' '+cy+'H'+cardX+'" stroke="'+(vago?'#a9bad4':Cline)+'" stroke-width="2" stroke-dasharray="'+(vago?'4,3':'0')+'"/>';
      svg += '<circle cx="'+lineX+'" cy="'+cy+'" r="4" fill="'+(vago?'#a9bad4':Cline)+'"/>';
      svg += '<rect x="'+cardX+'" y="'+y+'" width="'+cardWPdf+'" height="'+cardH+'" rx="13" fill="'+(vago?'#f6f9fc':'#fff')+'" stroke="'+(vago?'#c7d4e8':'#e2e8f0')+'" stroke-width="1.5"/>';
      if(!vago) svg += personIconOrg(cardX+avatarOffPdf,cy);
      svg += '<text x="'+(cardX+textXOffPdf)+'" y="'+(y+41)+'" font-size="'+nomeFsPdf+'" font-weight="700" fill="'+(vago?'#9aa8bd':Ctext)+'">'+escOrg(nome||'—')+'</text>';
      svg += '<text x="'+(cardX+textXOffPdf)+'" y="'+(y+64)+'" font-size="'+cargoFsPdf+'" font-weight="400" fill="'+(vago?'#b7c1d6':Cbright)+'">'+escOrg(cargo||'')+'</text>';
    }});
  }});
  svg += '<rect x="0" y="'+footerTop+'" width="1222" height="23" fill="#07457f"/><path d="M1230 '+(footerTop-8)+'H1536V'+(footerTop+23)+'H1210Z" fill="#176bc0"/><path d="M1225 '+(footerTop-8)+'H1240L1220 '+(footerTop+23)+'H1205Z" fill="#f8fafc"/>';
  svg += '</svg>';
  return svg;
}}

function gerarOrganogramaPdf(){{
  const svg       = gerarOrganogramaSvgTexto(false);
  const now       = new Date();
  const dtStr     = now.toLocaleDateString('pt-BR');
  const hrStr     = now.toLocaleTimeString('pt-BR', {{hour:'2-digit', minute:'2-digit'}});
  const emissor   = (typeof usuarioLogado !== 'undefined' && usuarioLogado) ? usuarioLogado.nome : 'Sistema';

  const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Organograma — LUFT LOGISTICS</title>
  <style>
    *{{box-sizing:border-box;margin:0;padding:0}}
    body{{font-family:'Segoe UI',Arial,sans-serif;background:#fff;color:#1a2a44;padding:28px 32px}}
    .header{{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #1a3a6b;padding-bottom:12px;margin-bottom:16px}}
    .brand-title{{font-size:20px;font-weight:900;color:#1a3a6b}}
    .brand-title span{{color:#22cc88}}
    .brand-sub{{font-size:10px;color:#5a6e8a;letter-spacing:1px;text-transform:uppercase;margin-top:3px}}
    .doc-info{{text-align:right;font-size:10px;color:#5a6e8a;line-height:1.6}}
    .doc-info strong{{display:block;font-size:13px;color:#1a3a6b;font-weight:800}}
    .svg-wrap{{width:100%}}
    .svg-wrap svg{{width:100%;height:auto;display:block}}
    .footer{{border-top:1px solid #dde6f4;margin-top:16px;padding-top:6px;display:flex;justify-content:space-between;font-size:8px;color:#9aaabb}}
    @media print{{body{{padding:10px 16px}}}}
  </style></head><body>
    <div class="header">
      <div>
        <div class="brand-title">LUFT<span> LOGISTICS</span></div>
        <div class="brand-sub">Sistema de Controle de Motoristas</div>
        <div class="brand-sub">Organograma da Equipe — Gestão de Frota</div>
      </div>
      <div class="doc-info">
        <strong>Organograma Oficial</strong>
        Emitido em: ${{dtStr}} às ${{hrStr}}<br>
        Emitido por: ${{emissor}}<br>
      </div>
    </div>
    <div class="svg-wrap">${{svg}}</div>
    <div class="footer">
      <span><strong style="color:#1a3a6b;">LUFT LOGISTICS</strong> — Documento interno e confidencial. Uso restrito à gestão operacional.</span>
      <span>Emitido por ${{emissor}} • ${{dtStr}} ${{hrStr}}</span>
    </div>
  </body></html>`;

  const blob = new Blob([html], {{type:'text/html;charset=utf-8'}});
  const url  = URL.createObjectURL(blob);
  const win  = window.open(url, '_blank');
  if(win) win.onload = () => {{ win.focus(); win.print(); }};
}}

// ── Login + Inicialização ──
// CREDENCIAIS já foi injetado mais acima (vem do secrets.toml)
let usuarioLogado = null;

function tentarLogin(){{
  const u = document.getElementById('loginUser').value.trim();
  const p = document.getElementById('loginPass').value.trim();
  const erro = document.getElementById('loginErro');
  const found = CREDENCIAIS.find(c => c.usuario === u && c.senha === p);
  if(found){{
    usuarioLogado = found;
    document.getElementById('loginBox').style.display  = 'none';
    document.getElementById('loadingBox').style.display = 'block';
    inicializar();
  }} else {{
    erro.textContent = 'Usuário ou senha incorretos.';
    document.getElementById('loginPass').value = '';
    document.getElementById('loginPass').focus();
    setTimeout(() => {{ erro.textContent = ''; }}, 3000);
  }}
}}

(function(){{
  setTimeout(() => {{ document.getElementById('loginUser').focus(); }}, 200);

  document.getElementById('btnEntrar').addEventListener('click', tentarLogin);

  document.getElementById('loginUser').addEventListener('keydown', e => {{
    if(e.key === 'Enter') tentarLogin();
  }});
  document.getElementById('loginPass').addEventListener('keydown', e => {{
    if(e.key === 'Enter') tentarLogin();
  }});
}})();

const splash    = document.getElementById('splash-screen');
const status    = document.getElementById('splashStatus');
const progress  = document.getElementById('splashProgress');
const progressB = document.getElementById('splashProgressBar');

function setStatus(msg, erro){{
  status.textContent = msg;
  status.className   = 'splash-status' + (erro ? ' erro' : '');
}}

function fecharSplashECarregar(total){{
  progressB.style.width = '100%';
  setStatus(`✓ Google Sheets — ${{total}} motorista(s) encontrado(s)`, false);
  setTimeout(()=>{{
    splash.style.transition = 'opacity .5s';
    splash.style.opacity    = '0';
    setTimeout(()=>{{
      splash.style.display='none';
      if(usuarioLogado){{
        const nomeEl = document.getElementById('topbarNomeUsuario');
        if(nomeEl) nomeEl.textContent = usuarioLogado.nome;
      }}
      atualizarDashboardCompleto();
      iniciarAutoRefresh();
    }}, 500);
  }}, 900);
}}

async function inicializar(){{
  progress.style.display = 'block';
  progressB.style.width  = '0%';
  setStatus('Conectando ao Google Sheets...', false);
  await new Promise(r => setTimeout(r, 400));
  progressB.style.width  = '40%';
  setStatus('Carregando motoristas...', false);
  await new Promise(r => setTimeout(r, 600));
  progressB.style.width  = '80%';
  await new Promise(r => setTimeout(r, 400));
  fecharSplashECarregar(motoristasDB.length);
}}
</script>
</body>
</html>
"""

HTML = HTML.replace("__ANO__", str(datetime.now().year))

# ─── Renderiza o HTML no Streamlit ────────────────────────────────────────────
components.html(HTML, height=800, scrolling=True)
