import io
import re
import requests
import pdfplumber
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# ─────────────────────────── SELIC (BCB SGS 4390) ───────────────────────────

SELIC_FALLBACK = {
    (2016,1):1.06,(2016,2):1.00,(2016,3):1.16,(2016,4):1.06,(2016,5):1.11,(2016,6):1.16,
    (2016,7):1.11,(2016,8):1.22,(2016,9):1.11,(2016,10):1.05,(2016,11):1.04,(2016,12):1.12,
    (2017,1):1.09,(2017,2):0.87,(2017,3):1.05,(2017,4):0.79,(2017,5):0.93,(2017,6):0.81,
    (2017,7):0.80,(2017,8):0.80,(2017,9):0.64,(2017,10):0.64,(2017,11):0.57,(2017,12):0.54,
    (2018,1):0.58,(2018,2):0.47,(2018,3):0.53,(2018,4):0.52,(2018,5):0.52,(2018,6):0.52,
    (2018,7):0.54,(2018,8):0.57,(2018,9):0.47,(2018,10):0.54,(2018,11):0.49,(2018,12):0.49,
    (2019,1):0.54,(2019,2):0.49,(2019,3):0.47,(2019,4):0.52,(2019,5):0.54,(2019,6):0.47,
    (2019,7):0.57,(2019,8):0.50,(2019,9):0.46,(2019,10):0.48,(2019,11):0.38,(2019,12):0.37,
    (2020,1):0.38,(2020,2):0.29,(2020,3):0.34,(2020,4):0.28,(2020,5):0.24,(2020,6):0.21,
    (2020,7):0.19,(2020,8):0.16,(2020,9):0.16,(2020,10):0.16,(2020,11):0.15,(2020,12):0.16,
    (2021,1):0.15,(2021,2):0.13,(2021,3):0.20,(2021,4):0.21,(2021,5):0.27,(2021,6):0.31,
    (2021,7):0.36,(2021,8):0.43,(2021,9):0.44,(2021,10):0.49,(2021,11):0.59,(2021,12):0.77,
    (2022,1):0.73,(2022,2):0.76,(2022,3):0.93,(2022,4):0.83,(2022,5):1.03,(2022,6):1.02,
    (2022,7):1.03,(2022,8):1.17,(2022,9):1.07,(2022,10):1.02,(2022,11):1.02,(2022,12):1.12,
    (2023,1):1.12,(2023,2):0.92,(2023,3):1.17,(2023,4):0.92,(2023,5):1.12,(2023,6):1.07,
    (2023,7):1.07,(2023,8):1.14,(2023,9):0.97,(2023,10):1.00,(2023,11):0.92,(2023,12):0.89,
    (2024,1):0.97,(2024,2):0.80,(2024,3):0.83,(2024,4):0.89,(2024,5):0.83,(2024,6):0.79,
    (2024,7):0.91,(2024,8):0.87,(2024,9):0.84,(2024,10):0.93,(2024,11):0.79,(2024,12):0.93,
    (2025,1):1.01,(2025,2):0.99,(2025,3):0.96,(2025,4):1.06,(2025,5):1.14,(2025,6):1.10,
    (2025,7):1.28,(2025,8):1.16,(2025,9):1.22,(2025,10):1.28,(2025,11):1.05,(2025,12):1.22,
    (2026,1):1.16,(2026,2):1.00,(2026,3):1.21,(2026,4):1.09,(2026,5):1.07,(2026,6):1.12,
    (2026,7):1.06,
}

MES_NOME = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',
            7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}

MES_MAP = {'jan':1,'fev':2,'mar':3,'abr':4,'mai':5,'jun':6,
           'jul':7,'ago':8,'set':9,'out':10,'nov':11,'dez':12}

@st.cache_data(ttl=3600)
def carregar_selic():
    try:
        url = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json'
        r = requests.get(url, timeout=10)
        dados = r.json()
        selic = {}
        for d in dados:
            partes = d['data'].split('/')
            m, a = int(partes[1]), int(partes[2])
            selic[(a, m)] = float(d['valor'].replace(',', '.'))
        return selic, True
    except Exception:
        return SELIC_FALLBACK, False

def ref_mais_recente(selic):
    return max(selic.keys())

def fator_selic(ano, mes, selic, ref):
    y, m = ano, mes + 1
    if m > 12:
        m, y = 1, y + 1
    soma = 0.0
    while (y, m) <= ref:
        soma += selic.get((y, m), 0.0)
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return round(1 + soma / 100, 8)

# ──────────────────────────── EXTRAÇÃO PDF ───────────────────────────────────

def extrair_mes(txt):
    m = re.search(r'Per[íi]odo\s*:\s*(\w+)\s*[àa]\s*\w+', txt, re.IGNORECASE)
    if m:
        palavra = m.group(1).lower()[:3]
        return MES_MAP.get(palavra)
    return None

def extrair_ano(txt, nome_arquivo=''):
    m = re.search(r'Exerc[íi]cio\s*[:\-]?\s*(\d{4})', txt, re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.search(r'Per[íi]odo[^\n]*(\d{4})', txt, re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.search(r'(\d{4})', nome_arquivo)
    if m:
        return int(m.group(1))
    return None

def br2float(s):
    return float(s.replace('.', '').replace(',', '.'))

def extrair_cosip(txt):
    lines = txt.splitlines()
    for i, l in enumerate(lines):
        gatilho = re.search(r'\b117\s*[-–]', l) or re.search(r'1\.751\.000', l)
        if gatilho and i + 1 < len(lines):
            nums = re.findall(r'[\d\.]+,\d{2}', lines[i + 1])
            if len(nums) >= 4:
                return br2float(nums[3])
    for l in lines:
        if re.search(r'\b117\b', l):
            nums = re.findall(r'[\d\.]+,\d{2}', l)
            if len(nums) >= 4:
                return br2float(nums[3])
    return None

def processar_pdf(arquivo_bytes, nome_arquivo):
    try:
        with pdfplumber.open(io.BytesIO(arquivo_bytes)) as pdf:
            txt = pdf.pages[0].extract_text() or ''
        mes = extrair_mes(txt)
        ano = extrair_ano(txt, nome_arquivo)
        val = extrair_cosip(txt)
        return {'mes': mes, 'ano': ano, 'valor': val, 'arquivo': nome_arquivo, 'ok': True}
    except Exception as e:
        return {'mes': None, 'ano': None, 'valor': None, 'arquivo': nome_arquivo, 'ok': False, 'erro': str(e)}

# ─────────────────────────── GERAR EXCEL ─────────────────────────────────────

def cor(hex_str):
    return PatternFill('solid', fgColor=hex_str)

def borda():
    s = Side(style='thin', color='BFBFBF')
    return Border(left=s, right=s, top=s, bottom=s)

def gerar_excel(linhas, selic, ref, municipio):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'CIP Corrigida'

    AZUL_ESC = '1F3864'
    AZUL_MED = '2E75B6'
    AZUL_CLA = 'DDEBF7'
    VERM     = 'C00000'

    # merge titulo para 9 colunas agora
    ws.merge_cells('A1:I1')
    tc = ws['A1']
    tc.value = f'CONFERÊNCIA CIP — {municipio.upper()} — Correção SELIC até {MES_NOME[ref[1]]}/{ref[0]}'
    tc.font = Font(bold=True, color='FFFFFF', size=13)
    tc.fill = cor(AZUL_ESC)
    tc.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    # ── cabeçalho colunas
    headers = ['Mês','Nome Mês','Ano','Data Ref.',
               'Valor Arrecadado (R$)','Valor CIP 5% (R$)',
               'Fator SELIC','CIP Atualizado (R$)','ART. 42 CDC (R$)']
    for c, h in enumerate(headers, 1):
        cell = ws.cell(2, c, h)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = cor(AZUL_MED)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = borda()
    ws.row_dimensions[2].height = 30

    col_w = [8, 14, 8, 14, 26, 22, 16, 24, 24]
    for i, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    BRL = '#,##0.00'
    FAT = '0.00000000'

    tot_arrecad = tot_cip = tot_h = 0.0
    lin = 3
    for item in linhas:
        mes, ano, val, fat = item['mes'], item['ano'], item['valor'], item['fator']
        nome_mes = MES_NOME.get(mes, '?')
        data_ref = f'{mes:02d}/{ano}' if mes and ano else ''
        val_cip  = round(val * 0.05, 2) if val else 0.0
        cip_corr = round(val_cip * fat, 2) if val else 0.0
        art42    = round(cip_corr * 2, 2)
        fill_row = cor('F2F7FC') if lin % 2 == 0 else cor('FFFFFF')

        vals = [mes, nome_mes, ano, data_ref,
                val or 0.0, val_cip,
                fat if val else '', cip_corr, art42]
        fmts = [None, None, None, None, BRL, BRL, FAT, BRL, BRL]
        for c, (v, f) in enumerate(zip(vals, fmts), 1):
            cell = ws.cell(lin, c, v)
            if f:
                cell.number_format = f
            cell.fill = fill_row
            cell.border = borda()
            cell.alignment = Alignment(horizontal='center')

        if val:
            tot_arrecad += val
            tot_cip     += val_cip
            tot_h       += cip_corr
        lin += 1

    # ── linha TOTAL
    ws.merge_cells(f'A{lin}:D{lin}')
    tc = ws.cell(lin, 1, 'TOTAL')
    tc.font = Font(bold=True, color='FFFFFF')
    tc.fill = cor(AZUL_ESC)
    tc.alignment = Alignment(horizontal='center', vertical='center')
    tc.border = borda()
    for c in range(2, 5):
        ws.cell(lin, c).fill = cor(AZUL_ESC)
        ws.cell(lin, c).border = borda()

    for c, v, f in [(5, tot_arrecad, BRL), (6, tot_cip, BRL), (8, tot_h, BRL), (9, round(tot_h * 2, 2), BRL)]:
        cell = ws.cell(lin, c, v)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = cor(AZUL_ESC)
        cell.number_format = f
        cell.border = borda()
        cell.alignment = Alignment(horizontal='center')
    ws.cell(lin, 7).fill = cor(AZUL_ESC)
    ws.cell(lin, 7).border = borda()

    lin += 1
    # ── ART. 42 CDC
    ws.merge_cells(f'A{lin}:G{lin}')
    tc = ws.cell(lin, 1, 'ART. 42 — CDC  ·  repetição do indébito EM DOBRO')
    tc.font = Font(bold=True, color='FFFFFF', size=11)
    tc.fill = cor(VERM)
    tc.alignment = Alignment(horizontal='center', vertical='center')
    tc.border = borda()
    for c in range(2, 8):
        ws.cell(lin, c).fill = cor(VERM)
        ws.cell(lin, c).border = borda()

    art_total = round(tot_h * 2, 2)
    cell = ws.cell(lin, 8, art_total)
    cell.font = Font(bold=True, color='FFFFFF', size=11)
    cell.fill = cor(VERM)
    cell.number_format = BRL
    cell.border = borda()
    cell.alignment = Alignment(horizontal='center')

    cell8 = ws.cell(lin, 9, art_total)
    cell8.font = Font(bold=True, color='FFFFFF', size=11)
    cell8.fill = cor(VERM)
    cell8.number_format = BRL
    cell8.border = borda()
    cell8.alignment = Alignment(horizontal='center')

    ws.row_dimensions[lin].height = 22
    ws.freeze_panes = 'A3'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), tot_arrecad, tot_cip, tot_h

# ─────────────────────────── STREAMLIT UI ────────────────────────────────────

import pandas as pd

st.set_page_config(
    page_title='Extrator CIP — SELIC',
    page_icon='⚖️',
    layout='wide',
    initial_sidebar_state='expanded',
)

_CSS = (
    "@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');"
    "html,body,[class*='css']{font-family:'Inter',-apple-system,sans-serif!important}"
    ".stApp{background:#F4F7FB!important}"
    ".main .block-container{padding:0 2rem 3rem!important;max-width:1380px!important}"
    "#MainMenu,footer,[data-testid='stDecoration'],[data-testid='stToolbar']{display:none!important}"
    "header[data-testid='stHeader']{background:transparent!important;height:0!important;min-height:0!important}"
    "[data-testid='stSidebar']{background:#FFFFFF!important;border-right:1px solid #E4E8EE!important}"
    "[data-testid='stSidebar']>div:first-child{padding:1.25rem 1.25rem 2rem!important}"
    "[data-testid='stSidebar'] label{font-size:0.68rem!important;font-weight:700!important;letter-spacing:0.09em!important;text-transform:uppercase!important;color:#8A96A8!important}"
    "[data-testid='stSidebar'] .stCheckbox label{font-size:0.78rem!important;font-weight:400!important;letter-spacing:0!important;text-transform:none!important;color:#3D4F6B!important}"
    "[data-testid='stSidebar'] p{font-size:0.78rem!important;color:#3D4F6B!important}"
    ".stTextInput>div>div>input{font-family:'Inter',sans-serif!important;font-size:0.875rem!important;color:#1A2232!important;background:#FFFFFF!important;border:1.5px solid #D0D8E4!important;border-radius:5px!important;padding:0.45rem 0.7rem!important}"
    ".stTextInput>div>div>input:focus{border-color:#2B5FAD!important;box-shadow:0 0 0 2.5px rgba(43,95,173,0.13)!important;outline:none!important}"
    "[data-testid='stFileUploaderDropzone']{background:#F8FAFE!important;border:2px dashed #BAC8DC!important;border-radius:7px!important;transition:border-color 0.15s!important}"
    "[data-testid='stFileUploaderDropzone']:hover{border-color:#2B5FAD!important}"
    "[data-testid='stFileUploaderDropzoneInstructions']>div{color:#5B6A80!important;font-size:0.8rem!important}"
    "[data-testid='stFileUploaderDropzoneInstructions'] small{color:#9BA8BB!important;font-size:0.7rem!important}"
    ".stButton>button{font-family:'Inter',sans-serif!important;font-weight:600!important;font-size:0.78rem!important;letter-spacing:0.03em!important;border-radius:5px!important;padding:0.45rem 1rem!important;border:1.5px solid transparent!important;transition:all 0.12s!important}"
    ".stButton>button[kind='primary'],.stButton>button:not([kind]){background:#1B3A6B!important;color:#FFFFFF!important;border-color:#1B3A6B!important}"
    ".stButton>button[kind='primary']:hover,.stButton>button:not([kind]):hover{background:#152d54!important;border-color:#152d54!important}"
    ".stButton>button[kind='secondary']{background:#FFFFFF!important;color:#3D4F6B!important;border-color:#C9D3E0!important}"
    ".stButton>button[kind='secondary']:hover{background:#F4F7FB!important;border-color:#A8B4C4!important}"
    ".stButton>button:disabled{background:#EDF0F5!important;color:#A4ADB9!important;border-color:#EDF0F5!important;cursor:not-allowed!important}"
    ".stDownloadButton>button{font-family:'Inter',sans-serif!important;font-weight:700!important;font-size:0.875rem!important;letter-spacing:0.03em!important;border-radius:5px!important;padding:0.6rem 1.5rem!important;background:#0E7A55!important;color:#FFFFFF!important;border:none!important;width:100%!important}"
    ".stDownloadButton>button:hover{background:#0b6145!important}"
    ".stProgress>div>div{background:#1B3A6B!important;border-radius:3px!important}"
    ".stProgress>div{background:#E4E8EE!important;border-radius:3px!important}"
    "[data-testid='stNotification']{border-radius:5px!important;font-size:0.8rem!important;font-family:'Inter',sans-serif!important}"
    "[data-testid='stExpander']{border:1px solid #E4E8EE!important;border-radius:6px!important;background:#FFFFFF!important}"
    "[data-testid='stExpander'] summary{font-size:0.78rem!important;font-weight:600!important;color:#5B6A80!important}"
    "[data-testid='stDataFrame']{border:1px solid #E4E8EE!important;border-radius:7px!important;overflow:hidden!important}"
    ".cip-topbar{background:#FFFFFF;border-bottom:1px solid #E4E8EE;padding:0.9rem 2rem;margin:0 -2rem 2rem -2rem;display:flex;align-items:center;gap:0.875rem}"
    ".cip-topbar-mark{width:34px;height:34px;background:#1B3A6B;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:0.95rem;color:#FFFFFF;flex-shrink:0}"
    ".cip-topbar-name{font-size:0.9rem;font-weight:800;color:#1A2232;letter-spacing:0.01em}"
    ".cip-topbar-desc{font-size:0.68rem;color:#8A96A8;font-weight:400;margin-top:2px}"
    ".cip-topbar-legal{margin-left:auto;font-size:0.62rem;color:#B0BACA;text-align:right;line-height:1.7;letter-spacing:0.04em;font-weight:500}"
    ".sec-marker{display:flex;align-items:center;gap:0.45rem;margin:1.5rem 0 0.65rem}"
    ".sec-marker-bar{width:3px;height:14px;background:#1B3A6B;border-radius:2px;flex-shrink:0}"
    ".sec-marker-label{font-size:0.65rem;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:#8A96A8}"
    ".sidebar-logo{padding:0 0 1.25rem;border-bottom:1px solid #F0F3F7;margin-bottom:1rem}"
    ".sidebar-logo-row{display:flex;align-items:center;gap:0.5rem;margin-bottom:0.3rem}"
    ".sidebar-logo-badge{width:26px;height:26px;background:#1B3A6B;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:0.75rem;color:#FFFFFF;flex-shrink:0}"
    ".sidebar-logo-name{font-size:0.8rem;font-weight:800;color:#1A2232;letter-spacing:0.01em}"
    ".sidebar-logo-sub{font-size:0.6rem;color:#9BA8BB;letter-spacing:0.07em;font-weight:600;margin-left:34px}"
    ".sidebar-status{margin-top:1.25rem;padding:0.75rem 0.875rem;background:#EBF2FC;border:1px solid #C7D8EF;border-radius:6px}"
    ".sidebar-status-title{font-size:0.6rem;font-weight:700;letter-spacing:0.09em;text-transform:uppercase;color:#4A6FA5;margin-bottom:0.35rem}"
    ".sidebar-status-val{font-size:0.8rem;font-weight:700;color:#1A2232}"
    ".sidebar-status-sub{font-size:0.68rem;color:#5B6A80;margin-top:0.15rem}"
    ".kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:0.875rem;margin:0.5rem 0 1.75rem}"
    "@media(max-width:1000px){.kpi-grid{grid-template-columns:repeat(2,1fr)}}"
    "@media(max-width:600px){.kpi-grid{grid-template-columns:1fr}}"
    ".kpi-card{background:#FFFFFF;border:1px solid #E4E8EE;border-radius:7px;padding:1.1rem 1.25rem;position:relative;overflow:hidden}"
    ".kpi-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:#1B3A6B;border-radius:7px 7px 0 0}"
    ".kpi-card.green::before{background:#0E7A55}"
    ".kpi-card.red::before{background:#B83024}"
    ".kpi-label{font-size:0.62rem;font-weight:700;text-transform:uppercase;letter-spacing:0.09em;color:#8A96A8;margin-bottom:0.45rem}"
    ".kpi-value{font-size:1.25rem;font-weight:700;color:#1A2232;font-variant-numeric:tabular-nums;letter-spacing:-0.01em;line-height:1.2}"
    ".kpi-card.green .kpi-value{color:#0E7A55}"
    ".kpi-card.red .kpi-value{color:#B83024}"
    ".kpi-card.navy .kpi-value{color:#1B3A6B}"
    ".table-meta{display:flex;flex-wrap:wrap;gap:1.25rem;padding:0.55rem 0.875rem;margin-bottom:0.75rem;background:#EBF2FC;border:1px solid #C7D8EF;border-radius:5px}"
    ".table-meta-item{display:flex;align-items:center;gap:0.3rem;font-size:0.72rem}"
    ".tmi-label{color:#7A8FA8;font-weight:500}"
    ".tmi-val{color:#1A2232;font-weight:700}"
    ".export-wrap{background:#FFFFFF;border:1px solid #E4E8EE;border-radius:7px;padding:1.25rem 1.5rem;margin-top:1.25rem}"
    ".export-title{font-size:0.62rem;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;color:#8A96A8;margin-bottom:0.65rem;padding-bottom:0.5rem;border-bottom:1px solid #F0F3F7}"
    ".empty-state{text-align:center;padding:5rem 2rem}"
    ".empty-state-icon{font-size:1.75rem;margin-bottom:0.75rem;opacity:0.35}"
    ".empty-state-title{font-size:0.875rem;font-weight:600;color:#5B6A80;margin-bottom:0.3rem}"
    ".empty-state-sub{font-size:0.75rem;color:#9BA8BB}"
    "hr{border:none;border-top:1px solid #E4E8EE!important;margin:1.25rem 0!important}"
)

_CSS_DARK = (
    ".stApp,.main{background:#111827!important}"
    ".main .block-container{background:#111827!important}"
    "[data-testid='stSidebar']{background:#1A2535!important;border-right:1px solid #243447!important}"
    "[data-testid='stSidebar'] *{color:#94A3B8!important}"
    "[data-testid='stSidebar'] .stCheckbox label{color:#CBD5E1!important}"
    "label,p,.stMarkdown p,.stCaption{color:#94A3B8!important}"
    "h1,h2,h3,h4{color:#F1F5F9!important}"
    ".stTextInput>div>div>input{color:#F1F5F9!important;background:#1E2D40!important;border-color:#2D4257!important}"
    ".stTextInput>div>div>input::placeholder{color:#475569!important}"
    ".stTextInput>div>div>input:focus{border-color:#3B82F6!important;box-shadow:0 0 0 2.5px rgba(59,130,246,0.2)!important}"
    ".stCheckbox>label>div:first-child{background:#1E2D40!important;border-color:#2D4257!important}"
    "[data-testid='stFileUploaderDropzone']{background:#1E2D40!important;border-color:#2D4257!important}"
    "[data-testid='stFileUploaderDropzoneInstructions']>div>span{color:#94A3B8!important}"
    "[data-testid='stFileUploaderDropzoneInstructions'] small{color:#64748B!important}"
    ".stButton>button{background:#1E2D40!important;color:#CBD5E1!important;border-color:#2D4257!important}"
    ".stButton>button[kind='primary']{background:#2563EB!important;color:#FFFFFF!important;border-color:#2563EB!important}"
    ".stButton>button[kind='primary']:hover{background:#1D4ED8!important;border-color:#1D4ED8!important}"
    ".stButton>button:not([kind='primary']):hover{background:#263549!important;border-color:#3D5470!important}"
    ".stButton>button:disabled{background:#1A2535!important;color:#374151!important;border-color:#1A2535!important}"
    ".stDownloadButton>button{background:#059669!important;color:#FFFFFF!important;border-color:#059669!important}"
    ".stDownloadButton>button:hover{background:#047857!important}"
    ".stProgress>div{background:#1E2D40!important}"
    ".stProgress>div>div{background:#3B82F6!important}"
    "[data-testid='stExpander']{border-color:#243447!important;background:#1A2535!important}"
    "[data-testid='stExpander'] summary{color:#94A3B8!important}"
    "[data-testid='stExpander'] summary:hover{color:#CBD5E1!important}"
    "[data-testid='stDataFrame']{border-color:#243447!important;background:#1A2535!important}"
    "[data-testid='stDataFrame'] *{color:#E2E8F0!important}"
    "[data-testid='stDataFrame'] th{background:#1E2D40!important;color:#94A3B8!important;border-color:#243447!important}"
    "[data-testid='stDataFrame'] td{border-color:#243447!important}"
    "[data-testid='stNotification']{background:#1E2D40!important;color:#CBD5E1!important}"
    ".cip-topbar{background:#1A2535!important;border-bottom-color:#243447!important}"
    ".cip-topbar-name{color:#F1F5F9!important}"
    ".cip-topbar-desc{color:#64748B!important}"
    ".cip-topbar-legal{color:#4A5568!important}"
    ".sec-marker-bar{background:#3B82F6!important}"
    ".sec-marker-label{color:#64748B!important}"
    ".sidebar-logo{border-bottom-color:#243447!important}"
    ".sidebar-logo-name{color:#F1F5F9!important}"
    ".sidebar-logo-sub{color:#4A5568!important}"
    ".sidebar-status{background:#1E2D40!important;border-color:#2D4257!important}"
    ".sidebar-status-title{color:#60A5FA!important}"
    ".sidebar-status-val{color:#F1F5F9!important}"
    ".sidebar-status-sub{color:#94A3B8!important}"
    ".kpi-card{background:#1A2535!important;border-color:#243447!important}"
    ".kpi-card::before{background:#3B82F6!important}"
    ".kpi-card.green::before{background:#10B981!important}"
    ".kpi-card.red::before{background:#EF4444!important}"
    ".kpi-label{color:#64748B!important}"
    ".kpi-value{color:#F1F5F9!important}"
    ".kpi-card.green .kpi-value{color:#10B981!important}"
    ".kpi-card.red .kpi-value{color:#EF4444!important}"
    ".kpi-card.navy .kpi-value{color:#60A5FA!important}"
    ".table-meta{background:#1E2D40!important;border-color:#2D4257!important}"
    ".tmi-label{color:#64748B!important}"
    ".tmi-val{color:#E2E8F0!important}"
    ".export-wrap{background:#1A2535!important;border-color:#243447!important}"
    ".export-title{color:#64748B!important;border-bottom-color:#243447!important}"
    ".empty-state-title{color:#64748B!important}"
    ".empty-state-sub{color:#4A5568!important}"
    "hr{border-top-color:#243447!important}"
)

# ── helpers ───────────────────────────────────────────────────────────────────

def brl(v):
    return f'R$ {v:,.2f}'.replace(',','X').replace('.',',').replace('X','.')

def sec(label):
    return f'<div class="sec-marker"><div class="sec-marker-bar"></div><div class="sec-marker-label">{label}</div></div>'

# ── session state ─────────────────────────────────────────────────────────────

for k, v in [('resultados',[]), ('selic',None), ('ref',None), ('upload_key',0), ('dark',False)]:
    if k not in st.session_state:
        st.session_state[k] = v

_full_css = _CSS + (_CSS_DARK if st.session_state.dark else "")
st.markdown(f"<style>{_full_css}</style>", unsafe_allow_html=True)

# ── TOPBAR (área principal) ───────────────────────────────────────────────────

st.markdown("""
<div class="cip-topbar">
  <div class="cip-topbar-mark">&#9878;</div>
  <div>
    <div class="cip-topbar-name">EXTRATOR CIP &#8212; SELIC</div>
    <div class="cip-topbar-desc">Análise e correção financeira de receitas COSIP &nbsp;&middot;&nbsp; Municípios &nbsp;&middot;&nbsp; Auditoria &nbsp;&middot;&nbsp; Exportação Excel</div>
  </div>
  <div class="cip-topbar-legal">Art. 39 §4º · Lei 9.430/96<br>BCB · SGS · Série 4390</div>
</div>
""", unsafe_allow_html=True)

# ── UPLOAD (área principal — sempre visível) ──────────────────────────────────

st.markdown(sec('Carregar Arquivos'), unsafe_allow_html=True)
col_up, col_cfg = st.columns([3, 1])

with col_up:
    arquivos = st.file_uploader(
        'PDFs SICOM',
        type='pdf',
        accept_multiple_files=True,
        key=f'uploader_{st.session_state.upload_key}',
        help='Relatório: Comparativo da Receita Prevista x Realizada',
        label_visibility='collapsed',
    )

with col_cfg:
    municipio  = st.text_input('Município', placeholder='Ex: Leopoldina')
    buscar_api = st.checkbox('SELIC via BCB', value=True)

col_b1, col_b2, col_b3, col_b4 = st.columns([2, 1, 1, 1])
with col_b1:
    processar   = st.button('Processar e adicionar', type='primary', disabled=not arquivos)
with col_b2:
    limpar_fila = st.button('Limpar fila', disabled=not arquivos)
with col_b3:
    limpar_tudo = st.button('Limpar tudo', disabled=not st.session_state.resultados)
with col_b4:
    _lbl = '☀ Modo claro' if st.session_state.dark else '◑ Modo escuro'
    if st.button(_lbl, use_container_width=True, key='toggle_dark'):
        st.session_state.dark = not st.session_state.dark
        st.rerun()

# ── SIDEBAR (apenas status) ───────────────────────────────────────────────────

with st.sidebar:
    st.markdown("""
<div class="sidebar-logo">
  <div class="sidebar-logo-row">
    <div class="sidebar-logo-badge">&#9878;</div>
    <div class="sidebar-logo-name">EXTRATOR CIP</div>
  </div>
  <div class="sidebar-logo-sub">SELIC · AUDITORIA · EXCEL</div>
</div>
""", unsafe_allow_html=True)

    if st.session_state.resultados:
        n    = len(st.session_state.resultados)
        anos = sorted({r['ano'] for r in st.session_state.resultados})
        st.markdown(f"""
<div class="sidebar-status">
  <div class="sidebar-status-title">Sessão atual</div>
  <div class="sidebar-status-val">{n} período(s) acumulado(s)</div>
  <div class="sidebar-status-sub">Anos: {', '.join(str(a) for a in anos)}</div>
</div>
""", unsafe_allow_html=True)
    else:
        st.markdown("""
<div style="font-size:0.75rem;color:#9BA8BB;line-height:1.6;margin-top:0.5rem">
  Faça upload dos PDFs SICOM na área principal e clique em <strong>Processar</strong>.
</div>
""", unsafe_allow_html=True)

# ── PROCESSAMENTO ─────────────────────────────────────────────────────────────

if limpar_fila:
    st.session_state.upload_key += 1
    st.rerun()

if limpar_tudo:
    st.session_state.resultados = []
    st.session_state.selic      = None
    st.session_state.ref        = None
    st.rerun()

if processar and arquivos:
    with st.spinner('Consultando Banco Central...'):
        selic, api_ok = carregar_selic() if buscar_api else (SELIC_FALLBACK, False)
        ref = ref_mais_recente(selic)
        st.session_state.selic = selic
        st.session_state.ref   = ref

    if api_ok:
        st.success(f'SELIC atualizada via BCB — referência: {MES_NOME[ref[1]]}/{ref[0]}')
    else:
        st.warning(f'BCB indisponível — tabela interna ({MES_NOME[ref[1]]}/{ref[0]})')

    novos  = []
    erros  = []
    chaves = {(r['ano'], r['mes']) for r in st.session_state.resultados}
    bar    = st.progress(0, text='Analisando arquivos...')

    for i, arq in enumerate(arquivos):
        res = processar_pdf(arq.read(), arq.name)
        if res['ok'] and res['mes'] and res['ano'] and res['valor'] is not None:
            chave = (res['ano'], res['mes'])
            if chave in chaves:
                erros.append({**res, 'erro': f'{res["mes"]:02d}/{res["ano"]} já carregado'})
            else:
                res['fator'] = fator_selic(res['ano'], res['mes'], selic, ref)
                novos.append(res)
                chaves.add(chave)
        else:
            erros.append(res)
        bar.progress((i + 1) / len(arquivos), text=f'{arq.name[:50]}')

    bar.empty()
    st.session_state.resultados.extend(novos)
    st.session_state.resultados.sort(key=lambda x: (x['ano'], x['mes']))

    if novos:
        st.success(f'{len(novos)} período(s) adicionado(s) — total: **{len(st.session_state.resultados)}**')
    if erros:
        with st.expander(f'Avisos — {len(erros)} arquivo(s)', expanded=not novos):
            for e in erros:
                st.write(f'• **{e["arquivo"]}** — {e.get("erro","COSIP não identificada")}')

# ── RESULTADOS ────────────────────────────────────────────────────────────────

if st.session_state.resultados:
    resultados = st.session_state.resultados
    selic      = st.session_state.selic or SELIC_FALLBACK
    ref        = st.session_state.ref   or ref_mais_recente(selic)
    nome_mun   = municipio.strip() or 'Município'

    tot_arr = sum(r['valor'] for r in resultados)
    tot_cip = round(tot_arr * 0.05, 2)
    tot_h   = round(sum(r['valor'] * 0.05 * r['fator'] for r in resultados), 2)
    tot_cdc = round(tot_h * 2, 2)

    st.markdown(sec('Resumo Executivo'), unsafe_allow_html=True)
    st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">Total Arrecadado</div>
    <div class="kpi-value">{brl(tot_arr)}</div>
  </div>
  <div class="kpi-card navy">
    <div class="kpi-label">Base CIP 5%</div>
    <div class="kpi-value">{brl(tot_cip)}</div>
  </div>
  <div class="kpi-card green">
    <div class="kpi-label">CIP Corrigida SELIC</div>
    <div class="kpi-value">{brl(tot_h)}</div>
  </div>
  <div class="kpi-card red">
    <div class="kpi-label">ART. 42 CDC · Dobro</div>
    <div class="kpi-value">{brl(tot_cdc)}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    anos     = sorted({r['ano'] for r in resultados})
    anos_str = ', '.join(str(a) for a in anos)

    st.markdown(sec('Períodos Processados'), unsafe_allow_html=True)
    st.markdown(f"""
<div class="table-meta">
  <div class="table-meta-item"><span class="tmi-label">Registros</span><span class="tmi-val">{len(resultados)} períodos</span></div>
  <div class="table-meta-item"><span class="tmi-label">Município</span><span class="tmi-val">{nome_mun.upper()}</span></div>
  <div class="table-meta-item"><span class="tmi-label">Anos</span><span class="tmi-val">{anos_str}</span></div>
  <div class="table-meta-item"><span class="tmi-label">Referência SELIC</span><span class="tmi-val">{MES_NOME[ref[1]]}/{ref[0]}</span></div>
  <div class="table-meta-item"><span class="tmi-label">Metodologia</span><span class="tmi-val">Art. 39 §4o · Lei 9.430/96</span></div>
</div>
""", unsafe_allow_html=True)

    df = pd.DataFrame([{
        'Mês':               r['mes'],
        'Nome':              MES_NOME.get(r['mes'], '?'),
        'Ano':               r['ano'],
        'Arrecadado (R$)':   r['valor'],
        'CIP 5% (R$)':       round(r['valor'] * 0.05, 2),
        'Fator SELIC':       r['fator'],
        'CIP Corrigida (R$)':round(r['valor'] * 0.05 * r['fator'], 2),
        'ART.42 CDC (R$)':   round(r['valor'] * 0.05 * r['fator'] * 2, 2),
    } for r in resultados])

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        height=min(36 * len(df) + 38, 680),
        column_config={
            'Mês':               st.column_config.NumberColumn('Mês',  format='%d',      width='small'),
            'Ano':               st.column_config.NumberColumn('Ano',  format='%d',      width='small'),
            'Arrecadado (R$)':   st.column_config.NumberColumn('Arrecadado (R$)',   format='R$ %,.2f', width='medium'),
            'CIP 5% (R$)':       st.column_config.NumberColumn('CIP 5% (R$)',       format='R$ %,.2f', width='medium'),
            'Fator SELIC':       st.column_config.NumberColumn('Fator SELIC',       format='%.8f',     width='medium'),
            'CIP Corrigida (R$)':st.column_config.NumberColumn('CIP Corrigida (R$)',format='R$ %,.2f', width='medium'),
            'ART.42 CDC (R$)':   st.column_config.NumberColumn('ART.42 CDC (R$)',   format='R$ %,.2f', width='medium'),
        }
    )

    excel_bytes, _a, _b, _c = gerar_excel(resultados, selic, ref, nome_mun)

    st.markdown(sec('Exportação'), unsafe_allow_html=True)
    st.markdown('<div class="export-wrap"><div class="export-title">Planilha Excel — dados completos com correção SELIC aplicada</div>', unsafe_allow_html=True)
    st.download_button(
        label='Exportar planilha Excel',
        data=excel_bytes,
        file_name=f'CIP {nome_mun} - SELIC.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    st.markdown('</div>', unsafe_allow_html=True)
