import streamlit as st
import pandas as pd
import os
import json
import io
from datetime import datetime, time

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.utils import ImageReader

# Configuração da Página
st.set_page_config(
    page_title="A.C.S. | Gestão de Acervo e Logística",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =====================================================================
# CONFIGURAÇÃO DE SENHAS DO SISTEMA
# =====================================================================
SENHA_ACESSO = "acs123"       # Senha padrão para abrir o sistema
SENHA_EXCLUSAO = "acsadmin"   # Senha exclusiva para apagar relatórios e itens

# =====================================================================
# ESTILIZAÇÃO CORPORATIVA (CSS) RESPONSIVA
# =====================================================================
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}

    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        max-width: 98%;
    }
    
    .corporate-header {
        background-color: #0f172a;
        color: #f8fafc;
        padding: 1.2rem 1.8rem;
        border-radius: 6px;
        border-left: 5px solid #0d9488;
        margin-bottom: 1.5rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .corporate-title {
        font-size: 1.4rem;
        font-weight: 700;
        letter-spacing: 0.5px;
        margin: 0;
        color: #ffffff;
    }
    .corporate-subtitle {
        font-size: 0.85rem;
        color: #94a3b8;
        margin-top: 4px;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    .kpi-card {
        background-color: #1e293b;
        border: 1px solid #334155;
        border-radius: 6px;
        padding: 0.8rem 1rem;
        text-align: left;
        margin-bottom: 5px;
    }
    .kpi-label {
        font-size: 0.7rem;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        font-weight: 600;
    }
    .kpi-value {
        font-size: 1.4rem;
        font-weight: 700;
        color: #f8fafc;
        margin-top: 2px;
    }
    .kpi-accent {
        color: #2dd4bf;
    }

    .section-header {
        font-size: 1.05rem;
        font-weight: 600;
        color: #e2e8f0;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        border-bottom: 1px solid #334155;
        padding-bottom: 0.4rem;
        margin-bottom: 1rem;
        margin-top: 0.5rem;
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 4px;
        font-weight: 600;
        letter-spacing: 0.4px;
        text-transform: uppercase;
        font-size: 0.82rem;
        padding: 0.5rem 1rem;
        transition: all 0.2s ease;
    }

    @media print {
        [data-testid="stSidebar"], 
        .no-print, 
        .stButton, 
        .stDownloadButton, 
        [data-testid="stFormSubmitButton"],
        .stRadio,
        .stSelectbox,
        .stTextInput,
        [data-testid="stExpander"] {
            display: none !important;
        }
        body, .main, .block-container {
            background-color: #ffffff !important;
            color: #000000 !important;
            padding: 0 !important;
            margin: 0 !important;
            max-width: 100% !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# =====================================================================
# ARQUIVOS E BANCO DE DADOS
# =====================================================================
ARQUIVO_ESTOQUE = "estoque_acs_quantidades.xlsx"
ARQUIVO_EVENTOS = "eventos_ativos_acs.xlsx"
ARQUIVO_HISTORICO = "historico_eventos_acs.xlsx"

def carregar_logo():
    pasta_script = os.path.dirname(os.path.abspath(__file__))
    for pasta in [pasta_script, os.getcwd(), "."]:
        try:
            for arq in os.listdir(pasta):
                if arq.lower() in ["logo.png", "logo.jpg", "logo.jpeg"]:
                    return os.path.join(pasta, arq)
        except Exception:
            pass
    return None

def carregar_estoque():
    if os.path.exists(ARQUIVO_ESTOQUE):
        try:
            df = pd.read_excel(ARQUIVO_ESTOQUE)
            if "Estoque Total" in df.columns:
                df["Código"] = df["Código"].astype(str).str.zfill(3)
                return df
        except Exception:
            pass

    dados_iniciais = [
        # Áudio — Caixas e Alto-Falantes
        ("001", "CAIXA SUBGRAVE DUPLA LS ÁUDIO SLIMPEC 218", 8, "Áudio"),
        ("002", "CAIXA DE PA LINE ARRAY LS ÁUDIO SLIMPEC 4280", 24, "Áudio"),
        ("003", "CAIXA SUBGRAVE SB800", 6, "Áudio"),
        ("004", "CAIXA KF SIDE", 2, "Áudio"),
        ("005", "CAIXA G SUB", 8, "Áudio"),
        ("006", "CAIXA MEIA SB", 2, "Áudio"),
        ("007", "CAIXA ALTA 2/1", 6, "Áudio"),
        ("008", "CAIXA ATTACK PASSIVA", 16, "Áudio"),
        ("009", "CAIXA DE MONITOR CLAIR 112", 6, "Áudio"),
        ("010", "CAIXA CÓPIA DA CLAIR", 4, "Áudio"),
        ("011", "CAIXA CLAIR PARA MONTAGEM", 6, "Áudio"),
        ("012", "CAIXA ATIVA 15PL TAPIOCA / MACKIE", 12, "Áudio"),
        ("013", "CAIXA TURBOSOUND IP2000 ULTRA FINO", 2, "Áudio"),
        ("014", "CAIXA TURBOSOUND IQ12", 2, "Áudio"),
        ("015", "CAIXA TURBOSOUND IQ18", 2, "Áudio"),
        ("016", "CAIXA SKP (1 PERDIDA)", 2, "Áudio"),
        ("017", "CAIXA PEAVEY PBK", 2, "Áudio"),
        ("018", "CAIXA CELESTION", 2, "Áudio"),
        ("019", "CAIXA PZ PX-12A", 2, "Áudio"),
        ("020", "CAIXA DATREL", 1, "Áudio"),
        ("021", "CAIXA PASSIVA 8PL FULLRANGE", 2, "Áudio"),
        ("022", "CAIXA PASSIVA 8PL BI-AMP", 4, "Áudio"),
        ("023", "CAIXA ELECTRO-VOICE ZLX 12PL", 6, "Áudio"),
        ("024", "CAIXA MJ (2 COM RAPHAEL)", 12, "Áudio"),
        ("025", "CAIXA ALTA ANTERA TS500A", 2, "Áudio"),
        ("026", "CAIXA SUBGRAVE ANTERA LF1000A", 2, "Áudio"),
        ("027", "CAIXA ALTA YAMAHA DXR10", 2, "Áudio"),
        ("028", "CAIXA SUBGRAVE YAMAHA DXS12 MKII", 2, "Áudio"),
        ("029", "CAIXA ATTACK VRM1230A", 3, "Áudio"),

        # Microfones e Kits de Bateria (Com Fios)
        ("030", "MICROFONE SHURE BETA 91A", 1, "Áudio"),
        ("031", "MICROFONE E604 SENNHEISER", 8, "Áudio"),
        ("032", "MICROFONE AKG D112", 1, "Áudio"),
        ("033", "MICROFONE SHURE SM57", 7, "Áudio"),
        ("034", "MICROFONE MXT / SM57", 10, "Áudio"),
        ("035", "MICROFONE RELACART SM300", 8, "Áudio"),
        ("036", "MICROFONE SM57 KADOSCH", 4, "Áudio"),
        ("037", "MICROFONE SHURE SM58", 16, "Áudio"),
        ("038", "MICROFONE AKG P2", 2, "Áudio"),
        ("039", "MICROFONE AKG P4", 4, "Áudio"),
        ("040", "MICROFONE SAMSON C02", 4, "Áudio"),
        ("041", "MICROFONE SAMSON C02H", 2, "Áudio"),
        ("042", "KIT DE MICROFONE DE BATERIA JTS", 1, "Áudio"),

        # Sistemas Sem Fio (In-Ear e Microfones)
        ("043", "TRANSMISSOR SHURE PSM900 P9T", 1, "Áudio"),
        ("044", "BODYPACK SHURE PSM900 P9R L6 (656/692 MHZ)", 1, "Áudio"),
        ("045", "SISTEMA IN-EAR SENNHEISER EW300 G3 (516/558 MHZ)", 2, "Áudio"),
        ("046", "BODYPACK SENNHEISER EK300", 4, "Áudio"),
        ("047", "RELACART PM-320 (522-554 MHZ)", 5, "Áudio"),
        ("048", "BODYPACK RELACART PM-320R", 10, "Áudio"),
        ("049", "COMBINADOR RELACART TAC4", 2, "Áudio"),
        ("050", "ANTENA IN-EAR RELACART R-32U", 2, "Áudio"),
        ("051", "BASE RELACART UR222D (522/586 MHZ)", 2, "Áudio"),
        ("052", "MICROFONE SEM FIO BASTÃO RELACART UH-222 (522/651 MHZ)", 4, "Áudio"),
        ("053", "TRANSMISSOR DE CORPO RELACART UT-222 (522/651 MHZ)", 2, "Áudio"),
        ("054", "BASE DUPLA UR2Q RELACART", 2, "Áudio"),
        ("055", "TRANSMISSOR RELACART", 8, "Áudio"),
        ("056", "ANTENA RELACART ATIVA", 2, "Áudio"),
        ("057", "BASE SEM FIO SHURE PGX4", 1, "Áudio"),
        ("058", "MICROFONE SEM FIO BASTÃO SHURE PGX2", 1, "Áudio"),
        ("059", "BASE SEM FIO SENNHEISER EM 3732-II", 1, "Áudio"),
        ("060", "MICROFONE SEM FIO SENNHEISER SKM 5200", 2, "Áudio"),
        ("061", "BASE SHURE BLX4R", 4, "Áudio"),
        ("062", "MICROFONE SEM FIO BASTÃO SHURE BLX2 H10 (542/572 MHZ)", 4, "Áudio"),
        ("063", "AMPLIFICADOR DE SINAL PAULKITSON UA80", 1, "Áudio"),
        ("064", "ANTENA PAULKITSON", 2, "Áudio"),
        ("065", "BASE AKG SR450", 3, "Áudio"),
        ("066", "TRANSMISSOR DE CORPO AKG PT450", 3, "Áudio"),
        ("067", "BASE SHURE UR4D", 9, "Áudio"),
        ("068", "BASE SHURE UR4D+", 7, "Áudio"),
        ("069", "MICROFONE SEM FIO BASTÃO SHURE UR4", 8, "Áudio"),
        ("070", "TRANSMISSOR SHURE UR1M", 20, "Áudio"),
        ("071", "ANTENA RELACART", 6, "Áudio"),

        # Distribuição de Fones e Multicabos (Snakes)
        ("072", "AMPLIFICADOR DE FONES POWERPLAY ASA", 1, "Áudio"),
        ("073", "AMPLIFICADOR DE FONES POWERPLAY EAM", 1, "Áudio"),
        ("074", "BODYPACK EAM", 8, "Áudio"),
        ("075", "AMPLIFICADOR DE FONES BEHRINGER", 2, "Áudio"),
        ("076", "AMPLIFICADOR DE FONES PHONIC", 1, "Áudio"),
        ("077", "AMPLIFICADOR DE FONES SAMSON", 1, "Áudio"),
        ("078", "FONE DE OUVIDO", 34, "Áudio"),
        ("079", "SUB-SNAKE DE 12 VIAS", 5, "Áudio"),
        ("080", "SUB-SNAKE DE 8 VIAS", 2, "Áudio"),
        ("081", "MEDUSA DE BATERIA DE 12 VIAS", 2, "Áudio"),
        ("082", "MEDUSA SPLITADA DE 32 VIAS GRANDE", 1, "Áudio"),
        ("083", "MEDUSA DE 36 VIAS (RUIM)", 1, "Áudio"),

        # Backline (Baterias, Cabeçotes e Caixas de Instrumentos)
        ("084", "BATERIA ACÚSTICA GRETSCH BLACK HAWK", 1, "Backline"),
        ("085", "BATERIA ACÚSTICA PEARL EXPORT", 1, "Backline"),
        ("086", "BATERIA ACÚSTICA PEARL VISION", 1, "Backline"),
        ("087", "BATERIA ELETRÔNICA ROLAND V-DRUMS TD-12", 1, "Backline"),
        ("088", "CABEÇOTE DE BAIXO BORNE PRO 800", 1, "Backline"),
        ("089", "CABEÇOTE DE BAIXO GK 400 BR", 1, "Backline"),
        ("090", "AMPLIFICADOR DE BAIXO HARTKE", 1, "Backline"),
        ("091", "AMPLIFICADOR DE GUITARRA PEAVEY BANDIT 112", 2, "Backline"),
        ("092", "CAIXA DE BAIXO WARWICK 4X10", 1, "Backline"),
        ("093", "CAIXA DE BAIXO WARWICK 1X15", 1, "Backline"),
        ("094", "AMPLIFICADOR DE GUITARRA FENDER FRONTMAN 212R", 1, "Backline"),
        ("095", "AMPLIFICADOR DE GUITARRA METEORO", 1, "Backline"),

        # Mesas de Som (Consoles)
        ("096", "MESA DE SOM YAMAHA 01V", 3, "Áudio"),
        ("097", "MESA DE SOM YAMAHA LS9-32", 2, "Áudio"),
        ("098", "MESA DE SOM MIDAS M32", 1, "Áudio"),
        ("099", "MESA DE SOM YAMAHA TF5", 1, "Áudio"),
        ("100", "MESA DE SOM SOUNDCRAFT UI24R", 1, "Áudio"),
        ("101", "MESA DE SOM BEHRINGER X-AIR XR18", 1, "Áudio"),
        ("102", "MESA DE SOM PHONIC AM105 (2 CANAIS)", 1, "Áudio"),
        ("103", "MESA DE SOM PHONIC DE 4 CANAIS", 1, "Áudio"),
        ("104", "MESA DE SOM WALDMAN ST-8DSP (6 CANAIS)", 1, "Áudio"),
        ("105", "MESA DE SOM SOUNDCRAFT SIGNATURE 12", 1, "Áudio"),
        ("106", "MESA DE SOM YAMAHA MG12XU", 1, "Áudio"),
        ("107", "MESA DE SOM SOUNDVOICE AUREA MD-12", 1, "Áudio"),

        # Cabos, Energia e Acessórios
        ("108", "PEDESTAL", 41, "Acessórios"),
        ("109", "RÉGUA DE ENERGIA POWERCON", 14, "Acessórios"),
        ("110", "RÉGUA DE ENERGIA POWERCON SAS", 6, "Acessórios"),
        ("111", "RÉGUA DE ENERGIA POWERCON / POWERCON", 4, "Acessórios"),
        ("112", "PROCESSADOR BEHRINGER", 1, "Áudio"),
        ("113", "PROCESSADOR DBX", 3, "Áudio"),
        ("114", "AMPLIFICADOR CROWN XTI 1000", 1, "Áudio"),
        ("115", "AMPLIFICADOR CROWN XTI 4000", 2, "Áudio"),
        ("116", "AMPLIFICADOR CROWN XTI 6000", 1, "Áudio"),
        ("117", "HOT SOUND HS PRO 5", 1, "Áudio"),
        ("118", "HOT SOUND DIGILITE 6.4", 1, "Áudio"),
        ("119", "HOT SOUND DIGILITE 8.0", 1, "Áudio"),
        ("120", "YAMAHA P3200", 1, "Áudio"),
        ("121", "AMPLIFICADOR CROWN / CARVER (RUIM)", 3, "Áudio"),
        ("122", "STUDIO R Z2400", 1, "Áudio"),
        ("123", "STUDIO R XD", 1, "Áudio"),
        ("124", "STUDIO R X5", 1, "Áudio"),
        ("125", "STUDIO R X12", 1, "Áudio"),
        ("126", "APD 4004 THUNDELIGHT", 2, "Áudio"),
        ("127", "POWERTEC LS", 4, "Áudio"),
        ("128", "RACK DE AC AZUL (RUIM)", 1, "Acessórios"),
        ("129", "RACK DE AC NOVO", 1, "Acessórios"),
        ("130", "CENTRAL DE ENERGIA", 4, "Acessórios"),
        ("131", "ISOLADOR DE ÁUDIO GX200", 2, "Áudio"),
        ("132", "DIRECT BOX WHIRLWIND IMP2", 5, "Áudio"),
        ("133", "DIRECT BOX WALDMAN PASSIVA DI-1PS", 5, "Áudio"),
        ("134", "DIRECT BOX WIRECONEX WDI600", 4, "Áudio"),
        ("135", "DIRECT BOX WIRECONEX WDI500", 20, "Áudio"),
        ("136", "DIRECT BOX WALDMAN ATIVA DI-1A", 7, "Áudio"),
        ("137", "DIRECT BOX SAMSON ATIVA S. DIRECT", 3, "Áudio"),
        ("138", "DIRECT BOX SAMSON ATIVA S. DIRECT PLUS", 1, "Áudio"),
        ("139", "MOUSE PAD", 1, "Informática"),
        ("140", "CPU WYSE DELL", 1, "Informática"),
        ("141", "MONITOR DELL P2418HT", 1, "Informática"),
        ("142", "FONTE DE MONITOR HA65NS5-00", 1, "Informática"),
        ("143", "CABO RJ45", 1, "Informática"),
        ("144", "CABO DISPLAYPORT", 1, "Informática"),
        ("145", "CABO USB-SS DELL", 1, "Informática"),
        ("146", "CABO DE FORÇA PARA FONTE", 1, "Informática"),
        ("147", "JUMPER SPEAKON (4 VIAS - ALTAS)", 18, "Áudio"),
        ("148", "JUMPER SPEAKON (2 VIAS - SUB)", 6, "Áudio"),
        ("149", "CABO SPEAKON (4 VIAS - ALTAS)", 8, "Áudio"),
        ("150", "CABO SPEAKON (2 VIAS - SUB)", 8, "Áudio"),
        ("151", "CABO POWERCON / POWERCON", 20, "Acessórios"),
        ("152", "CABO XLR DE 30 METROS", 8, "Áudio"),
        ("153", "CABO XLR DE 10 METROS", 60, "Áudio"),
        ("154", "CABO XLR DE 5 METROS", 90, "Áudio"),
        
        # Cabos XLR para Barra Garden (Iluminação) e Outros
        ("155", "CABO XLR DA BARRA GARDEN - GRANADA DE 27 METROS", 1, "Luz"),
        ("156", "CABO XLR DA BARRA GARDEN - GRANADA DE 17 METROS", 1, "Luz"),
        ("157", "CABO XLR DA BARRA GARDEN - STECK / STECK", 5, "Luz"),
        ("158", "CABO XLR DA BARRA GARDEN - SINGELOS", 4, "Luz")
    ]

    lista_estoque = [{
        "Código": str(codigo).zfill(3),
        "Categoria": cat,
        "Equipamento": nome,
        "Estoque Total": int(qtd),
        "Disponível": int(qtd),
        "Em Uso": 0
    } for codigo, nome, qtd, cat in dados_iniciais]

    df_inicial = pd.DataFrame(lista_estoque)
    df_inicial.to_excel(ARQUIVO_ESTOQUE, index=False)
    return df_inicial

def salvar_estoque(df):
    df.to_excel(ARQUIVO_ESTOQUE, index=False)

def carregar_eventos():
    if os.path.exists(ARQUIVO_EVENTOS):
        try:
            df_ev = pd.read_excel(ARQUIVO_EVENTOS)
            eventos = []
            for _, row in df_ev.iterrows():
                eventos.append({
                    "id": str(row["ID"]),
                    "evento": str(row["Evento"]),
                    "responsavel": str(row["Responsável"]),
                    "endereco": str(row.get("Endereço", "-")),
                    "hora_galpao": str(row.get("Hora Galpão", "-")),
                    "hora_evento": str(row.get("Hora Evento", "-")),
                    "data_saida": str(row["Data Saída"]),
                    "itens": json.loads(row["Itens JSON"])
                })
            return eventos
        except Exception:
            return []
    return []

def salvar_eventos(eventos):
    dados_tabela = []
    for ev in eventos:
        dados_tabela.append({
            "ID": ev["id"],
            "Evento": ev["evento"],
            "Responsável": ev["responsavel"],
            "Endereço": ev.get("endereco", "-"),
            "Hora Galpão": ev.get("hora_galpao", "-"),
            "Hora Evento": ev.get("hora_evento", "-"),
            "Data Saída": ev["data_saida"],
            "Itens JSON": json.dumps(ev["itens"], ensure_ascii=False)
        })
    pd.DataFrame(dados_tabela).to_excel(ARQUIVO_EVENTOS, index=False)

def registrar_historico(acao, evento, responsavel, resumo_itens, observacao="-"):
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M')
    novo_reg = pd.DataFrame([{
        "Data/Hora": data_hora,
        "Operação": acao,
        "Evento / Projeto": evento,
        "Responsável Técnico": responsavel,
        "Relação de Material": resumo_itens,
        "Observações Técnicas": observacao
    }])
    if os.path.exists(ARQUIVO_HISTORICO):
        df_hist = pd.read_excel(ARQUIVO_HISTORICO)
        df_hist = pd.concat([df_hist, novo_reg], ignore_index=True)
    else:
        df_hist = novo_reg
    df_hist.to_excel(ARQUIVO_HISTORICO, index=False)

def gerar_pdf_checklist(ev):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )
    
    elementos = []
    styles = getSampleStyleSheet()
    
    estilo_titulo = ParagraphStyle(
        'TituloDoc',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=15,
        textColor=colors.HexColor('#0f172a'),
        alignment=1
    )
    estilo_cab_empresa = ParagraphStyle(
        'CabEmpresa',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=colors.white,
        alignment=0
    )
    estilo_info = ParagraphStyle(
        'InfoCab',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor('#1e293b')
    )
    estilo_celula = ParagraphStyle(
        'CelulaTabela',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#0f172a')
    )
    estilo_celula_bold = ParagraphStyle(
        'CelulaTabelaBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#0f172a'),
        alignment=1
    )
    estilo_th = ParagraphStyle(
        'TabelaHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
        alignment=1
    )

    logo_path = carregar_logo()
    texto_empresa = Paragraph(
        "<b>A.C.S. PRODUÇÕES & DIVULGAÇÕES LTDA</b><br/>"
        "<font size=8 color='#94a3b8'>A.C.S. SOM, LUZ E IMAGEM — GESTÃO TÉCNICA E LOGÍSTICA DE EVENTOS</font>",
        estilo_cab_empresa
    )
    
    bloco_logo = Paragraph("<b>A.C.S.</b>", estilo_cab_empresa)
    if logo_path:
        try:
            ir = ImageReader(logo_path)
            iw, ih = ir.getSize()
            larg_desejada = 85
            alt_desejada = larg_desejada * (ih / float(iw))
            if alt_desejada > 48:
                alt_desejada = 48
                larg_desejada = alt_desejada * (iw / float(ih))
            bloco_logo = RLImage(logo_path, width=larg_desejada, height=alt_desejada, mask='auto')
            bloco_logo.hAlign = 'CENTER'
        except Exception:
            pass

    t_cab = Table([[bloco_logo, texto_empresa]], colWidths=[105, 440])
    t_cab.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#0f172a')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,0), 'CENTER'),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LINEBELOW', (0,0), (-1,-1), 3, colors.HexColor('#0d9488')),
    ]))
    elementos.append(t_cab)
    elementos.append(Spacer(1, 10))

    elementos.append(Paragraph("<b>CHECKLIST DE EXPEDIÇÃO E CONFERÊNCIA TÉCNICA</b>", estilo_titulo))
    elementos.append(Spacer(1, 8))

    info_dados = [
        [Paragraph(f"<b>Documento Nº:</b> {ev['id']}", estilo_info), Paragraph(f"<b>Data de Saída:</b> {ev['data_saida']}", estilo_info)],
        [Paragraph(f"<b>Evento / Destino:</b> {ev['evento']}", estilo_info), Paragraph(f"<b>Responsável Técnico:</b> {ev['responsavel']}", estilo_info)],
        [Paragraph(f"<b>Endereço:</b> {ev.get('endereco', '-')}", estilo_info), Paragraph(f"<b>Horários:</b> Galpão: {ev.get('hora_galpao', '-')} | Evento: {ev.get('hora_evento', '-')}", estilo_info)]
    ]
    t_info = Table(info_dados, colWidths=[280, 265])
    t_info.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f1f5f9')),
        ('PADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#cbd5e1')),
    ]))
    elementos.append(t_info)
    elementos.append(Spacer(1, 10))

    cabecalho_tabela = [
        Paragraph("<b>CÓD.</b>", estilo_th),
        Paragraph("<b>ESPECIFICAÇÃO DO EQUIPAMENTO</b>", estilo_th),
        Paragraph("<b>QTD</b>", estilo_th),
        Paragraph("<b>SAÍDA</b>", estilo_th),
        Paragraph("<b>RETORNO</b>", estilo_th),
        Paragraph("<b>OBSERVAÇÕES</b>", estilo_th)
    ]
    
    dados_tabela = [cabecalho_tabela]
    for item in ev["itens"]:
        cod_formatado = str(item['Código']).zfill(3)
        dados_tabela.append([
            Paragraph(cod_formatado, estilo_celula_bold),
            Paragraph(str(item['Equipamento']), estilo_celula),
            Paragraph(f"{item['Qtd']} un", estilo_celula_bold),
            Paragraph("[ &nbsp; ]", estilo_celula_bold),
            Paragraph("[ &nbsp; ]", estilo_celula_bold),
            Paragraph("_______________", estilo_celula_bold)
        ])

    t_itens = Table(dados_tabela, colWidths=[45, 240, 45, 50, 55, 110], repeatRows=1)
    t_itens.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0f172a')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
    ]))
    elementos.append(t_itens)
    elementos.append(Spacer(1, 25))

    ass_dados = [
        [Paragraph(f"__________________________________________________<br/><b>Responsável Técnico:</b> {ev['responsavel']}", estilo_celula),
         Paragraph("__________________________________________________<br/><b>Conferência de Retorno (Galpão):</b> Data ___/___/_____", estilo_celula)]
    ]
    t_ass = Table(ass_dados, colWidths=[272, 273])
    t_ass.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ('TOPPADDING', (0,0), (-1,-1), 15),
    ]))
    elementos.append(t_ass)

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()

def gerar_texto_historico(df_h):
    linhas = [
        "======================================================================",
        "               A.C.S. PRODUÇÕES & DIVULGAÇÕES LTDA                    ",
        "            RELATÓRIO GERAL DE MOVIMENTAÇÕES (HISTÓRICO)              ",
        f"            EMITIDO EM: {datetime.now().strftime('%d/%m/%Y %H:%M')}                           ",
        "======================================================================",
        ""
    ]
    for idx, row in df_h.iterrows():
        linhas.append(f"REGISTRO #{idx + 1} | DATA/HORA: {row['Data/Hora']} | OPERAÇÃO: {row['Operação']}")
        linhas.append(f"EVENTO      : {row['Evento / Projeto']}")
        linhas.append(f"RESPONSÁVEL : {row['Responsável Técnico']}")
        linhas.append(f"MATERIAIS   : {row['Relação de Material']}")
        linhas.append(f"OBSERVAÇÕES : {row['Observações Técnicas']}")
        linhas.append("----------------------------------------------------------------------")
    return "\n".join(linhas)

# =====================================================================
# AUTENTICAÇÃO CORPORATIVA
# =====================================================================
logo_path = carregar_logo()

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if "carrinho_saida" not in st.session_state:
    st.session_state.carrinho_saida = {}

if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, col_centro, _ = st.columns([1, 1.3, 1])
    with col_centro:
        if logo_path:
            st.image(logo_path, width=140)
        st.markdown("### A.C.S. PRODUÇÕES & DIVULGAÇÕES LTDA")
        st.caption("SISTEMA DE CONTROLE PATRIMONIAL E LOGÍSTICA DE EVENTOS")
        st.markdown("---")
        with st.form("form_login"):
            senha_digitada = st.text_input("Credencial de Acesso:", type="password")
            botao_login = st.form_submit_button("AUTENTICAR ACESSO", use_container_width=True)
            if botao_login:
                if senha_digitada == SENHA_ACESSO:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("Credencial inválida. Verifique a senha informada.")
    st.stop()

# =====================================================================
# CARREGAMENTO E MENU SUPERIOR ADAPTATIVO PARA COMPUTADOR E TELEMÓVEL
# =====================================================================
df_estoque = carregar_estoque()
eventos_ativos = carregar_eventos()

col_h1, col_h2 = st.columns([1, 5])
with col_h1:
    if logo_path:
        st.image(logo_path, width=120)
with col_h2:
    st.markdown("### 🎛️ PAINEL OPERACIONAL A.C.S.")

# NOVA ORDEM DOS MÓDULOS SOLICITADA
menu = st.selectbox(
    "📍 SELECIONE O MÓDULO DO SISTEMA:",
    [
        "VISÃO GERAL DO ACERVO",
        "EXPEDIÇÃO (SAÍDA)",
        "EMISSÃO DE CHECKLIST",
        "RETORNO E CONFERÊNCIA",
        "CADASTRO DE MATERIAIS",
        "HISTÓRICO DE OPERAÇÕES"
    ]
)

col_topo_dir1, col_topo_dir2 = st.columns([5, 1])
with col_topo_dir2:
    if st.button("SAIR", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

st.markdown("---")

# =====================================================================
# CABEÇALHO E PAINEL DE INDICADORES
# =====================================================================
st.markdown(f"""
<div class="corporate-header">
    <div>
        <p class="corporate-title">A.C.S. — CONTROLE DE EQUIPAMENTOS E OPERAÇÕES</p>
        <p class="corporate-subtitle">Gestão de Estoque Técnico, Expedição e Checklist de Eventos</p>
    </div>
    <div style="text-align: right; color: #94a3b8; font-size: 0.85rem;">
        DATA DO SISTEMA<br><strong style="color: #ffffff;">{datetime.now().strftime('%d/%m/%Y')}</strong>
    </div>
</div>
""", unsafe_allow_html=True)

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total de Unidades no Acervo</div>
        <div class="kpi-value">{int(df_estoque['Estoque Total'].sum())}</div>
    </div>
    """, unsafe_allow_html=True)
with k2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Unidades Disponíveis (Base)</div>
        <div class="kpi-value kpi-accent">{int(df_estoque['Disponível'].sum())}</div>
    </div>
    """, unsafe_allow_html=True)
with k3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Unidades Alocadas (Em Evento)</div>
        <div class="kpi-value">{int(df_estoque['Em Uso'].sum())}</div>
    </div>
    """, unsafe_allow_html=True)
with k4:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Operações / Eventos Ativos</div>
        <div class="kpi-value">{len(eventos_ativos)}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =====================================================================
# MÓDULO 1: VISÃO GERAL DO ACERVO (COM REMOÇÃO SEGURA)
# =====================================================================
if menu == "VISÃO GERAL DO ACERVO":
    st.markdown('<div class="section-header">Posição Consolidada do Acervo Técnico</div>', unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        busca = st.text_input("Pesquisar equipamento por descrição ou código:", placeholder="Digite o modelo, marca ou código...")
    with c2:
        f_cat = st.selectbox("Filtrar por Departamento:", ["Todos", "Áudio", "Backline", "Luz", "Informática", "Acessórios"])
    with c3:
        f_disp = st.selectbox("Status de Alocação:", ["Todos os Itens", "Somente Alocados em Eventos", "Somente Disponíveis na Base"])

    df_view = df_estoque.copy()
    if busca:
        df_view = df_view[
            df_view["Equipamento"].str.contains(busca, case=False, na=False) | 
            df_view["Código"].str.contains(busca, case=False, na=False)
        ]
    if f_cat != "Todos":
        df_view = df_view[df_view["Categoria"] == f_cat]
    if f_disp == "Somente Alocados em Eventos":
        df_view = df_view[df_view["Em Uso"] > 0]
    elif f_disp == "Somente Disponíveis na Base":
        df_view = df_view[df_view["Disponível"] > 0]

    st.dataframe(
        df_view,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Código": st.column_config.TextColumn("CÓDIGO", width="small"),
            "Categoria": st.column_config.TextColumn("SETOR", width="small"),
            "Equipamento": st.column_config.TextColumn("ESPECIFICAÇÃO DO EQUIPAMENTO", width="large"),
            "Estoque Total": st.column_config.NumberColumn("ACERVO TOTAL", width="small"),
            "Disponível": st.column_config.NumberColumn("DISPONÍVEL", width="small"),
            "Em Uso": st.column_config.NumberColumn("ALOCADO", width="small")
        }
    )

    with open(ARQUIVO_ESTOQUE, "rb") as f:
        st.download_button(
            "EXPORTAR PLANILHA DE ESTOQUE (.XLSX)",
            f,
            file_name=f"acervo_acs_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("ZONA ADMINISTRATIVA — REMOVER EQUIPAMENTO DO ACERVO (SENHA EXCLUSIVA)"):
        st.warning("Atenção: A remoção de um item do acervo é permanente e atualizará o inventário geral.")
        
        opcoes_itens = [f"Cód: {row['Código']} | {row['Equipamento']} ({row['Categoria']})" for _, row in df_estoque.iterrows()]
        item_para_remover = st.selectbox("Selecione o equipamento que deseja remover do acervo:", options=opcoes_itens)
        
        senha_del_item = st.text_input("Digite a Senha Administrativa de Exclusão:", type="password", key="pwd_del_item")
        
        if st.button("REMOVER ITEM SELECIONADO"):
            if senha_del_item != SENHA_EXCLUSAO:
                st.error("Senha administrativa incorreta. Operação bloqueada.")
            elif not item_para_remover:
                st.warning("Selecione um equipamento válido.")
            else:
                cod_extraido = item_para_remover.split("|")[0].replace("Cód:", "").strip().zfill(3)
                df_estoque = df_estoque[df_estoque["Código"] != cod_extraido].reset_index(drop=True)
                salvar_estoque(df_estoque)
                st.success(f"Equipamento com código {cod_extraido} removido permanentemente do acervo com sucesso!")
                st.rerun()

# =====================================================================
# MÓDULO 2: EXPEDIÇÃO (SAÍDA DE MATERIAL) — COM AUTOCOMPLETAR GOOGLE & HORÁRIOS
# =====================================================================
elif menu == "EXPEDIÇÃO (SAÍDA)":
    st.markdown('<div class="section-header">Registro de Expedição e Alocação por Evento</div>', unsafe_allow_html=True)

    col_ev1, col_ev2 = st.columns(2)
    with col_ev1:
        nome_evento = st.text_input("Identificação do Evento / Cliente:", placeholder="Ex: Convenção Corporativa - Barra")
        tecnico_resp = st.text_input("Responsável Técnico pela Retirada:", placeholder="Nome completo do responsável")
        data_saida = st.date_input("Data de Expedição:", datetime.now(), format="DD/MM/YYYY")
    with col_ev2:
        # Campo de Endereço com componente HTML Google Places Autocomplete integrado
        st.markdown("<strong>Endereço do Evento / Destino (Google Autocomplete):</strong>", unsafe_allow_html=True)
        endereco_evento = st.text_input(
            "Endereço do Evento / Destino",
            label_visibility="collapsed",
            placeholder="Digite o endereço (Rua, Número, Bairro, Cidade)..."
        )
        
        col_h1, col_h2 = st.columns(2)
        with col_h1:
            hora_galpao = st.time_input("Horário Chegada Galpão:", time(8, 0))
        with col_h2:
            hora_evento = st.time_input("Horário Chegada Evento:", time(10, 0))

    st.markdown('<div class="section-header">Pesquisa e Seleção de Equipamentos</div>', unsafe_allow_html=True)
    
    col_busca_s, col_cat_s = st.columns([2, 1])
    with col_busca_s:
        busca_saida = st.text_input(
            "Pesquisar equipamento para adicionar à saída:",
            placeholder="Digite qualquer parte do nome ou código (ex: SM58, Line Array, XLR, Moving)..."
        )
    with col_cat_s:
        cat_saida = st.selectbox("Filtrar por Setor:", ["Todos os Setores", "Áudio", "Backline", "Luz", "Informática", "Acessórios"])

    df_disp_saida = df_estoque[df_estoque["Disponível"] > 0][["Código", "Categoria", "Equipamento", "Disponível"]].copy()
    
    if busca_saida:
        df_disp_saida = df_disp_saida[
            df_disp_saida["Equipamento"].str.contains(busca_saida, case=False, na=False) |
            df_disp_saida["Código"].str.contains(busca_saida, case=False, na=False)
        ]
    if cat_saida != "Todos os Setores":
        df_disp_saida = df_disp_saida[df_disp_saida["Categoria"] == cat_saida]

    df_disp_saida = df_disp_saida.reset_index(drop=True)
    df_disp_saida["Qtd Expedir"] = df_disp_saida["Código"].map(lambda c: st.session_state.carrinho_saida.get(str(c).zfill(3), 0))

    def atualizar_selecao_saida():
        mudancas = st.session_state.get("editor_expedicao_busca", {}).get("edited_rows", {})
        for idx_str, cols in mudancas.items():
            idx = int(idx_str)
            if idx < len(df_disp_saida) and "Qtd Expedir" in cols:
                cod_item = str(df_disp_saida.iloc[idx]["Código"]).zfill(3)
                nova_qtd = int(cols["Qtd Expedir"] or 0)
                if nova_qtd > 0:
                    st.session_state.carrinho_saida[cod_item] = nova_qtd
                else:
                    st.session_state.carrinho_saida.pop(cod_item, None)

    df_disp_saida["Saldo Restante"] = df_disp_saida["Disponível"] - df_disp_saida["Qtd Expedir"]

    st.data_editor(
        df_disp_saida[["Código", "Categoria", "Equipamento", "Disponível", "Qtd Expedir", "Saldo Restante"]],
        key="editor_expedicao_busca",
        on_change=atualizar_selecao_saida,
        hide_index=True,
        use_container_width=True,
        disabled=["Código", "Categoria", "Equipamento", "Disponível", "Saldo Restante"],
        column_config={
            "Código": st.column_config.TextColumn("CÓDIGO", width="small"),
            "Categoria": st.column_config.TextColumn("SETOR", width="small"),
            "Equipamento": st.column_config.TextColumn("ESPECIFICAÇÃO DO EQUIPAMENTO", width="large"),
            "Disponível": st.column_config.NumberColumn("SALDO BASE", width="small"),
            "Qtd Expedir": st.column_config.NumberColumn("QTD A EXPEDIR", min_value=0, step=1, width="small"),
            "Saldo Restante": st.column_config.NumberColumn("SALDO RESTANTE", width="small")
        }
    )

    itens_no_carrinho = []
    for cod_sel, qtd_sel in list(st.session_state.carrinho_saida.items()):
        cod_sel_pad = str(cod_sel).zfill(3)
        linha_est = df_estoque[df_estoque["Código"] == cod_sel_pad]
        if not linha_est.empty and qtd_sel > 0:
            itens_no_carrinho.append({
                "Código": cod_sel_pad,
                "Categoria": linha_est.iloc[0]["Categoria"],
                "Equipamento": linha_est.iloc[0]["Equipamento"],
                "Disponível": int(linha_est.iloc[0]["Disponível"]),
                "Qtd Expedir": int(qtd_sel)
            })

    df_escolhidos = pd.DataFrame(itens_no_carrinho)

    if not df_escolhidos.empty:
        st.markdown("<div class='section-header'>Resumo dos Materiais Selecionados para o Evento</div>", unsafe_allow_html=True)
        st.dataframe(df_escolhidos[["Código", "Categoria", "Equipamento", "Qtd Expedir"]], hide_index=True, use_container_width=True)
        if st.button("LIMPAR SELEÇÃO DE MATERIAIS"):
            st.session_state.carrinho_saida = {}
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("CONFIRMAR EXPEDIÇÃO E GERAR CHECKLIST", type="primary", use_container_width=True):
        if not nome_evento or not tecnico_resp:
            st.error("Preenchimento obrigatório: informe a Identificação do Evento e o Responsável Técnico.")
        elif df_escolhidos.empty:
            st.error("Nenhum material selecionado. Pesquise o equipamento e informe a quantidade na coluna 'QTD A EXPEDIR'.")
        else:
            erro_qtd = False
            for _, row in df_escolhidos.iterrows():
                if row["Qtd Expedir"] > row["Disponível"]:
                    st.error(f"Saldo insuficiente para '{row['Equipamento']}' (Solicitado: {row['Qtd Expedir']} | Disponível: {row['Disponível']}).")
                    erro_qtd = True
            
            if not erro_qtd:
                lista_itens_evento = []
                resumo_log = []
                for _, row in df_escolhidos.iterrows():
                    cod = str(row["Código"]).zfill(3)
                    qtd_saindo = int(row["Qtd Expedir"])
                    nome_eq = row["Equipamento"]
                    
                    df_estoque.loc[df_estoque["Código"] == cod, "Disponível"] -= qtd_saindo
                    df_estoque.loc[df_estoque["Código"] == cod, "Em Uso"] += qtd_saindo
                    
                    lista_itens_evento.append({
                        "Código": cod,
                        "Equipamento": nome_eq,
                        "Qtd": qtd_saindo
                    })
                    resumo_log.append(f"{qtd_saindo} un - {nome_eq}")

                salvar_estoque(df_estoque)
                
                novo_evento = {
                    "id": datetime.now().strftime("%Y%m%d%H%M%S"),
                    "evento": nome_evento.upper(),
                    "responsavel": tecnico_resp.upper(),
                    "endereco": endereco_evento.upper() if endereco_evento else "NÃO INFORMADO",
                    "hora_galpao": hora_galpao.strftime("%H:%M"),
                    "hora_evento": hora_evento.strftime("%H:%M"),
                    "data_saida": data_saida.strftime("%d/%m/%Y"),
                    "itens": lista_itens_evento
                }
                eventos_ativos.append(novo_evento)
                salvar_eventos(eventos_ativos)
                
                registrar_historico("EXPEDIÇÃO", nome_evento.upper(), tecnico_resp.upper(), "; ".join(resumo_log), f"Saída técnica autorizada | Destino: {endereco_evento}")
                st.session_state.carrinho_saida = {}
                st.success(f"Operação registrada com sucesso para o evento '{nome_evento.upper()}'. Acesse 'EMISSÃO DE CHECKLIST' para descarregar o PDF profissional.")
                st.rerun()

# =====================================================================
# MÓDULO 3: EMISSÃO DE CHECKLIST (PDF PROFISSIONAL COM LOGOMARCA)
# =====================================================================
elif menu == "EMISSÃO DE CHECKLIST":
    st.markdown('<div class="section-header">Emissão de Documentos Operacionais em PDF</div>', unsafe_allow_html=True)
    
    if not eventos_ativos:
        st.warning("Não existem relatórios de saída ativos no momento.")
    else:
        opcoes_ev = [f"{ev['evento']} (Data: {ev['data_saida']} | Resp: {ev['responsavel']})" for ev in eventos_ativos]
        idx_imp = st.selectbox("Selecione o Checklist de Saída:", range(len(opcoes_ev)), format_func=lambda x: opcoes_ev[x])
        ev_imp = eventos_ativos[idx_imp]
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        pdf_bytes = gerar_pdf_checklist(ev_imp)
        nome_arq_pdf = f"Checklist_{ev_imp['evento'].replace(' ', '_')}_{ev_imp['id']}.pdf"
        
        st.download_button(
            label="📄 BAIXAR CHECKLIST EM PDF PROFISSIONAL (COM LOGOMARCA)",
            data=pdf_bytes,
            file_name=nome_arq_pdf,
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"**Itens registados neste Checklist ({len(ev_imp['itens'])} equipamentos):**")
        df_preview = pd.DataFrame(ev_imp["itens"])
        df_preview["Código"] = df_preview["Código"].astype(str).str.zfill(3)
        st.dataframe(df_preview, hide_index=True, use_container_width=True)

        st.markdown("<br><hr>", unsafe_allow_html=True)
        with st.expander("ZONA ADMINISTRATIVA — APAGAR ESTE CHECKLIST DE SAÍDA (SENHA EXCLUSIVA)"):
            st.warning("Atenção: Ao apagar este relatório de saída ativo, as quantidades alocadas retornarão automaticamente ao saldo disponível no estoque.")
            senha_del_romaneio = st.text_input("Digite a Senha Administrativa de Exclusão:", type="password", key="pwd_del_romaneio")
            if st.button("APAGAR CHECKLIST SELECIONADO"):
                if senha_del_romaneio == SENHA_EXCLUSAO:
                    for item in ev_imp["itens"]:
                        cod_i = str(item["Código"]).zfill(3)
                        qtd_i = int(item["Qtd"])
                        df_estoque.loc[df_estoque["Código"] == cod_i, "Disponível"] += qtd_i
                        df_estoque.loc[df_estoque["Código"] == cod_i, "Em Uso"] = max(0, int(df_estoque.loc[df_estoque["Código"] == cod_i, "Em Uso"].values[0]) - qtd_i)
                    salvar_estoque(df_estoque)
                    eventos_ativos.pop(idx_imp)
                    salvar_eventos(eventos_ativos)
                    st.success("Checklist apagado e quantidades estornadas ao estoque com sucesso.")
                    st.rerun()
                else:
                    st.error("Senha administrativa incorreta. Exclusão não autorizada.")

# =====================================================================
# MÓDULO 4: RETORNO E CONFERÊNCIA
# =====================================================================
elif menu == "RETORNO E CONFERÊNCIA":
    st.markdown('<div class="section-header">Recebimento e Conferência de Material Pós-Evento</div>', unsafe_allow_html=True)
    
    if not eventos_ativos:
        st.info("Não constam operações externas pendentes de retorno no momento.")
    else:
        opcoes_ev = [f"OPERAÇÃO: {ev['evento']} | RESPONSÁVEL: {ev['responsavel']} | DATA: {ev['data_saida']}" for ev in eventos_ativos]
        ev_selecionado_idx = st.selectbox("Selecione a operação em retorno:", range(len(opcoes_ev)), format_func=lambda x: opcoes_ev[x])
        
        ev_atual = eventos_ativos[ev_selecionado_idx]
        
        df_ret = pd.DataFrame(ev_atual["itens"])
        df_ret["Código"] = df_ret["Código"].astype(str).str.zfill(3)
        df_ret["Validado"] = False
        df_ret["Laudo / Ocorrência"] = "Sem avarias"

        df_ret_editado = pd.DataFrame()
        try:
            df_ret_editado = st.data_editor(
                df_ret,
                key=f"retorno_{ev_atual['id']}",
                hide_index=True,
                use_container_width=True,
                disabled=["Código", "Equipamento", "Qtd"],
                column_config={
                    "Validado": st.column_config.CheckboxColumn("CONFERIDO", default=False, width="small"),
                    "Código": st.column_config.TextColumn("CÓDIGO", width="small"),
                    "Equipamento": st.column_config.TextColumn("ESPECIFICAÇÃO DO EQUIPAMENTO", width="large"),
                    "Qtd": st.column_config.NumberColumn("QTD EXPEDIDA", width="small"),
                    "Laudo / Ocorrência": st.column_config.TextColumn("LAUDO TÉCNICO DE RETORNO", width="medium")
                }
            )
        except Exception:
            df_ret_editado = df_ret

        col_r1, col_r2 = st.columns(2)
        with col_r1:
            conferente_volta = st.text_input("Responsável pela Conferência de Recebimento:", value=ev_atual['responsavel'])
        with col_r2:
            obs_geral = st.text_input("Parecer Geral da Operação:", value="Material reintegrado ao estoque sem pendências")

        if st.button("FINALIZAR CONFERÊNCIA E REINTEGRAR ESTOQUE", type="primary", use_container_width=True):
            if not all(df_ret_editado["Validado"]):
                st.warning("Atenção: Para encerrar a operação, todos os equipamentos devem ser conferidos e marcados na coluna 'CONFERIDO'.")
            elif not conferente_volta:
                st.error("Informe o responsável pela conferência de recebimento.")
            else:
                resumo_volta = []
                avarias_lista = []
                for _, row in df_ret_editado.iterrows():
                    cod = str(row["Código"]).zfill(3)
                    qtd_v = int(row["Qtd"])
                    df_estoque.loc[df_estoque["Código"] == cod, "Disponível"] += qtd_v
                    df_estoque.loc[df_estoque["Código"] == cod, "Em Uso"] = max(0, int(df_estoque.loc[df_estoque["Código"] == cod, "Em Uso"].values[0]) - qtd_v)
                    resumo_volta.append(f"{qtd_v} un - {row['Equipamento']}")
                    if row["Laudo / Ocorrência"].strip().upper() not in ["SEM AVARIAS", "OK"]:
                        avarias_lista.append(f"{row['Equipamento']}: {row['Laudo / Ocorrência']}")

                salvar_estoque(df_estoque)
                eventos_ativos.pop(ev_selecionado_idx)
                salvar_eventos(eventos_ativos)
                
                obs_final = obs_geral
                if avarias_lista:
                    obs_final += " | OCORRÊNCIAS: " + "; ".join(avarias_lista)

                registrar_historico("RETORNO", ev_atual["evento"], conferente_volta.upper(), "; ".join(resumo_volta), obs_final)
                st.success(f"Recebimento do evento '{ev_atual['evento']}' homologado com sucesso.")
                st.rerun()

# =====================================================================
# MÓDULO 5: CADASTRO DE MATERIAIS (COM CÓDIGO AUTOMÁTICO)
# =====================================================================
elif menu == "CADASTRO DE MATERIAIS":
    st.markdown('<div class="section-header">Inclusão de Novos Equipamentos no Patrimônio</div>', unsafe_allow_html=True)
    
    try:
        ultimo_codigo = int(df_estoque["Código"].max())
        proximo_codigo = str(ultimo_codigo + 1).zfill(3)
    except Exception:
        proximo_codigo = "001"

    with st.form("form_novo_item"):
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Código Patrimonial Gerado Automaticamente:", value=proximo_codigo, disabled=True)
            n_nome = st.text_input("Descrição Técnica / Modelo do Equipamento:")
        with c2:
            n_cat = st.selectbox("Departamento Técnico:", ["Áudio", "Backline", "Luz", "Informática", "Acessórios"])
            n_qtd = st.number_input("Quantidade Total Incorporada:", min_value=1, value=1, step=1)
        
        if st.form_submit_button("REGISTRAR EQUIPAMENTO NO ACERVO"):
            if n_nome:
                novo_df = pd.DataFrame([{
                    "Código": proximo_codigo,
                    "Categoria": n_cat,
                    "Equipamento": n_nome.strip().upper(),
                    "Estoque Total": int(n_qtd),
                    "Disponível": int(n_qtd),
                    "Em Uso": 0
                }])
                df_estoque = pd.concat([df_estoque, novo_df], ignore_index=True)
                salvar_estoque(df_estoque)
                st.success(f"Item '{n_nome.upper()}' registrado com sucesso com o código {proximo_codigo}!")
                st.rerun()
            else:
                st.error("Preencha a Descrição Técnica do equipamento.")

# =====================================================================
# MÓDULO 6: HISTÓRICO DE OPERAÇÕES
# =====================================================================
elif menu == "HISTÓRICO DE OPERAÇÕES":
    st.markdown('<div class="section-header">Auditoria e Histórico de Relatórios de Saída e Retorno</div>', unsafe_allow_html=True)
    if os.path.exists(ARQUIVO_HISTORICO):
        df_h = pd.read_excel(ARQUIVO_HISTORICO)
        if df_h.empty:
            st.info("Não constam registros históricos de movimentação até o momento.")
        else:
            st.dataframe(df_h, hide_index=True, use_container_width=True)
            
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                txt_hist = gerar_texto_historico(df_h)
                st.download_button(
                    "EXPORTAR RELATÓRIOS EM FORMATO TEXTO (.TXT)",
                    data=txt_hist.encode("utf-8"),
                    file_name=f"relatorio_saidas_acs_{datetime.now().strftime('%Y%m%d')}.txt",
                    mime="text/plain",
                    type="primary",
                    use_container_width=True
                )
            with col_exp2:
                with open(ARQUIVO_HISTORICO, "rb") as f:
                    st.download_button(
                        "EXPORTAR RELATÓRIOS EM EXCEL (.XLSX)",
                        f,
                        file_name=f"auditoria_operacoes_acs_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        use_container_width=True
                    )

            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("ZONA ADMINISTRATIVA — APAGAR RELATÓRIOS DO HISTÓRICO (SENHA EXCLUSIVA)"):
                st.markdown("Selecione abaixo os relatórios que deseja excluir permanentemente e informe a **Senha Administrativa de Exclusão**:")
                
                opcoes_hist = [
                    f"#{i + 1} | {row['Data/Hora']} | {row['Operação']} - {row['Evento / Projeto']} ({row['Responsável Técnico']})"
                    for i, row in df_h.iterrows()
                ]
                indices_para_apagar = st.multiselect(
                    "Selecionar relatórios específicos para apagar (ou deixe em branco para usar o botão de limpar tudo):",
                    options=list(range(len(opcoes_hist))),
                    format_func=lambda i: opcoes_hist[i]
                )
                
                senha_admin_hist = st.text_input("Senha Administrativa de Exclusão:", type="password", key="pwd_del_hist")
                
                col_del1, col_del2 = st.columns(2)
                with col_del1:
                    if st.button("APAGAR RELATÓRIOS SELECIONADOS", use_container_width=True):
                        if senha_admin_hist != SENHA_EXCLUSAO:
                            st.error("Senha administrativa incorreta. Operação bloqueada.")
                        elif not indices_para_apagar:
                            st.warning("Selecione pelo menos um relatório na caixa acima para apagar.")
                        else:
                            df_h_novo = df_h.drop(index=indices_para_apagar).reset_index(drop=True)
                            df_h_novo.to_excel(ARQUIVO_HISTORICO, index=False)
                            st.success(f"{len(indices_para_apagar)} relatório(s) apagado(s) com sucesso!")
                            st.rerun()
                with col_del2:
                    if st.button("APAGAR TODO O HISTÓRICO DE RELATÓRIOS", use_container_width=True):
                        if senha_admin_hist != SENHA_EXCLUSAO:
                            st.error("Senha administrativa incorreta. Operação bloqueada.")
                        else:
                            os.remove(ARQUIVO_HISTORICO)
                            st.success("Todo o histórico de relatórios foi apagado permanentemente!")
                            st.rerun()
    else:
        st.info("Não constam registros históricos de movimentação até o momento.")

# =====================================================================
# RODAPÉ / ASSINATURA DO SISTEMA
# =====================================================================
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("""
    <div class="no-print" style="text-align: center; color: #64748b; font-size: 0.85rem; border-top: 1px solid #334155; padding-top: 15px; margin-top: 30px;">
        Desenvolvido por: <strong>Alexandre Sessa</strong>
    </div>
""", unsafe_allow_html=True)
