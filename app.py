import streamlit as st
import pandas as pd
import os
import json
import io
import requests
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
ARQUIVO_VISITAS = "visitas_tecnicas_acs.xlsx"

ENDERECO_GALPAO = "Rua São Severino, 16, Taquara, Rio de Janeiro, RJ"

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

def buscar_sugestoes_enderecos(termo):
    if not termo or len(termo) < 3:
        return []
    try:
        headers = {'User-Agent': 'ACS_Logistica_App/1.0'}
        url = f"https://nominatim.openstreetmap.org/search?q={requests.utils.quote(termo + ', Rio de Janeiro, RJ')}&format=json&limit=5"
        res = requests.get(url, headers=headers, timeout=4).json()
        return [item['display_name'] for item in res]
    except Exception:
        return []

def formatar_tempo(minutos):
    if not minutos:
        return "A calcular"
    if minutos < 60:
        return f"{minutos} min"
    horas = minutos // 60
    mins_restantes = minutos % 60
    if mins_restantes == 0:
        return f"{horas}h"
    return f"{horas}h {mins_restantes}min"

def calcular_rota_osm(endereco_destino):
    try:
        headers = {'User-Agent': 'ACS_Logistica_App/1.0'}
        url_galpao = f"https://nominatim.openstreetmap.org/search?q={requests.utils.quote(ENDERECO_GALPAO)}&format=json&limit=1"
        res_g = requests.get(url_galpao, headers=headers, timeout=4).json()
        if not res_g:
            return None, None
        lon_g, lat_g = float(res_g[0]['lon']), float(res_g[0]['lat'])
        
        url_dest = f"https://nominatim.openstreetmap.org/search?q={requests.utils.quote(endereco_destino)}&format=json&limit=1"
        res_d = requests.get(url_dest, headers=headers, timeout=4).json()
        if not res_d:
            return None, None
        lon_d, lat_d = float(res_d[0]['lon']), float(res_d[0]['lat'])
        
        url_osrm = f"http://router.project-osrm.org/route/v1/driving/{lon_g},{lat_g};{lon_d},{lat_d}?overview=false"
        res_o = requests.get(url_osrm, timeout=4).json()
        
        if res_o.get("code") == "Ok":
            rota = res_o["routes"][0]
            distancia_km = round(rota["distance"] / 1000.0, 1)
            duracao_min = round(rota["duration"] / 60.0)
            return distancia_km, duracao_min
    except Exception:
        pass
    return None, None

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

        # Iluminação
        ("108", "MESA DMX MINI AVOLITES 1024", 2, "Luz"),
        ("109", "MESA DE LUZ AVOLITES", 2, "Luz"),
        ("110", "MESA DE LUZ GRAND MA", 1, "Luz"),
        ("111", "DMX CONROLLER", 1, "Luz"),
        ("112", "STROBO LED RGBW FITA", 16, "Luz"),
        ("113", "MÁQUINA PEQUENA DE FUMAÇA FOG MACHINE", 1, "Luz"),
        ("114", "MAIN POWER GRANDE", 2, "Luz"),
        ("115", "MAIN POWER PEQUENO", 2, "Luz"),
        ("116", "SPLIT DE SINAL", 5, "Luz"),
        ("117", "SPLIT WIRELESS", 1, "Luz"),
        ("118", "PAR LED PEQUENO RGB ANTIGO", 40, "Luz"),
        ("119", "PAR LED RGBW NOVO", 60, "Luz"),
        ("120", "PAR LED GRANDE RGB ANTIGO", 29, "Luz"),
        ("121", "MINI BRUT DE LED", 4, "Luz"),
        ("122", "MINI CANHÃO DE LUZ", 1, "Luz"),
        ("123", "PAR LED BLINDADO", 31, "Luz"),
        ("124", "PAR 38", 60, "Luz"),
        ("125", "REFLETOR AR111", 70, "Luz"),
        ("126", "MOVING BEAM 230 BORDA DE LED", 6, "Luz"),
        ("127", "MOVING BEAM 9R", 13, "Luz"),
        ("128", "MOVING BEAM 7R", 5, "Luz"),
        ("129", "COB 200W LEVE", 2, "Luz"),
        ("130", "COB 200W BLINDADO", 8, "Luz"),
        ("131", "VARAS DE AR 1M", 8, "Luz"),
        ("132", "VARAS DE AR 2M", 8, "Luz"),
        ("133", "PAR LED SPOT PRETO", 10, "Luz"),
        ("134", "PAR LED SPOT BRANCO", 10, "Luz"),
        ("135", "REFLETOR DE 300W", 3, "Luz"),
        ("136", "REFLETOR VERDE", 10, "Luz"),
        ("137", "REFLETOR BRANCO QUENTE DE 100W", 2, "Luz"),
        ("138", "REFLETOR BRANCO QUENTE 200W", 1, "Luz"),
        ("139", "RIBALTA RGBW", 24, "Luz"),
        ("140", "CANHÃO SKY WALKER", 2, "Luz"),
        ("141", "RIBALTA ANTIGA", 10, "Luz"),
        ("142", "GARRA DE PAR LED", 76, "Luz"),
        ("143", "GARRA DE MOVING", 65, "Luz"),
        ("144", "MULTIVIAS DE XLR", 1, "Luz"),
        ("145", "MÁQUINA DE BOLHAS", 2, "Luz"),

        # Estruturas (Nova Seção)
        ("146", "BOX TRUSS 3.0M", 10, "Estruturas"),
        ("147", "BOX TRUSS 2.0M", 16, "Estruturas"),
        ("148", "BOX TRUSS 1.0M", 9, "Estruturas"),
        ("149", "BOX TRUSS 0.50M", 6, "Estruturas"),
        ("150", "CUBO", 15, "Estruturas"),
        ("151", "FACEADA", 6, "Estruturas"),
        ("152", "MEIA LUA", 4, "Estruturas"),
        ("153", "BASE", 12, "Estruturas"),
        ("154", "BASE ALTA", 2, "Estruturas"),
        ("155", "SLEEVE", 4, "Estruturas"),
        ("156", "PAU DE CARGA", 4, "Estruturas"),
        ("157", "CUMIERA", 6, "Estruturas"),
        ("158", "PARAFUSOS NOVOS", 150, "Estruturas"),
        ("159", "PARAFUSO DIVERSOS", 40, "Estruturas"),
        ("160", "PARAFUSO SEM PORCA", 30, "Estruturas"),
        ("161", "GREPO", 8, "Estruturas"),
        ("162", "PRATICÁVEL 2.0M X 1.00M", 49, "Estruturas"),
        ("163", "PRATICÁVEL 0.50M X 1.0M", 2, "Estruturas"),
        ("164", "PRATICÁVEL 1.0M X 1.0M", 2, "Estruturas"),

        # Cabos, Energia, Informática e Acessórios
        ("165", "PEDESTAL", 41, "Acessórios"),
        ("166", "RÉGUA DE ENERGIA POWERCON", 14, "Acessórios"),
        ("167", "RÉGUA DE ENERGIA POWERCON SAS", 6, "Acessórios"),
        ("168", "RÉGUA DE ENERGIA POWERCON / POWERCON", 4, "Acessórios"),
        ("169", "PROCESSADOR BEHRINGER", 1, "Áudio"),
        ("170", "PROCESSADOR DBX", 3, "Áudio"),
        ("171", "AMPLIFICADOR CROWN XTI 1000", 1, "Áudio"),
        ("172", "AMPLIFICADOR CROWN XTI 4000", 2, "Áudio"),
        ("173", "AMPLIFICADOR CROWN XTI 6000", 1, "Áudio"),
        ("174", "HOT SOUND HS PRO 5", 1, "Áudio"),
        ("175", "HOT SOUND DIGILITE 6.4", 1, "Áudio"),
        ("176", "HOT SOUND DIGILITE 8.0", 1, "Áudio"),
        ("177", "YAMAHA P3200", 1, "Áudio"),
        ("178", "AMPLIFICADOR CROWN / CARVER (RUIM)", 3, "Áudio"),
        ("179", "STUDIO R Z2400", 1, "Áudio"),
        ("180", "STUDIO R XD", 1, "Áudio"),
        ("181", "STUDIO R X5", 1, "Áudio"),
        ("182", "STUDIO R X12", 1, "Áudio"),
        ("183", "APD 4004 THUNDELIGHT", 2, "Áudio"),
        ("184", "POWERTEC LS", 4, "Áudio"),
        ("185", "RACK DE AC AZUL (RUIM)", 1, "Acessórios"),
        ("186", "RACK DE AC NOVO", 1, "Acessórios"),
        ("187", "CENTRAL DE ENERGIA", 4, "Acessórios"),
        ("188", "ISOLADOR DE ÁUDIO GX200", 2, "Áudio"),
        ("189", "DIRECT BOX WHIRLWIND IMP2", 5, "Áudio"),
        ("190", "DIRECT BOX WALDMAN PASSIVA DI-1PS", 5, "Áudio"),
        ("191", "DIRECT BOX WIRECONEX WDI600", 4, "Áudio"),
        ("192", "DIRECT BOX WIRECONEX WDI500", 20, "Áudio"),
        ("193", "DIRECT BOX WALDMAN ATIVA DI-1A", 7, "Áudio"),
        ("194", "DIRECT BOX SAMSON ATIVA S. DIRECT", 3, "Áudio"),
        ("195", "DIRECT BOX SAMSON ATIVA S. DIRECT PLUS", 1, "Áudio"),
        ("196", "MOUSE PAD", 1, "Informática"),
        ("197", "CPU WYSE DELL", 1, "Informática"),
        ("198", "MONITOR DELL P2418HT", 1, "Informática"),
        ("199", "FONTE DE MONITOR HA65NS5-00", 1, "Informática"),
        ("200", "CABO RJ45", 1, "Informática"),
        ("201", "CABO DISPLAYPORT", 1, "Informática"),
        ("202", "CABO USB-SS DELL", 1, "Informática"),
        ("203", "CABO DE FORÇA PARA FONTE", 1, "Informática"),
        ("204", "JUMPER SPEAKON (4 VIAS - ALTAS)", 18, "Áudio"),
        ("205", "JUMPER SPEAKON (2 VIAS - SUB)", 6, "Áudio"),
        ("206", "CABO SPEAKON (4 VIAS - ALTAS)", 8, "Áudio"),
        ("207", "CABO SPEAKON (2 VIAS - SUB)", 8, "Áudio"),
        ("208", "CABO POWERCON / POWERCON", 20, "Acessórios"),
        ("209", "CABO XLR DE 30 METROS", 8, "Áudio"),
        ("210", "CABO XLR DE 10 METROS", 60, "Áudio"),
        ("211", "CABO XLR DE 5 METROS", 90, "Áudio"),
        ("212", "CABO XLR DA BARRA GARDEN - GRANADA DE 27 METROS", 1, "Luz"),
        ("213", "CABO XLR DA BARRA GARDEN - GRANADA DE 17 METROS", 1, "Luz"),
        ("214", "CABO XLR DA BARRA GARDEN - STECK / STECK", 5, "Luz"),
        ("215", "CABO XLR DA BARRA GARDEN - SINGELOS", 4, "Luz"),
        ("216", "CABO XLR DE SINAL", 200, "Luz"),
        ("217", "CABO PP", 200, "Acessórios"),
        ("218", "CABO PARALELO Y", 14, "Acessórios"),
        ("219", "CENTOPÉIA GRANDE DE 2,5 MM", 11, "Acessórios"),
        ("220", "CENTOPÉIA DE 1,5 MM TAMANHOS VARIADOS", 10, "Acessórios"),

        # Equipamentos com Defeito
        ("221", "CANHÃO SKY WALKER COM DEFEITO", 2, "Equipamentos com Defeito"),
        ("222", "STROBO LED RGBW FITA COM DEFEITO", 2, "Equipamentos com Defeito"),
        ("223", "MÁQUINA PEQUENA DE FUMAÇA FOG MACHINE COM DEFEITO", 1, "Equipamentos com Defeito"),
        ("224", "MÁQUINA DE FUMAÇA HAZE COM DEFEITO", 2, "Equipamentos com Defeito"),
        ("225", "MOVING BEAM 230 BORDA DE LED COM DEFEITO", 2, "Equipamentos com Defeito"),
        ("226", "MOVING BEAM 9R COM DEFEITO", 3, "Equipamentos com Defeito"),
        ("227", "MOVING BEAM 7R COM DEFEITO", 8, "Equipamentos com Defeito"),
        ("228", "LASER COM DEFEITO", 2, "Equipamentos com Defeito"),
        ("229", "STROBO ANTIGO COM DEFEITO", 2, "Equipamentos com Defeito")
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
                    "distancia": str(row.get("Distância", "-")),
                    "tempo_estimado": str(row.get("Tempo Estimado", "-")),
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
            "Distância": ev.get("distancia", "-"),
            "Tempo Estimado": ev.get("tempo_estimado", "-"),
            "Hora Galpão": ev.get("hora_galpao", "-"),
            "Hora Evento": ev.get("hora_evento", "-"),
            "Data Saída": ev["data_saida"],
            "Itens JSON": json.dumps(ev["itens"], ensure_ascii=False)
        })
    pd.DataFrame(dados_tabela).to_excel(ARQUIVO_EVENTOS, index=False)

def carregar_visitas():
    if os.path.exists(ARQUIVO_VISITAS):
        try:
            return pd.read_excel(ARQUIVO_VISITAS).to_dict("records")
        except Exception:
            return []
    return []

def salvar_visitas(visitas):
    pd.DataFrame(visitas).to_excel(ARQUIVO_VISITAS, index=False)

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
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    elementos = []
    styles = getSampleStyleSheet()
    
    estilo_titulo = ParagraphStyle('TituloDoc', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, leading=15, textColor=colors.HexColor('#0f172a'), alignment=1)
    estilo_cab_empresa = ParagraphStyle('CabEmpresa', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=14, textColor=colors.white, alignment=0)
    estilo_info = ParagraphStyle('InfoCab', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=12, textColor=colors.HexColor('#1e293b'))
    estilo_celula = ParagraphStyle('CelulaTabela', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=11, textColor=colors.HexColor('#0f172a'))
    estilo_celula_bold = ParagraphStyle('CelulaTabelaBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=11, textColor=colors.HexColor('#0f172a'), alignment=1)
    estilo_th = ParagraphStyle('TabelaHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=11, textColor=colors.white, alignment=1)

    logo_path = carregar_logo()
    texto_empresa = Paragraph("<b>A.C.S. PRODUÇÕES & DIVULGAÇÕES LTDA</b><br/><font size=8 color='#94a3b8'>A.C.S. SOM, LUZ E IMAGEM — GESTÃO TÉCNICA E LOGÍSTICA DE EVENTOS</font>", estilo_cab_empresa)
    
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
        [Paragraph(f"<b>Endereço:</b> {ev.get('endereco', '-')}", estilo_info), Paragraph(f"<b>Distância / Tempo:</b> {ev.get('distancia', '-')} | Estimado: {ev.get('tempo_estimado', '-')}", estilo_info)],
        [Paragraph(f"<b>Horários:</b> Galpão: {ev.get('hora_galpao', '-')} | Evento: {ev.get('hora_evento', '-')}", estilo_info), Paragraph("", estilo_info)]
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

KITS_PRONTOS = {
    "CASE A IN EAR RELACART": [
        {"termo": "RELACART PM-320", "qtd": 3},
        {"termo": "BODYPACK RELACART PM-320R", "qtd": 6},
        {"termo": "COMBINADOR RELACART TAC4", "qtd": 1},
        {"termo": "RÉGUA DE ENERGIA POWERCON SAS", "qtd": 1},
        {"termo": "ANTENA IN-EAR RELACART R-32U", "qtd": 1},
        {"termo": "CABO POWERCON / POWERCON", "qtd": 1},
        {"termo": "ANTENA RELACART", "qtd": 1}
    ],
    "CASE B IN EAR RELACART": [
        {"termo": "RELACART PM-320", "qtd": 2},
        {"termo": "BODYPACK RELACART PM-320R", "qtd": 4},
        {"termo": "COMBINADOR RELACART TAC4", "qtd": 1},
        {"termo": "RÉGUA DE ENERGIA POWERCON SAS", "qtd": 1},
        {"termo": "ANTENA IN-EAR RELACART R-32U", "qtd": 1},
        {"termo": "CABO POWERCON / POWERCON", "qtd": 1},
        {"termo": "ANTENA RELACART", "qtd": 1}
    ],
    "CASE A MICROFONES SHURE UR4D+": [
        {"termo": "BASE SHURE UR4D+", "qtd": 3},
        {"termo": "BASE SHURE UR4D", "qtd": 1},
        {"termo": "MICROFONE SEM FIO BASTÃO SHURE UR4", "qtd": 4},
        {"termo": "TRANSMISSOR SHURE UR1M", "qtd": 4},
        {"termo": "MICROFONE SHURE SM58", "qtd": 4},
        {"termo": "ANTENA RELACART ATIVA", "qtd": 2},
        {"termo": "ANTENA RELACART", "qtd": 2},
        {"termo": "CABO POWERCON / POWERCON", "qtd": 1},
        {"termo": "SUB-SNAKE DE 8 VIAS", "qtd": 1},
        {"termo": "CABO USB-SS DELL", "qtd": 1},
        {"termo": "CABO RJ45", "qtd": 1},
        {"termo": "RÉGUA DE ENERGIA POWERCON SAS", "qtd": 1},
        {"termo": "JUMPER SPEAKON (4 VIAS - ALTAS)", "qtd": 8}
    ],
    "CASE B MICROFONES SHURE UR4D+": [
        {"termo": "BASE SHURE UR4D+", "qtd": 3},
        {"termo": "BASE SHURE UR4D", "qtd": 1},
        {"termo": "MICROFONE SEM FIO BASTÃO SHURE UR4", "qtd": 4},
        {"termo": "TRANSMISSOR SHURE UR1M", "qtd": 4},
        {"termo": "MICROFONE SHURE SM58", "qtd": 4},
        {"termo": "ANTENA RELACART ATIVA", "qtd": 2},
        {"termo": "ANTENA RELACART", "qtd": 2},
        {"termo": "CABO POWERCON / POWERCON", "qtd": 1},
        {"termo": "SUB-SNAKE DE 8 VIAS", "qtd": 1},
        {"termo": "CABO USB-SS DELL", "qtd": 1},
        {"termo": "CABO RJ45", "qtd": 1},
        {"termo": "RÉGUA DE ENERGIA POWERCON SAS", "qtd": 1},
        {"termo": "JUMPER SPEAKON (4 VIAS - ALTAS)", "qtd": 8}
    ],
    "CASE C MICROFONES SHURE UR4D+": [
        {"termo": "BASE SHURE UR4D+", "qtd": 3},
        {"termo": "BASE SHURE UR4D", "qtd": 1},
        {"termo": "TRANSMISSOR SHURE UR1M", "qtd": 8},
        {"termo": "ANTENA RELACART", "qtd": 2},
        {"termo": "ANTENA RELACART ATIVA", "qtd": 2},
        {"termo": "CABO POWERCON / POWERCON", "qtd": 1},
        {"termo": "SUB-SNAKE DE 8 VIAS", "qtd": 1},
        {"termo": "CABO USB-SS DELL", "qtd": 1},
        {"termo": "CABO RJ45", "qtd": 1},
        {"termo": "RÉGUA DE ENERGIA POWERCON SAS", "qtd": 1},
        {"termo": "JUMPER SPEAKON (4 VIAS - ALTAS)", "qtd": 8}
    ],
    "BAÚ DE CABOS A": [
        {"termo": "RÉGUA DE ENERGIA POWERCON", "qtd": 7},
        {"termo": "RÉGUA DE ENERGIA POWERCON SAS", "qtd": 1},
        {"termo": "DIRECT BOX WIRECONEX WDI500", "qtd": 10},
        {"termo": "DIRECT BOX WALDMAN ATIVA DI-1A", "qtd": 2},
        {"termo": "ISOLADOR DE ÁUDIO GX200", "qtd": 1},
        {"termo": "CABO XLR DE 10 METROS", "qtd": 20},
        {"termo": "CABO XLR DE 5 METROS", "qtd": 30},
        {"termo": "SUB-SNAKE DE 8 VIAS", "qtd": 2},
        {"termo": "SUB-SNAKE DE 12 VIAS", "qtd": 1},
        {"termo": "MEDUSA DE BATERIA DE 12 VIAS", "qtd": 2},
        {"termo": "MICROFONE SAMSON C02", "qtd": 2},
        {"termo": "MICROFONE SAMSON C02H", "qtd": 2},
        {"termo": "MICROFONE E604 SENNHEISER", "qtd": 4},
        {"termo": "MICROFONE SHURE BETA 91A", "qtd": 1}
    ],
    "BAÚ DE CABOS B": [
        {"termo": "RÉGUA DE ENERGIA POWERCON", "qtd": 7},
        {"termo": "RÉGUA DE ENERGIA POWERCON SAS", "qtd": 1},
        {"termo": "DIRECT BOX WIRECONEX WDI500", "qtd": 10},
        {"termo": "DIRECT BOX WALDMAN ATIVA DI-1A", "qtd": 2},
        {"termo": "ISOLADOR DE ÁUDIO GX200", "qtd": 1},
        {"termo": "CABO XLR DE 10 METROS", "qtd": 20},
        {"termo": "CABO XLR DE 5 METROS", "qtd": 30},
        {"termo": "SUB-SNAKE DE 8 VIAS", "qtd": 2},
        {"termo": "SUB-SNAKE DE 12 VIAS", "qtd": 1},
        {"termo": "MEDUSA DE BATERIA DE 12 VIAS", "qtd": 2},
        {"termo": "MICROFONE AKG D112", "qtd": 1},
        {"termo": "MICROFONE SAMSON C02", "qtd": 2},
        {"termo": "MICROFONE SAMSON C02H", "qtd": 2},
        {"termo": "MICROFONE E604 SENNHEISER", "qtd": 4}
    ]
}

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
# CARREGAMENTO E MENU SUPERIOR ADAPTATIVO
# =====================================================================
df_estoque = carregar_estoque()
eventos_ativos = carregar_eventos()
visitas_tecnicas = carregar_visitas()

col_h1, col_h2 = st.columns([1, 5])
with col_h1:
    if logo_path:
        st.image(logo_path, width=120)
with col_h2:
    st.markdown("### 🎛️ PAINEL OPERACIONAL A.C.S.")

menu = st.selectbox(
    "📍 SELECIONE O MÓDULO DO SISTEMA:",
    [
        "VISÃO GERAL DO ACERVO",
        "EXPEDIÇÃO (SAÍDA)",
        "EMISSÃO DE CHECKLIST",
        "RETORNO E CONFERÊNCIA",
        "CADASTRO DE MATERIAIS",
        "HISTÓRICO DE OPERAÇÕES",
        "VISITA TÉCNICA"
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
# MÓDULO 1: VISÃO GERAL DO ACERVO
# =====================================================================
if menu == "VISÃO GERAL DO ACERVO":
    st.markdown('<div class="section-header">Posição Consolidada do Acervo Técnico</div>', unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        busca = st.text_input("Pesquisar equipamento por descrição ou código:", placeholder="Digite o modelo, marca ou código...")
    with c2:
        f_cat = st.selectbox("Filtrar por Departamento:", ["Todos", "Áudio", "Backline", "Luz", "Estruturas", "Informática", "Acessórios", "Equipamentos com Defeito"])
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
# MÓDULO 2: EXPEDIÇÃO (SAÍDA)
# =====================================================================
elif menu == "EXPEDIÇÃO (SAÍDA)":
    st.markdown('<div class="section-header">Registro de Expedição e Alocação por Evento (Kits e Itens)</div>', unsafe_allow_html=True)
    st.caption(f"📍 Origem da Carga (Galpão A.C.S.): {ENDERECO_GALPAO}")

    col_ev1, col_ev2 = st.columns(2)
    with col_ev1:
        nome_evento = st.text_input("Identificação do Evento / Cliente:", placeholder="Ex: Convenção Corporativa - Barra")
        tecnico_resp = st.text_input("Responsável Técnico pela Retirada:", placeholder="Nome completo do responsável")
        data_saida = st.date_input("Data de Expedição:", datetime.now(), format="DD/MM/YYYY")
    with col_ev2:
        termo_busca_end = st.text_input(
            "Pesquisar Endereço de Destino (Autocompletar Online):",
            placeholder="Digite parte do local ou rua (ex: Riocentro, Copacabana)..."
        )
        
        sugestoes = []
        if termo_busca_end and len(termo_busca_end) >= 3:
            sugestoes = buscar_sugestoes_enderecos(termo_busca_end)

        endereco_evento = termo_busca_end
        if sugestoes:
            endereco_selecionado = st.selectbox("📍 Selecione o endereço correspondente:", options=["--- Digite / Selecione ---"] + sugestoes)
            if endereco_selecionado != "--- Digite / Selecione ---":
                endereco_evento = endereco_selecionado

        col_h1, col_h2 = st.columns(2)
        with col_h1:
            hora_galpao = st.time_input("Horário Chegada Galpão:", time(8, 0))
        with col_h2:
            hora_evento = st.time_input("Horário Chegada Evento:", time(10, 0))

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("📦 **ADICIONAR KITS PRONTOS AO CHECKLIST:**", unsafe_allow_html=True)
    col_kit1, col_kit2 = st.columns([3, 1])
    with col_kit1:
        kit_selecionado = st.selectbox("Selecione um Kit para adicionar:", options=["--- Selecione um Kit ---"] + list(KITS_PRONTOS.keys()))
    with col_kit2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("ADICIONAR KIT", use_container_width=True):
            if kit_selecionado != "--- Selecione um Kit ---":
                itens_kit = KITS_PRONTOS[kit_selecionado]
                for ik in itens_kit:
                    match_eq = df_estoque[df_estoque["Equipamento"].str.contains(ik["termo"], case=False, na=False)]
                    if not match_eq.empty:
                        cod_m = match_eq.iloc[0]["Código"]
                        atual = st.session_state.carrinho_saida.get(cod_m, 0)
                        st.session_state.carrinho_saida[cod_m] = atual + ik["qtd"]
                st.success(f"Kit '{kit_selecionado}' adicionado com sucesso ao carrinho!")
                st.rerun()

    dist_calc, tempo_min_calc = None, None
    if endereco_evento and len(endereco_evento) > 5:
        with st.spinner("Calculando distância e tempo de percurso a partir da Taquara..."):
            dist_calc, tempo_min_calc = calcular_rota_osm(endereco_evento)
        if dist_calc is not None and tempo_min_calc is not None:
            tempo_formatado = formatar_tempo(tempo_min_calc)
            st.success(f"🚗 Rota Calculada: **{dist_calc} km** | Tempo Estimado de Viagem: **~{tempo_formatado}**")
        else:
            st.info("ℹ️ Endereço em processamento. O checklist será gerado normalmente.")

    st.markdown('<div class="section-header">Pesquisa e Seleção Manual de Equipamentos</div>', unsafe_allow_html=True)
    
    col_busca_s, col_cat_s = st.columns([2, 1])
    with col_busca_s:
        busca_saida = st.text_input(
            "Pesquisar equipamento para adicionar à saída:",
            placeholder="Digite qualquer parte do nome ou código (ex: SM58, Line Array, XLR, Moving)..."
        )
    with col_cat_s:
        cat_saida = st.selectbox("Filtrar por Setor:", ["Todos os Setores", "Áudio", "Backline", "Luz", "Estruturas", "Informática", "Acessórios", "Equipamentos com Defeito"])

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
            st.error("Nenhum material selecionado. Pesquise o equipamento ou selecione um kit para expedir.")
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
                
                tempo_str = formatar_tempo(tempo_min_calc) if tempo_min_calc is not None else "A calcular"
                
                novo_evento = {
                    "id": datetime.now().strftime("%Y%m%d%H%M%S"),
                    "evento": nome_evento.upper(),
                    "responsavel": tecnico_resp.upper(),
                    "endereco": endereco_evento.upper() if endereco_evento else "NÃO INFORMADO",
                    "distancia": f"{dist_calc} km" if dist_calc else "A calcular",
                    "tempo_estimado": tempo_str,
                    "hora_galpao": hora_galpao.strftime("%H:%M"),
                    "hora_evento": hora_evento.strftime("%H:%M"),
                    "data_saida": data_saida.strftime("%d/%m/%Y"),
                    "itens": lista_itens_evento
                }
                eventos_ativos.append(novo_evento)
                salvar_eventos(eventos_ativos)
                
                registrar_historico("EXPEDIÇÃO", nome_evento.upper(), tecnico_resp.upper(), "; ".join(resumo_log), f"Destino: {endereco_evento} | Tempo est.: {tempo_str}")
                st.session_state.carrinho_saida = {}
                st.success(f"Operação registrada com sucesso para o evento '{nome_evento.upper()}'. Acesse 'EMISSÃO DE CHECKLIST' para descarregar o PDF profissional.")
                st.rerun()

# =====================================================================
# MÓDULO 3: EMISSÃO DE CHECKLIST
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
# MÓDULO 5: CADASTRO DE MATERIAIS
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
            n_cat = st.selectbox("Departamento Técnico:", ["Áudio", "Backline", "Luz", "Estruturas", "Informática", "Acessórios", "Equipamentos com Defeito"])
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
# MÓDULO 7: VISITA TÉCNICA (QUESTIONÁRIO / GABARITO)
# =====================================================================
elif menu == "VISITA TÉCNICA":
    st.markdown('<div class="section-header">Gabarito de Visita Técnica e Levantamento de Locação</div>', unsafe_allow_html=True)
    
    with st.form("form_visita_tecnica"):
        col_vt1, col_vt2 = st.columns(2)
        with col_vt1:
            vt_resp = st.text_input("Responsável pela Visita Técnica:")
            vt_local = st.text_input("Local do Evento:")
            vt_dia_horario = st.text_input("Dia do Evento, Horário e Passagem:", placeholder="Ex: Sábado, 20h, Soundcheck às 17h")
            vt_eletrica = st.text_input("Elétrica, Responsável e Distância:", placeholder="Ex: Trifásico, Seu João, próximo ao palco")
            vt_pre_montagem = st.text_input("Pré Montagem:", placeholder="Detalhes sobre a pré-montagem")
            vt_pe_direito = st.text_input("Pé Direito:", placeholder="Ex: 6 metros")
            vt_painel_led = st.text_input("Painel de LED:", placeholder="Especificações ou dimensões")
            vt_iluminacao = st.text_input("Iluminação:", placeholder="Pontos de luz, varas, etc.")
        with col_vt2:
            vt_sonorizacao = st.text_input("Sonorização:", placeholder="Acústica, pontos de PA, etc.")
            vt_riders = st.text_input("Riders e Contatos:", placeholder="Nome e telefone dos técnicos contratantes")
            vt_palco = st.text_input("Palco (Tamanho, Altura e Cor):", placeholder="Ex: 10x8m, altura 1.2m, preto")
            vt_doca = st.text_input("Doca (Se entra o caminhão):", placeholder="Sim / Não / Detalhes de acesso")
            vt_dificuldades = st.text_input("Escada, Corredor ou Outras Dificuldades:", placeholder="Ex: 2 lances de escada estreita")
            vt_quadro_energia = st.text_input("Distância do Quadro de Energia ou Gerador:", placeholder="Ex: 35 metros")
            vt_obs = st.text_area("Observações Gerais:")

        if st.form_submit_button("REGISTRAR E SALVAR VISITA TÉCNICA"):
            if not vt_local or not vt_resp:
                st.error("Preencha pelo menos os campos de Local e Responsável pela Visita.")
            else:
                nova_visita = {
                    "Data Registro": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "Responsável": vt_resp.upper(),
                    "Local": vt_local.upper(),
                    "Horário/Passagem": vt_dia_horario,
                    "Elétrica": vt_eletrica,
                    "Pré Montagem": vt_pre_montagem,
                    "Pé Direito": vt_pe_direito,
                    "Painel LED": vt_painel_led,
                    "Iluminação": vt_iluminacao,
                    "Sonorização": vt_sonorizacao,
                    "Riders/Contatos": vt_riders,
                    "Palco": vt_palco,
                    "Doca/Caminhão": vt_doca,
                    "Dificuldades": vt_dificuldades,
                    "Distância Energia": vt_quadro_energia,
                    "Observações": vt_obs
                }
                visitas_tecnicas.append(nova_visita)
                salvar_visitas(visitas_tecnicas)
                st.success(f"Visita técnica para '{vt_local.upper()}' salva com sucesso!")
                st.rerun()

    if visitas_tecnicas:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-header">Histórico de Visitas Técnicas Registadas</div>', unsafe_allow_html=True)
        df_v = pd.DataFrame(visitas_tecnicas)
        st.dataframe(df_v, hide_index=True, use_container_width=True)

        with open(ARQUIVO_VISITAS, "rb") as f:
            st.download_button(
                "EXPORTAR VISITAS TÉCNICAS EM EXCEL (.XLSX)",
                f,
                file_name=f"visitas_tecnicas_acs_{datetime.now().strftime('%Y%m%d')}.xlsx",
                use_container_width=True
            )

# =====================================================================
# RODAPÉ / ASSINATURA DO SISTEMA
# =====================================================================
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("""
    <div class="no-print" style="text-align: center; color: #64748b; font-size: 0.85rem; border-top: 1px solid #334155; padding-top: 15px; margin-top: 30px;">
        Desenvolvido por: <strong>Alexandre Sessa</strong>
    </div>
""", unsafe_allow_html=True)
