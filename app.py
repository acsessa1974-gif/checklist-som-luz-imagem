import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime

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
SENHA_EXCLUSAO = "acsadmin"   # Senha exclusiva para apagar relatórios de saída

# =====================================================================
# ESTILIZAÇÃO CORPORATIVA (CSS)
# =====================================================================
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}

    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 96%;
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
        padding: 1rem 1.2rem;
        text-align: left;
    }
    .kpi-label {
        font-size: 0.75rem;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        font-weight: 600;
    }
    .kpi-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #f8fafc;
        margin-top: 4px;
    }
    .kpi-accent {
        color: #2dd4bf;
    }

    .section-header {
        font-size: 1.1rem;
        font-weight: 600;
        color: #e2e8f0;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        border-bottom: 1px solid #334155;
        padding-bottom: 0.5rem;
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
ARQUIVO_EVENTOS = "eventos_ativos_acs.json"
ARQUIVO_HISTORICO = "historico_eventos_acs.xlsx"

def carregar_logo():
    for nome in ["Logo.png", "logo.png", "LOGO.png", "logo.jpg", "Logo.jpg"]:
        if os.path.exists(nome):
            return nome
    return None

def carregar_estoque():
    if os.path.exists(ARQUIVO_ESTOQUE):
        df = pd.read_excel(ARQUIVO_ESTOQUE)
        if "Estoque Total" in df.columns:
            df["Código"] = df["Código"].astype(str)
            return df

    dados_iniciais = [
        ("001", "CAIXA DE PA LINE ARRAY LS ÁUDIO SLIMPEC 4280", 24, "Áudio"),
        ("002", "CAIXA DE PA LINE ARRAY ATTACK VRD108", 16, "Áudio"),
        ("003", "SUBGRAVE DUPLO LS ÁUDIO SLIMPEC 218", 8, "Áudio"),
        ("004", "SUBGRAVE DUPLO GSUB", 8, "Áudio"),
        ("005", "SUBGRAVE DUPLO STANER", 6, "Áudio"),
        ("006", "SUBGRAVE SIMPLES STANER", 2, "Áudio"),
        ("007", "SUBGRAVE SIMPLES ACS", 2, "Áudio"),
        ("008", "CAIXA ATIVA EV ZLX 12", 6, "Áudio"),
        ("009", "CAIXA ATIVA MJ ÁUDIO 1000W", 12, "Áudio"),
        ("010", "CAIXA ATIVA MACKIE TH15", 12, "Áudio"),
        ("011", "CAIXA ATIVA YAMAHA DXR 10", 2, "Áudio"),
        ("012", "SUB ATIVO YAMAHA DXS 12", 2, "Áudio"),
        ("013", "SUB ATIVO ANTERA LF 1000", 2, "Áudio"),
        ("014", "MESA DE SOM 32 CANAIS YAMAHA LS9", 2, "Áudio"),
        ("015", "MESA DE SOM 32 CANAIS YAMAHA TF5", 1, "Áudio"),
        ("016", "MESA DE SOM 32 CANAIS MIDAS M32", 1, "Áudio"),
        ("017", "MESA DE SOM 48 CANAIS SOUNDCRAFT VI1", 1, "Áudio"),
        ("018", "MESA DE SOM 16 CANAIS SOUNDCRAFT UI16", 1, "Áudio"),
        ("019", "MESA DE SOM 16 CANAIS BEHRINGER XR18", 1, "Áudio"),
        ("020", "MESA DE SOM 12 CANAIS BEHRINGER XR12", 1, "Áudio"),
        ("021", "MESA DE SOM 12 CANAIS SOUNDVOICE AUREA", 1, "Áudio"),
        ("022", "MESA ANALÓGICA 12 CANAIS SOUNDCRAFT SIGNATURE 12", 1, "Áudio"),
        ("023", "MESA ANALÓGICA 12 CANAIS YAMAHA MG 12XU", 1, "Áudio"),
        ("024", "DIRECT BOX IMP2", 6, "Áudio"),
        ("025", "DIRECT BOX WIRECONEX", 20, "Áudio"),
        ("026", "DIRECT BOX SAMSON", 3, "Áudio"),
        ("027", "DIRECT BOX VALDMAN", 10, "Áudio"),
        ("028", "CABO XLR MICROFONE NEUTRIK", 500, "Áudio"),
        ("029", "AMPLIFICADOR DE GUITARRA FENDER FRONTMAN 212", 1, "Áudio"),
        ("030", "AMPLIFICADOR DE GUITARRA METERORO 212", 1, "Áudio"),
        ("031", "AMPLIFICADOR DE GUITARRA PEAVEY BANDIT", 2, "Áudio"),
        ("032", "AMPLIFICADOR DE GUITARRA BORNE 212", 1, "Áudio"),
        ("033", "AMPLIFICADOR DE BAIXO GK400", 1, "Áudio"),
        ("034", "AMPLIFICADOR DE BAIXO BORNE", 1, "Áudio"),
        ("035", "BATERIA PEARL EXPORT", 1, "Áudio"),
        ("036", "BATERIA PEARL VISION", 1, "Áudio"),
        ("037", "BATERIA GRETSCH BLACK HAWK", 1, "Áudio"),
        ("038", "BATERIA TAMMA", 1, "Áudio"),
        ("039", "BATERIA ROLAND TD12", 1, "Áudio"),
        ("040", "MICROFONE SEM FIO SHURE UR2/UR4D", 8, "Áudio"),
        ("041", "MICROFONE SEM FIO SHURE UR1M/UR4D", 24, "Áudio"),
        ("042", "MICROFONE SEM FIO RELACART", 8, "Áudio"),
        ("043", "MICROFONE SEM FIO SHURE BLX", 6, "Áudio"),
        ("044", "MICROFONE SEM FIO SENNHEISER 5200", 2, "Áudio"),
        ("045", "MICROFONE SHURE SM 58", 20, "Áudio"),
        ("046", "MICROFONE SHURE SM57", 20, "Áudio"),
        ("047", "MICROFONE SENNHEISER E604", 8, "Áudio"),
        ("048", "MICROFONE SAMSON C02", 8, "Áudio"),
        ("049", "KIT DE MICROFONE DE BATERIA AKG", 2, "Áudio"),
        ("050", "KIT DE MICROFONE DE BATERIA SAMSON", 2, "Áudio"),
        ("051", "CAIXA DE MONITOR CLAIR 112", 12, "Áudio"),
        ("052", "FONE DE OUVIDO AKG", 24, "Áudio"),
        ("053", "REFLETOR PAR LED RGBW", 100, "Luz"),
        ("054", "REFLETOR COB", 10, "Luz"),
        ("055", "REFLETOR RIBALTA LED RGBW", 40, "Luz"),
        ("056", "REFLETOR ESTROBO LED 1000W", 18, "Luz"),
        ("057", "MOVING BEAM 7R", 12, "Luz"),
        ("058", "MOVING BEAM 8R", 12, "Luz"),
        ("059", "MOVING BEAM 9R", 12, "Luz"),
        ("060", "MÁQUINA DE FUMAÇA HAZE", 4, "Luz"),
        ("061", "MESA DMX AVOLITES PEARL", 2, "Luz"),
        ("062", "MESA DMX AVOLITES MINI", 2, "Luz"),
        ("063", "MESA DMX GRAND MA", 1, "Luz"),
        ("064", "NOTEBOOK VÁRIOS MODELOS", 100, "Informática"),
        ("065", "TABLET IPAD", 12, "Informática"),
        ("066", "TABLET SAMSUNG", 4, "Informática"),
        ("067", "ROTEADOR SEM FIO DLINK", 8, "Acessórios")
    ]

    lista_estoque = [{
        "Código": codigo,
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
            with open(ARQUIVO_EVENTOS, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def salvar_eventos(eventos):
    with open(ARQUIVO_EVENTOS, "w", encoding="utf-8") as f:
        json.dump(eventos, f, ensure_ascii=False, indent=2)

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

def gerar_texto_romaneio(ev):
    """Gera o conteúdo formatado em texto (.txt) para exportação de saída de equipamentos."""
    linhas = [
        "======================================================================",
        "               A.C.S. PRODUÇÕES & DIVULGAÇÕES LTDA                    ",
        "         RELATÓRIO DE EXPEDIÇÃO E SAÍDA DE EQUIPAMENTOS               ",
        "======================================================================",
        f"DOCUMENTO Nº        : {ev['id']}",
        f"EVENTO / DESTINO    : {ev['evento']}",
        f"RESPONSÁVEL TÉCNICO : {ev['responsavel']}",
        f"DATA DE SAÍDA       : {ev['data_saida']}",
        "----------------------------------------------------------------------",
        f"{'CÓD.':<8} | {'QTD':<8} | {'ESPECIFICAÇÃO DO EQUIPAMENTO'}",
        "----------------------------------------------------------------------"
    ]
    for item in ev["itens"]:
        linhas.append(f"{str(item['Código']):<8} | {str(item['Qtd']) + ' un':<8} | {item['Equipamento']}")
    
    linhas.extend([
        "----------------------------------------------------------------------",
        f"TOTAL DE ITENS LISTADOS: {len(ev['itens'])}",
        "",
        "",
        "_____________________________________________________",
        f"ASSINATURA DO RESPONSÁVEL ({ev['responsavel']})",
        "",
        "",
        "_____________________________________________________",
        "CONFERÊNCIA DE RETORNO AO GALPÃO (DATA: ___/___/_____)",
        "======================================================================"
    ])
    return "\n".join(linhas)

def gerar_texto_historico(df_h):
    """Gera arquivo .txt com todos os relatórios registrados no histórico."""
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
# CARREGAMENTO E BARRA LATERAL
# =====================================================================
df_estoque = carregar_estoque()
eventos_ativos = carregar_eventos()

with st.sidebar:
    if logo_path:
        st.image(logo_path, width=130)
    st.markdown("**A.C.S. SOM, LUZ E IMAGEM**")
    st.caption("Módulo de Gestão Operacional")
    st.markdown("---")
    
    menu = st.radio(
        "MÓDULOS DO SISTEMA",
        [
            "VISÃO GERAL DO ACERVO",
            "EXPEDIÇÃO (SAÍDA)",
            "RETORNO E CONFERÊNCIA",
            "EMISSÃO DE ROMANEIO",
            "CADASTRO DE MATERIAIS",
            "HISTÓRICO DE OPERAÇÕES"
        ]
    )
    
    st.markdown("---")
    if st.button("ENCERRAR SESSÃO", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

# =====================================================================
# CABEÇALHO E PAINEL DE INDICADORES (OCULTOS NA IMPRESSÃO)
# =====================================================================
st.markdown('<div class="no-print">', unsafe_allow_html=True)

st.markdown(f"""
<div class="corporate-header">
    <div>
        <p class="corporate-title">A.C.S. — CONTROLE DE EQUIPAMENTOS E OPERAÇÕES</p>
        <p class="corporate-subtitle">Gestão de Estoque Técnico, Expedição e Romaneio de Eventos</p>
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

st.markdown("<br></div>", unsafe_allow_html=True)

# =====================================================================
# MÓDULO 1: VISÃO GERAL DO ACERVO
# =====================================================================
if menu == "VISÃO GERAL DO ACERVO":
    st.markdown('<div class="section-header">Posição Consolidada do Acervo Técnico</div>', unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        busca = st.text_input("Pesquisar equipamento por descrição ou código:", placeholder="Digite o modelo, marca ou código...")
    with c2:
        f_cat = st.selectbox("Filtrar por Departamento:", ["Todos", "Áudio", "Luz", "Informática", "Acessórios"])
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

# =====================================================================
# MÓDULO 2: EXPEDIÇÃO (SAÍDA DE MATERIAL COM PESQUISA RÁPIDA)
# =====================================================================
elif menu == "EXPEDIÇÃO (SAÍDA)":
    st.markdown('<div class="section-header">Registro de Expedição e Alocação por Evento</div>', unsafe_allow_html=True)

    col_ev1, col_ev2, col_ev3 = st.columns([2, 2, 1])
    with col_ev1:
        nome_evento = st.text_input("Identificação do Evento / Cliente:", placeholder="Ex: Convenção Corporativa - Barra")
    with col_ev2:
        tecnico_resp = st.text_input("Responsável Técnico pela Retirada:", placeholder="Nome completo do responsável")
    with col_ev3:
        data_saida = st.date_input("Data de Expedição:", datetime.now())

    st.markdown('<div class="section-header">Pesquisa e Seleção de Equipamentos</div>', unsafe_allow_html=True)
    
    # Caixa de Pesquisa e Filtro de Setor combinados
    col_busca_s, col_cat_s = st.columns([2, 1])
    with col_busca_s:
        busca_saida = st.text_input(
            "Pesquisar equipamento para adicionar à saída:",
            placeholder="Digite qualquer parte do nome ou código (ex: SM58, Line Array, XLR, Moving)..."
        )
    with col_cat_s:
        cat_saida = st.selectbox("Filtrar por Setor:", ["Todos os Setores", "Áudio", "Luz", "Informática", "Acessórios"])

    df_disp_saida = df_estoque[df_estoque["Disponível"] > 0][["Código", "Categoria", "Equipamento", "Disponível"]].copy()
    
    if busca_saida:
        df_disp_saida = df_disp_saida[
            df_disp_saida["Equipamento"].str.contains(busca_saida, case=False, na=False) |
            df_disp_saida["Código"].str.contains(busca_saida, case=False, na=False)
        ]
    if cat_saida != "Todos os Setores":
        df_disp_saida = df_disp_saida[df_disp_saida["Categoria"] == cat_saida]

    df_disp_saida = df_disp_saida.reset_index(drop=True)
    df_disp_saida["Qtd Expedir"] = df_disp_saida["Código"].map(lambda c: st.session_state.carrinho_saida.get(str(c), 0))

    def atualizar_selecao_saida():
        mudancas = st.session_state.get("editor_expedicao_busca", {}).get("edited_rows", {})
        for idx_str, cols in mudancas.items():
            idx = int(idx_str)
            if idx < len(df_disp_saida) and "Qtd Expedir" in cols:
                cod_item = str(df_disp_saida.iloc[idx]["Código"])
                nova_qtd = int(cols["Qtd Expedir"] or 0)
                if nova_qtd > 0:
                    st.session_state.carrinho_saida[cod_item] = nova_qtd
                else:
                    st.session_state.carrinho_saida.pop(cod_item, None)

    st.data_editor(
        df_disp_saida,
        key="editor_expedicao_busca",
        on_change=atualizar_selecao_saida,
        hide_index=True,
        use_container_width=True,
        disabled=["Código", "Categoria", "Equipamento", "Disponível"],
        column_config={
            "Código": st.column_config.TextColumn("CÓDIGO", width="small"),
            "Categoria": st.column_config.TextColumn("SETOR", width="small"),
            "Equipamento": st.column_config.TextColumn("ESPECIFICAÇÃO DO EQUIPAMENTO", width="large"),
            "Disponível": st.column_config.NumberColumn("SALDO EM BASE", width="small"),
            "Qtd Expedir": st.column_config.NumberColumn("QUANTIDADE A EXPEDIR", min_value=0, step=1, width="medium")
        }
    )

    # Monta resumo consolidado a partir do carrinho_saida
    itens_no_carrinho = []
    for cod_sel, qtd_sel in list(st.session_state.carrinho_saida.items()):
        linha_est = df_estoque[df_estoque["Código"] == cod_sel]
        if not linha_est.empty and qtd_sel > 0:
            itens_no_carrinho.append({
                "Código": cod_sel,
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
    if st.button("CONFIRMAR EXPEDIÇÃO E GERAR ORDEM DE SAÍDA", type="primary", use_container_width=True):
        if not nome_evento or not tecnico_resp:
            st.error("Preenchimento obrigatório: informe a Identificação do Evento e o Responsável Técnico.")
        elif df_escolhidos.empty:
            st.error("Nenhum material selecionado. Pesquise o equipamento e informe a quantidade na coluna 'QUANTIDADE A EXPEDIR'.")
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
                    cod = str(row["Código"])
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
                    "data_saida": data_saida.strftime("%d/%m/%Y"),
                    "itens": lista_itens_evento
                }
                eventos_ativos.append(novo_evento)
                salvar_eventos(eventos_ativos)
                
                registrar_historico("EXPEDIÇÃO", nome_evento.upper(), tecnico_resp.upper(), "; ".join(resumo_log), "Saída técnica autorizada")
                st.session_state.carrinho_saida = {}
                st.success(f"Operação registrada com sucesso para o evento '{nome_evento.upper()}'. Acesse 'EMISSÃO DE ROMANEIO' para imprimir ou exportar em .TXT.")
                st.rerun()

# =====================================================================
# MÓDULO 3: RETORNO E CONFERÊNCIA
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
        df_ret["Validado"] = True
        df_ret["Laudo / Ocorrência"] = "Sem avarias"

        df_ret_editado = st.data_editor(
            df_ret,
            key=f"retorno_{ev_atual['id']}",
            hide_index=True,
            use_container_width=True,
            disabled=["Código", "Equipamento", "Qtd"],
            column_config={
                "Validado": st.column_config.CheckboxColumn("CONFERIDO", default=True, width="small"),
                "Código": st.column_config.TextColumn("CÓDIGO", width="small"),
                "Equipamento": st.column_config.TextColumn("ESPECIFICAÇÃO DO EQUIPAMENTO", width="large"),
                "Qtd": st.column_config.NumberColumn("QTD EXPEDIDA", width="small"),
                "Laudo / Ocorrência": st.column_config.TextColumn("LAUDO TÉCNICO DE RETORNO", width="medium")
            }
        )

        col_r1, col_r2 = st.columns(2)
        with col_r1:
            conferente_volta = st.text_input("Responsável pela Conferência de Recebimento:", value=ev_atual['responsavel'])
        with col_r2:
            obs_geral = st.text_input("Parecer Geral da Operação:", value="Material reintegrado ao estoque sem pendências")

        if st.button("FINALIZAR CONFERÊNCIA E REINTEGRAR ESTOQUE", type="primary", use_container_width=True):
            if not all(df_ret_editado["Validado"]):
                st.warning("Para encerrar a operação, todos os itens devem estar marcados na coluna 'CONFERIDO'.")
            elif not conferente_volta:
                st.error("Informe o responsável pela conferência de recebimento.")
            else:
                resumo_volta = []
                avarias_lista = []
                for _, row in df_ret_editado.iterrows():
                    cod = str(row["Código"])
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
# MÓDULO 4: EMISSÃO DE ROMANEIO (IMPRESSÃO A4 + EXPORTAR TXT + EXCLUSÃO)
# =====================================================================
elif menu == "EMISSÃO DE ROMANEIO":
    st.markdown('<div class="no-print">', unsafe_allow_html=True)
    st.markdown('<div class="section-header">Emissão e Exportação de Documentos Operacionais</div>', unsafe_allow_html=True)
    
    tipo_impressao = st.radio(
        "Selecione o documento desejado:",
        ["ORDEM DE EXPEDIÇÃO / ROMANEIO DE EVENTO", "LISTA GERAL DE CONFERÊNCIA FÍSICA (INVENTÁRIO)"],
        horizontal=True
    )
    st.caption("Instrução: Utilize o botão abaixo para baixar em formato Texto (.TXT) ou pressione Ctrl + P para imprimir em folha A4.")
    st.markdown("</div>", unsafe_allow_html=True)

    if tipo_impressao == "ORDEM DE EXPEDIÇÃO / ROMANEIO DE EVENTO":
        if not eventos_ativos:
            st.warning("Não existem relatórios de saída ativos no momento.")
        else:
            st.markdown('<div class="no-print">', unsafe_allow_html=True)
            opcoes_ev = [f"{ev['evento']} (Data: {ev['data_saida']} | Resp: {ev['responsavel']})" for ev in eventos_ativos]
            idx_imp = st.selectbox("Selecione o Relatório de Saída:", range(len(opcoes_ev)), format_func=lambda x: opcoes_ev[x])
            ev_imp = eventos_ativos[idx_imp]
            
            # Botão de Exportação em Formato Texto (.TXT)
            conteudo_txt = gerar_texto_romaneio(ev_imp)
            nome_arq_txt = f"Saida_{ev_imp['evento'].replace(' ', '_')}_{ev_imp['id']}.txt"
            st.download_button(
                label="EXPORTAR RELATÓRIO DE SAÍDA EM FORMATO TEXTO (.TXT)",
                data=conteudo_txt.encode("utf-8"),
                file_name=nome_arq_txt,
                mime="text/plain",
                type="primary"
            )
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("---")
            c_logo, c_titulo = st.columns([1, 4])
            with c_logo:
                if logo_path:
                    st.image(logo_path, width=120)
            with c_titulo:
                st.markdown("### **A.C.S. PRODUÇÕES & DIVULGAÇÕES LTDA**")
                st.markdown("**ORDEM DE EXPEDIÇÃO E CONFERÊNCIA TÉCNICA DE EQUIPAMENTOS**")
                st.markdown(f"**DOCUMENTO Nº:** `{ev_imp['id']}` &nbsp;|&nbsp; **DATA DE SAÍDA:** `{ev_imp['data_saida']}`")
            
            st.markdown(f"""
            **EVENTO / DESTINO:** {ev_imp['evento']}  
            **RESPONSÁVEL TÉCNICO:** {ev_imp['responsavel']}
            """)
            st.markdown("---")
            
            tabela_romaneio = [{
                "CÓD.": item["Código"],
                "ESPECIFICAÇÃO DO EQUIPAMENTO": item["Equipamento"],
                "QTD": f"{item['Qtd']} un",
                "SAÍDA": "[    ]",
                "RETORNO": "[    ]",
                "OBSERVAÇÕES TÉCNICAS": "________________________"
            } for item in ev_imp["itens"]]
            st.table(pd.DataFrame(tabela_romaneio))

            st.markdown("<br><br><br>", unsafe_allow_html=True)
            ca1, ca2 = st.columns(2)
            with ca1:
                st.markdown("_____________________________________________________")
                st.markdown(f"**EXPEDIÇÃO / RESPONSÁVEL TÉCNICO ({ev_imp['responsavel']})**")
            with ca2:
                st.markdown("_____________________________________________________")
                st.markdown("**CONFERÊNCIA DE RECEBIMENTO (BASE A.C.S.) — DATA: ___/___/___**")

            # Zona Administrativa para apagar relatório ativo com senha exclusiva
            st.markdown('<div class="no-print"><br><hr>', unsafe_allow_html=True)
            with st.expander("ZONA ADMINISTRATIVA — APAGAR ESTE RELATÓRIO DE SAÍDA (SENHA EXCLUSIVA)"):
                st.warning("Atenção: Ao apagar este relatório de saída ativo, as quantidades alocadas retornarão automaticamente ao saldo disponível no estoque.")
                senha_del_romaneio = st.text_input("Digite a Senha Administrativa de Exclusão:", type="password", key="pwd_del_romaneio")
                if st.button("APAGAR RELATÓRIO DE SAÍDA SELECIONADO"):
                    if senha_del_romaneio == SENHA_EXCLUSAO:
                        for item in ev_imp["itens"]:
                            cod_i = str(item["Código"])
                            qtd_i = int(item["Qtd"])
                            df_estoque.loc[df_estoque["Código"] == cod_i, "Disponível"] += qtd_i
                            df_estoque.loc[df_estoque["Código"] == cod_i, "Em Uso"] = max(0, int(df_estoque.loc[df_estoque["Código"] == cod_i, "Em Uso"].values[0]) - qtd_i)
                        salvar_estoque(df_estoque)
                        eventos_ativos.pop(idx_imp)
                        salvar_eventos(eventos_ativos)
                        st.success("Relatório de saída apagado e quantidades estornadas ao estoque com sucesso.")
                        st.rerun()
                    else:
                        st.error("Senha administrativa incorreta. Exclusão não autorizada.")
            st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown('<div class="no-print">', unsafe_allow_html=True)
        cat_filtro_imp = st.selectbox("Filtrar Departamento:", ["Todos", "Áudio", "Luz", "Informática", "Acessórios"])
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        c_logo, c_titulo = st.columns([1, 4])
        with c_logo:
            if logo_path:
                st.image(logo_path, width=120)
        with c_titulo:
            st.markdown("### **A.C.S. PRODUÇÕES & DIVULGAÇÕES LTDA**")
            st.markdown("**FORMULÁRIO DE REQUISIÇÃO E CONFERÊNCIA DE ACERVO**")
            st.markdown("**EVENTO:** _______________________________________ &nbsp;|&nbsp; **DATA:** ___/___/2026")
        
        st.markdown("---")
        df_ficha = df_estoque.copy()
        if cat_filtro_imp != "Todos":
            df_ficha = df_ficha[df_ficha["Categoria"] == cat_filtro_imp]

        tabela_ficha = [{
            "CÓD.": row["Código"],
            "SETOR": row["Categoria"],
            "ESPECIFICAÇÃO DO EQUIPAMENTO": row["Equipamento"],
            "SALDO BASE": row["Disponível"],
            "QTD SAÍDA": "________",
            "CONFERÊNCIA": "[    ]"
        } for _, row in df_ficha.iterrows()]
        st.table(pd.DataFrame(tabela_ficha))

# =====================================================================
# MÓDULO 5: CADASTRO DE MATERIAIS
# =====================================================================
elif menu == "CADASTRO DE MATERIAIS":
    st.markdown('<div class="section-header">Inclusão de Novos Equipamentos no Patrimônio</div>', unsafe_allow_html=True)
    with st.form("form_novo_item"):
        c1, c2 = st.columns(2)
        with c1:
            n_cod = st.text_input("Código Patrimonial (Ex: 068):")
            n_nome = st.text_input("Descrição Técnica / Modelo do Equipamento:")
        with c2:
            n_cat = st.selectbox("Departamento Técnico:", ["Áudio", "Luz", "Informática", "Acessórios"])
            n_qtd = st.number_input("Quantidade Total Incorporada:", min_value=1, value=1, step=1)
        
        if st.form_submit_button("REGISTRAR EQUIPAMENTO NO ACERVO"):
            if n_cod and n_nome:
                novo_df = pd.DataFrame([{
                    "Código": n_cod.strip(),
                    "Categoria": n_cat,
                    "Equipamento": n_nome.strip().upper(),
                    "Estoque Total": int(n_qtd),
                    "Disponível": int(n_qtd),
                    "Em Uso": 0
                }])
                df_estoque = pd.concat([df_estoque, novo_df], ignore_index=True)
                salvar_estoque(df_estoque)
                st.success(f"Item '{n_nome.upper()}' registrado com sucesso no banco de dados.")
                st.rerun()
            else:
                st.error("Preencha o Código Patrimonial e a Descrição Técnica do equipamento.")

# =====================================================================
# MÓDULO 6: HISTÓRICO DE OPERAÇÕES (EXPORTAR TXT + APAGAR COM SENHA)
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
