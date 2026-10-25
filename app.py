import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime

# Configuração da Página
st.set_page_config(page_title="Checklist - A.C.S. Controle de Equipamentos", page_icon="🎛️", layout="wide")

# Estilo CSS para deixar a impressão (Ctrl+P) limpa em folha A4, escondendo menus e botões
st.markdown("""
<style>
@media print {
    [data-testid="stSidebar"], header, footer, .stButton, .stDownloadButton, [data-testid="stFormSubmitButton"] {
        display: none !important;
    }
    .main .block-container {
        padding-top: 1rem !important;
        max-width: 100% !important;
    }
}
</style>
""", unsafe_allow_html=True)

# ARQUIVOS DE DADOS (Nova estrutura por quantidade e eventos)
ARQUIVO_ESTOQUE = "estoque_acs_quantidades.xlsx"
ARQUIVO_EVENTOS = "eventos_ativos_acs.json"
ARQUIVO_HISTORICO = "historico_eventos_acs.xlsx"

def carregar_logo():
    for nome in ["logo.png", "Logo.png", "LOGO.png", "logo.JPG", "Logo.JPG"]:
        if os.path.exists(nome):
            return nome
    return None

def carregar_estoque():
    if os.path.exists(ARQUIVO_ESTOQUE):
        df = pd.read_excel(ARQUIVO_ESTOQUE)
        if "Estoque Total" in df.columns:
            return df

    # Inventário completo agrupado por quantidade total
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

    lista_estoque = []
    for codigo, nome, qtd, cat in dados_iniciais:
        lista_estoque.append({
            "Código": codigo,
            "Categoria": cat,
            "Equipamento": nome,
            "Estoque Total": int(qtd),
            "Disponível": int(qtd),
            "Em Uso": 0
        })

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
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    novo_reg = pd.DataFrame([{
        "Data/Hora": data_hora,
        "Tipo": acao,
        "Evento": evento,
        "Responsável": responsavel,
        "Itens e Quantidades": resumo_itens,
        "Observações": observacao
    }])
    if os.path.exists(ARQUIVO_HISTORICO):
        df_hist = pd.read_excel(ARQUIVO_HISTORICO)
        df_hist = pd.concat([df_hist, novo_reg], ignore_index=True)
    else:
        df_hist = novo_reg
    df_hist.to_excel(ARQUIVO_HISTORICO, index=False)

# --- SISTEMA DE LOGIN ---
logo_path = carregar_logo()

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    col_l1, col_l2 = st.columns([1, 4])
    with col_l1:
        if logo_path:
            st.image(logo_path, width=130)
    with col_l2:
        st.title("A.C.S. Som, Luz e Imagem")
    
    st.markdown("### 🔒 Acesso Restrito - Controle de Equipamentos")
    with st.form("form_login"):
        senha_digitada = st.text_input("Senha de Acesso:", type="password")
        botao_login = st.form_submit_button("Entrar no Sistema")
        if botao_login:
            if senha_digitada == "acs123":
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
    st.stop()

# --- CARREGAMENTO DE DADOS ---
df_estoque = carregar_estoque()
eventos_ativos = carregar_eventos()

# --- CABEÇALHO PRINCIPAL ---
col_topo1, col_topo2 = st.columns([1, 6])
with col_topo1:
    if logo_path:
        st.image(logo_path, width=110)
with col_topo2:
    st.title("🎛️ A.C.S. - Controle de Equipamentos")
    st.caption("Sistema Rápido de Romaneio, Checklist por Evento e Controle de Estoque")

# --- MENU LATERAL ---
if st.sidebar.button("🔒 Sair do Sistema"):
    st.session_state.autenticado = False
    st.rerun()

st.sidebar.markdown("---")
menu = st.sidebar.radio("📌 Navegação Principal", [
    "📤 1. Novo Checklist (Saída)",
    "📥 2. Conferir Retorno (Entrada)",
    "🖨️ 3. Imprimir Romaneio / Lista",
    "📊 4. Estoque Geral Disponível",
    "➕ 5. Cadastrar / Ajustar Item",
    "📜 6. Histórico de Eventos"
])

# Indicadores rápidos na barra lateral
st.sidebar.markdown("---")
st.sidebar.markdown("**Resumo Rápido:**")
st.sidebar.write(f"🎪 **Eventos na Rua:** {len(eventos_ativos)}")
st.sidebar.write(f"📦 **Unidades em Uso:** {int(df_estoque['Em Uso'].sum())}")
st.sidebar.write(f"✅ **Unidades Disponíveis:** {int(df_estoque['Disponível'].sum())}")

# =====================================================================
# 1. NOVO CHECKLIST DE SAÍDA (MONTAR ROMANEIO DO EVENTO)
# =====================================================================
if menu == "📤 1. Novo Checklist (Saída)":
    st.subheader("📤 Montar Checklist de Saída para Evento")
    st.markdown("Preencha os dados do evento e digite a **quantidade que vai sair** em cada categoria abaixo:")

    col_ev1, col_ev2, col_ev3 = st.columns([2, 2, 1])
    with col_ev1:
        nome_evento = st.text_input("🎪 Nome do Evento / Cliente:", placeholder="Ex: Show Shopping Carioca")
    with col_ev2:
        tecnico_resp = st.text_input("👷 Técnico / Responsável:", placeholder="Ex: João Silva")
    with col_ev3:
        data_saida = st.date_input("📅 Data de Saída:", datetime.now())

    st.markdown("#### Selecione os Equipamentos e Quantidades:")
    
    # Abas separadas por categoria para facilitar a montagem sem perder dados
    abas = st.tabs(["🔊 Áudio", "💡 Luz", "💻 Informática", "🔌 Acessórios"])
    categorias_nomes = ["Áudio", "Luz", "Informática", "Acessórios"]
    
    tabelas_editadas = []

    for aba, cat in zip(abas, categorias_nomes):
        with aba:
            df_cat = df_estoque[(df_estoque["Categoria"] == cat) & (df_estoque["Disponível"] > 0)][
                ["Código", "Equipamento", "Disponível"]
            ].copy()
            df_cat["Qtd Saída"] = 0
            
            editado = st.data_editor(
                df_cat,
                key=f"editor_saida_{cat}",
                hide_index=True,
                use_container_width=True,
                disabled=["Código", "Equipamento", "Disponível"],
                column_config={
                    "Código": st.column_config.TextColumn("Cód.", width="small"),
                    "Equipamento": st.column_config.TextColumn("Modelo do Equipamento", width="large"),
                    "Disponível": st.column_config.NumberColumn("Em Estoque", width="small"),
                    "Qtd Saída": st.column_config.NumberColumn("👉 Qtd Levar", min_value=0, step=1, width="small")
                }
            )
            tabelas_editadas.append(editado)

    df_selecao_total = pd.concat(tabelas_editadas, ignore_index=True)
    itens_escolhidos = df_selecao_total[df_selecao_total["Qtd Saída"] > 0]

    st.markdown("---")
    if not itens_escolhidos.empty:
        st.markdown("##### 📋 Resumo dos Itens Selecionados para este Evento:")
        st.dataframe(itens_escolhidos[["Código", "Equipamento", "Qtd Saída"]], hide_index=True, use_container_width=True)

    if st.button("✅ Confirmar Saída e Gerar Romaneio do Evento", type="primary", use_container_width=True):
        if not nome_evento or not tecnico_resp:
            st.error("⚠️ Preencha o Nome do Evento e o Técnico Responsável no topo da página.")
        elif itens_escolhidos.empty:
            st.error("⚠️ Informe a quantidade ('Qtd Levar' maior que zero) para pelo menos um equipamento.")
        else:
            erro_qtd = False
            for _, row in itens_escolhidos.iterrows():
                if row["Qtd Saída"] > row["Disponível"]:
                    st.error(f"⚠️ Quantidade indisponível para '{row['Equipamento']}' (Pedido: {row['Qtd Saída']}, Disponível: {row['Disponível']}).")
                    erro_qtd = True
            
            if not erro_qtd:
                lista_itens_evento = []
                resumo_log = []
                for _, row in itens_escolhidos.iterrows():
                    cod = row["Código"]
                    qtd_saindo = int(row["Qtd Saída"])
                    nome_eq = row["Equipamento"]
                    
                    # Atualiza estoque
                    df_estoque.loc[df_estoque["Código"] == cod, "Disponível"] -= qtd_saindo
                    df_estoque.loc[df_estoque["Código"] == cod, "Em Uso"] += qtd_saindo
                    
                    lista_itens_evento.append({
                        "Código": cod,
                        "Equipamento": nome_eq,
                        "Qtd": qtd_saindo
                    })
                    resumo_log.append(f"{qtd_saindo}x {nome_eq}")

                salvar_estoque(df_estoque)
                
                novo_evento = {
                    "id": datetime.now().strftime("%Y%m%d%H%M%S"),
                    "evento": nome_evento,
                    "responsavel": tecnico_resp,
                    "data_saida": data_saida.strftime("%d/%m/%Y"),
                    "itens": lista_itens_evento
                }
                eventos_ativos.append(novo_evento)
                salvar_eventos(eventos_ativos)
                
                registrar_historico("SAÍDA", nome_evento, tecnico_resp, "; ".join(resumo_log), "Saída de material confirmada")
                st.success(f"🎉 Checklist do evento '{nome_evento}' registrado com sucesso! Vá na aba '3. Imprimir Romaneio' se desejar imprimir a folha.")
                st.rerun()

# =====================================================================
# 2. CONFERIR RETORNO (CHECK-IN COM CHECKBOXES)
# =====================================================================
elif menu == "📥 2. Conferir Retorno (Entrada)":
    st.subheader("📥 Conferência de Retorno de Evento")
    
    if not eventos_ativos:
        st.info("✅ Não há nenhum evento na rua no momento. Todos os equipamentos estão no galpão!")
    else:
        opcoes_ev = [f"{ev['evento']} (Resp: {ev['responsavel']} - Data: {ev['data_saida']})" for ev in eventos_ativos]
        ev_selecionado_idx = st.selectbox("Selecione o Evento que está retornando ao galpão:", range(len(opcoes_ev)), format_func=lambda x: opcoes_ev[x])
        
        ev_atual = eventos_ativos[ev_selecionado_idx]
        
        st.markdown(f"**Evento:** {ev_atual['evento']} | **Responsável na Saída:** {ev_atual['responsavel']} | **Saída:** {ev_atual['data_saida']}")
        st.markdown("Marque a caixinha **☑️ Conferido** para cada equipamento que voltou ao armazém:")

        df_ret = pd.DataFrame(ev_atual["itens"])
        df_ret["Conferido"] = True
        df_ret["Observação / Avaria"] = "OK"

        df_ret_editado = st.data_editor(
            df_ret,
            key=f"retorno_{ev_atual['id']}",
            hide_index=True,
            use_container_width=True,
            disabled=["Código", "Equipamento", "Qtd"],
            column_config={
                "Conferido": st.column_config.CheckboxColumn("☑️ Conferido", default=True, width="small"),
                "Código": st.column_config.TextColumn("Cód.", width="small"),
                "Equipamento": st.column_config.TextColumn("Equipamento", width="large"),
                "Qtd": st.column_config.NumberColumn("Qtd que Saiu", width="small"),
                "Observação / Avaria": st.column_config.TextColumn("Estado / Observações na Volta", width="medium")
            }
        )

        col_r1, col_r2 = st.columns(2)
        with col_r1:
            conferente_volta = st.text_input("👷 Conferente de Entrada (Galpão):", value=ev_atual['responsavel'])
        with col_r2:
            obs_geral = st.text_input("📝 Observação Geral do Retorno:", value="Tudo conferido e guardado")

        if st.button("📥 Confirmar Retorno e Devolver Tudo ao Estoque", type="primary", use_container_width=True):
            if not all(df_ret_editado["Conferido"]):
                st.warning("⚠️ Marque a caixa '☑️ Conferido' em todos os itens da lista para finalizar a entrada completa do evento.")
            elif not conferente_volta:
                st.error("⚠️ Informe o nome do conferente de entrada.")
            else:
                resumo_volta = []
                avarias_lista = []
                for _, row in df_ret_editado.iterrows():
                    cod = row["Código"]
                    qtd_v = int(row["Qtd"])
                    df_estoque.loc[df_estoque["Código"] == cod, "Disponível"] += qtd_v
                    df_estoque.loc[df_estoque["Código"] == cod, "Em Uso"] = max(0, int(df_estoque.loc[df_estoque["Código"] == cod, "Em Uso"].values[0]) - qtd_v)
                    resumo_volta.append(f"{qtd_v}x {row['Equipamento']}")
                    if row["Observação / Avaria"].strip().upper() != "OK":
                        avarias_lista.append(f"{row['Equipamento']}: {row['Observação / Avaria']}")

                salvar_estoque(df_estoque)
                eventos_ativos.pop(ev_selecionado_idx)
                salvar_eventos(eventos_ativos)
                
                obs_final = obs_geral
                if avarias_lista:
                    obs_final += " | Detalhes: " + "; ".join(avarias_lista)

                registrar_historico("RETORNO", ev_atual["evento"], conferente_volta, "; ".join(resumo_volta), obs_final)
                st.success(f"✅ Retorno do evento '{ev_atual['evento']}' concluído! Todos os equipamentos voltaram ao estoque disponível.")
                st.rerun()

# =====================================================================
# 3. IMPRIMIR ROMANEIO / LISTA DE CHECKLIST
# =====================================================================
elif menu == "🖨️ 3. Imprimir Romaneio / Lista":
    st.subheader("🖨️ Impressão de Checklist / Romaneio")
    
    tipo_impressao = st.radio(
        "Escolha o modelo de folha que deseja gerar para impressão:",
        ["📄 Romaneio de um Evento Ativo (Apenas itens que vão para o evento)", "📋 Ficha Geral de Galpão (Lista para preenchimento manual à caneta)"],
        horizontal=True
    )
    st.info("💡 **Como imprimir:** Abaixo está a visualização pronta para papel A4. Pressione **Ctrl + P** no teclado (ou 'Imprimir / Salvar como PDF' no celular). O menu lateral escuro não sairá na folha!")
    st.markdown("---")

    if tipo_impressao.startswith("📄"):
        if not eventos_ativos:
            st.warning("Não há eventos ativos cadastrados no momento. Crie uma saída na aba '1. Novo Checklist' primeiro, ou escolha a 'Ficha Geral de Galpão' acima.")
        else:
            opcoes_ev = [f"{ev['evento']} ({ev['data_saida']})" for ev in eventos_ativos]
            idx_imp = st.selectbox("Selecione o Evento para imprimir:", range(len(opcoes_ev)), format_func=lambda x: opcoes_ev[x])
            ev_imp = eventos_ativos[idx_imp]

            # Cabeçalho do Documento Impresso
            c_logo, c_titulo = st.columns([1, 5])
            with c_logo:
                if logo_path:
                    st.image(logo_path, width=110)
            with c_titulo:
                st.markdown("## **A.C.S. SOM, LUZ E IMAGEM - ROMANEIO DE EVENTO**")
                st.markdown(f"**Evento / Destino:** {ev_imp['evento']} &nbsp;&nbsp;|&nbsp;&nbsp; **Técnico Resp.:** {ev_imp['responsavel']} &nbsp;&nbsp;|&nbsp;&nbsp; **Data:** {ev_imp['data_saida']}")
            
            st.markdown("---")
            
            tabela_romaneio = []
            for item in ev_imp["itens"]:
                tabela_romaneio.append({
                    "Código": item["Código"],
                    "Equipamento / Modelo": item["Equipamento"],
                    "Quantidade": f"{item['Qtd']} un",
                    "Check Saída": "[     ]",
                    "Check Retorno": "[     ]",
                    "Observações / Estado": "____________________"
                })
            st.table(pd.DataFrame(tabela_romaneio))

            st.markdown("<br><br>", unsafe_allow_html=True)
            ca1, ca2 = st.columns(2)
            with ca1:
                st.markdown("_________________________________________________")
                st.markdown(f"**Assinatura de Saída ({ev_imp['responsavel']})**")
            with ca2:
                st.markdown("_________________________________________________")
                st.markdown("**Conferência de Retorno ao Galpão (Data: ___/___/___)**")

    else:
        cat_filtro_imp = st.selectbox("Filtrar Categoria para a Ficha:", ["Todas", "Áudio", "Luz", "Informática", "Acessórios"])
        
        c_logo, c_titulo = st.columns([1, 5])
        with c_logo:
            if logo_path:
                st.image(logo_path, width=110)
        with c_titulo:
            st.markdown("## **A.C.S. SOM, LUZ E IMAGEM - FICHA DE CHECKLIST MANUAL**")
            st.markdown("**Evento:** ___________________________________ &nbsp;|&nbsp; **Técnico:** ______________________ &nbsp;|&nbsp; **Data:** ___/___/202__")
        
        st.markdown("---")
        df_ficha = df_estoque.copy()
        if cat_filtro_imp != "Todas":
            df_ficha = df_ficha[df_ficha["Categoria"] == cat_filtro_imp]

        tabela_ficha = []
        for _, row in df_ficha.iterrows():
            tabela_ficha.append({
                "Cód.": row["Código"],
                "Cat.": row["Categoria"],
                "Equipamento / Modelo": row["Equipamento"],
                "Disp.": row["Disponível"],
                "Qtd Saída": "_______",
                "Check Volta": "[    ]"
            })
        st.table(pd.DataFrame(tabela_ficha))

# =====================================================================
# 4. ESTOQUE GERAL DISPONÍVEL
# =====================================================================
elif menu == "📊 4. Estoque Geral Disponível":
    st.subheader("📊 Posição Atual do Estoque")
    
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        busca = st.text_input("🔍 Buscar equipamento por nome ou código:", placeholder="Ex: Shure, Line Array, Par LED...")
    with c2:
        f_cat = st.selectbox("Categoria:", ["Todas", "Áudio", "Luz", "Informática", "Acessórios"])
    with c3:
        f_disp = st.selectbox("Filtro:", ["Todos os Itens", "Somente em Uso na Rua", "Somente Disponíveis"])

    df_view = df_estoque.copy()
    if busca:
        df_view = df_view[df_view["Equipamento"].str.contains(busca, case=False, na=False) | df_view["Código"].str.contains(busca, case=False, na=False)]
    if f_cat != "Todas":
        df_view = df_view[df_view["Categoria"] == f_cat]
    if f_disp == "Somente em Uso na Rua":
        df_view = df_view[df_view["Em Uso"] > 0]
    elif f_disp == "Somente Disponíveis":
        df_view = df_view[df_view["Disponível"] > 0]

    st.dataframe(df_view, hide_index=True, use_container_width=True)

    with open(ARQUIVO_ESTOQUE, "rb") as f:
        st.download_button("📥 Baixar Planilha de Estoque (Excel)", f, file_name="estoque_acs.xlsx")

# =====================================================================
# 5. CADASTRAR OU AJUSTAR ITEM
# =====================================================================
elif menu == "➕ 5. Cadastrar / Ajustar Item":
    st.subheader("➕ Adicionar Novo Equipamento ao Inventário")
    with st.form("form_novo_item"):
        c1, c2 = st.columns(2)
        with c1:
            n_cod = st.text_input("Código (ex: 068):")
            n_nome = st.text_input("Nome / Modelo do Equipamento:")
        with c2:
            n_cat = st.selectbox("Categoria:", ["Áudio", "Luz", "Informática", "Acessórios"])
            n_qtd = st.number_input("Quantidade Total no Acervo:", min_value=1, value=1, step=1)
        
        if st.form_submit_button("💾 Salvar Novo Equipamento"):
            if n_cod and n_nome:
                novo_df = pd.DataFrame([{
                    "Código": n_cod,
                    "Categoria": n_cat,
                    "Equipamento": n_nome.upper(),
                    "Estoque Total": int(n_qtd),
                    "Disponível": int(n_qtd),
                    "Em Uso": 0
                }])
                df_estoque = pd.concat([df_estoque, novo_df], ignore_index=True)
                salvar_estoque(df_estoque)
                st.success(f"Equipamento '{n_nome.upper()}' cadastrado com sucesso!")
                st.rerun()
            else:
                st.error("Preencha o código e o nome do equipamento.")

# =====================================================================
# 6. HISTÓRICO DE EVENTOS
# =====================================================================
elif menu == "📜 6. Histórico de Eventos":
    st.subheader("📜 Histórico de Saídas e Retornos")
    if os.path.exists(ARQUIVO_HISTORICO):
        df_h = pd.read_excel(ARQUIVO_HISTORICO)
        st.dataframe(df_h, hide_index=True, use_container_width=True)
        with open(ARQUIVO_HISTORICO, "rb") as f:
            st.download_button("📥 Baixar Histórico Completo (Excel)", f, file_name="historico_eventos_acs.xlsx")
    else:
        st.info("Ainda não há registros no histórico.")
