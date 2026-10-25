import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuração da Página
st.set_page_config(page_title="Checklist - A.C.S. Som, Luz e Imagem", page_icon="🎛️", layout="wide")

# FICHEIROS DE DADOS
ARQUIVO_INVENTARIO = "inventario_equipamentos.xlsx"
ARQUIVO_HISTORICO = "historico_movimentacoes.xlsx"

def carregar_logo():
    for nome in ["logo.png", "Logo.png", "LOGO.png", "logo.JPG", "Logo.JPG"]:
        if os.path.exists(nome):
            return nome
    return None

def carregar_inventario():
    if os.path.exists(ARQUIVO_INVENTARIO):
        return pd.read_excel(ARQUIVO_INVENTARIO)
    else:
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

        lista_expandida = []
        for codigo, nome, qtd, cat in dados_iniciais:
            if qtd == 1:
                lista_expandida.append({"Código": f"{codigo}", "Equipamento": nome, "N° de Série": "-", "Categoria": cat, "Status": "Disponível", "Responsável": "-"})
            else:
                for i in range(1, qtd + 1):
                    lista_expandida.append({"Código": f"{codigo}-{i:03d}", "Equipamento": nome, "N° de Série": "-", "Categoria": cat, "Status": "Disponível", "Responsável": "-"})

        df_inicial = pd.DataFrame(lista_expandida)
        df_inicial.to_excel(ARQUIVO_INVENTARIO, index=False)
        return df_inicial

def salvar_inventario(df):
    df.to_excel(ARQUIVO_INVENTARIO, index=False)

def registrar_historico(acao, itens, responsavel, evento_destino):
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    novo_registro = pd.DataFrame([{"Data/Hora": data_hora, "Ação": acao, "Equipamentos": ", ".join(itens), "Responsável": responsavel, "Evento/Destino": evento_destino}])
    if os.path.exists(ARQUIVO_HISTORICO):
        df_hist = pd.read_excel(ARQUIVO_HISTORICO)
        df_hist = pd.concat([df_hist, novo_registro], ignore_index=True)
    else:
        df_hist = novo_registro
    df_hist.to_excel(ARQUIVO_HISTORICO, index=False)

# --- SISTEMA DE SENHA / LOGIN ---
logo_path = carregar_logo()

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    col_l1, col_l2 = st.columns([1, 4])
    with col_l1:
        if logo_path: st.image(logo_path, width=120)
    with col_l2:
        st.title("A.C.S. Som, Luz e Imagem")
    
    st.markdown("### 🔒 Acesso Restrito - Controle de Equipamentos")
    with st.form("form_login"):
        senha_digitada = st.text_input("Senha de Acesso:", type="password")
        botao_login = st.form_submit_button("Entrar")
        if botao_login:
            if senha_digitada == "acs123":
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
    st.stop()

# --- APLICATIVO PRINCIPAL ---
df = carregar_inventario()

col_topo1, col_topo2 = st.columns([1, 6])
with col_topo1:
    if logo_path: st.image(logo_path, width=100)
with col_topo2:
    st.title("🎛️ A.C.S. - Controle de Equipamentos")

st.markdown("---")

if st.sidebar.button("🔒 Sair do Sistema"):
    st.session_state.autenticado = False
    st.rerun()

menu = st.sidebar.selectbox("Menu Principal", [
    "📋 Inventário Geral", 
    "📤 Registrar Saída", 
    "📥 Registrar Retorno (Entrada)", 
    "🖨️ Lista para Impressão",
    "➕ Adicionar Equipamento", 
    "📜 Histórico de Movimentações"
])

# 1. INVENTÁRIO GERAL
if menu == "📋 Inventário Geral":
    st.subheader("📋 Estado Atual do Inventário")
    col1, col2 = st.columns(2)
    with col1: filtro_status = st.selectbox("Filtrar por Status", ["Todos", "Disponível", "Em uso"])
    with col2: filtro_cat = st.selectbox("Filtrar por Categoria", ["Todas", "Áudio", "Luz", "Informática", "Acessórios"])
    
    df_filtrado = df.copy()
    if filtro_status != "Todos": df_filtrado = df_filtrado[df_filtrado["Status"] == filtro_status]
    if filtro_cat != "Todas": df_filtrado = df_filtrado[df_filtrado["Categoria"] == filtro_cat]

    st.dataframe(df_filtrado, use_container_width=True)
    with open(ARQUIVO_INVENTARIO, "rb") as f:
        st.download_button("📥 Baixar Inventário em Excel", f, file_name="inventario_acs.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 2. REGISTRAR SAÍDA
elif menu == "📤 Registrar Saída":
    st.subheader("📤 Saída de Equipamentos para Evento")
    disponiveis = df[df["Status"] == "Disponível"]["Equipamento"].unique().tolist()
    if not disponiveis:
        st.warning("Não há equipamentos disponíveis no momento.")
    else:
        with st.form("form_saida"):
            item_escolhido = st.selectbox("Selecione o Equipamento:", disponiveis)
            itens_especificos = df[(df["Equipamento"] == item_escolhido) & (df["Status"] == "Disponível")]["Código"].tolist()
            unidades_selecionadas = st.multiselect("Selecione os Códigos/Unidades específicas:", itens_especificos)
            responsavel = st.text_input("Técnico / Responsável pela Retirada:")
            evento = st.text_input("Nome do Evento / Destino:")
            botao_enviar = st.form_submit_button("Confirmar Saída")
            
            if botao_enviar:
                if unidades_selecionadas and responsavel and evento:
                    for cod in unidades_selecionadas:
                        df.loc[df["Código"] == cod, "Status"] = "Em uso"
                        df.loc[df["Código"] == cod, "Responsável"] = f"{responsavel} ({evento})"
                    salvar_inventario(df)
                    registrar_historico("SAÍDA", unidades_selecionadas, responsavel, evento)
                    st.success(f"Saída registrada com sucesso para {len(unidades_selecionadas)} unidade(s)!")
                    st.rerun()
                else:
                    st.error("Por favor, preencha todos os campos obrigatórios.")

# 3. REGISTRAR RETORNO
elif menu == "📥 Registrar Retorno (Entrada)":
    st.subheader("📥 Retorno e Conferência de Equipamentos")
    em_uso = df[df["Status"] == "Em uso"]
    if em_uso.empty:
        st.info("Nenhum equipamento consta como 'Em uso' no momento.")
    else:
        opcoes_retorno = [f"{row['Código']} - {row['Equipamento']} (Com: {row['Responsável']})" for index, row in em_uso.iterrows()]
        with st.form("form_entrada"):
            itens_retorno_sel = st.multiselect("Selecione os itens/unidades que estão retornando:", opcoes_retorno)
            responsavel_retorno = st.text_input("Responsável pela Conferência de Entrada:")
            observacao = st.text_area("Observações de Retorno:", "Tudo OK")
            botao_retorno = st.form_submit_button("Confirmar Retorno ao Armazém")
            
            if botao_retorno:
                if itens_retorno_sel and responsavel_retorno:
                    codigos_retornando = [item.split(" - ")[0] for item in itens_retorno_sel]
                    for cod in codigos_retornando:
                        df.loc[df["Código"] == cod, "Status"] = "Disponível"
                        df.loc[df["Código"] == cod, "Responsável"] = "-"
                    salvar_inventario(df)
                    registrar_historico("ENTRADA", codigos_retornando, responsavel_retorno, f"Retorno - Obs: {observacao}")
                    st.success("Retorno registrado com sucesso!")
                    st.rerun()
                else:
                    st.error("Selecione ao menos um item e informe o responsável.")

# 4. LISTA PARA IMPRESSÃO
elif menu == "🖨️ Lista para Impressão":
    st.subheader("🖨️ Folha de Checklist para Impressão")
    with st.form("form_impressao"):
        nome_evento_imp = st.text_input("Nome do Evento / Show:", value="Evento Exemplo")
        responsavel_imp = st.text_input("Técnico Responsável:", value="Nome do Técnico")
        cat_imp = st.selectbox("Filtrar Categoria:", ["Todas", "Áudio", "Luz", "Informática", "Acessórios"])
        gerar_lista = st.form_submit_button("Gerar Lista")

    if gerar_lista:
        st.markdown("---")
        if logo_path: st.image(logo_path, width=120)
        st.markdown(f"## **A.C.S. SOM, LUZ E IMAGEM - CHECKLIST DE SAÍDA**")
        st.markdown(f"**Evento:** {nome_evento_imp}  |  **Técnico:** {responsavel_imp}  |  **Data:** {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        st.markdown("---")

        df_disp = df[df["Status"] == "Disponível"]
        if cat_imp != "Todas": df_disp = df_disp[df_disp["Categoria"] == cat_imp]
        resumo_imp = df_disp.groupby(["Categoria", "Equipamento"]).size().reset_index(name="Qtd Disponível")
        
        tabela_impressao = [{"Conferir": "[  ]", "Categoria": row["Categoria"], "Equipamento / Modelo": row["Equipamento"], "Qtd Estoque": row["Qtd Disponível"], "Qtd Leva": "_____", "Qtd Volta": "_____"} for _, row in resumo_imp.iterrows()]
        st.table(pd.DataFrame(tabela_impressao))
        st.info("💡 Pressione **Ctrl + P** no teclado para imprimir esta folha diretamente.")

# 5. ADICIONAR EQUIPAMENTO
elif menu == "➕ Adicionar Equipamento":
    st.subheader("➕ Adicionar Novo Item ao Inventário")
    with st.form("form_cadastro"):
        novo_codigo = st.text_input("Código do Item (ex: 068-001):")
        novo_nome = st.text_input("Nome / Descrição do Equipamento:")
        novo_serie = st.text_input("Número de Série:", value="-")
        nova_categoria = st.selectbox("Categoria:", ["Áudio", "Luz", "Informática", "Acessórios", "Outros"])
        salvar = st.form_submit_button("Salvar no Inventário")
        if salvar:
            if novo_codigo and novo_nome:
                novo_reg = pd.DataFrame([{"Código": novo_codigo, "Equipamento": novo_nome, "N° de Série": novo_serie, "Categoria": nova_categoria, "Status": "Disponível", "Responsável": "-"}])
                salvar_inventario(pd.concat([df, novo_reg], ignore_index=True))
                st.success(f"Equipamento '{novo_nome}' adicionado com sucesso!")
                st.rerun()
            else:
                st.error("Preencha o código e o nome.")

# 6. HISTÓRICO
elif menu == "📜 Histórico de Movimentações":
    st.subheader("📜 Registro de Histórico")
    if os.path.exists(ARQUIVO_HISTORICO):
        st.dataframe(pd.read_excel(ARQUIVO_HISTORICO), use_container_width=True)
    else:
        st.info("Sem registros de movimentações.")
