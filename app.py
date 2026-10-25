import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuração da Página
st.set_page_config(page_title="Checklist de Equipamentos - ACS", page_icon="🎛️", layout="wide")

# FICHEIROS DE DADOS
ARQUIVO_INVENTARIO = "inventario_equipamentos.xlsx"
ARQUIVO_HISTORICO = "historico_movimentacoes.xlsx"

def carregar_inventario():
    if os.path.exists(ARQUIVO_INVENTARIO):
        return pd.read_excel(ARQUIVO_INVENTARIO)
    else:
        # Inventário inicial padrão caso o Excel não exista
        df_inicial = pd.DataFrame([
            {"Código": "001", "Equipamento": "Microfone Shure SM58", "N° de Série": "SH-10293", "Categoria": "Áudio", "Status": "Disponível", "Responsável": "-"},
            {"Código": "002", "Equipamento": "Mesa de Som Behringer", "N° de Série": "BH-55421", "Categoria": "Áudio", "Status": "Disponível", "Responsável": "-"},
            {"Código": "003", "Equipamento": "Refletor LED Par 64", "N° de Série": "LD-88390", "Categoria": "Luz", "Status": "Disponível", "Responsável": "-"},
            {"Código": "004", "Equipamento": "Cabo BNC 10m", "N° de Série": "-", "Categoria": "Acessórios", "Status": "Disponível", "Responsável": "-"}
        ])
        df_inicial.to_excel(ARQUIVO_INVENTARIO, index=False)
        return df_inicial

def salvar_inventario(df):
    df.to_excel(ARQUIVO_INVENTARIO, index=False)

def registrar_historico(acao, itens, responsavel, evento_destino):
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    novo_registro = pd.DataFrame([{
        "Data/Hora": data_hora,
        "Ação": acao,
        "Equipamentos": ", ".join(itens),
        "Responsável": responsavel,
        "Evento/Destino": evento_destino
    }])
    
    if os.path.exists(ARQUIVO_HISTORICO):
        df_hist = pd.read_excel(ARQUIVO_HISTORICO)
        df_hist = pd.concat([df_hist, novo_registro], ignore_index=True)
    else:
        df_hist = novo_registro
    df_hist.to_excel(ARQUIVO_HISTORICO, index=False)

# --- SISTEMA DE SENHA / LOGIN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔒 Acesso Restrito - Controle de Equipamentos")
    st.markdown("Por favor, introduza a palavra-passe para aceder ao sistema.")
    
    with st.form("form_login"):
        senha_digitada = st.text_input("Palavra-passe de Acesso:", type="password")
        botao_login = st.form_submit_button("Entrar")
        
        if botao_login:
            if senha_digitada == "acs123":  # <--- Pode alterar a palavra-passe aqui se desejar
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Palavra-passe incorreta. Tente novamente.")
    st.stop()

# --- APLICATIVO PRINCIPAL ---
df = carregar_inventario()

st.title("🎛️ Gestão de Equipamentos - Som, Luz e Imagem")
st.markdown("Controlo profissional de inventário, saídas e entradas para eventos.")

# Botão para Sair / Bloquear novamente
if st.sidebar.button("🔒 Sair do Sistema"):
    st.session_state.autenticado = False
    st.rerun()

menu = st.sidebar.selectbox("Menu Principal", [
    "📋 Inventário Geral", 
    "📤 Registar Saída", 
    "📥 Registar Retorno (Entrada)", 
    "➕ Adicionar Equipamento", 
    "📜 Histórico de Movimentações"
])

# 1. INVENTÁRIO GERAL
if menu == "📋 Inventário Geral":
    st.subheader("📋 Estado Atual do Inventário")
    
    filtro_status = st.selectbox("Filtrar por Status", ["Todos", "Disponível", "Em uso"])
    
    if filtro_status != "Todos":
        df_filtrado = df[df["Status"] == filtro_status]
    else:
        df_filtrado = df

    st.dataframe(df_filtrado, use_container_width=True)
    
    # Botão para transferir o inventário em Excel
    with open(ARQUIVO_INVENTARIO, "rb") as f:
        st.download_button("📥 Descarregar Inventário em Excel", f, file_name="inventario_atualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 2. REGISTAR SAÍDA
elif menu == "📤 Registar Saída":
    st.subheader("📤 Saída de Equipamentos para Evento")
    
    disponiveis = df[df["Status"] == "Disponível"]["Equipamento"].tolist()
    
    if not disponiveis:
        st.warning("Não há equipamentos disponíveis no momento.")
    else:
        with st.form("form_saida"):
            itens_selecionados = st.multiselect("Selecione os Equipamentos:", disponiveis)
            responsavel = st.text_input("Técnico / Responsável pela Retirada:")
            evento = st.text_input("Nome do Evento / Destino:")
            
            botao_enviar = st.form_submit_button("Confirmar Saída")
            
            if botao_enviar:
                if itens_selecionados and responsavel and evento:
                    for item in itens_selecionados:
                        df.loc[df["Equipamento"] == item, "Status"] = "Em uso"
                        df.loc[df["Equipamento"] == item, "Responsável"] = f"{responsavel} ({evento})"
                    
                    salvar_inventario(df)
                    registrar_historico("SAÍDA", itens_selecionados, responsavel, evento)
                    
                    st.success(f"Saída registada com sucesso para {len(itens_selecionados)} item(ns)!")
                    st.rerun()
                else:
                    st.error("Por favor, preencha todos os campos obrigatórios.")

# 3. REGISTAR RETORNO (ENTRADA)
elif menu == "📥 Registar Retorno (Entrada)":
    st.subheader("📥 Retorno e Conferência de Equipamentos")
    
    em_uso = df[df["Status"] == "Em uso"]["Equipamento"].tolist()
    
    if not em_uso:
        st.info("Nenhum equipamento consta como 'Em uso' no momento.")
    else:
        with st.form("form_entrada"):
            itens_retorno = st.multiselect("Selecione os itens que estão a retornar:", em_uso)
            responsavel_retorno = st.text_input("Responsável pela Conferência de Entrada:")
            observacao = st.text_area("Observações de Retorno (ex: avarias, cabos danificados, estado geral):", "Tudo OK")
            
            botao_retorno = st.form_submit_button("Confirmar Retorno ao Armazém")
            
            if botao_retorno:
                if itens_retorno and responsavel_retorno:
                    for item in itens_retorno:
                        df.loc[df["Equipamento"] == item, "Status"] = "Disponível"
                        df.loc[df["Equipamento"] == item, "Responsável"] = "-"
                    
                    salvar_inventario(df)
                    registrar_historico("ENTRADA", itens_retorno, responsavel_retorno, f"Retorno - Obs: {observacao}")
                    
                    st.success("Retorno, conferência e histórico registados com sucesso!")
                    st.rerun()
                else:
                    st.error("Selecione ao menos um equipamento e informe quem conferiu.")

# 4. ADICIONAR EQUIPAMENTO
elif menu == "➕ Adicionar Equipamento":
    st.subheader("➕ Adicionar Novo Item ao Inventário")
    
    with st.form("form_cadastro"):
        novo_codigo = st.text_input("Código do Item (ex: 005):")
        novo_nome = st.text_input("Nome / Descrição do Equipamento:")
        novo_serie = st.text_input("Número de Série (opcional):", value="-")
        nova_categoria = st.selectbox("Categoria:", ["Áudio", "Luz", "Imagem", "Acessórios", "Outros"])
        
        salvar = st.form_submit_button("Guardar no Inventário")
        
        if salvar:
            if novo_codigo and novo_nome:
                novo_registro = pd.DataFrame([{
                    "Código": novo_codigo, 
                    "Equipamento": novo_nome, 
                    "N° de Série": novo_serie, 
                    "Categoria": nova_categoria, 
                    "Status": "Disponível", 
                    "Responsável": "-"
                }])
                df_atualizado = pd.concat([df, novo_registro], ignore_index=True)
                salvar_inventario(df_atualizado)
                st.success(f"Equipamento '{novo_nome}' adicionado com sucesso ao inventário!")
                st.rerun()
            else:
                st.error("Preencha obrigatoriamente o código e o nome do equipamento.")

# 5. HISTÓRICO DE MOVIMENTAÇÕES
elif menu == "📜 Histórico de Movimentações":
    st.subheader("📜 Registo de Histórico de Saídas e Entradas")
    
    if os.path.exists(ARQUIVO_HISTORICO):
        df_hist = pd.read_excel(ARQUIVO_HISTORICO)
        st.dataframe(df_hist, use_container_width=True)
        
        with open(ARQUIVO_HISTORICO, "rb") as f:
            st.download_button("📥 Descarregar Histórico em Excel", f, file_name="historico_movimentacoes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Ainda não existem registos de movimentações guardados.")
