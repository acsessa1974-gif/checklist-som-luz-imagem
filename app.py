import streamlit as st
import pandas as pd
import os

# Configuração da Página
st.set_page_config(page_title="Checklist de Equipamentos", page_icon="🎛️", layout="wide")

# ARQUIVO DE BANCO DE DADOS LOCAL (Salva automaticamente no computador/nuvem)
ARQUIVO_DADOS = "inventario_equipamentos.csv"

def carregar_dados():
    if os.path.exists(ARQUIVO_DADOS):
        return pd.read_csv(ARQUIVO_DADOS)
    else:
        # Dados iniciais padrão caso o arquivo não exista
        df_inicial = pd.DataFrame([
            {"Código": "001", "Equipamento": "Microfone Shure SM58", "Categoria": "Áudio", "Status": "Disponível", "Responsável": "-"},
            {"Código": "002", "Equipamento": "Mesa de Som Behringer", "Categoria": "Áudio", "Status": "Disponível", "Responsável": "-"},
            {"Código": "003", "Equipamento": "Refletor LED Par 64", "Categoria": "Luz", "Status": "Disponível", "Responsável": "-"},
            {"Código": "004", "Equipamento": "Cabo BNC 10m", "Categoria": "Acessórios", "Status": "Disponível", "Responsável": "-"}
        ])
        df_inicial.to_csv(ARQUIVO_DADOS, index=False)
        return df_inicial

def salvar_dados(df):
    df.to_csv(ARQUIVO_DADOS, index=False)

# --- SISTEMA DE SENHA / LOGIN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔒 Acesso Restrito - Controle de Equipamentos")
    st.markdown("Por favor, digite a senha para acessar o sistema.")
    
    with st.form("form_login"):
        senha_digitada = st.text_input("Senha de Acesso:", type="password")
        botao_login = st.form_submit_button("Entrar")
        
        if botao_login:
            # Defina aqui a senha que você e seus funcionários vão usar
            if senha_digitada == "acs123":  # <--- Altere a senha aqui se quiser
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
    
    # Para o código aqui se não estiver logado
    st.stop()

# --- APLICATIVO PRINCIPAL (Aparece apenas após a senha correta) ---
df = carregar_dados()

st.title("🎛️ Controle de Saída e Entrada de Equipamentos")
st.markdown("Gerencie o inventário e faça o checklist de forma rápida e segura.")

# Botão para Sair / Bloquear novamente
if st.sidebar.button("🔒 Sair do Sistema"):
    st.session_state.autenticado = False
    st.rerun()

menu = st.sidebar.selectbox("Menu Principal", ["Inventário & Status", "Registrar Saída", "Registrar Retorno (Entrada)", "Cadastrar Novo Item"])

# 1. INVENTÁRIO & STATUS
if menu == "Inventário & Status":
    st.subheader("📋 Status Atual dos Equipamentos")
    
    filtro_status = st.selectbox("Filtrar por Status", ["Todos", "Disponível", "Em uso"])
    
    if filtro_status != "Todos":
        df_filtrado = df[df["Status"] == filtro_status]
    else:
        df_filtrado = df

    st.dataframe(df_filtrado, use_container_width=True)

# 2. REGISTRAR SAÍDA
elif menu == "Registrar Saída":
    st.subheader("📤 Registrar Saída de Equipamentos")
    
    disponiveis = df[df["Status"] == "Disponível"]["Equipamento"].tolist()
    
    if not disponiveis:
        st.warning("Não há equipamentos disponíveis no momento.")
    else:
        with st.form("form_saida"):
            itens_selecionados = st.multiselect("Selecione os Equipamentos:", disponiveis)
            responsavel = st.text_input("Técnico / Responsável pelo Evento:")
            evento = st.text_input("Nome do Evento / Destino:")
            
            botao_enviar = st.form_submit_button("Confirmar Saída")
            
            if botao_enviar:
                if itens_selecionados and responsavel and evento:
                    for item in itens_selecionados:
                        df.loc[df["Equipamento"] == item, "Status"] = "Em uso"
                        df.loc[df["Equipamento"] == item, "Responsável"] = f"{responsavel} ({evento})"
                    salvar_dados(df)
                    st.success(f"Saída registrada com sucesso para {len(itens_selecionados)} item(ns)!")
                    st.rerun()
                else:
                    st.error("Por favor, preencha todos os campos.")

# 3. REGISTRAR RETORNO (ENTRADA)
elif menu == "Registrar Retorno (Entrada)":
    st.subheader("📥 Registrar Retorno de Equipamentos")
    
    em_uso = df[df["Status"] == "Em uso"]["Equipamento"].tolist()
    
    if not em_uso:
        st.info("Nenhum equipamento consta como 'Em uso' no momento.")
    else:
        with st.form("form_entrada"):
            itens_retorno = st.multiselect("Selecione os itens que estão retornando:", em_uso)
            observacao = st.text_area("Observações sobre o estado (avarias, cabos danificados, etc):", "Tudo OK")
            
            botao_retorno = st.form_submit_button("Confirmar Retorno")
            
            if botao_retorno:
                if itens_retorno:
                    for item in itens_retorno:
                        df.loc[df["Equipamento"] == item, "Status"] = "Disponível"
                        df.loc[df["Equipamento"] == item, "Responsável"] = "-"
                    salvar_dados(df)
                    st.success("Retorno e conferência registrados com sucesso!")
                    st.rerun()
                else:
                    st.error("Selecione ao menos um equipamento.")

# 4. CADASTRAR NOVO ITEM
elif menu == "Cadastrar Novo Item":
    st.subheader("➕ Adicionar Novo Equipamento ao Inventário")
    
    with st.form("form_cadastro"):
        novo_codigo = st.text_input("Código do Item (ex: 005):")
        novo_nome = st.text_input("Nome/Descrição do Equipamento:")
        nova_categoria = st.selectbox("Categoria:", ["Áudio", "Luz", "Imagem", "Acessórios", "Outros"])
        
        salvar = st.form_submit_button("Salvar no Inventário")
        
        if salvar:
            if novo_codigo and novo_nome:
                novo_registro = pd.DataFrame([{"Código": novo_codigo, "Equipamento": novo_nome, "Categoria": nova_categoria, "Status": "Disponível", "Responsável": "-"}])
                df_atualizado = pd.concat([df, novo_registro], ignore_index=True)
                salvar_dados(df_atualizado)
                st.success(f"Equipamento '{novo_nome}' cadastrado com sucesso!")
                st.rerun()
            else:
                st.error("Preencha o código e o nome do equipamento.")