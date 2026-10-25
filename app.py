import streamlit as st
import pandas as pd
import os
import math
import requests
from datetime import datetime

# Configuração da Página
st.set_page_config(page_title="A.C.S. - Sistema Integrado de Gestão e Orçamentos", page_icon="🎛️", layout="wide")

# FICHEIROS DE DADOS
ARQUIVO_INVENTARIO = "inventario_equipamentos.xlsx"
ARQUIVO_HISTORICO = "historico_movimentacoes.xlsx"
CHAVE_API = "AIzaSyBySuGQBc55DWAJo6H5Hsv2y8A7rFnYjNY"

# Inicializar lista global para o Orçamento Total Consolidado na sessão do Streamlit
if "orcamento_total_itens" not in st.session_state:
    st.session_state.orcamento_total_itens = []

def carregar_logo():
    for nome in ["logo.png", "Logo.png", "LOGO.png", "logo.JPG", "Logo.JPG"]:
        if os.path.exists(nome):
            return nome
    return None

def formatar_br(valor):
    """Formata um float para o padrão monetário brasileiro: 1.234,56"""
    try:
        val_str = f"{valor:,.2f}"
        val_str = val_str.replace(",", "X").replace(".", ",").replace("X", ".")
        return val_str
    except Exception:
        return f"{valor:.2f}"

def extrair_valor_total(texto_orcamento):
    """Extrai corretamente o valor monetário final do texto gerado pelo módulo."""
    try:
        linhas = texto_orcamento.strip().split('\n')
        for linha in reversed(linhas):
            if "R$" in linha:
                partes = linha.split("R$")
                trecho = partes[-1].strip()
                trecho_limpo = trecho.replace('.', '').replace(',', '.')
                return float(trecho_limpo)
    except Exception:
        pass
    return 0.0

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
    
    st.markdown("### 🔒 Acesso Restrito - Sistema Integrado")
    with st.form("form_login"):
        senha_digitada = st.text_input("Palavra-passe de Acesso:", type="password")
        botao_login = st.form_submit_button("Entrar")
        if botao_login:
            if senha_digitada == "acs123":
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Palavra-passe incorreta. Tente novamente.")
    st.stop()

# --- APLICATIVO PRINCIPAL ---
df = carregar_inventario()

col_topo1, col_topo2 = st.columns([1, 6])
with col_topo1:
    if logo_path: st.image(logo_path, width=100)
with col_topo2:
    st.title("🎛️ A.C.S. - Sistema Integrado")

st.markdown("---")

if st.sidebar.button("🔒 Sair do Sistema"):
    st.session_state.autenticado = False
    st.rerun()

menu = st.sidebar.selectbox("Menu Principal", [
    "📋 Inventário Geral", 
    "📤 Registar Saída", 
    "📥 Registar Retorno (Entrada)", 
    "🖨️ Lista para Impressão",
    "➕ Adicionar Equipamento", 
    "📜 Histórico de Movimentações",
    "📦 Orç: Palco e Forração",
    "⚡ Orç: Gerador",
    "🚚 Orç: Logística",
    "💰 Orç: Custos",
    "🖼️ Orç: Backdrop",
    "🖥️ Orç: Painel de LED",
    "🔊 Orç: Sonorização",
    "📄 Orçamento Total Consolidado"
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
        st.download_button("📥 Descarregar Inventário em Excel", f, file_name="inventario_acs.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 2. REGISTAR SAÍDA
elif menu == "📤 Registar Saída":
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
                    st.success(f"Saída registada com sucesso para {len(unidades_selecionadas)} unidade(s)!")
                    st.rerun()
                else:
                    st.error("Por favor, preencha todos os campos obrigatórios.")

# 3. REGISTAR RETORNO
elif menu == "📥 Registar Retorno (Entrada)":
    st.subheader("📥 Retorno e Conferência de Equipamentos")
    em_uso = df[df["Status"] == "Em uso"]
    if em_uso.empty:
        st.info("Nenhum equipamento consta como 'Em uso' no momento.")
    else:
        opcoes_retorno = [f"{row['Código']} - {row['Equipamento']} (Com: {row['Responsável']})" for index, row in em_uso.iterrows()]
        with st.form("form_entrada"):
            itens_retorno_sel = st.multiselect("Selecione os itens/unidades que estão a retornar:", opcoes_retorno)
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
                    st.success("Retorno registado com sucesso!")
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
        st.info("💡 Prima **Ctrl + P** no teclado para imprimir esta folha diretamente.")

# 5. ADICIONAR EQUIPAMENTO
elif menu == "➕ Adicionar Equipamento":
    st.subheader("➕ Adicionar Novo Item ao Inventário")
    with st.form("form_cadastro"):
        novo_codigo = st.text_input("Código do Item (ex: 068-001):")
        novo_nome = st.text_input("Nome / Descrição do Equipamento:")
        novo_serie = st.text_input("Número de Série:", value="-")
        nova_categoria = st.selectbox("Categoria:", ["Áudio", "Luz", "Informática", "Acessórios", "Outros"])
        salvar = st.form_submit_button("Guardar no Inventário")
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
    st.subheader("📜 Registo de Histórico")
    if os.path.exists(ARQUIVO_HISTORICO):
        st.dataframe(pd.read_excel(ARQUIVO_HISTORICO), use_container_width=True)
    else:
        st.info("Sem registos de movimentações.")

# ==========================================
# MÓDULOS DE ORÇAMENTO (ADAPTADOS DO SEU CÓDIGO)
# ==========================================

# MÓDULO: PALCO E FORRAÇÃO
elif menu == "📦 Orç: Palco e Forração":
    st.subheader("📦 Módulo de Cálculo: Palco e Forração")
    with st.form("form_palco"):
        c1, c2 = st.columns(2)
        with c1:
            largura_palco = st.number_input("Largura do Palco (m):", value=8.0)
            profundidade = st.number_input("Profundidade do Palco (m):", value=6.0)
            altura_saia = st.number_input("Altura do Palco / Saia (m):", value=0.8)
            largura_bobina = st.number_input("Largura da Bobina do Carpete (m):", value=2.0)
            cor_carpete = st.text_input("Cor do Carpete:", value="preto")
        with c2:
            valor_carpete_metro = st.number_input("Valor Carpete (metro linear R$):", value=24.0)
            valor_praticavel_m2 = st.number_input("Valor Praticável (m² R$):", value=50.0)
            valor_frete = st.number_input("Valor do Frete (R$):", value=0.0)
            valor_mao_obra_m2 = st.number_input("Mão de Obra (m² R$):", value=45.15)
            percentual_imposto = st.number_input("Imposto (% sobre total):", value=10.0)
            percentual_desconto = st.number_input("Desconto Aplicável (%):", value=0.0)
        
        calc_palco = st.form_submit_button("Calcular Orçamento de Palco")

    if calc_palco:
        larg_prat = math.ceil(largura_palco / 0.5) * 0.5
        prof_prat = math.ceil(profundidade / 0.5) * 0.5

        def simular_montagem(L, P):
            q_2x1 = int(L // 2) * int(P // 1)
            resto_L = L % 2
            P_inteiro = int(P // 1)
            q_1x1 = 0; q_05x1 = 0
            if resto_L >= 1:
                q_2x1 += int(P_inteiro // 2) * int(resto_L // 1)
                resto_P_da_faixa = P_inteiro % 2
                if resto_P_da_faixa >= 1: q_1x1 += int(resto_P_da_faixa // 1) * int(resto_L // 1)
            resto_L_05 = resto_L % 1
            if resto_L_05 >= 0.5: q_05x1 += int(P_inteiro // 1) * int(resto_L_05 / 0.5)
            resto_P_05 = P % 1
            if resto_P_05 >= 0.5:
                q_05x1 += int(L // 1)
                if (L % 1) >= 0.5: q_05x1 += 1
            return q_2x1, q_1x1, q_05x1

        op1_2x1, op1_1x1, op1_05x1 = simular_montagem(larg_prat, prof_prat)
        op2_2x1, op2_1x1, op2_05x1 = simular_montagem(prof_prat, larg_prat)
        qtd_2x1, qtd_1x1, qtd_05x1 = (op2_2x1, op2_1x1, op2_05x1) if op2_2x1 > op1_2x1 else (op1_2x1, op1_1x1, op1_05x1)

        total_modulos = qtd_2x1 + qtd_1x1 + qtd_05x1
        total_pes = total_modulos * 4
        area_praticaveis_real = (qtd_2x1 * 2) + (qtd_1x1 * 1) + (qtd_05x1 * 0.5)
        custo_praticaveis = area_praticaveis_real * valor_praticavel_m2
        custo_mao_obra = area_praticaveis_real * valor_mao_obra_m2
        
        tiras_piso = math.ceil(largura_palco / largura_bobina)
        metragem_piso = tiras_piso * profundidade
        perimetro_saia = 2 * (largura_palco + profundidade)
        metragem_saia = (perimetro_saia / math.floor(largura_bobina / altura_saia)) if altura_saia > 0 and math.floor(largura_bobina / altura_saia) > 0 else 0
        
        metragem_total = metragem_piso + metragem_saia
        metragem_final = metragem_total * 1.10
        custo_carpete = metragem_final * valor_carpete_metro
        
        subtotal = custo_carpete + custo_praticaveis + valor_frete + custo_mao_obra
        valor_imposto = subtotal * (percentual_imposto / 100)
        custo_total_bruto = subtotal + valor_imposto
        valor_desconto_aplicado = custo_total_bruto * (percentual_desconto / 100)
        custo_total_geral = custo_total_bruto - valor_desconto_aplicado

        res_txt = (
            f"--- MEDIDAS DO PROJETO ---\nPalco: {largura_palco:.2f}m x {profundidade:.2f}m | Saia: {altura_saia:.2f}m\n\n"
            f"--- ESTRUTURA (PRATICÁVEIS) ---\nMódulos (2x1: {qtd_2x1} | 1x1: {qtd_1x1} | 0.5x1: {qtd_05x1})\n"
            f"Total de Pés: {total_pes} un | Área: {area_praticaveis_real:.2f} m² | Custo: R$ {formatar_br(custo_praticaveis)}\n\n"
            f"--- FORRAÇÃO ({cor_carpete.upper()}) ---\nTotal (+10%): {metragem_final:.2f}m | Custo: R$ {formatar_br(custo_carpete)}\n\n"
            f"SUBTOTAL: R$ {formatar_br(subtotal)}\nIMPOSTO ({percentual_imposto}%): R$ {formatar_br(valor_imposto)}\n"
            f"TOTAL GERAL COM DESCONTO: R$ {formatar_br(custo_total_geral)}"
        )
        st.text_area("Resultado Gerado:", res_txt, height=200)
        
        if st.button("➕ Enviar para Orçamento Total Consolidado"):
            st.session_state.orcamento_total_itens.append(("Palco e Forração", custo_total_geral))
            st.success("Enviado com sucesso para o Orçamento Total!")

# MÓDULO: GERADOR
elif menu == "⚡ Orç: Gerador":
    st.subheader("⚡ Módulo de Cálculo: Gerador")
    with st.form("form_gerador"):
        carga_watts = st.number_input("Carga Total do Evento (Watts):", value=15000.0)
        margem_seguranca = st.number_input("Margem de Segurança (%):", value=30.0)
        calc_ger = st.form_submit_button("Calcular Gerador")

    if calc_ger:
        potencia_kw = carga_watts / 1000
        potencia_kva = potencia_kw / 0.8
        kva_necessario = potencia_kva * (1 + (margem_seguranca / 100))
        gerador_rec = math.ceil(kva_necessario)
        
        # Estimar custo base por kVA (ex: R$ 35 por kVA de locação diária como referência base ou valor estipulado)
        custo_gerador = gerador_rec * 35.0 

        res_txt = (
            f"--- DIMENSIONAMENTO ---\nCarga Total: {formatar_br(carga_watts)} W ({formatar_br(potencia_kw)} kW)\n"
            f"Potência Base: {formatar_br(potencia_kva)} kVA | Margem: {margem_seguranca}%\n"
            f"GERADOR RECOMENDADO: {gerador_rec} kVA\nEstimação de Locação Base: R$ {formatar_br(custo_gerador)}"
        )
        st.text_area("Resultado Gerado:", res_txt, height=150)
        
        if st.button("➕ Enviar para Orçamento Total Consolidado"):
            st.session_state.orcamento_total_itens.append((f"Gerador {gerador_rec} kVA", custo_gerador))
            st.success("Enviado com sucesso!")

# MÓDULO: LOGÍSTICA
elif menu == "🚚 Orç: Logística":
    st.subheader("🚚 Módulo de Cálculo: Logística e Rotas")
    with st.form("form_log"):
        origem = st.text_input("Endereço de Origem:", value="Jacarepaguá, Rio de Janeiro - RJ")
        destino = st.text_input("Endereço de Destino:", value="Copacabana, Rio de Janeiro - RJ")
        c1, c2, c3 = st.columns(3)
        with c1: consumo_kml = st.number_input("Consumo (km/L):", value=7.0)
        with c2: preco_combustivel = st.number_input("Preço Combustível (R$):", value=7.08)
        with c3: valor_depreciacao = st.number_input("Depreciação/km (R$):", value=0.65)
        calc_log = st.form_submit_button("Calcular Rota e Custos")

    if calc_log:
        url = "https://routes.googleapis.com/directions/v2:computeRoutes"
        headers = {"Content-Type": "application/json", "X-Goog-Api-Key": CHAVE_API, "X-Goog-FieldMask": "routes.distanceMeters,routes.travelAdvisory.tollInfo"}
        payload = {"origin": {"address": origem}, "destination": {"address": destino}, "travelMode": "DRIVE", "extraComputations": ["TOLLS"]}
        
        try:
            resp = requests.post(url, json=payload, headers=headers)
            dados = resp.json()
            if 'routes' in dados and len(dados['routes']) > 0:
                rota = dados['routes'][0]
                distancia_km = rota.get('distanceMeters', 0) / 1000
                distancia_total = distancia_km * 2
                
                pedagio_api_ida = 0.0
                if 'travelAdvisory' in rota and 'tollInfo' in rota['travelAdvisory']:
                    toll_info = rota['travelAdvisory']['tollInfo']
                    if 'estimatedPrice' in toll_info:
                        for price in toll_info['estimatedPrice']:
                            if price.get('currencyCode') == 'BRL':
                                pedagio_api_ida += float(price.get('units', 0)) + (float(price.get('nanos', 0)) / 1e9)
                
                pedagio_total = pedagio_api_ida * 2
                litros = distancia_total / consumo_kml
                custo_comb = litros * preco_combustivel
                custo_dep = distancia_total * valor_depreciacao
                custo_total_rota = custo_comb + pedagio_total + custo_dep

                res_txt = (
                    f"--- DADOS DA ROTA ---\nDistância (Ida): {formatar_br(distancia_km)} km | Total (Ida+Volta): {formatar_br(distancia_total)} km\n"
                    f"Combustível Necessário: {formatar_br(litros)} L | Custo Combustível: R$ {formatar_br(custo_comb)}\n"
                    f"Pedágios: R$ {formatar_br(pedagio_total)} | Depreciação: R$ {formatar_br(custo_dep)}\n"
                    f"TOTAL DA LOGÍSTICA: R$ {formatar_br(custo_total_rota)}"
                )
                st.text_area("Resultado Gerado:", res_txt, height=180)
                if st.button("➕ Enviar para Orçamento Total Consolidado"):
                    st.session_state.orcamento_total_itens.append(("Logística / Transporte", custo_total_rota))
                    st.success("Enviado com sucesso!")
            else:
                st.error("Rota não encontrada pela API do Google.")
        except Exception as e:
            st.error(f"Erro ao calcular rota: {e}")

# MÓDULO: CUSTOS
elif menu == "💰 Orç: Custos":
    st.subheader("💰 Centro de Custos da Equipa")
    with st.form("form_custos"):
        val_freelancers = st.number_input("Free Lancers (Mão de Obra) - R$:", value=0.0)
        val_alimentacao = st.number_input("Alimentação da Equipa - R$:", value=0.0)
        val_combustivel = st.number_input("Combustível/Portagens - R$:", value=0.0)
        val_hospedagem = st.number_input("Hospedagem - R$:", value=0.0)
        val_outros = st.number_input("Outros Gastos / Imprevistos - R$:", value=0.0)
        calc_custos = st.form_submit_button("Calcular Custos")

    if calc_custos:
        total_custos = val_freelancers + val_alimentacao + val_combustivel + val_hospedagem + val_outros
        res_txt = (
            f"--- RELATÓRIO DE CUSTOS ---\nFree Lancers: R$ {formatar_br(val_freelancers)}\n"
            f"Alimentação: R$ {formatar_br(val_alimentacao)} | Hospedagem: R$ {formatar_br(val_hospedagem)}\n"
            f"CUSTO TOTAL DE OPERAÇÃO: R$ {formatar_br(total_custos)}"
        )
        st.text_area("Resultado:", res_txt, height=150)
        if st.button("➕ Enviar para Orçamento Total Consolidado"):
            st.session_state.orcamento_total_itens.append(("Centro de Custos / Equipa", total_custos))
            st.success("Enviado com sucesso!")

# MÓDULO: BACKDROP
elif menu == "🖼️ Orç: Backdrop":
    st.subheader("🖼️ Módulo de Cálculo: Backdrop")
    with st.form("form_bd"):
        largura_bd = st.number_input("Largura do Backdrop (m):", value=4.0)
        altura_bd = st.number_input("Altura do Backdrop (m):", value=3.0)
        valor_ponto = st.number_input("Valor do Ponto (R$):", value=50.0)
        desconto_bd = st.number_input("Desconto (%):", value=0.0)
        fechado = st.checkbox("Quadro completo (fechado na base)", value=True)
        calc_bd = st.form_submit_button("Calcular Backdrop")

    if calc_bd:
        # Lógica simplificada de pontos de truss proporcional à área/perímetro
        perimetro = (largura_bd * 2) + (altura_bd * (2 if fechado else 1))
        total_pontos = math.ceil(perimetro * 1.5)
        custo_bruto = total_pontos * valor_ponto
        custo_final = custo_bruto * (1 - (desconto_bd / 100))
        
        res_txt = f"Backdrop {largura_bd}m x {altura_bd}m\nPontos Estimados: {total_pontos} un\nCUSTO DO BACKDROP: R$ {formatar_br(custo_final)}"
        st.text_area("Resultado:", res_txt, height=120)
        if st.button("➕ Enviar para Orçamento Total Consolidado"):
            st.session_state.orcamento_total_itens.append(("Backdrop / Estrutura", custo_final))
            st.success("Enviado com sucesso!")

# MÓDULO: PAINEL DE LED
elif menu == "🖥️ Orç: Painel de LED":
    st.subheader("🖥️ Módulo de Cálculo: Painel de LED")
    with st.form("form_led"):
        c1, c2 = st.columns(2)
        with c1:
            larg_led = st.number_input("Largura (m):", value=4.0)
            alt_led = st.number_input("Altura (m):", value=3.0)
            pitch = st.selectbox("Pixel Pitch:", ["P2.9", "P3.9"])
        with c2:
            valor_m2 = st.number_input("Valor de Locação por m² (R$):", value=350.0)
            qtd_tec = st.number_input("Qtd Técnicos de Vídeo:", value=1)
            val_tec = st.number_input("Valor Unit. Técnico (R$):", value=400.0)
        calc_led = st.form_submit_button("Calcular LED")

    if calc_led:
        area_m2 = larg_led * alt_led
        custo_eq = area_m2 * valor_m2
        custo_tec = qtd_tec * val_tec
        total_led = custo_eq + custo_tec

        res_txt = f"Painel LED {pitch} ({area_m2:.2f} m²)\nEquipamento: R$ {formatar_br(custo_eq)}\nTécnicos: R$ {formatar_br(custo_tec)}\nTOTAL LED: R$ {formatar_br(total_led)}"
        st.text_area("Resultado:", res_txt, height=130)
        if st.button("➕ Enviar para Orçamento Total Consolidado"):
            st.session_state.orcamento_total_itens.append((f"Painel de LED {pitch}", total_led))
            st.success("Enviado com sucesso!")

# MÓDULO: SONORIZAÇÃO
elif menu == "🔊 Orç: Sonorização":
    st.subheader("🔊 Módulo de Cálculo: Sonorização e Multimédia")
    with st.form("form_som"):
        q_caixa = st.number_input("Caixas Ativas c/ Pedestal:", value=2)
        v_caixa = st.number_input("Valor Unit. Caixa (R$):", value=300.0)
        q_mic = st.number_input("Microfones Sem Fio:", value=4)
        v_mic = st.number_input("Valor Unit. Mic (R$):", value=350.0)
        v_frete_som = st.number_input("Frete Fixo (R$):", value=400.0)
        calc_som_mod = st.form_submit_button("Calcular Sonorização")

    if calc_som_mod:
        t_caixas = q_caixa * v_caixa
        t_mics = q_mic * v_mic
        total_som = t_caixas + t_mics + v_frete_som
        res_txt = f"Sonorização Pequeno/Médio Porte\nCaixas: R$ {formatar_br(t_caixas)} | Mics: R$ {formatar_br(t_mics)}\nTOTAL SOM: R$ {formatar_br(total_som)}"
        st.text_area("Resultado:", res_txt, height=120)
        if st.button("➕ Enviar para Orçamento Total Consolidado"):
            st.session_state.orcamento_total_itens.append(("Sonorização e Multimédia", total_som))
            st.success("Enviado com sucesso!")

# MÓDULO: ORÇAMENTO CONSOLIDADO TOTAL
elif menu == "📄 Orçamento Total Consolidado":
    st.subheader("📄 Orçamento Total Consolidado da A.C.S.")
    
    if not st.session_state.orcamento_total_itens:
        st.info("Ainda não existem itens adicionados ao orçamento consolidado. Visite os módulos de cálculo anteriores e clique em 'Enviar para Orçamento Total'.")
    else:
        soma_geral = 0.0
        texto_cons = "--- CONSOLIDADO DO EVENTO ---\n\n"
        for idx, (nome, valor) in enumerate(st.session_state.orcamento_total_itens, 1):
            texto_cons += f"{idx}. {nome}: R$ {formatar_br(valor)}\n"
            soma_geral += valor
        
        texto_cons += f"\n-------------------------------------\nVALOR TOTAL GERAL: R$ {formatar_br(soma_geral)}"
        
        st.text_area("Resumo Geral:", texto_cons, height=300)
        
        if st.button("🗑️ Limpar Todos os Itens do Consolidado"):
            st.session_state.orcamento_total_itens.clear()
            st.success("Orçamento total limpo com sucesso!")
            st.rerun()
