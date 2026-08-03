import streamlit as st
import pandas as pd
import os
from datetime import date
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Mapa de Restrições", layout="wide")

# 1. Função de Extração de Dados
@st.cache_data
def carregar_dados_planilha():
    caminho_raiz = os.path.dirname(os.path.dirname(__file__))
    arquivo = os.path.join(caminho_raiz, "base.xlsx")
    
    try:
        df_ugs = pd.read_excel(arquivo, sheet_name="ug")
        df_restricoes = pd.read_excel(arquivo, sheet_name="restrições")
        
        # Prevenção de renderização de zeros em campos nulos
        df_ugs = df_ugs.fillna("").replace(0, "")
        df_restricoes = df_restricoes.fillna("").replace(0, "")
        
        lista_ugs = [str(ug) for ug in df_ugs.iloc[:, 0].tolist() if str(ug).strip() != ""]
        dict_restricoes = dict(zip(df_restricoes.iloc[:, 0], df_restricoes.iloc[:, 1]))
        
        return lista_ugs, dict_restricoes
        
    except Exception as e:
        st.error(f"Erro ao ler o arquivo {arquivo}. Verifique se as abas 'UG' e 'RESTRIÇÕES' existem. Detalhes: {e}")
        return [], {}

st.title("Mapa de Restrições por UG")

lista_ugs, dict_restricoes = carregar_dados_planilha()

if not lista_ugs or not dict_restricoes:
    st.warning("Aguardando carregamento da base de dados...")
    st.stop() 

# Inicialização da memória de estado
if 'restricoes_aplicadas' not in st.session_state:
    st.session_state.restricoes_aplicadas = {}

st.divider()

# Data de Fechamento
col_data1, col_data2 = st.columns(2)
with col_data1:
    data_fechamento = st.date_input("Data de Fechamento", value=date.today(), format="DD/MM/YYYY")
    
data_anterior = data_fechamento - relativedelta(months=1)
with col_data2:
    st.info(f"**Mês de Referência para Análise:** {data_anterior.strftime('%B / %Y')}")

st.divider()

# Interface de Seleção
col_selecao, col_restricoes = st.columns(2)

with col_selecao:
    ug_selecionada = st.selectbox("Selecione a UG:", lista_ugs)
    
with col_restricoes:
    # Novo campo de busca para destacar restrições
    busca = st.text_input("🔍 Pesquisar restrição (por nome ou descrição):", "")
    
    st.write("Marque as restrições desta UG:")
    restricoes_marcadas = []
    
    for restricao, descricao in dict_restricoes.items():
        # Lógica de destaque visual
        destaque = ""
        if busca:
            # Verifica se o texto digitado está no nome ou na descrição (ignorando maiúsculas/minúsculas)
            if busca.lower() in str(restricao).lower() or busca.lower() in str(descricao).lower():
                destaque = " 👈 (ENCONTRADA)"
        
        # O checkbox renderiza a restrição e o destaque (se houver). 
        # A 'key' agora inclui a ug_selecionada para resetar a tela ao trocar de UG.
        if st.checkbox(str(restricao) + destaque, help=str(descricao), key=f"chk_{ug_selecionada}_{restricao}"):
            restricoes_marcadas.append(restricao)
            
# Confirmação Individual
if st.button("Confirmar UG", type="primary"):
    if restricoes_marcadas:
        st.session_state.restricoes_aplicadas[ug_selecionada] = restricoes_marcadas
    else:
        st.session_state.restricoes_aplicadas[ug_selecionada] = ["SEM RESTRIÇÃO"]
        
    st.success(f"Log registrado para: {ug_selecionada}")

# Área de exibição de Log
if st.session_state.restricoes_aplicadas:
    st.write("### UGs processadas até o momento")
    resumo_texto = ""
    for ug, rests in st.session_state.restricoes_aplicadas.items():
        resumo_texto += f"• UG {ug}: {', '.join(rests)}\n"
    
    st.text_area("Acompanhamento:", value=resumo_texto, height=150, disabled=True)

st.divider()

# Finalização e Relatório
if st.button("Gerar Dashboard de Restrições", use_container_width=True):
    dados_dashboard = []
    
    for ug in lista_ugs:
        restricoes = st.session_state.restricoes_aplicadas.get(ug, ["SEM RESTRIÇÃO"])
            
        dados_dashboard.append({
            "UG": ug,
            "Situação": "Sem Restrição" if restricoes == ["SEM RESTRIÇÃO"] else "Com Restrição",
            "Restrições Aplicadas": ", ".join(restricoes)
        })
        
    df_dashboard = pd.DataFrame(dados_dashboard)
    
    # Tratamento final antes da exibição
    df_dashboard = df_dashboard.fillna("").replace(0, "")
    
    total_analisadas = len(df_dashboard)
    total_com_restricao = len(df_dashboard[df_dashboard["Situação"] == "Com Restrição"])
    total_sem_restricao = len(df_dashboard[df_dashboard["Situação"] == "Sem Restrição"])
    
    st.subheader("Painel Geral de Conformidade")
    
    col_metrica1, col_metrica2, col_metrica3 = st.columns(3)
    col_metrica1.metric("Total de UGs da Base", total_analisadas)
    col_metrica2.metric("Com Restrição", total_com_restricao)
    col_metrica3.metric("Sem Restrição", total_sem_restricao)
    
    st.dataframe(df_dashboard, use_container_width=True, hide_index=True)
    
    if st.button("Nova Análise (Limpar Memória)"):
        st.session_state.restricoes_aplicadas = {}
        st.rerun()
