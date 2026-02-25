import streamlit as st

# Configuração da página inicial
st.set_page_config(
    page_title="Hub de Automações", 
    page_icon="⚙️", 
    layout="centered"
)

# ==========================================
# 1. ESCONDER COMPLETAMENTE O MENU E A BARRA LATERAL
# ==========================================
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            [data-testid="stSidebar"] {display: none;}
            [data-testid="collapsedControl"] {display: none;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ==========================================
# 2. CONTEÚDO DA PÁGINA PRINCIPAL
# ==========================================
st.title("⚙️ Menu Central de Ferramentas")
st.write("Bem-vindo! Selecione abaixo a automação que deseja utilizar:")
st.markdown("---")

# Criando botões grandes que levam para as outras páginas
# st.page_link("pages/1_Preparar planilha Siafi RMB.py", label="📊 1. VBA RMB", icon="▶️")
st.page_link("pages/2.1_Conciliador_RMB_x_SIAFI.py", label="📝 1. Conciliador RMB x SIAFI", icon="▶️")

# st.page_link("pages/3_Preparar planilha Siafi Depreciação.py", label="🔧 3. VBA Depreciação", icon="▶️")
st.page_link("pages/4.1_Conciliador_Depreciação_x_SIAFI.py", label="📁 2. Conciliador Depreciação x SIAFI", icon="▶️")

st.page_link("pages/5_Conciliador_Almoxarifado_x_SIAFI.py", label="💼 3. Conciliador Almoxarifado x SIAFI", icon="▶️")

st.page_link("pages/test.py", label="💼 3. test", icon="▶️")
