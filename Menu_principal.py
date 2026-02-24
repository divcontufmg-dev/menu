import streamlit as st

# Configuração da página inicial
st.set_page_config(page_title="Hub de Automações", page_icon="⚙️", layout="centered")

# ==========================================
# 1. ESCONDER O MENU AUTOMÁTICO DO STREAMLIT
# ==========================================
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] {display: none;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CRIAR O SEU MENU LATERAL PERSONALIZADO
# ==========================================
with st.sidebar:
    st.markdown("### 🧭 Navegação")
    # Coloque aqui apenas as páginas que você quer que apareçam no menu lateral
    st.page_link("pages/2.1_Conciliador_RMB_x_SIAFI.py", label="1. Conciliador RMB x SIAFI", icon="📝")
    st.page_link("pages/4.1_Conciliador_Depreciação_x_SIAFI.py", label="2. Conciliador Depreciação", icon="📁")
    st.page_link("pages/5_Conciliador_Almoxarifado_x_SIAFI.py", label="3. Conciliador Almoxarifado", icon="💼")


# ==========================================
# 3. CONTEÚDO DA PÁGINA PRINCIPAL
# ==========================================
st.title("⚙️ Menu Central de Ferramentas")
st.write("Bem-vindo! Selecione abaixo a automação que deseja utilizar:")

# Criando botões grandes que levam para as outras páginas
# st.page_link("pages/1_Preparar planilha Siafi RMB.py", label="📊 1. VBA RMB", icon="▶️")
st.page_link("pages/2.1_Conciliador_RMB_x_SIAFI.py", label="📝 1. Conciliador RMB x SIAFI", icon="▶️")
# st.page_link("pages/3_Preparar planilha Siafi Depreciação.py", label="🔧 3. VBA Depreciação", icon="▶️")
st.page_link("pages/4.1_Conciliador_Depreciação_x_SIAFI.py", label="📁 2. Conciliador Depreciação x SIAFI", icon="▶️")
st.page_link("pages/5_Conciliador_Almoxarifado_x_SIAFI.py", label="💼 3. Conciliador Almoxarifado x SIAFI", icon="▶️")

st.divider()
st.info("💡 **Dica:** Você também pode usar o menu lateral esquerdo para voltar a esse menu.")
