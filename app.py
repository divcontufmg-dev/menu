import streamlit as st

# Configuração da página inicial
st.set_page_config(page_title="Hub de Automações", page_icon="⚙️", layout="centered")

st.title("⚙️ Menu Central de Ferramentas")
st.write("Bem-vindo! Selecione abaixo a automação que deseja utilizar:")

# Criando botões grandes que levam para as outras páginas
st.page_link("vbarmb.py", label="📊 1. VBA RMB", icon="▶️")
st.page_link("vbadep.py", label="🔧 2. VBA Depreciação", icon="▶️")
st.page_link("pages/rmb.py", label="📝 3. Conciliador RMB x SIAFI", icon="▶️")
st.page_link("dep.py", label="📁 4. Conciliador Depreciação x SIAFI", icon="▶️")

st.divider()
st.info("💡 **Dica:** Você também pode usar o menu lateral esquerdo para navegar entre as ferramentas e voltar para esta tela inicial a qualquer momento.")
