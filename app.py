import streamlit as st

# Configuração da página inicial
st.set_page_config(page_title="Hub de Automações", page_icon="⚙️", layout="centered")

st.title("⚙️ Menu Central de Ferramentas")
st.write("Bem-vindo! Selecione abaixo a automação que deseja utilizar:")

# Criando botões grandes que levam para as outras páginas
st.page_link("pages/1_Depreciacao.py", label="📊 1. Automação de Depreciação", icon="▶️")
st.page_link("pages/2_Ferramenta_Dois.py", label="🔧 2. Nome da Ferramenta Dois", icon="▶️")
st.page_link("pages/3_Ferramenta_Tres.py", label="📝 3. Nome da Ferramenta Três", icon="▶️")
st.page_link("pages/4_Ferramenta_Quatro.py", label="📁 4. Nome da Ferramenta Quatro", icon="▶️")

st.divider()
st.info("💡 **Dica:** Você também pode usar o menu lateral esquerdo para navegar entre as ferramentas e voltar para esta tela inicial a qualquer momento.")
