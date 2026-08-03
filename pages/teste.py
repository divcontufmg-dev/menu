import streamlit as st
import re

# Oculta marcas do Streamlit e a barra lateral
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
st.page_link("Menu_principal.py", label="⬅️ Voltar ao Menu Inicial")

st.title("🤖 Automação SIAFI: Inserção em Lote")
st.markdown("Como o sistema roda em nuvem e o SIAFI bloqueia scripts de navegador, esta ferramenta gera um executável nativo do Windows para digitar as UGs pelo seu teclado.")

st.divider()

col_input, col_output = st.columns(2)

with col_input:
    st.subheader("1. Cole os dados do PDF")
    texto_ugs = st.text_area(
        "Cole a lista de UGs (pode conter vírgulas, espaços e quebras de linha):", 
        height=250,
        placeholder="Ex: 153254, 153255, 153256,\n153280, 153281..."
    )
    
    gerar = st.button("Gerar Arquivo de Automação", type="primary", use_container_width=True)

with col_output:
    st.subheader("2. Arquivo Gerado")
    
    if gerar and texto_ugs:
        # Extrai apenas números com 5 ou 6 dígitos
        lista_ugs = re.findall(r'\b\d{5,6}\b', texto_ugs)
        
        if not lista_ugs:
            st.warning("Nenhuma UG válida encontrada no texto.")
        else:
            st.success(f"✅ {len(lista_ugs)} UGs identificadas!")
            
            # --- CONSTRUÇÃO DO SCRIPT NATIVO DO WINDOWS (VBS) ---
            vbs_code = 'Set WshShell = WScript.CreateObject("WScript.Shell")\n'
            vbs_code += 'WScript.Sleep 5000\n' # 5 segundos de pausa inicial para você clicar no SIAFI
            
            for ug in lista_ugs:
                vbs_code += f'WshShell.SendKeys "{ug}"\n'
                vbs_code += 'WScript.Sleep 500\n'         # Pausa antes do Enter
                vbs_code += 'WshShell.SendKeys "{ENTER}"\n'
                vbs_code += 'WScript.Sleep 1000\n'        # Pausa de 1 segundo para o SIAFI carregar
                
            # Converte a string para bytes para o botão de download
            vbs_bytes = vbs_code.encode('utf-8')
            
            st.download_button(
                label="⚙️ Baixar Robô de Digitação (.vbs)",
                data=vbs_bytes,
                file_name="digitar_siafi.vbs",
                mime="text/plain",
                type="primary",
                use_container_width=True
            )
            
            st.info("📋 **Como usar:** \n"
                    "1. Baixe o arquivo `digitar_siafi.vbs`.\n"
                    "2. Dê **dois cliques** no arquivo baixado.\n"
                    "3. Você terá **5 segundos** para abrir o navegador e clicar dentro do campo de destinatário no SIAFI.\n"
                    "4. Solte o mouse e aguarde. O Windows fará a digitação sozinho.")
