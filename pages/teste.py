import streamlit as st
import re

# Oculta marcas do Streamlit e a barra lateral (Padrão do seu repositório)
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
st.markdown("Como o sistema roda na nuvem, utilize esta ferramenta para gerar um script seguro que digitará as UGs automaticamente no módulo Comunica.")

st.divider()

col_input, col_output = st.columns(2)

with col_input:
    st.subheader("1. Cole os dados do PDF")
    texto_ugs = st.text_area(
        "Cole a lista de UGs (pode conter vírgulas, espaços e quebras de linha):", 
        height=250,
        placeholder="Ex: 153254, 153255, 153256,\n153280, 153281..."
    )
    
    gerar = st.button("Gerar Código de Automação", type="primary", use_container_width=True)

with col_output:
    st.subheader("2. Código Gerado")
    
    if gerar and texto_ugs:
        # Extrai apenas números que tenham pelo menos 5 dígitos (para evitar pegar números de página soltos)
        lista_ugs = re.findall(r'\b\d{5,6}\b', texto_ugs)
        
        if not lista_ugs:
            st.warning("Nenhuma UG válida encontrada no texto.")
        else:
            st.success(f"✅ {len(lista_ugs)} UGs identificadas!")
            
            # Formata a lista para o JavaScript
            array_js = str(lista_ugs)
            
            # Script JavaScript que fará o papel do pyautogui direto no navegador
            script_js = f"""
// 1. Clique dentro do campo de destinatário no SIAFI antes de rodar!
var ugs = {array_js};
var i = 0;

var timer = setInterval(function() {{
    if(i >= ugs.length) {{
        clearInterval(timer);
        alert("Automação Concluída! " + i + " UGs inseridas.");
        return;
    }}
    
    var campo = document.activeElement;
    campo.value = ugs[i];
    
    // Simula a digitação para o sistema reconhecer
    campo.dispatchEvent(new Event("input", {{ bubbles: true }}));
    
    // Simula o aperto da tecla Enter
    campo.dispatchEvent(new KeyboardEvent("keydown", {{ bubbles: true, key: "Enter", keyCode: 13 }}));
    
    i++;
}}, 800); // 800 milissegundos de pausa entre cada UG
            """
            
            st.code(script_js, language="javascript")
            
            st.info("📋 **Como usar:** \n"
                    "1. Clique no botão de copiar no canto superior direito do código acima.\n"
                    "2. Vá para a tela do SIAFI Web e clique dentro do campo de destinatário.\n"
                    "3. Aperte **F12** no teclado para abrir as ferramentas do desenvolvedor.\n"
                    "4. Clique na aba **Console**, cole o código e aperte **Enter**.")
