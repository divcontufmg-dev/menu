import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io
import os
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image
from pytesseract import Output

# ==========================================
# CONFIGURAÇÃO INICIAL E MEMÓRIA
# ==========================================
st.set_page_config(
    page_title="Conciliação Patrimonial: RMB x SIAFI",
    page_icon="📊",
    layout="wide"
)

# Inicializa a memória do Streamlit
if 'dados_processados' not in st.session_state:
    st.session_state.dados_processados = False
if 'dados_ug' not in st.session_state:
    st.session_state.dados_ug = {}
if 'avisos_usuario' not in st.session_state:
    st.session_state.avisos_usuario = []

# Ocultar marcas do Streamlit e o menu lateral automático para um visual mais limpo
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

# Botão para retornar à tela inicial solto no topo da tela
st.page_link("Menu_principal.py", label="⬅️ Voltar ao Menu Inicial")

# ==========================================
# FUNÇÕES DE PROCESSAMENTO (BASTIDORES)
# ==========================================
def limpar_valor(v):
    if v is None or pd.isna(v) or str(v).strip() == '': return 0.0
    if isinstance(v, (int, float)): return float(v)
    v = str(v).replace('"', '').replace("'", "").strip()
    if re.search(r',\d{1,2}$', v): v = v.replace('.', '').replace(',', '.')
    elif re.search(r'\.\d{1,2}$', v): v = v.replace(',', '')
    try: return float(re.sub(r'[^\d.-]', '', v))
    except: return 0.0

def extract_excel_data(df_raw):
    extracted_data = []
    for idx, row in df_raw.iterrows():
        if row.isna().all(): continue
        val_0 = str(row.iloc[0]).strip().replace('.0', '')
        
        if val_0.startswith('123'):
            codigo = val_0
            desc = "SEM DESCRIÇÃO"
            val = 0.0
            cols = [c for c in row.iloc[1:] if pd.notna(c) and str(c).strip() != '']
            
            if len(cols) >= 2:
                desc = str(cols[0]).strip().upper()
                val = limpar_valor(cols[1])
            elif len(cols) == 1:
                parsed_val = limpar_valor(cols[0])
                if parsed_val != 0.0 or str(cols[0]).strip() in ['0', '0.0']:
                    val = parsed_val
                else:
                    desc = str(cols[0]).strip().upper()
                    
            extracted_data.append({'Conta': codigo, 'Descricao': desc, 'Valor': val})
    return pd.DataFrame(extracted_data)

def get_chave_vinculo(conta, dict_matriz):
    conta = str(conta).strip()
    if conta in dict_matriz:
        val_matriz = str(dict_matriz[conta])
        match = re.search(r'(\d+)$', val_matriz)
        if match:
            digits = match.group(1)
            return int(digits[-2:]) if len(digits) >= 2 else int(digits)
    return None

def formatar_real(valor):
    sinal = "-" if valor < -0.001 else ""
    return f"{sinal}{abs(valor):,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

class PDF_Report(FPDF):
    def header(self):
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de Conferência Patrimonial', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(5)
    def footer(self):
        self.set_y(-15); self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='C')

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================
st.title("📊 Sistema de Conciliação: RMB x SIAFI")

with st.expander("📘 GUIA DE USO (Clique para abrir)", expanded=False):
    st.markdown("📌 **Orientações de Uso**")
    st.markdown("""
    1. Anexe a **Planilha SIAFI** e todos os **Relatórios PDF (RMB)** correspondentes na área abaixo.
    2. O sistema somará tudo. **Poderá corrigir valores divergentes. As edições serão registadas no PDF!**
    """)

# Área de Upload Unificada
arquivos_enviados = st.file_uploader(
    "📂 Arraste ou selecione a Planilha SIAFI (.xlsx) e os PDFs do RMB de uma só vez", 
    accept_multiple_files=True, 
    type=['xlsx', 'pdf']
)

# ==========================================
# ETAPA 1: PROCESSAMENTO DE DADOS
# ==========================================
if st.button("🚀 Gerar Relatório de Conciliação", type="primary", use_container_width=True):
    
    if not os.path.exists("MATRIZ.xlsx"):
        st.error("❌ O arquivo de configuração interno ('MATRIZ.xlsx') não foi encontrado. Contate o suporte técnico.")
        st.stop()
        
    if not arquivos_enviados:
        st.warning("⚠️ Por favor, insira seus arquivos para que possamos realizar a conciliação.")
        st.stop()

    uploaded_siafi = None
    uploaded_pdfs = []
    
    for arquivo in arquivos_enviados:
        if arquivo.name.lower().endswith('.xlsx'):
            if uploaded_siafi is None:
                uploaded_siafi = arquivo
            else:
                st.info(f"ℹ️ O sistema utilizará a planilha '{uploaded_siafi.name}' como base principal.")
        elif arquivo.name.lower().endswith('.pdf'):
            uploaded_pdfs.append(arquivo)

    if not uploaded_siafi:
        st.error("❌ A Planilha SIAFI (.xlsx) não foi encontrada entre os arquivos enviados.")
        st.stop()
    if not uploaded_pdfs:
        st.error("❌ Nenhum relatório RMB (.pdf) foi encontrado entre os arquivos enviados.")
        st.stop()

    progresso = st.progress(0)
    status_text = st.empty()
    status_text.text("Preparando ambiente de conciliação...")

    try:
        df_matriz = pd.read_excel("MATRIZ.xlsx", header=None)
        dict_matriz = {}
        for i in range(len(df_matriz)):
            c0 = str(df_matriz.iloc[i, 0]).strip().replace('.0', '')
            c1 = str(df_matriz.iloc[i, 1]).strip().replace('.0', '')
            if c0.startswith('123'): dict_matriz[c0] = c1
            elif c1.startswith('123'): dict_matriz[c1] = c0
    except Exception as e:
        st.error("❌ Ocorreu um erro ao ler a Matriz de configuração. Verifique o arquivo MATRIZ.xlsx.")
        st.stop()

    pdfs = {f.name: f for f in uploaded_pdfs}
    pares = []
    avisos_usuario = []

    try:
        xls_file = pd.ExcelFile(uploaded_siafi)
        for sheet_name in xls_file.sheet_names:
            if sheet_name.upper() == "MATRIZ": continue
            match = re.search(r'^(\d+)', sheet_name)
            if match:
                ug = match.group(1)
                pdf_match = next((f for n, f in pdfs.items() if n.startswith(ug)), None)
                if pdf_match: 
                    pares.append({'ug': ug, 'sheet_name': sheet_name, 'pdf': pdf_match})
                else: 
                    avisos_usuario.append(f"Falta PDF: A Unidade Gestora {ug} está na planilha, mas o PDF correspondente não foi enviado.")
    except Exception as e:
        st.error("❌ Não foi possível ler a Planilha SIAFI. Certifique-se de que o arquivo não está corrompido.")
        st.stop()

    if not pares:
        st.error("❌ Não foi possível encontrar pares correspondentes (Aba do Excel + PDF com o mesmo número de UG). Verifique o nome dos arquivos.")
        st.stop()

    dados_ug = {}
    
    for idx, par in enumerate(pares):
        ug = par['ug']
        status_text.text(f"Lendo e analisando dados da Unidade Gestora: {ug}...")
        
        # Leitura Excel
        df_padrao = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa'])
        saldo_estoque = 0.0
        tem_estoque_com_saldo = False
        
        try:
            df_raw = pd.read_excel(xls_file, sheet_name=par['sheet_name'], header=None)
            df_dados = extract_excel_data(df_raw)
            
            if not df_dados.empty:
                if '123110801' in df_dados['Conta'].values:
                    saldo_estoque = df_dados[df_dados['Conta'] == '123110801']['Valor'].sum()
                    if abs(saldo_estoque) > 0.0: tem_estoque_com_saldo = True
                
                contas_ignoradas = ['123110703', '123110402', '123119910', '123110801']
                df_dados = df_dados[~df_dados['Conta'].isin(contas_ignoradas)].copy()
                df_dados['Chave_Vinculo'] = df_dados['Conta'].apply(lambda c: get_chave_vinculo(c, dict_matriz))
                df_valid = df_dados.dropna(subset=['Chave_Vinculo']).copy()
                
                if not df_valid.empty:
                    df_valid['Chave_Vinculo'] = df_valid['Chave_Vinculo'].astype(int)
                    df_padrao = df_valid.groupby('Chave_Vinculo').agg({'Valor': 'sum', 'Descricao': 'first'}).reset_index()
                    df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']
        except Exception as e:
            avisos_usuario.append(f"Erro ao processar os dados da planilha para a UG {ug}.")

        # Leitura PDF
        df_pdf_final = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_PDF'])
        dados_pdf = []
        try:
            par['pdf'].seek(0)
            pdf_bytes = par['pdf'].read()
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as p_doc:
                for page in p_doc.pages:
                    txt = page.extract_text()
                    is_ocr = False
                    if not txt or len(txt) < 50:
                        is_ocr = True
                        try:
                            imagens = convert_from_bytes(pdf_bytes, first_page=page.page_number, last_page=page.page_number, dpi=300)
                            if imagens: txt = pytesseract.image_to_string(imagens[0], lang='por', config='--psm 6')
                        except: pass

                    if not txt: continue
                    if "DE ENTRADAS" in txt.upper() or "DE SAÍDAS" in txt.upper(): continue

                    for line in txt.split('\n'):
                        line = line.strip()
                        if re.match(r'^"?\d+', line):
                            vals = []
                            if is_ocr:
                                vals_raw = re.findall(r'([\d\.\s]+,\d{2})', line)
                                vals = [v.replace(' ', '') for v in vals_raw]
                            else:
                                vals = re.findall(r'([0-9]{1,3}(?:[.,][0-9]{3})*[.,]\d{2})', line)
                            
                            if len(vals) >= 4:
                                chave_match = re.match(r'^"?(\d+)', line)
                                if chave_match:
                                    chave_raw = chave_match.group(1)
                                    chave_final = int(chave_raw[-2:]) if len(chave_raw) >= 4 else int(chave_raw)
                                    dados_pdf.append({'Chave_Vinculo': chave_final, 'Saldo_PDF': limpar_valor(vals[-4])})
            if dados_pdf:
                df_pdf_final = pd.DataFrame(dados_pdf).groupby('Chave_Vinculo')['Saldo_PDF'].sum().reset_index()
        except Exception as e: 
            avisos_usuario.append(f"Erro ao ler o documento PDF da UG {ug}.")

        dados_ug[ug] = {
            'df_excel': df_padrao,
            'df_pdf': df_pdf_final.copy(),
            'df_pdf_original': df_pdf_final.copy(), # Memória fotográfica
            'saldo_estoque': saldo_estoque,
            'tem_estoque': tem_estoque_com_saldo
        }
        
        progresso.progress((idx + 1) / len(pares))

    st.session_state.dados_ug = dados_ug
    st.session_state.avisos_usuario = avisos_usuario
    st.session_state.dados_processados = True
    progresso.empty()
    status_text.empty()

# ==========================================
# ETAPA 2: REVISÃO CIRÚRGICA E PDF FINAL
# ==========================================
if st.session_state.get('dados_processados'):
    st.markdown("---")
    st.subheader("🔍 Resultados da Conciliação & Revisão")
    st.info("💡 **Ação Cirúrgica:** Altere os valores na coluna dos Relatórios caso o sistema tenha interpretado algo incorretamente. O cálculo refaz-se na hora.")

    pdf_out = PDF_Report()
    pdf_out.add_page()
    dados_ug = st.session_state.dados_ug

    for ug, info in dados_ug.items():
        # Cruzamento Original (para descobrir onde a edição deve ser liberada permanentemente)
        original = pd.merge(info['df_pdf_original'], info['df_excel'], on='Chave_Vinculo', how='outer').fillna(0)
        original['Diferenca_Original'] = original['Saldo_PDF'] - original['Saldo_Excel']
        chaves_com_erro = original[abs(original['Diferenca_Original']) > 0.05]['Chave_Vinculo'].tolist()

        # Atualiza df_pdf com base nas edições feitas pelo usuário
        for chave in chaves_com_erro:
            key_input = f"edit_rmb_{ug}_{int(chave)}"
            if key_input in st.session_state:
                novo_val = st.session_state[key_input]
                if chave in info['df_pdf']['Chave_Vinculo'].values:
                    info['df_pdf'].loc[info['df_pdf']['Chave_Vinculo'] == chave, 'Saldo_PDF'] = novo_val
                else:
                    info['df_pdf'] = pd.concat([info['df_pdf'], pd.DataFrame([{'Chave_Vinculo': chave, 'Saldo_PDF': novo_val}])], ignore_index=True)

        # Cruzamento Atualizado
        final = pd.merge(info['df_pdf'], info['df_excel'], on='Chave_Vinculo', how='outer').fillna(0)
        final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if pd.notna(x['Descricao_Completa']) and str(x['Descricao_Completa']).strip() != '0' else "ITEM SEM DESCRIÇÃO NO SIAFI", axis=1)
        final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
        
        soma_pdf = final['Saldo_PDF'].sum()
        soma_excel = final['Saldo_Excel'].sum()
        dif_total = soma_pdf - soma_excel

        with st.container():
            st.markdown(f"### 🏢 Unidade Gestora: {ug}")
            
            # Métricas
            col1, col2, col3 = st.columns(3)
            col1.metric("Total relatório (PDF)", f"R$ {formatar_real(soma_pdf)}")
            col2.metric("Total SIAFI (Excel)", f"R$ {formatar_real(soma_excel)}")
            col3.metric("Diferença Encontrada", f"R$ {formatar_real(dif_total)}", delta_color="inverse" if abs(dif_total) > 0.05 else "normal")
            
            if chaves_com_erro:
                tem_erro_atual = abs(dif_total) > 0.05
                titulo_expander = f"⚠️ Contas com Divergência" if tem_erro_atual else "✅ Corrigido Manualmente"
                
                with st.expander(titulo_expander, expanded=tem_erro_atual):
                    df_view = final[final['Chave_Vinculo'].isin(chaves_com_erro)].copy()
                    df_view = df_view[['Chave_Vinculo', 'Descricao', 'Saldo_PDF', 'Saldo_Excel', 'Diferenca']]
                    df_view.rename(columns={'Saldo_PDF': 'Saldo relatório', 'Saldo_Excel': 'Saldo SIAFI'}, inplace=True)
                    
                    st.dataframe(df_view.style.format({
                        "Saldo relatório": lambda x: f"R$ {formatar_real(x)}", 
                        "Saldo SIAFI": lambda x: f"R$ {formatar_real(x)}", 
                        "Diferença": lambda x: f"R$ {formatar_real(x)}"
                    }), use_container_width=True)

                    st.markdown("**✏️ Correção Manual por Item:**")
                    cols = st.columns(3)
                    for idx, chave in enumerate(chaves_com_erro):
                        val_atual = final.loc[final['Chave_Vinculo'] == chave, 'Saldo_PDF'].sum()
                        with cols[idx % 3]:
                            st.number_input(f"Item {int(chave)}", value=float(val_atual), step=100.0, key=f"edit_rmb_{ug}_{int(chave)}")
            else: 
                st.success("✅ Conciliado com sucesso! Nenhuma divergência de valores foi encontrada.")

            if info['tem_estoque']: 
                st.info(f"Aviso Contábil: A Conta de Estoque Interno (123110801) possui saldo de R$ {formatar_real(info['saldo_estoque'])}.")
            st.markdown("---")

            # ==========================================
            # GERAÇÃO DO PDF PARA A UG
            # ==========================================
            pdf_out.set_font("helvetica", 'B', 11)
            pdf_out.set_fill_color(240, 240, 240)
            pdf_out.cell(0, 10, text=f"Unidade Gestora: {ug}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
            
            mask_mostrar = (abs(final['Diferenca']) > 0.05) | (final['Chave_Vinculo'].isin(chaves_com_erro))
            itens_para_mostrar = final[mask_mostrar].copy()

            if not itens_para_mostrar.empty:
                pdf_out.set_font("helvetica", 'B', 9)
                pdf_out.set_fill_color(255, 200, 200)
                pdf_out.cell(15, 8, "Item", 1, fill=True)
                pdf_out.cell(85, 8, "Descrição da Conta", 1, fill=True)
                pdf_out.cell(30, 8, "Saldo relatório", 1, fill=True)
                pdf_out.cell(30, 8, "Saldo SIAFI", 1, fill=True)
                pdf_out.cell(30, 8, "Diferença", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                pdf_out.set_font("helvetica", '', 8)
                alertas_auditoria = []

                for _, row in itens_para_mostrar.iterrows():
                    chave = row['Chave_Vinculo']
                    val_original = info['df_pdf_original'].loc[info['df_pdf_original']['Chave_Vinculo'] == chave, 'Saldo_PDF'].sum()
                    editado = abs(row['Saldo_PDF'] - val_original) > 0.01

                    str_saldo_relatorio = formatar_real(row['Saldo_PDF']) + (" *" if editado else "")
                    
                    pdf_out.cell(15, 7, str(int(chave)), 1)
                    pdf_out.cell(85, 7, str(row['Descricao'])[:48], 1)
                    pdf_out.cell(30, 7, str_saldo_relatorio, 1, align='R')
                    pdf_out.cell(30, 7, formatar_real(row['Saldo_Excel']), 1, align='R')
                    if abs(row['Diferenca']) > 0.05: pdf_out.set_text_color(200, 0, 0)
                    pdf_out.cell(30, 7, formatar_real(row['Diferenca']), 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)

                    if editado:
                        alertas_auditoria.append(f"* ALERTA: O valor do Item {int(chave)} foi alterado manualmente pelo utilizador. (Valor original lido do PDF: R$ {formatar_real(val_original)})")

                # Impressão dos alertas de auditoria
                if alertas_auditoria:
                    pdf_out.set_font("helvetica", 'I', 7)
                    pdf_out.set_text_color(180, 0, 0)
                    for alerta in alertas_auditoria:
                        pdf_out.cell(0, 5, alerta, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)
            else:
                pdf_out.set_font("helvetica", 'I', 9)
                pdf_out.cell(0, 8, "Nenhuma divergência encontrada entre SIAFI e os Relatórios.", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            if info['tem_estoque']:
                pdf_out.ln(2)
                pdf_out.set_font("helvetica", 'B', 9)
                pdf_out.set_fill_color(255, 255, 200)
                pdf_out.cell(100, 8, "SALDO ESTOQUE INTERNO (123110801)", 1, fill=True)
                pdf_out.cell(90, 8, f"R$ {formatar_real(info['saldo_estoque'])}", 1, fill=True, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf_out.ln(2)
            pdf_out.set_font("helvetica", 'B', 9)
            pdf_out.set_fill_color(220, 230, 241)
            pdf_out.cell(100, 8, "RESUMO DOS TOTAIS", 1, fill=True)
            pdf_out.cell(30, 8, formatar_real(soma_pdf), 1, fill=True, align='R')
            pdf_out.cell(30, 8, formatar_real(soma_excel), 1, fill=True, align='R')
            if abs(dif_total) > 0.05: pdf_out.set_text_color(200, 0, 0)
            pdf_out.cell(30, 8, formatar_real(dif_total), 1, fill=True, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf_out.set_text_color(0, 0, 0)
            pdf_out.ln(5)

    if st.session_state.avisos_usuario:
        st.warning("⚠️ **Avisos do Sistema:**")
        for aviso in st.session_state.avisos_usuario:
            st.write(f"- {aviso}")
    
    try:
        pdf_bytes = bytes(pdf_out.output())
        st.download_button(
            label="📄 Fazer Download do Relatório Completo (PDF)", 
            data=pdf_bytes, 
            file_name="Relatorio_Conciliacao_Patrimonial_RMB.pdf", 
            mime="application/pdf", 
            type="primary", 
            use_container_width=True
        )
    except Exception as e: 
        st.error("Ocorreu um erro ao gerar o arquivo PDF para download.")
