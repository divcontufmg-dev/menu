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
import copy

# ==========================================
# CONFIGURAÇÃO INICIAL E MEMÓRIA
# ==========================================
st.set_page_config(
    page_title="Conciliador: Almoxarifado x SIAFI",
    page_icon="📊",
    layout="wide"
)

# Inicializa a memória do Streamlit
if 'dados_processados' not in st.session_state:
    st.session_state.dados_processados = False
if 'dados_ug' not in st.session_state:
    st.session_state.dados_ug = {}
if 'logs' not in st.session_state:
    st.session_state.logs = []

# Oculta marcas do Streamlit e o menu lateral automático para um visual mais limpo
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
# FUNÇÕES E CLASSES (BASTIDORES)
# ==========================================
def limpar_valor(v):
    if v is None or pd.isna(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    v = str(v).replace('"', '').replace("'", "").strip()
    if re.search(r',\d{1,2}$', v): v = v.replace('.', '').replace(',', '.')
    elif re.search(r'\.\d{1,2}$', v): v = v.replace(',', '')
    try: return float(re.sub(r'[^\d.-]', '', v))
    except: return 0.0

def formatar_real(valor):
    sinal = "-" if valor < -0.001 else ""
    return f"{sinal}{abs(valor):,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

class PDF_Report(FPDF):
    def header(self):
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de Conferência: Almoxarifado x SIAFI', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(5)
    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='C')

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================
st.title("📊 Conciliador: Almoxarifado x SIAFI")

with st.expander("📘 GUIA DE USO (Clique para abrir)", expanded=False):
    st.markdown("📌 **Orientações de Uso**")
    st.markdown("""
    1. Anexe **o relatório SIAFI em excel** e todos os **arquivos PDF correspondentes** de uma só vez na área abaixo. (O nome dos arquivos deverá ser o código da UG correspondente)
    2. O sistema fará a leitura inicial. **Poderá corrigir valores divergentes manualmente e as edições serão registadas no PDF!**
    """)

# Área de Upload Unificada
uploaded_files = st.file_uploader(
    "📂 Arraste a Planilha e os PDFs para esta área", 
    accept_multiple_files=True,
    type=['pdf', 'xlsx', 'xls']
)

# ==========================================
# ETAPA 1: PROCESSAMENTO DE DADOS
# ==========================================
if st.button("🚀 Gerar relatório", use_container_width=True, type="primary"):
    
    if not uploaded_files:
        st.warning("⚠️ Por favor, insira seus arquivos para que possamos realizar a conciliação.")
    else:
        progresso = st.progress(0)
        status_text = st.empty()
        
        # Separa os PDFs e pega o primeiro arquivo Excel encontrado
        pdfs = {f.name: f for f in uploaded_files if f.name.lower().endswith('.pdf')}
        excel_files = [f for f in uploaded_files if f.name.lower().endswith(('.xlsx', '.xls'))]
        
        pares = []
        logs = []

        if not excel_files:
            st.error("❌ Nenhuma planilha Excel (.xlsx ou .xls) foi encontrada no upload.")
            st.stop()
        else:
            planilha_mestre = excel_files[0] 
            
            try:
                planilha_mestre.seek(0)
                xls = pd.ExcelFile(planilha_mestre)
                nome_abas = xls.sheet_names
                
                for aba in nome_abas:
                    match = re.search(r'(\d+)', aba)
                    if match:
                        ug = match.group(1)
                        pdf_match = next((f for n, f in pdfs.items() if n.startswith(ug)), None)
                        
                        if pdf_match:
                            pares.append({'ug': ug, 'nome_aba': aba, 'pdf': pdf_match})
                        else:
                            logs.append(f"⚠️ Aba '{aba}' (UG {ug}): Planilha encontrada, mas falta o PDF correspondente.")
                    else:
                        logs.append(f"ℹ️ Aba '{aba}' ignorada: Não foi encontrado um número de UG no nome da aba.")
                        
            except Exception as e:
                st.error(f"❌ Erro ao ler a estrutura da planilha Excel: {e}")
                st.stop()
        
        if not pares:
            st.error("❌ Nenhum par completo (Aba do Excel + PDF) foi identificado.")
            st.stop()

        dados_ug = {}

        for idx, par in enumerate(pares):
            ug = par['ug']
            nome_aba = par['nome_aba']
            status_text.text(f"Processando Aba '{nome_aba}' (UG: {ug})...")
            
            # === 1. LEITURA EXCEL ===
            df_padrao = pd.DataFrame()
            try:
                planilha_mestre.seek(0)
                df_raw = pd.read_excel(planilha_mestre, sheet_name=nome_aba, header=None)
                
                idx_cabecalho = df_raw[df_raw.apply(lambda r: r.astype(str).str.contains('Conta Corrente', case=False).any(), axis=1)].index
                
                if not idx_cabecalho.empty:
                    df_raw.columns = df_raw.iloc[idx_cabecalho[0]]
                    df = df_raw.iloc[idx_cabecalho[0]+1:].reset_index(drop=True)
                    
                    df['Conta_Corrente'] = df.iloc[:, 0].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.zfill(2)
                    df['Valor_Limpo'] = df.iloc[:, 4].apply(limpar_valor)
                    df['Descricao_Excel'] = "Conta Corrente " + df['Conta_Corrente']
                    
                    df = df[df['Conta_Corrente'] != 'NAN']
                    df = df[df['Conta_Corrente'].str.isnumeric()]
                    
                    df_padrao = df.groupby('Conta_Corrente').agg({'Valor_Limpo': 'sum', 'Descricao_Excel': 'first'}).reset_index()
                    df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']
                else:
                    logs.append(f"⚠️ Aba '{nome_aba}' (UG {ug}): Cabeçalho 'Conta Corrente' não encontrado.")
            except Exception as e:
                logs.append(f"❌ Erro na leitura da Aba '{nome_aba}' (UG {ug}): {e}")

            # === 2. LEITURA PDF ===
            df_pdf_final = pd.DataFrame()
            dados_pdf = []
            try:
                par['pdf'].seek(0)
                pdf_bytes = par['pdf'].read()
                
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as p_doc:
                    for page in p_doc.pages:
                        txt = page.extract_text()
                        is_ocr = False
                        
                        tem_dados_validos = False
                        if txt and re.search(r'\d{1,3}(?:[.,]\d{3})*[.,]\d{2}', txt):
                            tem_dados_validos = True
                        
                        if not txt or not tem_dados_validos or len(txt) < 50:
                            is_ocr = True
                            try:
                                imagens = convert_from_bytes(pdf_bytes, first_page=page.page_number, last_page=page.page_number, dpi=300)
                                if imagens:
                                    img = imagens[0]
                                    try:
                                        osd = pytesseract.image_to_osd(img, output_type=Output.DICT)
                                        if osd['rotate'] != 0:
                                            img = img.rotate(-osd['rotate'], expand=True)
                                    except: pass
                                    txt = pytesseract.image_to_string(img, lang='por', config='--psm 6')
                            except Exception: pass

                        if not txt: continue

                        for line in txt.split('\n'):
                            if re.match(r'^"?\d+"?\s*[-]?\s*[A-Za-z]', line) or re.match(r'^"?\d+', line):
                                vals = []
                                if is_ocr:
                                    vals_raw = re.findall(r'([\d\.\s]+,\d{2})', line)
                                    vals = [v.replace(' ', '') for v in vals_raw]
                                else:
                                    vals = re.findall(r'([0-9]{1,3}(?:[.,][0-9]{3})*[.,]\d{2})', line)
                                
                                if len(vals) >= 1:
                                    chave_match = re.match(r'^"?(\d+)', line)
                                    if chave_match:
                                        conta_contabil_completa = chave_match.group(1)
                                        chave_vinculo = str(conta_contabil_completa[-2:]).zfill(2)
                                        saldo_atual_pdf = limpar_valor(vals[-1])
                                        
                                        dados_pdf.append({'Chave_Vinculo': chave_vinculo, 'Saldo_PDF': saldo_atual_pdf})
                
                if dados_pdf:
                    df_pdf_final = pd.DataFrame(dados_pdf).groupby('Chave_Vinculo')['Saldo_PDF'].sum().reset_index()
            except Exception as e:
                logs.append(f"❌ Erro Leitura PDF UG {ug}: {e}")

            # === 3. PREPARAÇÃO PARA MEMÓRIA ===
            if df_padrao.empty: df_padrao = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa'])
            if df_pdf_final.empty: df_pdf_final = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_PDF'])

            df_padrao['Chave_Vinculo'] = df_padrao['Chave_Vinculo'].astype(str)
            df_pdf_final['Chave_Vinculo'] = df_pdf_final['Chave_Vinculo'].astype(str)

            cruzamento_temp = pd.merge(df_pdf_final, df_padrao, on='Chave_Vinculo', how='outer').fillna(0)
            chaves_com_erro = cruzamento_temp[abs(cruzamento_temp['Saldo_PDF'] - cruzamento_temp['Saldo_Excel']) > 0.05]['Chave_Vinculo'].tolist()

            dados_ug[ug] = {
                'nome_aba': nome_aba,
                'df_excel': df_padrao,
                'df_pdf': df_pdf_final.copy(),
                'df_pdf_orig': df_pdf_final.copy(), # Memória Fotográfica
                'chaves_com_erro': chaves_com_erro,
                'erro_original': len(chaves_com_erro) > 0
            }

            progresso.progress((idx + 1) / len(pares))

        st.session_state.dados_ug = dados_ug
        st.session_state.logs = logs
        st.session_state.dados_processados = True
        progresso.empty()
        status_text.empty()

# ==========================================
# ETAPA 2: REVISÃO CIRÚRGICA E PDF FINAL
# ==========================================
if st.session_state.get('dados_processados'):
    st.markdown("---")
    st.subheader("🔍 Resultados da Análise & Revisão")
    st.info("💡 **Ação Cirúrgica:** Apenas as contas com divergências permitem edição. As edições ficarão registadas no PDF Final.")

    pdf_out = PDF_Report()
    pdf_out.add_page()
    dados_ug = st.session_state.dados_ug

    for ug, info in dados_ug.items():
        nome_aba = info['nome_aba']
        
        # Lógica de atualização a partir dos inputs do utilizador
        if info['erro_original']:
            for chave in info['chaves_com_erro']:
                key_input = f"edit_almox_{ug}_{chave}"
                if key_input in st.session_state:
                    novo_val = st.session_state[key_input]
                    if chave in info['df_pdf']['Chave_Vinculo'].values:
                        info['df_pdf'].loc[info['df_pdf']['Chave_Vinculo'] == chave, 'Saldo_PDF'] = novo_val
                    else:
                        novo_df = pd.DataFrame([{'Chave_Vinculo': chave, 'Saldo_PDF': novo_val}])
                        info['df_pdf'] = pd.concat([info['df_pdf'], novo_df], ignore_index=True)

        # Recálculo Final
        final = pd.merge(info['df_pdf'], info['df_excel'], on='Chave_Vinculo', how='outer').fillna(0)
        final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if x['Descricao_Completa'] != 0 else f"Conta {x['Chave_Vinculo']} (Sem no Excel)", axis=1)
        final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
        
        soma_pdf = final['Saldo_PDF'].sum()
        soma_excel = final['Saldo_Excel'].sum()
        dif_total = soma_pdf - soma_excel

        # Exibição na Interface
        with st.container():
            st.markdown(f"### 🏢 Unidade Gestora: {ug} (Aba: {nome_aba})")
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Saldo relatório", f"R$ {formatar_real(soma_pdf)}")
            col2.metric("Total Saldo SIAFI", f"R$ {formatar_real(soma_excel)}")
            col3.metric("Diferença Total", f"R$ {formatar_real(dif_total)}", delta_color="inverse" if abs(dif_total) > 0.05 else "normal")
            
            tem_erro_atual = abs(dif_total) > 0.05
            titulo_expander = f"⚠️ Contas com Divergência" if tem_erro_atual else ("✅ Corrigido Manualmente" if info['erro_original'] else "✅ Tudo certo! Nenhuma divergência.")
            
            with st.expander(titulo_expander, expanded=tem_erro_atual):
                if info['chaves_com_erro']:
                    df_view = final[final['Chave_Vinculo'].isin(info['chaves_com_erro'])].copy()
                    df_view = df_view[['Chave_Vinculo', 'Descricao', 'Saldo_PDF', 'Saldo_Excel', 'Diferenca']]
                    df_view.rename(columns={'Saldo_PDF': 'Saldo relatório', 'Saldo_Excel': 'Saldo SIAFI'}, inplace=True)
                    
                    st.dataframe(df_view.style.format({
                        "Saldo relatório": lambda x: f"R$ {formatar_real(x)}", 
                        "Saldo SIAFI": lambda x: f"R$ {formatar_real(x)}", 
                        "Diferença": lambda x: f"R$ {formatar_real(x)}"
                    }), use_container_width=True)

                    st.markdown("**✏️ Correção Direta por Conta:**")
                    cols = st.columns(3)
                    for idx, chave in enumerate(info['chaves_com_erro']):
                        val_atual = final.loc[final['Chave_Vinculo'] == chave, 'Saldo_PDF'].sum()
                        with cols[idx % 3]:
                            st.number_input(f"Conta {chave}", value=float(val_atual), step=100.0, key=f"edit_almox_{ug}_{chave}")
                elif not tem_erro_atual:
                    st.success("Tudo certo! Nenhuma divergência encontrada entre as contas do relatório e SIAFI.")

            st.markdown("---")

            # ==========================================
            # GERAÇÃO DO PDF PARA A UG
            # ==========================================
            pdf_out.set_font("helvetica", 'B', 11)
            pdf_out.set_fill_color(240, 240, 240)
            pdf_out.cell(0, 10, text=f"Unidade Gestora: {ug} (Aba: {nome_aba})", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
            
            mask_mostrar = (abs(final['Diferenca']) > 0.05) | (final['Chave_Vinculo'].isin(info['chaves_com_erro']))
            itens_para_mostrar = final[mask_mostrar].copy()

            if not itens_para_mostrar.empty:
                pdf_out.set_font("helvetica", 'B', 9)
                pdf_out.set_fill_color(255, 200, 200)
                pdf_out.cell(15, 8, "Chave", 1, fill=True)
                pdf_out.cell(85, 8, "Descrição (Conta Corrente)", 1, fill=True)
                pdf_out.cell(30, 8, "Saldo relatório", 1, fill=True, align='C')
                pdf_out.cell(30, 8, "Saldo SIAFI", 1, fill=True, align='C')
                pdf_out.cell(30, 8, "Diferença", 1, fill=True, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                pdf_out.set_font("helvetica", '', 8)
                alertas_auditoria = []

                for _, row in itens_para_mostrar.iterrows():
                    chave = row['Chave_Vinculo']
                    # Busca o valor original da fotografia
                    if chave in info['df_pdf_orig']['Chave_Vinculo'].values:
                        val_original = info['df_pdf_orig'].loc[info['df_pdf_orig']['Chave_Vinculo'] == chave, 'Saldo_PDF'].sum()
                    else:
                        val_original = 0.0
                        
                    editado = abs(row['Saldo_PDF'] - val_original) > 0.01

                    str_saldo_relatorio = formatar_real(row['Saldo_PDF']) + (" *" if editado else "")
                    
                    pdf_out.cell(15, 7, str(chave), 1)
                    pdf_out.cell(85, 7, str(row['Descricao'])[:48], 1)
                    pdf_out.cell(30, 7, str_saldo_relatorio, 1, align='R')
                    pdf_out.cell(30, 7, formatar_real(row['Saldo_Excel']), 1, align='R')
                    if abs(row['Diferenca']) > 0.05: pdf_out.set_text_color(200, 0, 0)
                    pdf_out.cell(30, 7, formatar_real(row['Diferenca']), 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)

                    if editado:
                        alertas_auditoria.append(f"* ALERTA: O valor da Conta {chave} foi alterado manualmente pelo utilizador. (Valor original lido do relatório: R$ {formatar_real(val_original)})")

                # Impressão dos alertas de auditoria no PDF
                if alertas_auditoria:
                    pdf_out.set_font("helvetica", 'I', 7)
                    pdf_out.set_text_color(180, 0, 0)
                    for alerta in alertas_auditoria:
                        pdf_out.cell(0, 5, alerta, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)
            else:
                pdf_out.set_font("helvetica", 'I', 9)
                pdf_out.cell(0, 8, "Nenhuma divergência encontrada nas contas.", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf_out.ln(2)
            pdf_out.set_font("helvetica", 'B', 9)
            pdf_out.set_fill_color(220, 230, 241)
            pdf_out.cell(100, 8, "TOTAIS", 1, fill=True)
            pdf_out.cell(30, 8, formatar_real(soma_pdf), 1, fill=True, align='R')
            pdf_out.cell(30, 8, formatar_real(soma_excel), 1, fill=True, align='R')
            if abs(dif_total) > 0.05: pdf_out.set_text_color(200, 0, 0)
            pdf_out.cell(30, 8, formatar_real(dif_total), 1, fill=True, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf_out.set_text_color(0, 0, 0)
            pdf_out.ln(5)

    if st.session_state.logs:
        with st.expander("⚠️ Avisos do Sistema (Arquivos e Abas Ignorados)"):
            for log in st.session_state.logs: st.write(log)
    
    try:
        pdf_bytes = bytes(pdf_out.output())
        st.download_button(
            label="📄 BAIXAR RELATÓRIO DE CONCILIAÇÃO (.PDF)", 
            data=pdf_bytes, 
            file_name="RELATORIO_ALMOXARIFADO_SIAFI.pdf", 
            mime="application/pdf", 
            type="primary", 
            use_container_width=True
        )
    except Exception as e:
        st.error(f"Erro ao gerar o download: {e}")
