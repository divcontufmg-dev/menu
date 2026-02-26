import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io
import os

# ==========================================
# CONFIGURAÇÃO INICIAL
# ==========================================
st.set_page_config(
    page_title="Conciliador: Acervo Bibliográfico",
    page_icon="📚",
    layout="wide"
)

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

# Botão para retornar à tela inicial solto no topo da tela
st.page_link("Menu_principal.py", label="⬅️ Voltar ao Menu Inicial")

# ==========================================
# FUNÇÕES E CLASSES (BASTIDORES)
# ==========================================
def formatar_real(valor):
    sinal = "-" if valor < -0.001 else ""
    return f"{sinal}{abs(valor):,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

# FUNÇÃO EXCLUSIVA PARA O EXCEL
def limpar_valor_excel(v):
    if pd.isna(v) or v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    
    v_str = str(v).strip()
    if v_str == '': return 0.0
    
    v_str = re.sub(r'[^\d\.,\-]', '', v_str)
    
    if ',' in v_str and '.' in v_str:
        if v_str.rfind(',') > v_str.rfind('.'):
            v_str = v_str.replace('.', '').replace(',', '.')
        else:
            v_str = v_str.replace(',', '')
    elif ',' in v_str:
        v_str = v_str.replace(',', '.')
        
    try:
        return float(v_str)
    except:
        return 0.0

# FUNÇÃO EXCLUSIVA PARA O PDF
def limpar_valor_pdf(v):
    v = re.sub(r'[^\d\.,]', '', str(v))
    if not v: return 0.0
    
    if len(v) >= 3 and v[-3] in ['.', ',']:
        inteiro = v[:-3].replace('.', '').replace(',', '')
        decimal = v[-2:]
        return float(f"{inteiro}.{decimal}")
    elif len(v) >= 2 and v[-2] in ['.', ',']:
        inteiro = v[:-2].replace('.', '').replace(',', '')
        decimal = v[-1:]
        return float(f"{inteiro}.{decimal}")
    else:
        return float(v.replace('.', '').replace(',', '.'))

# MOTOR DE EXTRAÇÃO CIRÚRGICO - AGORA COM VARREDURA TOTAL
def extrair_valor_pdf(pdf_bytes, texto_busca, texto_abrev=None, is_dep=False):
    texto_completo = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"
    except Exception:
        pass
    
    linhas = texto_completo.split('\n')
    valor_final = 0.0
    
    for i, line in enumerate(linhas):
        line_clean = line.strip().replace('"', '') 
        if not line_clean: continue
        
        condicao_extenso = line_clean.upper().startswith(texto_busca.upper())
        condicao_abrev = texto_abrev and line_clean.upper().startswith(texto_abrev.upper())
        
        if condicao_extenso or condicao_abrev:
            bloco_texto = line_clean
            
            # Reconstrói a tabela caso o PDF tenha vindo quebrado em várias linhas
            for j in range(i + 1, min(i + 30, len(linhas))):
                proxima = linhas[j].strip().replace('"', '')
                if not proxima: continue
                
                # Critério de parada de leitura (Cobre meses, rodapé, paginação e sistema)
                if not is_dep:
                    if re.match(r'^(Janeiro|Fevereiro|Março|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro|Jan\.?|Fev\.?|Mar\.?|Abr\.?|Mai\.?|Jun\.?|Jul\.?|Ago\.?|Set\.?|Out\.?|Nov\.?|Dez\.?|TOTAL|Pag\.|Página|Pergamum|Sistema|Emissão|Data)', proxima, re.IGNORECASE):
                        break
                else:
                    if re.match(r'^(\d{2}/\d{4}|TOTAL|Pag\.|Página|Pergamum|Sistema|Emissão|Data)', proxima, re.IGNORECASE):
                        break
                        
                bloco_texto += " " + proxima
                
            # CORREÇÃO CIRÚRGICA DE ESPAÇOS MILHARES
            bloco_texto = re.sub(r'(\.\d{3})\s+(?=\d{3}[.,]\d{2}(?!\d))', r'\1', bloco_texto)
            
            matches = re.findall(r'[\d\.,]+', bloco_texto)
            if len(matches) >= 1:
                novo_valor = limpar_valor_pdf(matches[-1])
                # Grava o valor apenas se for válido. Como varre até o fim, pega a última linha de TOTAL (o Saldo Final real)
                if novo_valor != 0.0 or valor_final == 0.0:
                    valor_final = novo_valor 
                
    return valor_final

class PDF_Report(FPDF):
    def header(self):
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de Conferência: Acervo Bibliográfico x Pergamum', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(5)
    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='C')

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================
st.title("📚 Conciliador: Acervo Bibliográfico")

with st.expander("📘 GUIA DE USO (Clique para abrir)", expanded=False):
    st.markdown("📌 **Orientações de Uso**")
    st.markdown("""
    1. Selecione o **Mês** e o **Ano** exatos que deseja conciliar.
    2. Anexe a **Planilha Excel (Conf. RMB)** e todos os **arquivos PDF (Pergamum)** de uma só vez.
    3. **Nomenclatura dos PDFs:** - **Acervo:** Número da UG (ex: `153289.pdf`). *Se houver mais de um, use `a`, `a1`, `a2` no final (ex: `153289a.pdf`, `153289a2.pdf`).*
       - **Depreciação:** Número da UG com 'd' no final (ex: `153289d.pdf`). *Se houver mais de um, use `d2`, `d3` (ex: `153289d2.pdf`).*
    4. O sistema somará todos os relatórios da mesma categoria automaticamente. Clique em "Iniciar Conciliação".
    """)

# Seleção de Data
col_mes, col_ano = st.columns(2)
meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

with col_mes:
    mes_selecionado = st.selectbox("Selecione o Mês:", meses)
with col_ano:
    ano_selecionado = st.number_input("Digite o Ano:", min_value=2000, max_value=2100, value=2026, step=1)

idx_mes = meses.index(mes_selecionado)
mes_num = f"{idx_mes + 1:02d}"

# ALTERAÇÃO DE LÓGICA: O Acervo agora caça o TOTAL Final, a Depreciação caça o Mês/Ano
texto_busca_acervo = "TOTAL"           
texto_abrev_acervo = None      
texto_busca_dep = f"{mes_num}/{ano_selecionado}" 

# Área de Upload Unificada
uploaded_files = st.file_uploader(
    "📂 Arraste a Planilha do Tesouro e os PDFs do Pergamum para esta área", 
    accept_multiple_files=True,
    type=['pdf', 'xlsx', 'xls', 'csv']
)

# ==========================================
# EXECUÇÃO DO SISTEMA
# ==========================================
if st.button("🚀 Iniciar Conciliação", use_container_width=True, type="primary"):
    
    if not uploaded_files:
        st.warning("⚠️ Por favor, insira seus arquivos para que possamos realizar a conciliação.")
    else:
        progresso = st.progress(0)
        status_text = st.empty()
        
        pdfs = {f.name.lower(): f for f in uploaded_files if f.name.lower().endswith('.pdf')}
        excel_files = [f for f in uploaded_files if f.name.lower().endswith(('.xlsx', '.xls', '.csv'))]
        
        if not excel_files:
            st.error("❌ A planilha base em Excel não foi encontrada no upload.")
            st.stop()
            
        planilha_mestre = excel_files[0]
        dados_ug = {}
        logs = []

        status_text.text("Lendo os dados da Planilha Base...")
        try:
            planilha_mestre.seek(0)
            if planilha_mestre.name.lower().endswith('.csv'):
                df = pd.read_csv(planilha_mestre)
            else:
                df = pd.read_excel(planilha_mestre, header=None)
            
            for idx, row in df.iterrows():
                val0 = str(row[0]).strip()
                if val0.isdigit() and len(val0) >= 5:
                    ug = val0
                    nome = str(row[1]).strip()
                    saldo_acervo = limpar_valor_excel(row[2]) if len(row) > 2 else 0.0
                    saldo_dep = limpar_valor_excel(row[3]) if len(row) > 3 else 0.0
                    
                    dados_ug[ug] = {
                        'nome': nome,
                        'ex_acervo': saldo_acervo,
                        'ex_dep': abs(saldo_dep), 
                        'pdf_acervo': 0.0,
                        'pdf_dep': 0.0,
                        'achou_pdf_acervo': False,
                        'achou_pdf_dep': False,
                        'arquivos_acervo_somados': 0,
                        'arquivos_dep_somados': 0
                    }
        except Exception as e:
            st.error(f"❌ Erro ao ler a estrutura da planilha: {e}")
            st.stop()

        status_text.text("Processando e cruzando os documentos PDF...")
        total_ugs = len(dados_ug)
        if total_ugs == 0:
            st.warning("⚠️ Nenhuma Unidade Gestora (UG) foi encontrada na primeira coluna da planilha.")
            st.stop()

        for i, (ug, info) in enumerate(dados_ug.items()):
            
            # 1. Busca Múltiplos PDFs de Acervo
            padrao_acervo = re.compile(rf"^{ug}(a\d*)?\.pdf$")
            achou_algum_acervo = False
            
            for nome_arquivo, arquivo_obj in pdfs.items():
                if padrao_acervo.match(nome_arquivo):
                    achou_algum_acervo = True
                    info['arquivos_acervo_somados'] += 1
                    arquivo_obj.seek(0)
                    valor_extraido = extrair_valor_pdf(
                        arquivo_obj.read(), 
                        texto_busca_acervo, 
                        texto_abrev=texto_abrev_acervo, 
                        is_dep=False
                    )
                    info['pdf_acervo'] += valor_extraido
            
            if achou_algum_acervo:
                info['achou_pdf_acervo'] = True
            else:
                logs.append(f"⚠️ UG {ug}: Faltou o PDF do Acervo (esperado {ug}.pdf).")

            # 2. Busca Múltiplos PDFs de Depreciação
            padrao_dep = re.compile(rf"^{ug}d\d*\.pdf$")
            achou_algum_dep = False
            
            for nome_arquivo, arquivo_obj in pdfs.items():
                if padrao_dep.match(nome_arquivo):
                    achou_algum_dep = True
                    info['arquivos_dep_somados'] += 1
                    arquivo_obj.seek(0)
                    valor_extraido = extrair_valor_pdf(
                        arquivo_obj.read(), 
                        texto_busca_dep, 
                        texto_abrev=None, 
                        is_dep=True
                    )
                    info['pdf_dep'] += valor_extraido
            
            if achou_algum_dep:
                info['achou_pdf_dep'] = True
            else:
                logs.append(f"⚠️ UG {ug}: Faltou o PDF de Depreciação (esperado {ug}d.pdf).")
                
            progresso.progress((i + 1) / total_ugs)

        # ==========================================
        # GERAÇÃO DO RELATÓRIO E EXIBIÇÃO
        # ==========================================
        pdf_out = PDF_Report()
        pdf_out.add_page()
        
        st.markdown("---")
        st.subheader("🔍 Resultados da Análise")
        
        total_ex_acervo = total_ex_dep = total_pdf_acervo = total_pdf_dep = 0.0

        for ug, info in dados_ug.items():
            dif_acervo = info['pdf_acervo'] - info['ex_acervo']
            dif_dep = info['pdf_dep'] - info['ex_dep']
            
            total_ex_acervo += info['ex_acervo']
            total_pdf_acervo += info['pdf_acervo']
            total_ex_dep += info['ex_dep']
            total_pdf_dep += info['pdf_dep']
            
            tem_erro = abs(dif_acervo) > 0.05 or abs(dif_dep) > 0.05
            
            texto_ug = f"Unidade Gestora: {ug} - {info['nome'][:50]}"
            avisos_soma = []
            if info['arquivos_acervo_somados'] > 1:
                avisos_soma.append(f"{info['arquivos_acervo_somados']} Acervos")
            if info['arquivos_dep_somados'] > 1:
                avisos_soma.append(f"{info['arquivos_dep_somados']} Depreciações")
            
            if avisos_soma:
                texto_ug += f" (+{' e '.join(avisos_soma)})"
                
            pdf_out.set_font("helvetica", 'B', 10)
            pdf_out.set_fill_color(240, 240, 240)
            pdf_out.cell(0, 8, text=texto_ug, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
            
            pdf_out.set_font("helvetica", 'B', 8)
            pdf_out.set_fill_color(220, 230, 241)
            pdf_out.cell(46, 7, "Conta", 1, fill=True)
            pdf_out.cell(48, 7, "Saldo PDF (Pergamum)", 1, fill=True, align='C')
            pdf_out.cell(48, 7, "Saldo Excel (SIAFI)", 1, fill=True, align='C')
            pdf_out.cell(48, 7, "Diferença", 1, fill=True, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            pdf_out.set_font("helvetica", '', 8)
            
            pdf_out.cell(46, 7, "Acervo Bibliográfico", 1)
            pdf_out.cell(48, 7, f"R$ {formatar_real(info['pdf_acervo'])}", 1, align='R')
            pdf_out.cell(48, 7, f"R$ {formatar_real(info['ex_acervo'])}", 1, align='R')
            if abs(dif_acervo) > 0.05: pdf_out.set_text_color(200, 0, 0)
            pdf_out.cell(48, 7, f"R$ {formatar_real(dif_acervo)}", 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf_out.set_text_color(0, 0, 0)
            
            pdf_out.cell(46, 7, "Depreciação Acumulada", 1)
            pdf_out.cell(48, 7, f"R$ {formatar_real(info['pdf_dep'])}", 1, align='R')
            pdf_out.cell(48, 7, f"R$ {formatar_real(info['ex_dep'])}", 1, align='R')
            if abs(dif_dep) > 0.05: pdf_out.set_text_color(200, 0, 0)
            pdf_out.cell(48, 7, f"R$ {formatar_real(dif_dep)}", 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf_out.set_text_color(0, 0, 0)
            
            pdf_out.ln(5)
            
            if tem_erro:
                aviso_extra = f" (⚠️ {' e '.join(avisos_soma)} somados)" if avisos_soma else ""
                with st.expander(f"⚠️ UG {ug}: Divergências Encontradas {aviso_extra}", expanded=True):
                    df_view = pd.DataFrame([
                        {"Conta": "Acervo Bibliográfico", "PDF": info['pdf_acervo'], "Excel": info['ex_acervo'], "Diferença": dif_acervo},
                        {"Conta": "Depreciação Acumulada", "PDF": info['pdf_dep'], "Excel": info['ex_dep'], "Diferença": dif_dep}
                    ])
                    st.dataframe(df_view.style.format({"PDF": "R$ {:,.2f}", "Excel": "R$ {:,.2f}", "Diferença": "R$ {:,.2f}"}))

        dif_total_acervo = total_pdf_acervo - total_ex_acervo
        dif_total_dep = total_pdf_dep - total_ex_dep
        
        st.markdown("### Resumo Geral da Conciliação")
        c1, c2, c3 = st.columns(3)
        c1.metric("Diferença Total (Acervo)", f"R$ {dif_total_acervo:,.2f}", delta_color="inverse" if abs(dif_total_acervo) > 0.05 else "normal")
        c2.metric("Diferença Total (Depreciação)", f"R$ {dif_total_dep:,.2f}", delta_color="inverse" if abs(dif_total_dep) > 0.05 else "normal")
        
        status_text.success("Conciliação concluída com sucesso!")
        progresso.empty()
        
        if logs:
            with st.expander("⚠️ Avisos de Ficheiros Ausentes", expanded=False):
                for log in logs: st.write(log)
        
        try:
            pdf_bytes = bytes(pdf_out.output())
            st.download_button(
                label="📄 BAIXAR RELATÓRIO DE CONCILIAÇÃO (.PDF)", 
                data=pdf_bytes, 
                file_name=f"RELATORIO_ACERVO_BIBLIOGRAFICO_{mes_selecionado}_{ano_selecionado}.pdf", 
                mime="application/pdf", 
                type="primary", 
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Erro ao gerar o download: {e}")
