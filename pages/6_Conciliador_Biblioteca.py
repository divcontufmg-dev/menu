import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io

st.set_page_config(
    page_title="Conciliador: Acervo Bibliográfico",
    page_icon="📚",
    layout="wide"
)

if 'dados_processados' not in st.session_state:
    st.session_state.dados_processados = False
if 'dados_ug' not in st.session_state:
    st.session_state.dados_ug = {}
if 'logs' not in st.session_state:
    st.session_state.logs = []
if 'arquivos_lidos' not in st.session_state:
    st.session_state.arquivos_lidos = set()

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


def formatar_real(valor):
    sinal = "-" if valor < -0.001 else ""
    return f"{sinal}{abs(valor):,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

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

def limpar_valor_pdf(v):
    v = re.sub(r'[^\d\.,]', '', str(v)).rstrip('.,')
    if not any(c.isdigit() for c in v): return 0.0
    if len(v) >= 3 and v[-3] in ['.', ',']:
        inteiro = v[:-3].replace('.', '').replace(',', '')
        decimal = v[-2:]
        try:
            return float(f"{inteiro}.{decimal}")
        except:
            return 0.0
    return 0.0

def extrair_valor_pdf(pdf_bytes, texto_busca, texto_abrev=None, is_dep=False):
    texto_completo = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"
    except Exception:
        pass
    
    linhas = texto_completo.split('\n')
    valores_encontrados = []
    encontrou_mes = False
    
    for i, line in enumerate(linhas):
        line_clean = line.strip().replace('"', '') 
        if not line_clean: continue
        
        condicao_mes = False
        condicao_total = False
        
        if not is_dep:
            padrao_mes = rf'^[\d\s\W]*({texto_busca.upper()}'
            if texto_abrev:
                padrao_mes += rf'|{texto_abrev.upper()}'
            padrao_mes += r')\b'
            
            if re.search(padrao_mes, line_clean.upper()):
                condicao_mes = True
            elif encontrou_mes and re.match(r'^[\d\s\W]*TOTAL\b', line_clean.upper()):
                condicao_total = True
        else:
            if line_clean.upper().startswith(texto_busca.upper()):
                condicao_mes = True
                
        if condicao_mes or condicao_total:
            if condicao_mes:
                encontrou_mes = True 
                
            bloco_texto = line_clean
            
            for j in range(i + 1, min(i + 30, len(linhas))):
                proxima = linhas[j].strip().replace('"', '')
                if not proxima: continue
                if not is_dep:
                    if re.match(r'^(Janeiro|Fevereiro|Março|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro|Jan\.?|Fev\.?|Mar\.?|Abr\.?|Mai\.?|Jun\.?|Jul\.?|Ago\.?|Set\.?|Out\.?|Nov\.?|Dez\.?|TOTAL|Pag\.|Página|Pergamum|Sistema|Emissão|Data)', proxima, re.IGNORECASE):
                        break
                else:
                    if re.match(r'^(\d{2}/\d{4}|TOTAL|Pag\.|Página|Pergamum|Sistema|Emissão|Data)', proxima, re.IGNORECASE):
                        break
                bloco_texto += " " + proxima
                
            bloco_texto = re.sub(r'(\.\d{3})\s+(?=\d{3}[.,]\d{2}(?!\d))', r'\1', bloco_texto)
            matches = [m for m in re.findall(r'[\d\.,]+', bloco_texto) if any(c.isdigit() for c in m)]
            
            for m in reversed(matches):
                v_clean = re.sub(r'[^\d\.,]', '', m).rstrip('.,')
                if len(v_clean) >= 3 and v_clean[-3] in ['.', ',']:
                    valor_real = limpar_valor_pdf(v_clean)
                    valores_encontrados.append(valor_real)
                    break 
                    
    if not valores_encontrados:
        return 0.0
    if is_dep:
        return valores_encontrados[-1]
    else:
        return max(valores_encontrados)

class PDF_Report(FPDF):
    def header(self):
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de Conferência Biblioteca', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(5)
    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='C')


st.title("📚 Conciliador: Acervo Bibliográfico")

with st.expander("📘 GUIA DE USO (Clique para abrir)", expanded=False):
    st.markdown("📌 **Orientações de Uso**")
    st.markdown("""
    1. Selecione o **Mês** e o **Ano** exatos a serem conciliados.
    2. Anexe a **Planilha do Tesouro** e os **Relatórios PDF** emitidos.
    3. **Identificação dos Arquivos:**
       - **Acervo:** Renomeie o PDF com o número da UG (ex: `153289.pdf`). Se existirem complementos, adicione letras (ex: `153289a.pdf`).
       - **Depreciação:** Renomeie com o número da UG seguido da letra 'd' (ex: `153289d.pdf`).
    4. O sistema irá agrupar as informações. Caso encontre inconsistências, poderá corrigi-las manualmente.
    5. **Esqueceu de algum relatório?** Basta arrastar os PDFs esquecidos para a mesma caixa de anexo e clicar em atualizar!
    """)

col_mes, col_ano = st.columns(2)
meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
meses_abrev = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

with col_mes:
    mes_selecionado = st.selectbox("Selecione o Mês:", meses)
with col_ano:
    ano_selecionado = st.number_input("Digite o Ano:", min_value=2000, max_value=2100, value=2026, step=1)

idx_mes = meses.index(mes_selecionado)
mes_num = f"{idx_mes + 1:02d}"

texto_busca_acervo = mes_selecionado           
texto_abrev_acervo = meses_abrev[idx_mes]      
texto_busca_dep = f"{mes_num}/{ano_selecionado}" 

uploaded_files = st.file_uploader(
    "📂 Arraste os seus documentos (Planilha e PDFs) para esta área", 
    accept_multiple_files=True,
    type=['pdf', 'xlsx', 'xls', 'csv']
)

is_atualizacao = st.session_state.dados_processados
botao_label = "🔄 Inserir Ficheiros Esquecidos" if is_atualizacao else "🚀 Iniciar Conciliação"

if st.button(botao_label, use_container_width=True, type="primary"):
    if not uploaded_files and not is_atualizacao:
        st.warning("⚠️ Por favor, anexe a planilha e os relatórios para começar.")
    else:
        progresso = st.progress(0)
        status_text = st.empty()
        
        todos_pdfs_inseridos = {f.name.lower(): f for f in uploaded_files if f.name.lower().endswith('.pdf')}
        
        if not is_atualizacao:
            excel_files = [f for f in uploaded_files if f.name.lower().endswith(('.xlsx', '.xls', '.csv'))]
            if not excel_files:
                st.error("❌ A planilha contendo os dados do Tesouro não foi encontrada. Verifique os anexos.")
                st.stop()
                
            planilha_mestre = excel_files[0]
            dados_ug = {}
            logs = []
            pdfs_para_analise = todos_pdfs_inseridos 

            status_text.text("A preparar a base de dados da Planilha...")
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
                            'original_pdf_acervo': 0.0, 
                            'original_pdf_dep': 0.0,    
                            'arquivos_acervo_somados': 0,
                            'arquivos_dep_somados': 0,
                            'detalhes_acervo': {}, 
                            'detalhes_dep': {},
                            'erro_original_acervo': False,
                            'erro_original_dep': False
                        }
            except Exception as e:
                st.error(f"❌ Não foi possível ler a estrutura da planilha: {e}")
                st.stop()
        else:
            dados_ug = st.session_state.dados_ug
            logs = st.session_state.logs
            pdfs_para_analise = {nome: f for nome, f in todos_pdfs_inseridos.items() if nome not in st.session_state.arquivos_lidos}
            
            if not pdfs_para_analise:
                st.warning("⚠️ Não foram detetados relatórios novos na área de anexo. Insira os PDFs que estão em falta.")
                progresso.empty()
                status_text.empty()
                st.stop()

        status_text.text("A verificar e cruzar as informações dos Relatórios (PDF)...")
        total_ugs = len(dados_ug) if len(dados_ug) > 0 else 1
        adicionados_agora = 0

        for i, (ug, info) in enumerate(dados_ug.items()):
            padrao_acervo = re.compile(rf"^{ug}(a\d*)?\.pdf$")
            padrao_dep = re.compile(rf"^{ug}d\d*\.pdf$")
            
            for nome_arquivo, arquivo_obj in pdfs_para_analise.items():
                if padrao_acervo.match(nome_arquivo) and nome_arquivo not in info['detalhes_acervo']:
                    info['arquivos_acervo_somados'] += 1
                    arquivo_obj.seek(0)
                    valor_extraido = extrair_valor_pdf(arquivo_obj.read(), texto_busca_acervo, texto_abrev_acervo, False)
                    info['pdf_acervo'] += valor_extraido
                    info['detalhes_acervo'][nome_arquivo] = valor_extraido
                    info['original_pdf_acervo'] += valor_extraido
                    info['erro_original_acervo'] = abs(info['pdf_acervo'] - info['ex_acervo']) > 0.05
                    st.session_state.arquivos_lidos.add(nome_arquivo)
                    adicionados_agora += 1

                elif padrao_dep.match(nome_arquivo) and nome_arquivo not in info['detalhes_dep']:
                    info['arquivos_dep_somados'] += 1
                    arquivo_obj.seek(0)
                    valor_extraido = extrair_valor_pdf(arquivo_obj.read(), texto_busca_dep, None, True)
                    info['pdf_dep'] += valor_extraido
                    info['detalhes_dep'][nome_arquivo] = valor_extraido
                    info['original_pdf_dep'] += valor_extraido
                    info['erro_original_dep'] = abs(info['pdf_dep'] - info['ex_dep']) > 0.05
                    st.session_state.arquivos_lidos.add(nome_arquivo)
                    adicionados_agora += 1
                    
            if not is_atualizacao:
                if info['arquivos_acervo_somados'] == 0: logs.append(f"⚠️ UG {ug}: Relatório de Acervo encontra-se ausente.")
                if info['arquivos_dep_somados'] == 0: logs.append(f"⚠️ UG {ug}: Relatório de Depreciação encontra-se ausente.")
                
            progresso.progress((i + 1) / total_ugs)
            
        if is_atualizacao and adicionados_agora == 0:
            st.warning("⚠️ Os relatórios anexados não correspondem a nenhuma Unidade Gestora da conciliação em andamento.")
            progresso.empty()
            status_text.empty()
            st.stop()
        
        st.session_state.dados_ug = dados_ug
        st.session_state.logs = logs
        st.session_state.dados_processados = True
        progresso.empty()
        
        if is_atualizacao:
            st.success(f"✔️ {adicionados_agora} novo(s) relatório(s) inserido(s) com sucesso na conciliação!")
            st.rerun() 
        else:
            status_text.empty()


if st.session_state.get('dados_processados'):
    st.markdown("---")
    st.subheader("🔍 Painel de Conferência e Ajustes")
    st.info("💡 **Atenção:** As secções abaixo apresentam as Unidades Gestoras analisadas. Valores com divergência exibem um campo aberto para ajuste manual do utilizador.")

    dados_ug = st.session_state.dados_ug
    total_ex_acervo = total_ex_dep = total_pdf_acervo = total_pdf_dep = 0.0

    pdf_out = PDF_Report()
    pdf_out.add_page()

    for ug, info in dados_ug.items():
        if info['erro_original_acervo']:
            if info['detalhes_acervo']:
                soma = 0.0
                for arq in info['detalhes_acervo'].keys():
                    key = f"edit_ac_{ug}_{arq}"
                    if key in st.session_state: info['detalhes_acervo'][arq] = st.session_state[key]
                    soma += info['detalhes_acervo'][arq]
                info['pdf_acervo'] = soma
            else:
                key = f"edit_ac_{ug}_total"
                if key in st.session_state: info['pdf_acervo'] = st.session_state[key]

        if info['erro_original_dep']:
            if info['detalhes_dep']:
                soma = 0.0
                for arq in info['detalhes_dep'].keys():
                    key = f"edit_dp_{ug}_{arq}"
                    if key in st.session_state: info['detalhes_dep'][arq] = st.session_state[key]
                    soma += info['detalhes_dep'][arq]
                info['pdf_dep'] = soma
            else:
                key = f"edit_dp_{ug}_total"
                if key in st.session_state: info['pdf_dep'] = st.session_state[key]

        dif_acervo_final = info['pdf_acervo'] - info['ex_acervo']
        dif_dep_final = info['pdf_dep'] - info['ex_dep']
        
        total_ex_acervo += info['ex_acervo']
        total_pdf_acervo += info['pdf_acervo']
        total_ex_dep += info['ex_dep']
        total_pdf_dep += info['pdf_dep']

        editado_acervo = abs(info['pdf_acervo'] - info['original_pdf_acervo']) > 0.01
        editado_dep = abs(info['pdf_dep'] - info['original_pdf_dep']) > 0.01

        mostrar_expander = info['erro_original_acervo'] or info['erro_original_dep'] or info['arquivos_acervo_somados'] > 1 or info['arquivos_dep_somados'] > 1
        
        if mostrar_expander:
            tem_erro_atual = abs(dif_acervo_final) > 0.05 or abs(dif_dep_final) > 0.05
            
            if tem_erro_atual:
                titulo = f"⚠️ UG {ug}: Exige Verificação"
            elif info['erro_original_acervo'] or info['erro_original_dep']:
                titulo = f"✅ UG {ug}: Confirmado/Ajustado pelo Utilizador"
            else:
                titulo = f"🔍 UG {ug}: Relatórios Agrupados com Sucesso"
                
            with st.expander(titulo, expanded=tem_erro_atual):
                df_view = pd.DataFrame([
                    {"Conta": "Acervo", "Saldo Relatórios": info['pdf_acervo'], "Saldo Tesouro": info['ex_acervo'], "Diferença": dif_acervo_final},
                    {"Conta": "Depreciação", "Saldo Relatórios": info['pdf_dep'], "Saldo Tesouro": info['ex_dep'], "Diferença": dif_dep_final}
                ])
                
                st.dataframe(df_view.style.format({
                    "Saldo Relatórios": lambda x: f"R$ {formatar_real(x)}", 
                    "Saldo Tesouro": lambda x: f"R$ {formatar_real(x)}", 
                    "Diferença": lambda x: f"R$ {formatar_real(x)}"
                }), use_container_width=True)
                
                if info['erro_original_acervo'] or info['erro_original_dep']:
                    st.markdown("---")
                    st.markdown("**✏️ Inserir Valor Correto:**")
                    
                    if info['erro_original_acervo']:
                        st.markdown("**🔹 Acervo Bibliográfico**")
                        if info['detalhes_acervo']:
                            cols = st.columns(2)
                            for idx, (arq, val) in enumerate(info['detalhes_acervo'].items()):
                                with cols[idx % 2]:
                                    st.number_input(f"Ajuste - Relatório {arq}", value=float(val), step=100.0, key=f"edit_ac_{ug}_{arq}")
                        else:
                            st.number_input(f"Ajuste - Acervo Total (Relatório Ausente)", value=float(info['pdf_acervo']), step=100.0, key=f"edit_ac_{ug}_total")

                    if info['erro_original_dep']:
                        st.markdown("**🔸 Depreciação Acumulada**")
                        if info['detalhes_dep']:
                            cols = st.columns(2)
                            for idx, (arq, val) in enumerate(info['detalhes_dep'].items()):
                                with cols[idx % 2]:
                                    st.number_input(f"Ajuste - Relatório {arq}", value=float(val), step=100.0, key=f"edit_dp_{ug}_{arq}")
                        else:
                            st.number_input(f"Ajuste - Depreciação Total (Relatório Ausente)", value=float(info['pdf_dep']), step=100.0, key=f"edit_dp_{ug}_total")

        texto_ug = f"Unidade Gestora: {ug} - {info['nome'][:50]}"
        avisos_soma = []
        if info['arquivos_acervo_somados'] > 1: avisos_soma.append(f"{info['arquivos_acervo_somados']} Acervos")
        if info['arquivos_dep_somados'] > 1: avisos_soma.append(f"{info['arquivos_dep_somados']} Depreciações")
        if avisos_soma: texto_ug += f" (+{' e '.join(avisos_soma)})"
            
        pdf_out.set_font("helvetica", 'B', 10)
        pdf_out.set_fill_color(240, 240, 240)
        pdf_out.cell(0, 8, text=texto_ug, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
        
        pdf_out.set_font("helvetica", 'B', 8)
        pdf_out.set_fill_color(220, 230, 241)
        pdf_out.cell(46, 7, "Conta", 1, fill=True)
        pdf_out.cell(48, 7, "Saldo Relatórios", 1, fill=True, align='C') 
        pdf_out.cell(48, 7, "Saldo Tesouro", 1, fill=True, align='C')      
        pdf_out.cell(48, 7, "Diferença", 1, fill=True, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        pdf_out.set_font("helvetica", '', 8)
        
        str_pdf_acervo = f"R$ {formatar_real(info['pdf_acervo'])}" + (" *" if editado_acervo else "")
        str_pdf_dep = f"R$ {formatar_real(info['pdf_dep'])}" + (" *" if editado_dep else "")

        pdf_out.cell(46, 7, "Acervo Bibliográfico", 1)
        pdf_out.cell(48, 7, str_pdf_acervo, 1, align='R')
        pdf_out.cell(48, 7, f"R$ {formatar_real(info['ex_acervo'])}", 1, align='R')
        if abs(dif_acervo_final) > 0.05: pdf_out.set_text_color(200, 0, 0)
        pdf_out.cell(48, 7, f"R$ {formatar_real(dif_acervo_final)}", 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf_out.set_text_color(0, 0, 0)
        
        pdf_out.cell(46, 7, "Depreciação Acumulada", 1)
        pdf_out.cell(48, 7, str_pdf_dep, 1, align='R')
        pdf_out.cell(48, 7, f"R$ {formatar_real(info['ex_dep'])}", 1, align='R')
        if abs(dif_dep_final) > 0.05: pdf_out.set_text_color(200, 0, 0)
        pdf_out.cell(48, 7, f"R$ {formatar_real(dif_dep_final)}", 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf_out.set_text_color(0, 0, 0)
        
        if editado_acervo or editado_dep:
            pdf_out.set_font("helvetica", 'I', 7)
            pdf_out.set_text_color(180, 0, 0) 
            if editado_acervo:
                pdf_out.cell(0, 5, f"* NOTA DE AUDITORIA: Valor do Acervo ajustado pelo utilizador. (Original lido do PDF: R$ {formatar_real(info['original_pdf_acervo'])})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            if editado_dep:
                pdf_out.cell(0, 5, f"* NOTA DE AUDITORIA: Valor da Depreciação ajustado pelo utilizador. (Original lido do PDF: R$ {formatar_real(info['original_pdf_dep'])})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf_out.set_text_color(0, 0, 0)
        
        pdf_out.ln(5)

    dif_total_acervo = total_pdf_acervo - total_ex_acervo
    dif_total_dep = total_pdf_dep - total_ex_dep
    
    st.markdown("### Resumo Global (Atualizado Constantemente)")
    c1, c2, c3 = st.columns(3)
    c1.metric("Divergência Total (Acervo)", f"R$ {formatar_real(dif_total_acervo)}", delta_color="inverse" if abs(dif_total_acervo) > 0.05 else "normal")
    c2.metric("Divergência Total (Depreciação)", f"R$ {formatar_real(dif_total_dep)}", delta_color="inverse" if abs(dif_total_dep) > 0.05 else "normal")
    
    if st.session_state.logs:
        with st.expander("⚠️ Alertas de Relatórios Incompletos", expanded=False):
            for log in st.session_state.logs: st.write(log)
    
    try:
        pdf_bytes = bytes(pdf_out.output())
        st.download_button(
            label="📄 BAIXAR RELATÓRIO FINAL E COMPLETO (.PDF)", 
            data=pdf_bytes, 
            file_name=f"RELATORIO_ACERVO_BIBLIOGRAFICO_{mes_selecionado}_{ano_selecionado}.pdf", 
            mime="application/pdf", 
            type="primary", 
            use_container_width=True
        )
    except Exception as e:
        st.error(f"Não foi possível concluir a emissão do documento: {e}")
