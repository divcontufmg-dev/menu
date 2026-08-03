import streamlit as st
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
from fpdf import FPDF
import openpyxl

# ==========================================
# CONFIGURAÇÃO INICIAL E MEMÓRIA
# ==========================================
st.set_page_config(page_title="Mapa de Restrições", layout="wide")

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

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

# --- MÓDULO DE ADMINISTRAÇÃO (TOPO DIREITO) ---
col_titulo, col_toggle = st.columns([8, 2])

with col_titulo:
    st.title("Mapa de Restrições por UG")

with col_toggle:
    st.write("") # Espaçamento para alinhar com o título verticalmente
    admin_mode = st.toggle("⚙️ Modo Admin")

if admin_mode:
    st.markdown("### Gerenciador do Banco de Dados (Excel)")
    senha = st.text_input("Senha de acesso:", type="password", key="senha_admin")
    
    if senha == "dcfdc":
        try:
            arquivo = "base.xlsx"
            df_ug_bruto = pd.read_excel(arquivo, sheet_name="ug")
            df_rest_bruto = pd.read_excel(arquivo, sheet_name="restrições")
            
            aba_rest, aba_ug = st.tabs(["Restrições", "UGs"])
            
            with aba_rest:
                st.info("Edite os campos, adicione linhas no final ou exclua selecionando a lateral e apertando Delete.")
                df_rest_editado = st.data_editor(
                    df_rest_bruto, 
                    num_rows="dynamic", 
                    use_container_width=True,
                    height=400,
                    key="editor_rest"
                )
                
            with aba_ug:
                df_ug_editado = st.data_editor(
                    df_ug_bruto, 
                    num_rows="dynamic", 
                    use_container_width=True, 
                    height=400,
                    key="editor_ug"
                )
            
            if st.button("💾 Salvar Alterações na Planilha", type="primary", use_container_width=True):
                try:
                    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
                        df_ug_editado.to_excel(writer, sheet_name="ug", index=False)
                        df_rest_editado.to_excel(writer, sheet_name="restrições", index=False)
                        
                    st.cache_data.clear()
                    st.success("Planilha atualizada permanentemente!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro crítico ao salvar o Excel: {e}")
                    
        except Exception as e:
            st.error(f"Erro ao ler a planilha base: {e}")
            
    st.divider()


# ==========================================
# EXTRAÇÃO DE DADOS PRINCIPAL
# ==========================================
@st.cache_data
def carregar_dados_planilha():
    arquivo = "base.xlsx"
    try:
        df_ugs = pd.read_excel(arquivo, sheet_name="ug")
        df_restricoes = pd.read_excel(arquivo, sheet_name="restrições")
        
        df_ugs = df_ugs.fillna("").replace(0, "")
        df_restricoes = df_restricoes.fillna("").replace(0, "")
        
        lista_ugs = [str(ug) for ug in df_ugs.iloc[:, 0].tolist() if str(ug).strip() != ""]
        dict_restricoes = {str(k): str(v) for k, v in zip(df_restricoes.iloc[:, 0], df_restricoes.iloc[:, 1])}
        
        return lista_ugs, dict_restricoes
        
    except Exception as e:
        return [], {}


# ==========================================
# GERAÇÃO DO PDF
# ==========================================
def gerar_pdf_mensagens(df_dash, dict_rest, data_ref_str, mes_ant_nome, ano_ant_str):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    ugs_com_restricao = df_dash[df_dash["Situação"] == "Com Restrição"]
    ugs_sem_restricao = df_dash[df_dash["Situação"] == "Sem Restrição"]["UG"].tolist()
    
    pdf.add_page()
    pdf.set_font("helvetica", size=11)
    
    for _, row in ugs_com_restricao.iterrows():
        ug_numero = row["UG"]
        codigos_aplicados = str(row["Restrições Aplicadas"])
        lista_codigos = [c.strip() for c in codigos_aplicados.split(',')]
        
        restricoes_com_desc = ""
        for cod in lista_codigos:
            descricao = dict_rest.get(cod, "Descrição não encontrada")
            restricoes_com_desc += f"* {cod} -> {descricao}\n"
            
        mensagem = (
            f"UG {ug_numero}\n"
            f"Conforme estabelecido no calendário de fechamento mensal, disponível na transação CONFECMES do SIAFI Web, "
            f"a data limite para o registro da Conformidade Contábil de UG do mês de {mes_ant_nome.lower()} de {ano_ant_str} é {data_ref_str}\n\n"
            f"Esclarecemos que a respectiva conformidade deverá ser registrada no sistema SIAFIWeb {ano_ant_str} no dia {data_ref_str}, "
            f"de preferência até as 14:00 horas, por meio da transação CONCONFCON, com os códigos: {codigos_aplicados}.\n\n"
            f"{restricoes_com_desc}\n"
            f"Lembramos que, na nova funcionalidade do SIAFI Web (transação \"CONCONFCON\"), o registro da Conformidade Contábil só é finalizado "
            f"com o registro posterior à inclusão das restrições (em caso de ocorrência). Isto é, além de adicionar as restrições e salvar, "
            f"é preciso fazer a confirmação clicando no ícone \"Registrar Conformidade\" no rodapé da transação.\n\n"
            f"O passo a passo para registro da Conformidade Contábil poderá ser acessado por meio do link: "
            f"https://www.ufmg.br/proplan/wp-content/uploads/2024/02/Manual-Conformidade-Contabil-SIAFI-WEB-2024.pdf\n\n"
            f"Atenciosamente,\n\n"
            f"ELÍZIO MARCOS DOS REIS\n"
            f"Diretor do Departamento de Contabilidade e Finanças /PROPLAN/UFMG\n\n"
            f"- ATENÇÃO: Para sua segurança, sempre verifique a autenticidade de links.\n"
        )
        
        pdf.multi_cell(0, 5, mensagem)
        pdf.ln(10)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(10)
    
    pdf.add_page()
    pdf.set_font("helvetica", style="B", size=14)
    pdf.cell(0, 10, "Relação de UGs Sem Restrição", ln=True, align="C")
    pdf.ln(10)
    
    pdf.set_font("helvetica", size=12)
    if ugs_sem_restricao:
        lista_formatada = ", ".join(ugs_sem_restricao)
        pdf.multi_cell(0, 6, lista_formatada)
    else:
        pdf.cell(0, 10, "Nenhuma UG classificada como Sem Restrição.")
        
    return bytes(pdf.output())


# ==========================================
# FLUXO PRINCIPAL DA APLICAÇÃO
# ==========================================
lista_ugs, dict_restricoes = carregar_dados_planilha()

if not lista_ugs or not dict_restricoes:
    st.warning("Não foi possível carregar a base de dados principal. Verifique o arquivo base.xlsx.")
    st.stop() 

if 'restricoes_aplicadas' not in st.session_state:
    st.session_state.restricoes_aplicadas = {}

if not admin_mode:
    st.divider()

col_data1, col_data2 = st.columns(2)
with col_data1:
    data_fechamento = st.date_input("Data da Conformidade Contábil", value=date.today(), format="DD/MM/YYYY")
    
data_anterior = data_fechamento - relativedelta(months=1)
mes_anterior_pt = MESES_PT[data_anterior.month]
ano_anterior_str = str(data_anterior.year)
data_fech_str = data_fechamento.strftime('%d/%m/%Y')

with col_data2:
    st.info(f"**Mês de Referência para Análise:** {mes_anterior_pt} / {ano_anterior_str}")

st.divider()

col_selecao, col_busca = st.columns(2)
with col_selecao:
    ug_selecionada = st.selectbox("Selecione a UG:", lista_ugs)
    
with col_busca:
    busca = st.text_input("🔍 Pesquisar restrição (por nome ou descrição):", "")

st.write(f"**Marque as restrições para a UG {ug_selecionada}:**")
restricoes_marcadas = []

colunas_checkbox = st.columns(3)

for indice, (restricao, descricao) in enumerate(dict_restricoes.items()):
    coluna_atual = colunas_checkbox[indice % 3]
    
    destaque = ""
    if busca:
        if busca.lower() in restricao.lower() or busca.lower() in descricao.lower():
            destaque = " 👈 (ENCONTRADA)"
    
    with coluna_atual:
        if st.checkbox(restricao + destaque, help=descricao, key=f"chk_{ug_selecionada}_{restricao}"):
            restricoes_marcadas.append(restricao)
            
st.write("") 
if st.button("Confirmar UG", type="primary"):
    if restricoes_marcadas:
        st.session_state.restricoes_aplicadas[ug_selecionada] = restricoes_marcadas
    else:
        st.session_state.restricoes_aplicadas[ug_selecionada] = ["SEM RESTRIÇÃO"]
    st.success(f"Log registrado para a UG: {ug_selecionada}")

if st.session_state.restricoes_aplicadas:
    st.write("### UGs processadas até o momento")
    resumo_texto = ""
    for ug, rests in st.session_state.restricoes_aplicadas.items():
        resumo_texto += f"• UG {ug}: {', '.join(str(r) for r in rests)}\n"
    
    st.text_area("Acompanhamento:", value=resumo_texto, height=150, disabled=True)

st.divider()

if st.button("Gerar Dashboard e Relatórios", use_container_width=True):
    dados_dashboard = []
    
    for ug in lista_ugs:
        restricoes = st.session_state.restricoes_aplicadas.get(ug, ["SEM RESTRIÇÃO"])
            
        dados_dashboard.append({
            "UG": ug,
            "Situação": "Sem Restrição" if restricoes == ["SEM RESTRIÇÃO"] else "Com Restrição",
            "Restrições Aplicadas": ", ".join(str(r) for r in restricoes)
        })
        
    df_dashboard = pd.DataFrame(dados_dashboard).fillna("").replace(0, "")
    
    st.subheader("Painel Geral de Conformidade")
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Total de UGs da Base", len(df_dashboard))
    col_m2.metric("Com Restrição", len(df_dashboard[df_dashboard["Situação"] == "Com Restrição"]))
    col_m3.metric("Sem Restrição", len(df_dashboard[df_dashboard["Situação"] == "Sem Restrição"]))
    
    st.dataframe(df_dashboard, use_container_width=True, hide_index=True)
    
    # --- DOWNLOAD DO PDF DIRETO ---
    st.markdown("### Exportar Documentos")
    
    pdf_bytes = gerar_pdf_mensagens(
        df_dash=df_dashboard,
        dict_rest=dict_restricoes,
        data_ref_str=data_fech_str,
        mes_ant_nome=mes_anterior_pt,
        ano_ant_str=ano_anterior_str
    )
    
    st.download_button(
        label="📄 Baixar Mensagens (PDF)",
        data=pdf_bytes,
        file_name=f"Mensagens_Conformidade_{mes_anterior_pt}_{ano_anterior_str}.pdf",
        mime="application/pdf",
        type="primary"
    )
    
    st.write("") 
    if st.button("Nova Análise (Limpar Memória)"):
        st.session_state.restricoes_aplicadas = {}
        st.rerun()
