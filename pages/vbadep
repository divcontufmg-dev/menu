
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from copy import copy
import io
import zipfile
import os

# Configuração da página Web
st.set_page_config(page_title="Automação de Depreciação", page_icon="📊", layout="centered")

st.title("📊 Atualizador e Divisor de Planilhas")
st.write("A base **MATRIZ** já está carregada no sistema. Faça o upload apenas da **Base de UG** que deseja automatizar.")

# Área de Upload apenas para o arquivo alvo
arquivo_alvo = st.file_uploader("Envie a planilha a ser processada (com as abas das UGs)", type=["xlsx"])

if arquivo_alvo:
    if st.button("Processar e Dividir Planilhas", type="primary"):
        with st.spinner("Lendo a MATRIZ do sistema e aplicando automação..."):
            try:
                # --- 1) Ler os dados da MATRIZ direto do repositório (nuvem) ---
                caminho_matriz = "MATRIZ.xlsx"
                
                if not os.path.exists(caminho_matriz):
                    st.error(f"Erro: O arquivo '{caminho_matriz}' não foi encontrado no repositório do GitHub.")
                    st.stop()
                    
                wb_matriz = openpyxl.load_workbook(caminho_matriz, data_only=True)
                ws_matriz = wb_matriz.active
                
                # Cria um dicionário para fazer a função do PROCV
                dicionario_matriz = {}
                for row in ws_matriz.iter_rows(min_row=1, max_col=2, values_only=True):
                    if row[0] is not None:
                        dicionario_matriz[str(row[0]).strip()] = row[1]
                
                # --- 2) Carregar a planilha alvo que foi enviada pelo usuário ---
                wb_alvo = openpyxl.load_workbook(arquivo_alvo)
                
                # Definir estilos base para reutilizar
                borda_fina = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                fonte_arial_9 = Font(name='Arial', size=9)
                alinhamento_centro = Alignment(vertical='center')

                # Iniciar o processamento de cada aba
                abas_processadas = []
                for sheet_name in wb_alvo.sheetnames:
                    if sheet_name == "MATRIZ":
                        continue
                        
                    abas_processadas.append(sheet_name)
                    ws = wb_alvo[sheet_name]
                    
                    # Passo 2: Inserir uma coluna A
                    ws.insert_cols(1)
                    ultima_linha = ws.max_row
                    
                    # Remover linhas e ajustar valores
                    for r in range(ultima_linha, 8, -1):
                        celula_b = ws.cell(row=r, column=2)
                        valor_b = str(celula_b.value).strip() if celula_b.value else ""
                        
                        # Passo 5: Excluir linha se B for 123110402
                        if valor_b == "123110402":
                            ws.delete_rows(r)
                            continue
                            
                        # Passo 4: Converter para número
                        if valor_b.replace('.','',1).isdigit():
                            celula_b.value = float(valor_b) if '.' in valor_b else int(valor_b)
                            celula_b.number_format = 'General'
                            
                        # Passo 3: PROCV na Coluna A
                        ws.cell(row=r, column=1).value = dicionario_matriz.get(valor_b, "#N/D")

                    ultima_linha = ws.max_row
                    
                    # Passo 8: Classificar
                    if ultima_linha >= 9:
                        dados_para_ordenar = []
                        for r in range(9, ultima_linha + 1):
                            linha_dados = [ws.cell(row=r, column=c).value for c in range(1, 5)]
                            dados_para_ordenar.append(linha_dados)
                            
                        dados_para_ordenar.sort(key=lambda x: str(x[0]) if x[0] is not None else "")
                        
                        for i, linha_dados in enumerate(dados_para_ordenar):
                            linha_atual = 9 + i
                            for col_idx, valor in enumerate(linha_dados):
                                ws.cell(row=linha_atual, column=col_idx + 1).value = valor
                    
                    # Passo 6: Somatório
                    linha_total = ws.max_row + 1
                    ws.cell(row=linha_total, column=2).value = "TOTAL"
                    ws.cell(row=linha_total, column=3).value = f"=SUM(C9:C{linha_total - 1})"
                    ws.cell(row=linha_total, column=3).number_format = "#,##0.00"
                    ws.cell(row=linha_total, column=4).value = f"=SUM(D9:D{linha_total - 1})"
                    ws.cell(row=linha_total, column=4).number_format = "#,##0.00"
                    
                    # Passo 9: Nat Desp
                    celula_a8 = ws.cell(row=8, column=1)
                    celula_a8.value = "Nat Desp"
                    if ws.cell(row=8, column=4).font:
                        celula_a8.font = copy(ws.cell(row=8, column=4).font)
                    
                    # Formatar Colunas e Células
                    ws.column_dimensions['A'].width = 15
                    ws.column_dimensions['B'].width = 15
                    ws.column_dimensions['C'].width = 20
                    ws.column_dimensions['D'].width = 20

                    for r in range(6, linha_total + 1):
                        for c in range(1, 5):
                            celula = ws.cell(row=r, column=c)
                            celula.font = fonte_arial_9
                            celula.alignment = alinhamento_centro
                            if r >= 9:
                                celula.border = borda_fina
                                
                # Salvar o arquivo COMPLETO processado em memória
                output_completo = io.BytesIO()
                wb_alvo.save(output_completo)
                output_completo.seek(0)
                
                # --- LÓGICA: DIVIDIR AS ABAS EM ARQUIVOS E ZIPAR ---
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for sheet_name in abas_processadas:
                        output_completo.seek(0)
                        wb_temp = openpyxl.load_workbook(output_completo)
                        
                        for nome_aba in wb_temp.sheetnames:
                            if nome_aba != sheet_name:
                                del wb_temp[nome_aba]
                        
                        single_output = io.BytesIO()
                        wb_temp.save(single_output)
                        zip_file.writestr(f"{sheet_name}.xlsx", single_output.getvalue())
                        
                zip_buffer.seek(0)
                output_completo.seek(0)
                
                st.success("Processo concluído! As abas foram separadas e estão prontas para download.")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="📥 Baixar Planilha Consolidada",
                        data=output_completo,
                        file_name="Planilha_Completa_Atualizada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                with col2:
                    st.download_button(
                        label="🗂️ Baixar Arquivo ZIP (Abas Separadas)",
                        data=zip_buffer,
                        file_name="Abas_Separadas.zip",
                        mime="application/zip"
                    )
                
            except Exception as e:
                st.error(f"Ocorreu um erro durante o processamento: {e}")
