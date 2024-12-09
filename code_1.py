from tabula import read_pdf # type: ignore
import pandas as pd # type: ignore
import os
from pathlib import Path
import fitz # type: ignore

pasta_arquivo = str(Path.home()/"Downloads/PDF")
arquivos_pdf = [f for f in os.listdir(pasta_arquivo) if f.endswith('.pdf')]
arquivo_pdf = arquivos_pdf[0]
pdf = fitz.open(pasta_arquivo + '/' + arquivo_pdf) 
retangulo = fitz.Rect(0, 0, 600, 110)
for pagina in pdf:
    pagina.draw_rect(retangulo, color=(1, 1, 1), fill=(1, 1, 1))  # Branco (RGB: 1,1,1)
pdf.save(pasta_arquivo+"/resultado.pdf")
pdf.close()

arquivo_pdf_novo = pasta_arquivo+"/resultado.pdf"
arquivo_excel = pasta_arquivo+"/tabelas_extraidas.xlsx"
dados_unificados = pd.DataFrame()
tabelas = read_pdf(arquivo_pdf_novo, pages="all", multiple_tables=True, pandas_options={"header": None}, lattice=False)
for tabela in tabelas:
    tabela.columns = tabela.iloc[0]
    tabela = tabela.drop(0)
    tabela[['conta', 'digito']] = tabela.iloc[:, 0].str.split('-', expand=True)
    tabela = tabela.drop(tabela.columns[0], axis=1)
    tabela['id']='1'
    nova_ordem = ['Agência', 'conta', 'digito', 'Nome do Funcionário', 'CPF', 'id', 'Líquidos']
    tabela = tabela[nova_ordem]
    tabela['Líquidos'] = tabela['Líquidos'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    tabela['Líquidos'] = pd.to_numeric(tabela['Líquidos'], errors='coerce')
    dados_unificados = pd.concat([dados_unificados, tabela], ignore_index=True)
   
print("Iniciando processamento...")
print("Salvando arquivo Excel em:", arquivo_excel)

try:
    dados_unificados.to_excel(arquivo_excel, index=False)
    print("Arquivo Excel salvo com sucesso.")
except Exception as e:
    print("Erro ao salvar o Excel:", e)
