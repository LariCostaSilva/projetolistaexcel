'''
Criar scritpt para obter arquivo excel a partir de informações retiradas de um documento pdf
'''
import os
from PyPDF2 import PdfReader
from openpyxl import Workbook

# Diretório onde estão os arquivos PDF (pasta "lista" na área de trabalho) ---- Trocar a pasta para uma sua
# Dentro deve ser colocado os PDFs que voce deseja gerar uma lista 
PDF_DIRECTORY = r'C:\Users\larissa\Desktop\lista'

# Iniciar um novo arquivo Excel
workbook = Workbook()
sheet = workbook.active

# Adicionar cabeçalhos às colunas do arquivo Excel
sheet.append(["Número da Nota Fiscal", "Data de Competência/Emissão", "Vl. Total dos Serviços"])

# Loop através dos arquivos PDF na pasta
for filename in os.listdir(PDF_DIRECTORY):
    if filename.endswith(".pdf"):
        pdf_file = os.path.join(PDF_DIRECTORY, filename)
        
        # Extrair texto de todas as páginas do PDF
        pdf = PdfReader(pdf_file)
        
        for page in pdf.pages:
            pdf_text = page.extract_text()
            
            # Encontre as informações diretamente abaixo dos cabeçalhos
            valores = pdf_text.split("\n")
            
            # Processe os valores
            numero_nota_fiscal = ""
            data_emissao = ""
            valor_total_servicos = ""
            
            for i, valor in enumerate(valores):
                if "Número da Nota Fiscal" in valor:
                    numero_nota_fiscal = valores[i + 1]
                if "Data de Competência/Emissão" in valor:
                    data_emissao = valores[i + 1]
                if "Vl. Total dos Serviços" in valor:
                    # Divida o valor com base no espaço em branco e pegue a parte até o segundo espaço
                    partes = valores[i + 1].split()
                    if len(partes) >= 2:
                        valor_total_servicos = partes[0] + " " + partes[1]

            # Adicione as informações extraídas ao arquivo Excel
            sheet.append([numero_nota_fiscal, data_emissao, valor_total_servicos])

# Salve o arquivo Excel na pasta "lista" da área de trabalho
EXCEL_FILE = os.path.join(PDF_DIRECTORY, "resultado.xlsx")
workbook.save(EXCEL_FILE)
