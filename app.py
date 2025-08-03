from flask import Flask, request, render_template, send_from_directory
import pdfplumber
import pandas as pd
import re
import os

app = Flask(__name__)
PLANILHA_DIR = "static/planilhas"
os.makedirs(PLANILHA_DIR, exist_ok=True)

def extrair_dados_fgts(arquivo_pdf):
    dados_fgts = []
    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                linhas = texto.split('\n')
                for i in range(len(linhas)):
                    if "Nome/Razão Social do Empregador" in linhas[i]:
                        if i + 1 < len(linhas):
                            nome_empregador = linhas[i + 1].strip()
                    if "Valor a recolher" in linhas[i]:
                        if i + 1 < len(linhas):
                            valor_recolher = linhas[i + 1].strip()
                            dados_fgts.append({
                                "Nome/Razão Social do Empregador": nome_empregador,
                                "Valor a recolher": valor_recolher
                            })
    return dados_fgts

def extrair_dados_inss(arquivo_pdf):
    dados_inss = []
    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                linhas = texto.split('\n')
                razao_social = ""
                valor_total = ""
                codigo_barras = ""

                for i in range(len(linhas)):
                    if "Razão Social" in linhas[i]:
                        if i + 1 < len(linhas):
                            razao_social = linhas[i + 1].strip()
                    if "Valor Total do Documento" in linhas[i]:
                        if i + 1 < len(linhas):
                            valor_total = linhas[i + 1].strip()
                    if "Documento de Arrecadação de Receitas Federais" in linhas[i]:
                        if i + 1 < len(linhas):
                            codigo_barras = linhas[i + 1].strip()[:55]
                            dados_inss.append({
                                "Razão Social": razao_social,
                                "Valor Total do Documento": valor_total,
                                "Código de Barras": codigo_barras
                            })
    return dados_inss

def extrair_dados_vinculo(arquivo_pdf):
    dados_vinculo = []
    palavras_ignorar = ["Situação:", "Trabalhando", "CPF:", "Adm:", "Doença"]

    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                linhas = texto.split('\n')
                funcionario = ""
                tipo_vinculo = ""
                valor_liquido = ""

                for i in range(len(linhas)):
                    if "Empr.:" in linhas[i]:
                        funcionario = linhas[i].split("Empr.:")[1].strip()
                        for palavra in palavras_ignorar:
                            funcionario = funcionario.replace(palavra, "")
                        funcionario = re.sub(r'[^a-zA-Z\s]', '', funcionario)
                        funcionario = re.sub(r'\s+', ' ', funcionario).strip()

                    if "Vínculo:" in linhas[i]:
                        tipo_vinculo = linhas[i].split("Vínculo:")[1].strip()
                        if "celetista" in tipo_vinculo.lower():
                            tipo_vinculo = "Celetista"

                    if "Líquido:" in linhas[i]:
                        valor_liquido = linhas[i].split("Líquido:")[1].strip()

                        if tipo_vinculo == "Celetista":
                            dados_vinculo.append({
                                "Funcionário": funcionario,
                                "Vínculo": tipo_vinculo,
                                "Líquido": valor_liquido
                            })
    return dados_vinculo

def exportar_para_planilhas(dados_fgts, dados_inss, dados_vinculo):
    pd.DataFrame(dados_fgts).to_excel(os.path.join(PLANILHA_DIR, "dados_fgts.xlsx"), index=False)
    pd.DataFrame(dados_inss).to_excel(os.path.join(PLANILHA_DIR, "dados_inss.xlsx"), index=False)
    pd.DataFrame(dados_vinculo).to_excel(os.path.join(PLANILHA_DIR, "dados_vinculo.xlsx"), index=False)

@app.route("/", methods=["GET", "POST"])
def index():
    arquivos_gerados = False
    if request.method == "POST":
        arquivos_pdf1 = request.files.getlist("pdf1")
        arquivos_pdf2 = request.files.getlist("pdf2")
        arquivos_pdf3 = request.files.getlist("pdf3")

        todos_dados_fgts = []
        todos_dados_inss = []
        todos_dados_vinculo = []

        for arquivo_pdf in arquivos_pdf1:
            todos_dados_fgts.extend(extrair_dados_fgts(arquivo_pdf))

        for arquivo_pdf in arquivos_pdf2:
            todos_dados_inss.extend(extrair_dados_inss(arquivo_pdf))

        for arquivo_pdf in arquivos_pdf3:
            todos_dados_vinculo.extend(extrair_dados_vinculo(arquivo_pdf))

        exportar_para_planilhas(todos_dados_fgts, todos_dados_inss, todos_dados_vinculo)
        arquivos_gerados = True

    return render_template("index.html", arquivos_gerados=arquivos_gerados)

@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(PLANILHA_DIR, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
