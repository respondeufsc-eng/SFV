import os
import tempfile
from flask import Flask, send_file, request
from openpyxl import Workbook

app = Flask(__name__)

@app.route('/generate_excel')
def generate_excel():
    # Criar a planilha com os campos conforme solicitado
    wb = Workbook()
    ws = wb.active
    ws.title = "Análise Módulos FV"

    # Cabeçalhos conforme o fluxograma e dados elétricos solicitados
    headers = [
        "ID do Módulo", "Fabricante", "Modelo", "Ano", "Bifacial/Monofacial",
        "Vidro Quebrado?", "Reparável?", "Danos Graves?", "Resistência de Isolamento (MΩ·m²)",
        "Idade Conhecida?", "Potência (% da original)", "Voc Medido (V)", "Isc Medido (A)",
        "Pmáx Medido (W)", "Fill Factor Medido (%)", "Fill Factor Original (%)",
        "Rachaduras?", ">50% Células Danificadas?", "Resultado"
    ]

    ws.append(headers)

    # Obter a quantidade de módulos da requisição (padrão: 1)
    quantity_str = request.args.get('quantity', '1')
    try:
        quantity = int(quantity_str)
    except ValueError:
        quantity = 1 # Garante que seja um número inteiro

    # Adicionar linhas vazias para cada módulo, com um ID sequencial
    for i in range(quantity):
        # Criar uma linha com o ID do módulo e o restante vazio
        row_data = [f"Módulo {i+1}"] + ["" for _ in range(len(headers) - 1)]
        ws.append(row_data)


    # Salvar o arquivo em um local temporário
    # Usamos tempfile para criar um arquivo temporário seguro
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        file_path = tmp.name
        wb.save(file_path)

    # Enviar o arquivo para download
    return send_file(file_path, as_attachment=True, download_name="Planilha_Modulos_FV_Segunda_Vida.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    app.run(debug=True)