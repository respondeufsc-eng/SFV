from flask import Flask, send_file, request, render_template, redirect, url_for, flash
from openpyxl import Workbook
import pandas as pd
import tempfile

app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = 'segredo'

# Página inicial: gerar planilha
@app.route('/')
def index():
    return render_template('index.html')

# Rota para gerar a planilha
@app.route('/generate_excel')
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Análise Módulos FV"

    headers = [
        "ID do Módulo", "Fabricante", "Modelo", "Ano", "Bifacial/Monofacial",
        "Vidro Quebrado?", "Reparável?", "Danos Graves?", "Resistência de Isolamento (MΩ·m²)",
        "Idade Conhecida?", "Potência (% da original)", "Voc Medido (V)", "Isc Medido (A)",
        "Pmáx Medido (W)", "Fill Factor Medido (%)", "Fill Factor Original (%)",
        "Rachaduras?", ">50% Células Danificadas?", "Resultado"
    ]
    ws.append(headers)

    try:
        quantity = int(request.args.get('quantity', '1'))
    except ValueError:
        quantity = 1

    for i in range(quantity):
        row_data = [f"Módulo {i+1}"] + ["" for _ in range(len(headers) - 1)]
        ws.append(row_data)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        file_path = tmp.name
        wb.save(file_path)

    return send_file(
        file_path,
        as_attachment=True,
        download_name="Planilha_Modulos_FV_Segunda_Vida.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Página de upload e análise
@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files.get('file')
        if file and file.filename.endswith('.xlsx'):
            df = pd.read_excel(file)

            # Avaliação com base no fluxograma
            def avaliar_modulo(row):
                if row["Vidro Quebrado?"] == "Sim" and row["Reparável?"] == "Não":
                    return "R"
                if row["Vidro Quebrado?"] == "Não" and row["Danos Graves?"] == "Sim":
                    return "R"
                if row["Resistência de Isolamento (MΩ·m²)"] <= 40:
                    return "R"
                if row["Idade Conhecida?"] == "Sim":
                    if row["Potência (% da original)"] <= 10:
                        return "R"
                else:
                    if row["Potência (% da original)"] < 60:
                        return "R"
                if row["Rachaduras?"] == "Sim":
                    return "R"
                if row[">50% Células Danificadas?"] == "Sim":
                    return "R"
                return "SL"

            df["Resultado"] = df.apply(avaliar_modulo, axis=1)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                output_path = tmp.name
                df.to_excel(output_path, index=False)

            return send_file(output_path, as_attachment=True, download_name="Resultado_Analise_Modulos_FV.xlsx")
        else:
            flash('Por favor, envie um arquivo .xlsx válido.')

    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)
