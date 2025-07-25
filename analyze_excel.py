import pandas as pd

# Caminho do arquivo gerado anteriormente
file_path = "/mnt/data/Planilha_Modulos_FV_Segunda_Vida.xlsx"

# Carregar a planilha
df = pd.read_excel(file_path)

# Função para aplicar o fluxo de decisão em cada linha
def avaliar_modulo(row):
    # Etapas do fluxograma baseadas nos dados da planilha
    if row["Vidro Quebrado?"] == "Sim":
        if row["Reparável?"] == "Não":
            return "R"
    else:
        if row["Danos Graves?"] == "Sim":
            return "R"

    if row["Resistência de Isolamento (MΩ·m²)"] <= 40:
        return "R"

    if row["Idade Conhecida?"] == "Sim":
        if row["Potência (% da original)"] <= 10:
            return "R"
    else:
        if row["Potência (% da original)"] < 60:
            return "R"

    if row["Potência (% da original)"] >= 100:
        classe = "Class A"
    else:
        classe = "Class B"

    if row["Rachaduras?"] == "Sim":
        return "R"

    if row[">50% Células Danificadas?"] == "Sim":
        return "R"

    return "SL"

# Aplicar a função a cada linha
df["Resultado"] = df.apply(avaliar_modulo, axis=1)

# Salvar o novo arquivo com os resultados
output_path = "/mnt/data/Resultado_Analise_Modulos_FV.xlsx"
df.to_excel(output_path, index=False)

output_path