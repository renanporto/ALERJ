import pandas as pd

# Caminho para o arquivo Excel
arquivo_excel = "02_-_Ordem_cronologica_de_pagamentos_realizados_-_Fevereiro.25.xlsx"

df = pd.read_excel(arquivo_excel, sheet_name="Table 1", header=2)

# Mostra as colunas para garantir
print("Colunas detectadas:", df.columns.tolist())

# Localiza a coluna de Despesas Pagas
coluna_despesas = [col for col in df.columns if "Despesas Pagas" in str(col)][0]
print(f"Coluna usada: {coluna_despesas}")

# Limpa e converte os valores
df[coluna_despesas] = (
    df[coluna_despesas]
    .astype(str)
    .str.extract(r"([\d\.,]+)").iloc[:, 0]
    .fillna("0")
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
    .astype(float)
)

# Ordena e seleciona os 50 maiores
df = df[df[coluna_despesas].notna()]
top50 = df.sort_values(by=coluna_despesas, ascending=False).head(50)

# Salva em novo arquivo Excel
top50.to_excel("top50_despesas_pagas.xlsx", index=False)

print("✅ Arquivo salvo como: top50_despesas_pagas.xlsx")
