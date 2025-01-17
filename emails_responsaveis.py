import pandas as pd

# Ler o arquivo com os dados
df = pd.read_excel("unificacao.xlsx")

# Normalizar os emails (remover espaços e padronizar para letras minúsculas)
df["EMAIL"] = df["EMAIL"].str.strip().str.lower()

# Criar um DataFrame que associa cada email às empresas que o utilizam
email_to_companies = df.groupby("EMAIL").agg({
    "N°": lambda x: list(x),  # Lista dos números das empresas
    "NOME": lambda x: list(x)  # Lista dos nomes das empresas
}).reset_index()

# Expandir os dados para criar uma tabela amigável
expanded_emails = email_to_companies.apply(
    lambda row: pd.Series({
        "EMAIL": row["EMAIL"],
        **{f"EMPRESA_{i+1}": f"{num} - {nome}" for i, (num, nome) in enumerate(zip(row["N°"], row["NOME"]))}
    }),
    axis=1
)

# Salvar o resultado em um arquivo Excel
expanded_emails.to_excel("emails_responsaveis.xlsx", index=False)

print("Arquivo gerado: 'emails_responsaveis.xlsx'")