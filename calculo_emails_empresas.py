import pandas as pd

# Ler a planilha consolidada e a planilha de emails responsáveis
consolidada_df = pd.read_excel("planilha_consolidada.xlsx")
emails_responsaveis_df = pd.read_excel("emails_responsaveis.xlsx")

# Calcular o total de empresas (da tabela consolidada)
total_empresas = consolidada_df["N°"].nunique()

# Calcular o total de emails únicos (da tabela emails responsáveis)
total_emails_unicos = emails_responsaveis_df["EMAIL"].nunique()

# Criar um DataFrame com os resultados
resultados_df = pd.DataFrame({
    "Descrição": ["Total de Empresas", "Total de Emails Únicos"],
    "Quantidade": [total_empresas, total_emails_unicos]
})

# Salvar os resultados em um arquivo Excel
resultados_df.to_excel("resultados_totais.xlsx", index=False)

# Salvar também em um arquivo TXT
with open("resultados_totais.txt", "w") as txt_file:
    txt_file.write(f"Total de Empresas: {total_empresas}\n")
    txt_file.write(f"Total de Emails Únicos: {total_emails_unicos}\n")

print("Cálculos concluídos.")
print(f"Total de Empresas: {total_empresas}")
print(f"Total de Emails Únicos: {total_emails_unicos}")
print("Arquivos gerados: 'resultados_totais.xlsx' e 'resultados_totais.txt'")
