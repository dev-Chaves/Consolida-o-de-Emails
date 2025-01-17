import pandas as pd

df = pd.read_excel("unificacao.xlsx") 

df_grouped = df.groupby("N°").agg({
    "NOME": "first",  
    "EMAIL": lambda x: list(set(x)) 
}).reset_index()

emails_df = pd.DataFrame(df_grouped["EMAIL"].tolist())


df_grouped = pd.concat([df_grouped.drop(columns=["EMAIL"]), emails_df], axis=1)

df_grouped.to_excel("planilha_consolidada.xlsx", index=False)

print("Consolidação concluída. A planilha 'planilha_consolidada.xlsx' foi gerada.")