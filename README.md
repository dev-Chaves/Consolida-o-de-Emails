# Consolidação de Contatos por Empresa

Este script foi desenvolvido para consolidar e organizar contatos (emails) associados a diferentes empresas em uma planilha Excel. Ele resolve problemas relacionados à duplicação de emails, agrupando dados de maneira estruturada e garantindo a consistência das informações.

## Funcionalidades

- **Agrupamento de Contatos por Empresa**:  
  Agrupa emails de uma mesma empresa com base no número único (`N°`) associado a cada empresa.

- **Verificação de Emails Duplicados**:  
  Identifica e lista emails que aparecem em mais de uma empresa. Esses casos são salvos em um arquivo separado para revisão (`emails_duplicados.xlsx`).

- **Normalização de Emails**:  
  Remove espaços extras e normaliza todos os emails para letras minúsculas, evitando problemas relacionados a maiúsculas/minúsculas.

- **Geração de Planilha Consolidada**:  
  Cria uma nova planilha Excel (`planilha_consolidada.xlsx`) onde os emails de cada empresa são organizados em colunas distintas, com no máximo um email por célula.

## Como Funciona

O script processa um arquivo Excel ou CSV contendo as seguintes colunas obrigatórias:

- **N°**: Número único de identificação da empresa.
- **NOME**: Nome da empresa.
- **EMAIL**: Endereço de email associado à empresa.

Ele:
1. Identifica duplicações de emails entre diferentes empresas e os separa para análise.
2. Remove emails duplicados dentro de uma mesma empresa.
3. Agrupa emails únicos por empresa e os distribui em colunas na planilha de saída.

## Arquivos Gerados

1. **`planilha_consolidada.xlsx`**:  
   Arquivo com os dados consolidados, onde cada linha representa uma empresa e os emails estão organizados em colunas separadas.

2. **`emails_duplicados.xlsx`** (opcional):  
   Arquivo com a lista de emails compartilhados entre diferentes empresas, para revisão manual.

## Exemplo de Saída

### Entrada

| N°   | NOME                                  | EMAIL                      |  
|------|---------------------------------------|----------------------------|  
| 1348 | VERJ ARMAZEM LTDA                     | verjmercado@gmail.com       |  
| 1349 | LEITE E MAIA SUPERMERCADO LTDA        | verjmercado@gmail.com       |  
| 1351 | SUPER VERJ SUPERMERCADO LTDA          | verjmercado@gmail.com       |  
| 1358 | CAMPAGNOLA RESTAURANTE LTDA           | lacampagnola@uol.com.br     |  

### Saída (planilha consolidada)

| N°   | NOME                                  | EMAIL_1                  | EMAIL_2            |  
|------|---------------------------------------|--------------------------|--------------------|  
| 1348 | VERJ ARMAZEM LTDA                     | verjmercado@gmail.com     | NaN                |  
| 1358 | CAMPAGNOLA RESTAURANTE LTDA           | lacampagnola@uol.com.br   | NaN                |  

### Saída (emails duplicados)

| N°   | NOME                                  | EMAIL                      |  
|------|---------------------------------------|----------------------------|  
| 1348 | VERJ ARMAZEM LTDA                     | verjmercado@gmail.com       |  
| 1349 | LEITE E MAIA SUPERMERCADO LTDA        | verjmercado@gmail.com       |  
| 1351 | SUPER VERJ SUPERMERCADO LTDA          | verjmercado@gmail.com       |  

## Requisitos

- **Python 3.8+**
- Bibliotecas necessárias:
  - `pandas`
  - `openpyxl` (para trabalhar com arquivos Excel)

## Como Usar

1. Instale as dependências:
   ```bash
   pip install pandas openpyxl

## Edite o arquivo 
df = pd.read_excel("seu_arquivo.xlsx")

## Execute o script

python consolidar_empresas.py

## Verifique os arquivos gerados
planilha_consolidada.xlsx para os dados consolidados.
emails_duplicados.xlsx (se aplicável) para emails compartilhados.
