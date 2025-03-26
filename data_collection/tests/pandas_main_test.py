#MID - Make It Dumb
import pandas as pd
from openpyxl import load_workbook

df = pd.read_excel('D:/coding/2025/solarproject/data_collection/tests/teste.xlsx',sheet_name='Página1')

valor1 = df.iloc[0,0]
valor2 = df.iloc[1,0]
valor3 = df.iloc[2,0]
valor4 = df.iloc[3,0]


valores1 = df.iloc[0,1]
valores2 = df.iloc[1,1]
valores3 = df.iloc[2,1]
valores4 = df.iloc[3,1]


novo_df = pd.DataFrame({
    "Valores copiados": [valor1,valor2,valor3,valor4],
    "Idade": [valores1,valores2,valores3,valores4]
})




# with pd.ExcelWriter('teste.xlsx', engine='openpyxl', mode='a') as writer:
#     novo_df.to_excel(writer, sheet_name='Página3', index=False)