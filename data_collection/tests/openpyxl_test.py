# Adicionado dados a uma planilha existente.
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = load_workbook('D:/coding/2025/solarproject/data_collection/tests/dados16.xlsx')

sheet = wb['page1']


# tabela = Table(displayName="Tabela1", ref="A1:Q50")

# style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
# tabela.tableStyleInfo = style



sheet['A53'] = 'Média'
sheet['B53'] = '=AVERAGE(B2:B50)'

# sheet.append(["Raul",15,"Solteiro"])


wb.save('D:/coding/2025/solarproject/data_collection/tests/dados16.xlsx')