from openpyxl import load_workbook

wb = load_workbook('D:/coding/2025/solarproject/data_collection/tests/dados16.xlsx')

sheet = wb['page1']

def average():
   return sheet.append(['Média','=AVERAGE(B2:B50)','=AVERAGE(C2:C50)', '=AVERAGE(D2:D50)', '=AVERAGE(E2:E50)' ,'=AVERAGE(F2:F50)', '=AVERAGE(G2:G50)','=AVERAGE(H2:H50)', '=AVERAGE(I2:I50)', '=AVERAGE(J2:J50)', '=AVERAGE(K2:K50)', '=AVERAGE(L2:L50)', '=AVERAGE(M2:M50)', '=AVERAGE(N2:N50)', '=AVERAGE(O2:O50)', '=AVERAGE(P2:P50)', '=AVERAGE(Q2:Q50)'])