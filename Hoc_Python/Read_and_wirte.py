import xlwings as xw
import pandas as pd
# Open an existing Workbook
from pandas import DataFrame

wb = xw.Book('G2LC_LTP_Analyze_BSP_5.10.xlsx')
sheet = wb.sheets['Sheet1']

# read and write values from the worksheet

sheet.range('A5').value = 'Foo'
print(sheet.range('A5').value)

df = pd.DataFrame([[1, 2], [3, 4]], columns=['a', 'b'])

sht.range('A5').value = df

sht.range('A5').options(pd.DataFrame, expand='table').value
