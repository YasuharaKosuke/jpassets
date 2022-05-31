from cProfile import label
from operator import index
import subprocess as sp
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import sys,os
if os.path.exists('./I011BJ_m.xlsx'):
    wb = openpyxl.load_workbook('I011BJ_m.xlsx')
else:
    sp.call("wget https://www.toushin.or.jp/tws/toukei_dw/I011BJ_m.xlsx",shell="True" )
    wb = openpyxl.load_workbook('I011BJ_m.xlsx')
sheet=wb['月次']
sheet.delete_rows(sheet.min_row,2)
wb.save('result.xlsx')
d=pd.read_excel('result.xlsx',engine='openpyxl',sheet_name='月次')

mouth=[]
# print(d.iloc[7:127,3])

for i in d.iloc[7:127,3]:
    mouth.append(i)

x=[]
for i in range(7,127):
    x.append(i)

y = d.iloc[7:127,4]

def main():
    plt.plot(x, y, label = "Total Net Assets-Structure of Investment Trusts")
    plt.savefig('result.png')
    plt.legend()
    plt.show()

if __name__ == "__main__":
    main()


