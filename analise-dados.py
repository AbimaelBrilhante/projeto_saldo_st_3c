import pandas as pd

excel = r"C:\Users\abimaelsoares\Desktop\CONCILIAÇÃO\Conciliação Fiscal 4ª Semana - ST29 06.2022.xlsx"
plan = pd.read_excel(excel, sheet_name="Conciliação Entradas")
plan_edit = plan.loc[plan['Vlr ICMS']!=0]

plan.groupby('Tributação')
