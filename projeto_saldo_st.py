import sqlite3
import pandas as pd
import xlsxwriter


cxn = sqlite3.connect('bd_saldo_icmsst.db')
cursor = cxn.cursor()

def importa():
    print("Importanto Planilhas")
    wb1 = pd.read_excel(r'C:\Users\abimaelsoares\Desktop\projeto_saldost\saidas pt1.xlsx',sheet_name = 'Análise 1')
    wb1.to_sql(name='SAIDAS_3C',con=cxn,if_exists='append',index=False)

    wb1 = pd.read_excel(r'C:\Users\abimaelsoares\Desktop\projeto_saldost\saidas pt2.xlsx',sheet_name = 'Análise 1')
    wb1.to_sql(name='SAIDAS_3C',con=cxn,if_exists='append',index=False)

    wb1 = pd.read_excel(r'C:\Users\abimaelsoares\Desktop\projeto_saldost\saidas pt3.xlsx',sheet_name = 'Análise 1')
    wb1.to_sql(name='SAIDAS_3C',con=cxn,if_exists='append',index=False)

    wb2 = pd.read_excel(r'C:\Users\abimaelsoares\Desktop\projeto_saldost\entradas.xlsx', sheet_name='Entradas')
    wb2.to_sql(name='ENTRADAS_3C', con=cxn, if_exists='append', index=False)

    wb3 = pd.read_excel(r'C:\Users\abimaelsoares\Desktop\projeto_saldost\entradas.xlsx', sheet_name='Saldo Anterior')
    wb3.to_sql(name='SALDO_ANTERIOR', con=cxn, if_exists='append', index=False)

    cxn.commit()


def saldo_atual_provisorio():
    print("Calculando ST e Saldo total das entradas")
    cursor.execute("""CREATE table saldo_atual_provisorio AS SELECT Empresa, Centro, Divisão, Material, "Descrição Material",
    UM, SUM("Saldo Qtd") as Saldo_Qtd, SUM("ICMS ST Total Atualizado"), (SUM("ICMS ST Total Atualizado")/SUM("Saldo Qtd")) as Valor_unit_ST
    FROM(
	SELECT Empresa, Centro, Divisão, Material, "Descrição Material",UM,"ICMS ST Total Atualizado", "Saldo Qtd"  FROM SALDO_ANTERIOR
		UNION ALL
	SELECT Empresa, Centro, Divisão, Material, "Descrição Material",UM,"Valor ICMS ST", Quantidade FROM ENTRADAS_3C
	) AS Total
    GROUP BY Material, Empresa, Centro, "Divisão" """)



def sintetiza_dados():
    print("Calculando ST das entradas para as saidas")
    cursor.execute("""CREATE TABLE saidas_sinteticas AS SELECT Material1,"Descrição Material1", "CFOP1",
    "Tipo de Avaliação1",Empresa1,Centro1, SUM(Quantidade1), AVG(Valor_unit_ST) AS unit_st,(AVG(Valor_unit_ST))*SUM(Quantidade1) as total_st_entrada 
	FROM saldo_atual_provisorio
    INNER JOIN SAIDAS_3C ON saldo_atual_provisorio.Material = SAIDAS_3C.Material1 AND
    saldo_atual_provisorio.Empresa = SAIDAS_3C.Empresa1 AND saldo_atual_provisorio.Centro = SAIDAS_3C.Centro1
    GROUP BY Material1,Empresa1,Centro1 """)
    cxn.commit()

def planilha_modelo_template():
    print("Gerando planilha Template")
    cursor.execute("""CREATE table modelo_template AS SELECT "ID do Cenário", "Data Lançamento", "Material", 
    "Tipo de Avaliação","Docnum", "Empresa","Centro","Divisão","Valor ICMS","Valor ICMS ST",
    "Valor IPI" FROM ENTRADAS_3C WHERE TIPO = "CALCULADO NA ENTRADA" """)
    cxn.commit()
    df = pd.read_sql("select * from planilha_modelo_template",cxn)
    df.to_excel("saldo_atual.xlsx", index = False)


def saldo_consistido():
    print("Consolidando Saldo Atual")
    cursor.execute("""create table SALDO_ATUAL as SELECT 
	Empresa,Centro,Material,"Descrição Material",UM, 
    SUM(Saldo_Qtd) AS qtd_entradas_sldanterior,SUM(Valor_unit_ST * Saldo_Qtd) as total_st, 
    SUM(Valor_unit_ST) as unt_st, (SUM(Saldo_Qtd) - SUM("SUM(Quantidade1)")) AS saldo_atualizado, 
    sum(Valor_unit_ST) * SUM(Saldo_Qtd) as total_st_atualizado,  sum(Valor_unit_ST) as unt_st_atualizado

	FROM 
		saidas_sinteticas
INNER JOIN 
		saldo_atual_provisorio ON saidas_sinteticas.Material1 = saldo_atual_provisorio.Material AND 
		saidas_sinteticas.Empresa1 = saldo_atual_provisorio.Empresa AND 
		saidas_sinteticas.Centro1 = saldo_atual_provisorio.Centro
GROUP BY 
		Material,Empresa, Centro""")
    exclui_saldo_provisorio()

def exclui_saldo_provisorio():
    cursor.execute("""drop table saldo_atual_provisorio""")

def exportar_saldo_atual():
    writer = pd.ExcelWriter('saldo_atual.xlsx', engine='xlsxwriter')
    df = pd.read_sql("select * from SALDO_ATUAL",cxn)
    df.to_excel(writer, index = False, sheet_name="Saldo_atual_detalhado")
    df2 = pd.read_sql("SELECT EMPRESA, saldo_atualizado, total_st_atualizado FROM SALDO_ATUAL GROUP BY EMPRESA", cxn)
    df2.to_excel(writer, index=False, sheet_name="saldo_atual_por_empresa")
    df3 = pd.read_sql("SELECT EMPRESA,Centro, saldo_atualizado, total_st_atualizado FROM SALDO_ATUAL GROUP BY Centro, EMPRESA", cxn)
    df3.to_excel(writer, index=False, sheet_name="saldo_atual_por_filial")
    writer.save()


if __name__ == "__main__":
    pass
    # importa()
    # saldo_atual_provisorio()
    # sintetiza_dados()
    # saldo_consistido()
    # planilha_modelo_template()
    exportar_saldo_atual()
    # cxn.close

