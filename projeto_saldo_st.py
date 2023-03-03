import sqlite3
import pandas as pd
import xlsxwriter
import time
from tkinter import filedialog

cxn = sqlite3.connect('bd_saldo_icmsst.db')
cursor = cxn.cursor()


def importa_saidas():
    filename_saida = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                          filetypes=(("Text files", "*.*"), ("all files", "*.*")))
    print("Importanto Planilhas")
    wb1 = pd.read_excel(filename_saida, sheet_name='Análise 1')
    wb1.to_sql(name='SAIDAS_3C', con=cxn, if_exists='append', index=False)

def importa_entradas():
    filename_entrada = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                          filetypes=(("Text files", "*.*"), ("all files", "*.*")))

    wb2 = pd.read_excel(filename_entrada, sheet_name='Entradas')
    wb2.to_sql(name='ENTRADAS_3C', con=cxn, if_exists='append', index=False)

    wb3 = pd.read_excel(r'C:\Users\abimaelsoares\Desktop\projeto_saldost\Entradas 02.2023.xlsx', sheet_name='Saldo Anterior')
    wb3.to_sql(name='SALDO_ANTERIOR', con=cxn, if_exists='append', index=False)

    cxn.commit()

def importa_devolucoes():
    filename_entrada = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                                      filetypes=(("Text files", "*.*"), ("all files", "*.*")))

    wb2 = pd.read_excel(filename_entrada, sheet_name='Devolucoes')
    wb2.to_sql(name='DEVOLUCOES_3C', con=cxn, if_exists='append', index=False)

def importa_ressarcimento_TIMP():
    filename_ressarcimento_timp = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                          filetypes=(("Text files", "*.*"), ("all files", "*.*")))
    wb4 = pd.read_excel(filename_ressarcimento_timp)
    wb4.to_sql(name="RESS_TIMP", con=cxn, if_exists='append', index=False)

def saldo_atual_provisorio():
    print("Calculando ST e Saldo total das entradas")
    cursor.execute("""CREATE table saldo_atual_provisorio AS SELECT Empresa, Centro, Divisão, Material, "Descrição Material" as Descricao_Material,
    UM, SUM("Saldo Qtd") as Saldo_Qtd, SUM("ICMS ST Total Atualizado" + "Valor ICMS") AS total_st_bruto_atualizado, (SUM("ICMS ST Total Atualizado" + "Valor ICMS")/SUM("Saldo Qtd")) as Valor_unit_ST
        
    FROM(
	SELECT Empresa, Centro, Divisão, Material, "Descrição Material",UM,"ICMS ST Total Atualizado", "Saldo Qtd","Valor ICMS"  FROM SALDO_ANTERIOR
		UNION ALL
	SELECT Empresa, Centro, Divisão, Material, "Descrição Material",UM,"Valor ICMS ST", Quantidade, "Valor ICMS" FROM ENTRADAS_3C  WHERE TIPO = "CALCULADO NA ENTRADA"
	) AS Total
    GROUP BY Material, Empresa, Centro, "Divisão" """)

def criar_coluna_tipo_contabilizacao_saidas():
    cursor.execute("""ALTER TABLE SAIDAS_3C ADD COLUMN tipo_contabilizacao""")
    tipo_contabilizacao_saidas()

def tipo_contabilizacao_saidas():
    cursor.execute("""UPDATE SAIDAS_3C
SET 
    tipo_contabilizacao = 
    CASE WHEN SUBSTRING(CFOP1, 1,1) = "5" THEN "SEM RESSARCIMENTO" 
	WHEN SUBSTRING(CFOP1, 1,4) = "6949" THEN "COM RESSARCIMENTO"
	WHEN SUBSTRING(CFOP1, 1,2) = "69" THEN "SEM RESSARCIMENTO"

	ELSE "COM RESSARCIMENTO" END""")

def sintetiza_dados():
    print("Calculando ST das entradas para as saidas")
    cursor.execute("""CREATE TABLE saidas_sinteticas AS SELECT SUM(Quantidade1) as saldo_saidas,*, AVG(Valor_unit_ST) 
    AS unit_st,(AVG(Valor_unit_ST))*(sum(Quantidade1)) as total_st_entrada 
	FROM saldo_atual_provisorio
    INNER JOIN SAIDAS_3C ON saldo_atual_provisorio.Material = SAIDAS_3C.Material1 AND
    saldo_atual_provisorio.Empresa = SAIDAS_3C.Empresa1 AND saldo_atual_provisorio.Centro = SAIDAS_3C.Centro1
    GROUP BY Docnum1,Material1,Empresa1,Centro1,CFOP1,"Tipo de Avaliação1" ,"tipo_contabilizacao" """)
    cxn.commit()

def sintetiza_dados_devolucoes():
    cursor.execute("""CREATE TABLE devolucoes_sinteticas AS SELECT SUM(Quantidade99) as saldo_saidas,*, AVG(Valor_unit_ST) 
    AS unit_st,(AVG(Valor_unit_ST))*(sum(Quantidade99)) as total_st_entrada 
	FROM saldo_atual_provisorio
    INNER JOIN DEVOLUCOES_3C ON saldo_atual_provisorio.Material = DEVOLUCOES_3C.Material99 AND
    saldo_atual_provisorio.Empresa = DEVOLUCOES_3C.Empresa99 AND saldo_atual_provisorio.Centro = DEVOLUCOES_3C.Centro99
    GROUP BY Docnum99,Material99,Empresa99,Centro99,CFOP99,"Tipo de Avaliação99" """)

def planilha_modelo_template_entradas():
    print("Gerando planilha Template")
    try:
        cursor.execute("""CREATE table modelo_template_entradas AS SELECT "ID do Cenário", "Data Lançamento", "Material", 
        "Tipo de Avaliação","Docnum", "Empresa","Centro","Divisão","Valor ICMS","Valor ICMS ST",
        "Valor IPI" FROM ENTRADAS_3C WHERE TIPO <> "DESTACADO NA NF" """)
        cxn.commit()
        df = pd.read_sql("select * from modelo_template_entradas", cxn)
        df.to_excel("planilha_modelo_template_entradas.xlsx", index=False)

    except:
        df = pd.read_sql("select * from modelo_template_entradas", cxn)
        df.to_excel("planilha_modelo_template_entradas.xlsx", index=False)

def planilha_modelo_template_saidas():
    print("Gerando planilha Template")
    try:
        cursor.execute("""CREATE table modelo_template_saidas AS select substring(CFOP1,1,1)
        as "ID do Cenário", "Data de Lançamento",Material1,"Tipo de Avaliação1",Docnum1,Empresa1,
        Centro1,"Divisão1","ICMS1",total_st_entrada,"IPI1", "tipo_contabilizacao" FROM saidas_sinteticas ORDER BY "tipo_contabilizacao" """)
        cxn.commit()
        df = pd.read_sql("select * from modelo_template_saidas", cxn)
        df.to_excel("planilha_modelo_template_saidas.xlsx", index=False)

        cursor.execute("""CREATE table modelo_template_devolucoes AS select "10"
        as "ID do Cenário", "Data Lançamento99",Material99,"Tipo de Avaliação99",Docnum99,Empresa99,
        Centro99,"Divisão99","Valor ICMS99",total_st_entrada,"Valor IPI99" FROM devolucoes_sinteticas """)
        cxn.commit()
        df2 = pd.read_sql("select * from modelo_template_devolucoes", cxn)
        df2.to_excel("planilha_modelo_template_devolucoes.xlsx", index=False)

    except:
        pass
        df = pd.read_sql("select * from modelo_template_saidas", cxn)
        df.to_excel("planilha_modelo_template_saidas.xlsx", index=False)

        df2 = pd.read_sql("select * from modelo_template_devolucoes", cxn)
        df2.to_excel("planilha_modelo_template_devolucoes.xlsx", index=False)

def saldo_consistido():
    print("Consolidando Saldo Atual")
    cursor.execute("""create table SALDO_ATUAL as select sap.Empresa, sap.Centro,sap.Divisão,
    sap.Material,sap.Descricao_Material,sap.UM, (sap.Saldo_Qtd - sum(saldo_saidas)) AS "Saldo Qtd", 
    avg(sap.total_st_bruto_atualizado) as "ICMS ST Total Atualizado",
    AVG(sap.Valor_unit_ST) as "ICMS ST Unit Atualizado", "Valor ICMS"
    from saldo_atual_provisorio sap
    left join saidas_sinteticas on sap.Material = saidas_sinteticas.Material1 and sap.Empresa = saidas_sinteticas.Empresa and sap.Centro = saidas_sinteticas.Centro and sap."Divisão" = saidas_sinteticas.Divisão
	
    group by
        sap.Material,sap.Empresa,sap.Centro, sap.Divisão""")
    #exclui_saldo_provisorio()

def exclui_saldo_provisorio():
    cursor.execute("""drop table saldo_atual_provisorio""")

def exportar_saldo_atual():
    writer = pd.ExcelWriter('saldo_atual.xlsx', engine='xlsxwriter')
    df = pd.read_sql("select * from SALDO_ATUAL", cxn)
    df.to_excel(writer, index=False, sheet_name="Saldo_atual_detalhado")
    df2 = pd.read_sql("SELECT EMPRESA, 'Saldo Qtd', 'ICMS ST Total Atualizado' FROM SALDO_ATUAL GROUP BY EMPRESA", cxn)
    df2.to_excel(writer, index=False, sheet_name="saldo_atual_por_empresa")
    df3 = pd.read_sql(
        "SELECT EMPRESA,Centro, 'Saldo Qtd', 'ICMS ST Total Atualizado' FROM SALDO_ATUAL GROUP BY Centro, EMPRESA", cxn)
    df3.to_excel(writer, index=False, sheet_name="saldo_atual_por_filial")
    writer.save()

if __name__ == "__main__":
    pass
    # importa_entradas()
    # importa_saidas()
    # criar_coluna_tipo_contabilizacao_saidas()
    # saldo_atual_provisorio()
    # sintetiza_dados()
    saldo_consistido()
    # planilha_modelo_template_entradas()
    # planilha_modelo_template_saidas()
    exportar_saldo_atual()
    # importa_ressarcimento_TIMP()
    cxn.close


#### parametrizar id saidas
#### conferir e definir layouts finais
#### cabeçalho dos relatorios
#### mensagens de erro
#### barra de progresso

#### alterar id das saidas para 8 ou 9
