import sqlite3
import pandas as pd
from tkinter import filedialog, simpledialog
import os
import logging
from pathlib import Path
import openpyxl
import xlsxwriter
from tkinter import messagebox

try:
    caminho = r"C:\temp"
    if not(os.path.exists(caminho)):
        os.mkdir(caminho)

    logging.basicConfig(filename=r'C:\temp\logfile.log', level=logging.DEBUG,
                        format='%(asctime)s - %(levelname)s - %(message)s')


    # cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    # cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    # cursor = cxn.cursor()
    # cursor_consolidado = cxn_consolidado.cursor()
except Exception as e:
    logging.error(str(e), exc_info=True)

## IMPORTAÇÃO ##
def importa_saidas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    try:
        filename_saida = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                              filetypes=(("Text files", "*.*"), ("all files", "*.*")))
        wb1 = pd.read_excel(filename_saida, sheet_name='Análise 1')
        import_time = pd.Timestamp.now()
        wb1['DataHoraImportacao'] = import_time
        wb1.to_sql(name='SAIDAS_3C', con=cxn, if_exists='append', index=False)
        cxn.commit()
        logging.info('Arquivo de saida importado no sistema')

        wb4 = pd.read_excel(filename_saida, sheet_name='Análise 1')
        import_time = pd.Timestamp.now()
        wb4['DataHoraImportacao'] = import_time
        wb4.to_sql(name='SAIDAS_3C', con=cxn_consolidado, if_exists='append', index=False)
        cxn_consolidado.commit()
        logging.info('Arquivo de saida importado no sistema')
        mensagem_importa_sucesso()

    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()
def importa_entradas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor_consolidado = cxn_consolidado.cursor()

    try:
        filename_entrada = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                              filetypes=(("Text files", "*.*"), ("all files", "*.*")))

        wb2 = pd.read_excel(filename_entrada, sheet_name='Entradas')
        import_time = pd.Timestamp.now()
        wb2['DataHoraImportacao'] = import_time
        wb2.to_sql(name='ENTRADAS_3C', con=cxn, if_exists='append', index=False)

        wb3 = pd.read_excel(filename_entrada, sheet_name='Entradas')
        import_time = pd.Timestamp.now()
        wb3['DataHoraImportacao'] = import_time
        wb3.to_sql(name='ENTRADAS_3C', con=cxn_consolidado, if_exists='append', index=False)
        cria_coluna_ano()
        preenche_coluna_ano()




        # wb4 = pd.read_excel(r'C:\Users\abimaelsoares\Desktop\projeto_saldost\Entradas 02.2023.xlsx', sheet_name='Saldo Anterior')
        # import_time = pd.Timestamp.now()
        # wb4['DataHoraImportacao'] = import_time
        # wb4.to_sql(name='SALDO_ANTERIOR', con=cxn, if_exists='append', index=False)
        # logging.info('Arquivo de entrada importado no sistema')
        cxn_consolidado.commit()
        cxn.commit()
        cxn.close()
        cxn_consolidado.close()
        mensagem_importa_sucesso()

    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()
def cria_coluna_ano():
    try:
        cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
        cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
        cursor = cxn.cursor()
        cursor_consolidado = cxn_consolidado.cursor()
        cursor.execute("""ALTER TABLE ENTRADAS_3C
                                ADD COLUMN ANO TEXT; """)
        cursor_consolidado.execute("""ALTER TABLE ENTRADAS_3C
                                ADD COLUMN ANO TEXT; """)
        cxn_consolidado.commit()
        cxn.commit()
    except:
        pass
def preenche_coluna_ano():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(
        r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor_consolidado = cxn_consolidado.cursor()
    cursor.execute("""UPDATE entradas_3c
                    SET ANO = SUBSTR("Data Lançamento", -4)""")
    cursor_consolidado.execute("""UPDATE entradas_3c
                    SET ANO = SUBSTR("Data Lançamento", -4)""")
    cxn_consolidado.commit()
    cxn.commit()

## EXCLUSÃO ##
def exclui_dados_entradas():

    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor_consolidado = cxn_consolidado.cursor()
    entrada_filial = simpledialog.askstring("Entrada", "Digite o Centro:")*1
    entrada_periodo = simpledialog.askstring("Entrada", "Digite o Periodo:")*1
    entrada_ano = simpledialog.askstring("Entrada", "Digite o Ano:")
    try:
        cursor.execute((f"DELETE FROM ENTRADAS_3C WHERE Centro = {entrada_filial} AND PERÍODO = {entrada_periodo} AND ANO = {entrada_ano}"))
        cursor_consolidado.execute((f"DELETE FROM ENTRADAS_3C WHERE Centro = {entrada_filial} AND PERÍODO = {entrada_periodo} AND ANO = {entrada_ano}"))
        cxn.commit()
        cxn_consolidado.commit()
        logging.info(f'Arquivo de entrada da filial {entrada_filial} do período {entrada_periodo} e do ano {entrada_ano} foi excluido do sistema')
        caminho_arquivo = Path(r"C:\TEMP\planilha_modelo_template_entradas.xlsx")
        if caminho_arquivo.exists():
            caminho_arquivo.unlink()
            mensagem_exclui_sucesso()
    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()

## EXPORTAÇÃO ##
def planilha_modelo_template_entradas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor.execute("""DROP TABLE IF EXISTS modelo_template_entradas;""")
    try:
        try:
            cursor.execute("""CREATE table modelo_template_entradas AS SELECT "ID do Cenário", "Data Lançamento", "Material", 
            "Tipo de Avaliação","Docnum", "Empresa","Centro","Divisão", "Valor ICMS", CASE WHEN TIPO = "DESTACADO NA NF" THEN "Valor ICMS ST"=""  ELSE "Valor ICMS ST" end as "Valor ICMS ST",
			"Valor IPI" FROM ENTRADAS_3C  """)
            cxn.commit()
            df = pd.read_sql("select * from modelo_template_entradas", cxn)
            df.to_excel(r"C:\TEMP\planilha_modelo_template_entradas.xlsx", index=False)
            logging.info('planilha template entradas exportada')
            transforma_dados()
            mensagem_exporta_sucesso()

        except:
            df = pd.read_sql("select * from modelo_template_entradas", cxn)
            df.to_excel(r"C:\TEMP\planilha_modelo_template_entradas.xlsx", index=False)
            logging.info('planilha template entradas exportada')
            transforma_dados()
            mensagem_exporta_sucesso()
    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()
def obsoleto_planilha_modelo_template_saidas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(
        r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    try:
        try:
            cursor_consolidado.execute("""CREATE table modelo_template_saidas AS select 
	CASE
		WHEN substring(CFOP1,1,1)="5" THEN "05 - Saída Interna"
		ELSE "06 - Ressarcimento ICMS"
    END AS  "CodigoCenario", "Data de Lançamento" as "Data",Material1 as "Material","Tipo de Avaliação1" as "TipoAvaliacao",
	CASE
		WHEN CFOP1 = '6152' THEN ''
		WHEN CFOP1 = '5152' THEN ''
		WHEN CFOP1 = '5409' THEN ''
		WHEN CFOP1 = '6409' THEN ''
		ELSE Docnum1
		END AS Docnum,Empresa1 as "Empresa",Centro1 as "CodigoCentro","Divisão" as "Divisao",total_st_entrada as "ValorIcms",ICMS1 as "ValorICMSST","IPI1" AS "ValorIPI" 
		FROM saidas_sinteticas ORDER BY "tipo_contabilizacao" """)
            cxn.commit()
            df = pd.read_sql("select * from modelo_template_saidas", cxn)
            df.to_excel("C:\TEMP\planilha_modelo_template_saidas.xlsx", index=False)
            logging.info('planilha template saidas exportada')
            mensagem_exporta_sucesso()

        except:
            pass
            df = pd.read_sql("select * from modelo_template_saidas", cxn)
            df.to_excel("C:\TEMP\planilha_modelo_template_saidas.xlsx", index=False)
            logging.info('planilha template saidas exportada no except')
            mensagem_exporta_sucesso()


    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()

def planilha_modelo_template_saidas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(
        r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    cursor = cxn.cursor()
    cursor.execute("""DROP TABLE IF EXISTS modelo_template_saidas;""")
    try:
        try:
            cursor.execute("""create table modelo_template_saidas as SELECT "ID do Cenário" as CodigoCenario, "Data de Lançamento" as Data,
             "Material1" as Material, "Tipo de avaliação1" as TipoAvaliacao, Docnum1 as Docnum,
        empresa1 as Empresa, Centro1 as CodigoCentro, Divisao1 as Divisao, ICMS1 as ValorICMS,ST1 AS ValorICMSST, IPI1 as ValorIPI from saidas_3c""")
            cxn.commit()
            df = pd.read_sql("select * from modelo_template_saidas", cxn)
            df.to_excel("C:\TEMP\planilha_modelo_template_saidas.xlsx", index=False)
            logging.info('planilha template saidas exportada')
            mensagem_exporta_sucesso()

        except:
            pass
            df = pd.read_sql("select * from modelo_template_saidas", cxn)
            df.to_excel("C:\TEMP\planilha_modelo_template_saidas.xlsx", index=False)
            logging.info('planilha template saidas exportada no except')
            mensagem_exporta_sucesso()


    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()
def exportar_saldo_atual():
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    try:
        writer = pd.ExcelWriter('saldo_atual.xlsx', engine='xlsxwriter')
        df = pd.read_sql("select * from SALDO_ATUAL", cxn_consolidado)
        df.to_excel(writer, index=False, sheet_name="Saldo_atual_detalhado")
        df2 = pd.read_sql("SELECT EMPRESA, 'Saldo Qtd', 'ICMS ST Total Atualizado' FROM SALDO_ATUAL GROUP BY EMPRESA", cxn_consolidado)
        df2.to_excel(writer, index=False, sheet_name="saldo_atual_por_empresa")
        df3 = pd.read_sql(
            "SELECT EMPRESA,Centro, 'Saldo Qtd', 'ICMS ST Total Atualizado' FROM SALDO_ATUAL GROUP BY Centro, EMPRESA", cxn_consolidado)
        df3.to_excel(writer, index=False, sheet_name="saldo_atual_por_filial")
        writer.save()
        mensagem_exporta_sucesso()
    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()

## TRANSFORMAÇÃO E MODIFICAÇÃO ##
def transforma_dados():
    arquivo_entrada = r"C:\TEMP\planilha_modelo_template_entradas.xlsx"
    workbook = openpyxl.load_workbook(arquivo_entrada)
    nome_planilha = 'Sheet1'
    planilha = workbook[nome_planilha]

    novo_cabecalho = ["CodigoCenario","Data","Material","TipoAvaliacao","Docnum","Empresa","CodigoCentro","Divisao","ValorICMS","ValorICMSST","ValorIPI"]
    planilha['A1'].value = novo_cabecalho[0]
    planilha['B1'].value = novo_cabecalho[1]
    planilha['C1'].value = novo_cabecalho[2]
    planilha['D1'].value = novo_cabecalho[3]
    planilha['E1'].value = novo_cabecalho[4]
    planilha['F1'].value = novo_cabecalho[5]
    planilha['G1'].value = novo_cabecalho[6]
    planilha['H1'].value = novo_cabecalho[7]
    planilha['I1'].value = novo_cabecalho[8]
    planilha['J1'].value = novo_cabecalho[9]
    planilha['K1'].value = novo_cabecalho[10]


    workbook.save(arquivo_entrada)
def criar_coluna_tipo_contabilizacao_saidas():
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    try:
        cursor_consolidado.execute("""ALTER TABLE SAIDAS_3C ADD COLUMN tipo_contabilizacao""")
        tipo_contabilizacao_saidas()
    except:
        tipo_contabilizacao_saidas()
def tipo_contabilizacao_saidas():
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    cursor_consolidado.execute("""UPDATE SAIDAS_3C
    SET 
        tipo_contabilizacao = 
        CASE WHEN SUBSTRING(CFOP1, 1,1) = "5" THEN "SEM RESSARCIMENTO" 
    	ELSE "COM RESSARCIMENTO" END""")
def saldo_atual_provisorio():
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    try:
        cursor_consolidado.execute("""CREATE table saldo_atual_provisorio AS SELECT Empresa, Centro, Divisão, Material, "Descrição Material" as Descricao_Material,
        UM, SUM("Saldo Qtd") as Saldo_Qtd, SUM("ICMS ST Total Atualizado" + "Valor ICMS") AS total_st_bruto_atualizado, (SUM("ICMS ST Total Atualizado" + "Valor ICMS")/SUM("Saldo Qtd")) as Valor_unit_ST

        FROM(
        SELECT Empresa, Centro, Divisão, Material, "Descrição Material",UM,"ICMS ST Total Atualizado", "Saldo Qtd","Valor ICMS"  FROM SALDO_ANTERIOR
            UNION ALL
        SELECT Empresa, Centro, Divisão, Material, "Descrição Material",UM,"Valor ICMS ST", Quantidade, "Valor ICMS" FROM ENTRADAS_3C  WHERE TIPO = "CALCULADO NA ENTRADA" OR "DESTACADO NA NF"
        ) AS Total
        GROUP BY Material, Empresa, Centro, "Divisão" """)
        logging.info('saldo_atual_provisorio criado')
        cxn_consolidado.commit()
    except Exception as e:
        logging.error(str(e), exc_info=True)
def exclui_saldo_provisorio():
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    cursor_consolidado.execute("""drop table saldo_atual_provisorio""")

## adicionar logs ##

def sintetiza_dados():
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    try:
        cursor_consolidado.execute("""CREATE TABLE saidas_sinteticas AS SELECT SUM(Quantidade1) as saldo_saidas,*, AVG(Valor_unit_ST) 
        AS unit_st,(AVG(Valor_unit_ST))*(sum(Quantidade1)) as total_st_entrada 
        FROM saldo_atual_provisorio
        INNER JOIN SAIDAS_3C ON saldo_atual_provisorio.Material = SAIDAS_3C.Material1 AND
        saldo_atual_provisorio.Empresa = SAIDAS_3C.Empresa1 AND saldo_atual_provisorio.Centro = SAIDAS_3C.Centro1
        GROUP BY Docnum1,Material1,Empresa1,Centro1,CFOP1,"Tipo de Avaliação1" ,"tipo_contabilizacao" """)
        cxn_consolidado.commit()
    except Exception as e:
        logging.error(str(e), exc_info=True)
def saldo_consistido():
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    try:
        cursor_consolidado.execute("""create table SALDO_ATUAL as select sap.Empresa, sap.Centro,sap.Divisão,
        sap.Material,sap.Descricao_Material,sap.UM, (sap.Saldo_Qtd - sum(saldo_saidas)) AS "Saldo Qtd", 
        avg(sap.total_st_bruto_atualizado) as "ICMS ST Total Atualizado",
        AVG(sap.Valor_unit_ST) as "ICMS ST Unit Atualizado", "Valor ICMS"
        from saldo_atual_provisorio sap
        left join saidas_sinteticas on sap.Material = saidas_sinteticas.Material1 and sap.Empresa = saidas_sinteticas.Empresa and sap.Centro = saidas_sinteticas.Centro and sap."Divisão" = saidas_sinteticas.Divisão
        
        group by
            sap.Material,sap.Empresa,sap.Centro, sap.Divisão""")
        #exclui_saldo_provisorio()
    except:
        pass
def saldo_final_refatorado():
    cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    cursor_consolidado.execute("""CREATE TABLE saldo_final AS
    SELECT
        "Material",
        "TIPO",
        "Descrição Material",
        "Empresa",
        "Centro",
        "Divisão",
        SUM("Quantidade") AS "Quantidade",
        "UM",
        SUM("Valor ICMS") AS "Valor ICMS",
        SUM("Valor ICMS ST") AS "Valor ICMS ST",
        SUM("Valor IPI") AS "Valor IPI",
        "PERÍODO",
        "ANO"
    FROM entradas_3c
    WHERE "TIPO" <> "AJUSTE APURAÇÃO"
    GROUP BY "Material";
;""")
    cursor_consolidado.execute("""ALTER TABLE saldo_final ADD COLUMN vlr_unit_icms REAL;""")
    cursor_consolidado.execute("""ALTER TABLE saldo_final ADD COLUMN vlr_unit_st REAL;""")
    cursor_consolidado.execute("""UPDATE saldo_final SET vlr_unit_icms = "Valor ICMS" / "Quantidade";""")
    cursor_consolidado.execute("""UPDATE saldo_final SET vlr_unit_st = "Valor ICMS ST" / "Quantidade";""")
    cursor_consolidado.execute("""UPDATE saidas_3c
    SET icms1 = 
        CASE WHEN SUBSTR(CFOP1, 1, 1) = '5'
             THEN COALESCE(
                    (SELECT SUM(sf."vlr_unit_icms" + sf."vlr_unit_st") * saidas_3c.Quantidade1
                     FROM saldo_final sf
                     WHERE sf."Material" = saidas_3c."Material1"
                       AND sf."Centro" = saidas_3c."Centro1"),
                0)
         ELSE icms1
    END;""")
    cursor_consolidado.execute("""UPDATE saldo_final
        SET quantidade = quantidade - COALESCE(
            (SELECT SUM(s."quantidade1")
             FROM saidas_3c s
             WHERE s."Material1" = saldo_final."Material"
               AND s."centro1" = saldo_final."centro",
               AND SUBSTR(saidas_3c.CFOP1, 1, 1) = "6"),
            0
        );
        """)
    cursor_consolidado.execute("""UPDATE saldo_final
        SET "vALOR icms" = "vALOR icms" - COALESCE(
            (SELECT SUM(s."ICMS1")
             FROM saidas_3c s
             WHERE s."Material1" = saldo_final."Material"
               AND s."centro1" = saldo_final."centro"
               AND SUBSTR(s.CFOP1, 1, 1) = "6"),
            0
        );""")
    cursor_consolidado.execute("""UPDATE saldo_final
        SET "vALOR icms ST" = "vALOR icms ST" - COALESCE(
            (SELECT SUM(s."ST1")
             FROM saidas_3c s
             WHERE s."Material1" = saldo_final."Material"
               AND s."centro1" = saldo_final."centro"
               AND SUBSTR(s.CFOP1, 1, 1) = "6"),
            0
        );""")
    cursor_consolidado.execute("""UPDATE saldo_final SET vlr_unit_icms = "Valor ICMS" / "Quantidade";""")
    cursor_consolidado.execute("""UPDATE saldo_final SET vlr_unit_st = "Valor ICMS ST" / "Quantidade";""")

def planilha_modelo_template_saidas_internas():
    cxn_consolidado = sqlite3.connect(
        r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    cursor_consolidado.execute("""DROP TABLE IF EXISTS modelo_template_saidas;""")
    try:
        try:
            cursor_consolidado.execute("""create table modelo_template_saidas as SELECT "ID do Cenário" as CodigoCenario, "Data de Lançamento" as Data,
             "Material1" as Material, "Tipo de avaliação1" as TipoAvaliacao, Docnum1 as Docnum,
        empresa1 as Empresa, Centro1 as CodigoCentro, Divisao1 as Divisao, ICMS1 as ValorICMS,ST1 AS ValorICMSST, IPI1 as ValorIPI from saidas_3c""")
            cxn_consolidado.commit()
            df = pd.read_sql("select * from modelo_template_saidas", cxn_consolidado)
            df.to_excel("C:\TEMP\planilha_modelo_template_saidas.xlsx", index=False)
            logging.info('planilha template saidas exportada')
            mensagem_exporta_sucesso()
        except Exception as e:
            logging.error(str(e), exc_info=True)
            mensagem_erro()
    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()

def mensagem_exporta_sucesso():
        messagebox.showinfo("Exportar para Excel", "Planilhas Geradas com Sucesso !")
def mensagem_exclui_sucesso():
        messagebox.showinfo("Exclusão", "Dados Excluídos com Sucesso!")
def mensagem_importa_sucesso():
        messagebox.showinfo("Importação", "Planilha importada com Sucesso !!")
def mensagem_processamento_sucesso():
        messagebox.showinfo("Processado !", "Dados processados com sucesso !!")
def mensagem_erro():
    messagebox.showinfo("Erro!", "Ação não execultada. Contate o Administrador")

if __name__ == "__main__":
    pass
    # importa_entradas()
    # importa_saidas()
    # criar_coluna_tipo_contabilizacao_saidas()
    # saldo_atual_provisorio()
    # sintetiza_dados()
    # saldo_consistido()
    # planilha_modelo_template_entradas()
    # transforma_dados()
    # planilha_modelo_template_saidas()
    # exportar_saldo_atual()
    # importa_ressarcimento_TIMP()
    # exclui_dados_entradas()



