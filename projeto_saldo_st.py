import sqlite3
import tkinter as tk
import pandas as pd
from tkinter import filedialog,ttk
import os
import logging
from tkinter import messagebox
import getpass
import re

try:
    caminho = r"C:\temp"
    if not(os.path.exists(caminho)):
        os.mkdir(caminho)

    logging.basicConfig(filename=r'C:\temp\logfile.log', level=logging.DEBUG,
                        format='%(asctime)s - %(levelname)s - %(message)s')
except Exception as e:
    logging.error(str(e), exc_info=True)

## IMPORTAÇÃO ##
def importa_entradas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cxn.execute("""DROP TABLE IF EXISTS ENTRADAS_3C""")
    #cxn.execute("""CREATE TABLE ENTRADAS_3C""")

    try:
        filename_entrada = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                                      filetypes=(("Text files", "*.*"), ("all files", "*.*")))

        wb2 = pd.read_excel(filename_entrada, sheet_name='Entradas')
        import_time = pd.Timestamp.now()
        usuario = getpass.getuser()
        wb2['Usuário'] = usuario
        wb2['DataHoraImportacao'] = import_time


        for index, row in wb2.iterrows():
            primeiro_digito = str(row[0])[:2]
            coluna_14 = row[13]
            coluna_15 = row[14]
            coluna_00 = row[0]
            coluna_02 = row[2]
            coluna_03 = row[3]
            coluna_08 = row[8]
            coluna_10 = row[10]
            coluna_11 = row[11]

            wb2['Data Lançamento'] = wb2['Data Lançamento'].apply(transforma_data)


            if primeiro_digito in ('01', '02', '18', '21') and (
                    coluna_14 == 0 or coluna_15 == 0 or pd.isnull(coluna_14) or pd.isnull(coluna_15)):
                mensagem_erro()
                return

            if coluna_00 != "" and (
                    coluna_02 == 0 or coluna_03 == 0 or coluna_08 == 0 or coluna_10 == 0 or coluna_11 == 0
                    or pd.isnull(coluna_02) or pd.isnull(coluna_03) or pd.isnull(coluna_08) or pd.isnull(coluna_10)or pd.isnull(coluna_11)):
                mensagem_erro()
                return


        wb2.to_sql(name='ENTRADAS_3C', con=cxn, if_exists='append', index=False)

        wb3 = pd.read_excel(filename_entrada, sheet_name='Entradas')
        import_time = pd.Timestamp.now()
        wb3['DataHoraImportacao'] = import_time
        wb3['Usuário'] = usuario
        wb3['Data Lançamento'] = wb3['Data Lançamento'].apply(transforma_data)
        wb3.to_sql(name='ENTRADAS_3C', con=cxn_consolidado, if_exists='append', index=False)

        preenche_coluna_mes_ano()

        cxn_consolidado.commit()
        cxn.commit()
        cxn.close()
        cxn_consolidado.close()
        mensagem_importa_sucesso()

    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()
def transforma_data(data):
    # Verifica se a data está no formato "dd/mm/aaaa" usando uma expressão regular
    if pd.notnull(data) and re.match(r'\d{4}-\d{2}-\d{2}', str(data)):
        # Transforma a data no formato "dd.mm.aaaa"
        return pd.to_datetime(data, format='%d/%m/%Y').strftime('%d.%m.%Y')
    else:
        return data
def preenche_coluna_mes_ano():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor_consolidado = cxn_consolidado.cursor()
    try:
        cursor.execute("""ALTER TABLE ENTRADAS_3C
                                ADD COLUMN MES_ANO TEXT; """)
        cursor_consolidado.execute("""ALTER TABLE ENTRADAS_3C
                                ADD COLUMN MES_ANO TEXT; """)
        cursor.execute("""UPDATE entradas_3c
                        SET MES_ANO = SUBSTR("Data Lançamento", -7)""")
        cursor_consolidado.execute("""UPDATE entradas_3c
                        SET MES_ANO = SUBSTR("Data Lançamento", -7)""")
    except:
        cursor.execute("""UPDATE entradas_3c
                        SET MES_ANO = SUBSTR("Data Lançamento", -7)""")
        cursor_consolidado.execute("""UPDATE entradas_3c
                        SET MES_ANO = SUBSTR("Data Lançamento", -7)""")
    cxn_consolidado.commit()
    cxn.commit()
def importa_saidas():
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    try:
        filename_saida = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                                    filetypes=(("Text files", "*.*"), ("all files", "*.*")))

        wb4 = pd.read_excel(filename_saida)
        import_time = pd.Timestamp.now()
        wb4['DataHoraImportacao'] = import_time
        wb4.to_sql(name='SAIDAS_3C', con=cxn_consolidado, if_exists='append', index=False)
        # cursor_consolidado.execute("""create table SAIDAS_RESUMIDA AS SELECT *
        #     FROM saidas_3c
        #     GROUP BY material, "Tipo de Avaliação", Centro, CFOP;
        #     """)
        # cxn_consolidado.commit()
        logging.info('Arquivo de saida importado no sistema')
        trata_saidas()
        mensagem_importa_sucesso()
    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()
def trata_saidas():
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cxn_consolidado.execute("""ALTER TABLE ENTRADAS_3C
                                ADD COLUMN SALDO_QTD INTERGER; """)
    cxn_consolidado.execute("""ALTER TABLE SAIDAS_3C
                                ADD COLUMN UNITARIO_ST REAL;""")
    cxn_consolidado.execute("""ALTER TABLE SAIDAS_3C
                                ADD COLUMN TOTAL_ST INTERGER; """)
    cxn_consolidado.execute("""ALTER TABLE SAIDAS_3C
                                ADD COLUMN SALDO_QTD INTERGER; """)
    cxn_consolidado.execute("""UPDATE ENTRADAS_3C
                            SET SALDO_QTD = Quantidade""")
    cxn_consolidado.execute("""UPDATE SAIDAS_3C
                            SET SALDO_QTD = Quantidade""")
    cxn_consolidado.execute("""ALTER TABLE ENTRADAS_3C
                            ADD COLUMN ICMS_SUPORTADO INTERGER; """)
    cxn_consolidado.execute("""UPDATE ENTRADAS_3C
                            SET ICMS_SUPORTADO = ("Valor ICMS" + "Valor ICMS ST") / Quantidade;""")
    cxn_consolidado.execute("""DELETE FROM ENTRADAS_3C
                            WHERE SUBSTRING(`ID do Cenário`, 1, 2) NOT IN ('01', '02', '03', '04', '09');""")
    cxn_consolidado.commit()
def importa_devolucoes():
    cxn_consolidado = sqlite3.connect(
            r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    try:
        filename_devolucao = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                                        filetypes=(("Text files", "*.*"), ("all files", "*.*")))

        wb4 = pd.read_excel(filename_devolucao)
        import_time = pd.Timestamp.now()
        wb4['DataHoraImportacao'] = import_time
        wb4.to_sql(name='DEVOLUCOES_3C', con=cxn_consolidado, if_exists='append', index=False)
        logging.info('Arquivo de devoluções'
                     ' importado no sistema')
        trata_devolucoes()
        mensagem_importa_sucesso()

    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()
def importa_saldo_inicial():
    cxn_consolidado = sqlite3.connect(
            r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    try:
        filename_devolucao = filedialog.askopenfilename(initialdir="/home", title="Select a File",
                                                        filetypes=(("Text files", "*.*"), ("all files", "*.*")))

        wb4 = pd.read_excel(filename_devolucao)
        import_time = pd.Timestamp.now()
        wb4['DataHoraImportacao'] = import_time
        wb4.to_sql(name='SALDO_INICIAL', con=cxn_consolidado, if_exists='replace', index=False)
        logging.info('Arquivo de devoluções'
                     ' importado no sistema')
        mensagem_importa_sucesso()

    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()
def trata_devolucoes():
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cxn_consolidado.execute("""ALTER TABLE DEVOLUCOES_3C
                                ADD COLUMN UNITARIO_ST REAL;""")
    cxn_consolidado.execute("""ALTER TABLE DEVOLUCOES_3C
                                ADD COLUMN TOTAL_ST INTERGER; """)

## EXCLUSÃO ##
def get_entries():
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal

    dialog = tk.Toplevel()
    dialog.title("Exclusão")
    dialog.geometry("400x200")

    # Variáveis para armazenar os valores inseridos pelo usuário
    centro_var = tk.StringVar()
    periodo_var = tk.StringVar()
    id_var = tk.StringVar()

    # Estilo ttk para os widgets
    style = ttk.Style()
    style.configure("TLabel", font=("Helvetica", 10), foreground="#333")
    style.configure("TEntry", font=("Helvetica", 10), padding=5)
    style.configure("TButton", font=("Helvetica", 10), foreground="white", background="#007acc")

    # Label e campo de entrada para o Centro
    centro_label = ttk.Label(dialog, text="Digite o Centro:")
    centro_label.pack(pady=5)
    centro_entry = ttk.Entry(dialog, textvariable=centro_var)
    centro_entry.pack()

    # Label e campo de entrada para o ID
    id_label = ttk.Label(dialog, text="Digite o ID e sua descrição")
    id_label.pack(pady=5)
    id_entry = ttk.Entry(dialog, textvariable=id_var)
    id_entry.pack()

    # Botão OK para confirmar a entrada
    ok_button = ttk.Button(dialog, text="OK", command=dialog.destroy)
    ok_button.pack(pady=25)

    dialog.wait_window(dialog)  # Esperar até que a janela seja fechada

    centro = centro_var.get()
    periodo = periodo_var.get()
    id = id_var.get()

    return int(centro), id
def exclui_dados_entradas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor_consolidado = cxn_consolidado.cursor()
    entrada_filial, entrada_id = get_entries()

    try:
        cursor.execute("DELETE FROM ENTRADAS_3C WHERE Centro = ? AND 'ID do Cenário' = ?", (entrada_filial, entrada_id,))
        cursor_consolidado.execute('DELETE FROM ENTRADAS_3C WHERE Centro = ? AND "ID do Cenário" = ?', (entrada_filial, entrada_id,))

        cxn.commit()
        cxn_consolidado.commit()

        rows_deleted = cursor_consolidado.rowcount
        if rows_deleted == 0:
            logging.info(
                f'Nada foi excluído. Revise os dados e tente novamente')
            mensagem_erro_exclusao()
        else:
            logging.info(
                f'Arquivo de entrada da filial {entrada_filial} com o ID {entrada_id} foi excluído do sistema')
            messagebox.showinfo("Exclusão", f'Arquivo de entrada da filial {entrada_filial} com o ID {entrada_id} foi excluído do sistema')
    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()

## EXPORTAÇÃO ##
def planilha_modelo_template_entradas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor.execute("""DROP TABLE IF EXISTS modelo_template_entradas;""")

    cursor.execute("""CREATE TABLE modelo_template_entradas AS 
        SELECT 
            "ID do Cenário" AS "CodigoCenario",
            "Data Lançamento" AS "Data",
            "Material",
            "Tipo de Avaliação" AS "TipoAvaliacao",
            "Docnum",
            "Empresa",
            "Centro" AS "CodigoCentro",
            "Divisão" AS "Divisao",
            "Valor ICMS" AS ValorICMS,
            CASE 
                WHEN substr("ID do Cenário", 1, 2) IN ('01', '02') THEN "" 
                ELSE "Valor ICMS ST" 
            END AS "ValorICMSST",
            "Valor IPI" AS ValorIPI
        FROM ENTRADAS_3C;
        """)
    cxn.commit()
    df = pd.read_sql("select * from modelo_template_entradas", cxn)
    df.to_excel(r"C:\TEMP\planilha_modelo_template_entradas.xlsx", index=False)
    logging.info('planilha template entradas exportada')
    mensagem_exporta_sucesso()
def planilha_modelo_template_saidas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor.execute("""DROP TABLE IF EXISTS modelo_template_saidas;""")

    cursor.execute("""CREATE TABLE modelo_template_saidas (
            "CodigoCenario" TEXT,
            "Data" TEXT,
            "Material" TEXT,
            "TipoAvaliacao" TEXT,
            "Docnum" TEXT,
            "Empresa" TEXT,
            "CodigoCentro" TEXT,
            "Divisao" TEXT,
            ValorICMS REAL,
			ValorICMSST REAL,
            ValorIPI TEXT)""")
    cursor.execute("""INSERT INTO MODELO_TEMPLATE_SAIDAS (CodigoCenario,Data, MATERIAL, TIPOAVALIACAO, EMPRESA, CODIGOCENTRO)
    SELECT CASE WHEN substr(CFOP,1,1) = '5' THEN 'INTERNA' ELSE 'INTERSTADUAL' END,strftime('%d.%m.%Y', 'now','start of month', '-1 day'), MATERIAL, "Tipo de Avaliação", EMPRESA, CENTRO 
    FROM SAIDAS_3C
    GROUP BY MATERIAL""")
    cxn.commit()
    df = pd.read_sql("select * from modelo_template_saidas", cxn)
    df.to_excel(r"C:\TEMP\planilha_modelo_template_saidas.xlsx", index=False)
    logging.info('planilha template devolucoes exportada')
    mensagem_exporta_sucesso()
def planilha_modelo_template_devolucoes():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor.execute("""DROP TABLE IF EXISTS modelo_template_devolucoes;""")

    cursor.execute("""CREATE TABLE modelo_template_devolucoes (
            "CodigoCenario" TEXT,
            "Data" TEXT,
            "Material" TEXT,
            "TipoAvaliacao" TEXT,
            "Docnum" TEXT,
            "Empresa" TEXT,
            "CodigoCentro" TEXT,
            "Divisao" TEXT,
            ValorICMS REAL
			ValorICMSST REAL,
            ValorIPI TEXT)""")
    cxn.commit()
    df = pd.read_sql("select * from modelo_template_devolucoes", cxn)
    df.to_excel(r"C:\TEMP\planilha_modelo_template_devolucoes.xlsx", index=False)
    logging.info('planilha template devolucoes exportada')
    mensagem_exporta_sucesso()

## MENSAGENS ##
def mensagem_exporta_sucesso():
        messagebox.showinfo("Exportar para Excel", "Planilhas Geradas com Sucesso !")
def mensagem_exclui_sucesso():
        messagebox.showinfo("Exclusão", "Dados Excluídos com Sucesso!")
def mensagem_importa_sucesso():
        messagebox.showinfo("Importação", "Planilha importada com Sucesso !!")
def mensagem_processamento_sucesso():
        messagebox.showinfo("Processado !", "Dados processados com sucesso !!")
def mensagem_erro():
    messagebox.showinfo("Erro!", "Ação não executada. Verifique o preenchimento da planilha e tente novamente !")
def mensagem_erro_exclusao():
    messagebox.showinfo("Erro!", "Nada foi excluído. Verifique as entradas e tente novamente")
def exporta_consolidado():
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()

    planilha_excel = pd.read_excel(r'C:\Users\abimaelsoares\Desktop\[S_4-E001] Estrutura Organizacional Completa - Centros - Hana (1).xlsx',sheet_name='Estrutura as Unidades - DEPARA')
    planilha_excel.to_sql(name='tabela_externa_organizacao', con=cxn_consolidado, if_exists='replace', index=False)

    try:
        cursor_consolidado.execute("""ALTER TABLE ENTRADAS_3C ADD COLUMN LOCAL_NEGOCIOS TEXT;""")
    except:
        pass

    cursor_consolidado.execute("""UPDATE "ENTRADAS_3C"
            SET LOCAL_NEGOCIOS = (
                SELECT "Unnamed: 11"
                FROM "tabela_externa_organizacao"
                WHERE ENTRADAS_3C.Centro = "tabela_externa_organizacao"."Unnamed: 14" );""")

    cxn_consolidado.commit()

    df = pd.read_sql("select * from ENTRADAS_3C", cxn_consolidado)
    df.to_excel("C:\TEMP\ICMS_ST_APROPRIAR.xlsx", index=False)
    logging.info('planilha exportada')
    mensagem_exporta_sucesso()

## TRATA DADOS ##
def trata_dados():
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    try:
        cursor_consolidado.execute("""DROP TABLE SALDO_MES""")
        cursor_consolidado.execute("""DROP TABLE RESUMO_POR_EMPRESA""")
        cursor_consolidado.execute("""DROP TABLE RESUMO_POR_CENTRO""")
        cursor_consolidado.execute("""DROP TABLE RESUMO_POR_MATERIAL""")
    except:
        pass
    cursor_consolidado.execute("""CREATE TABLE SALDO_MES_2 (
    Material TEXT,
    DESCRICAO_MATERIAL TEXT,
    TIPO_AVALIACAO TEXT,
    Empresa TEXT,
    Centro TEXT,
	UM TEXT )""")
    cursor_consolidado.execute("""INSERT INTO SALDO_MES_2 (Material, DESCRICAO_MATERIAL, TIPO_AVALIACAO, Empresa, Centro, UM)
    SELECT MATERIAL, DESCRICAO_MATERIAL, TIPO_AVALIACAO, EMPRESA, CENTRO, UM
    FROM SALDO_INICIAL
    GROUP BY Material, DESCRICAO_MATERIAL, TIPO_AVALIACAO, Empresa, Centro, UM;""")
    cursor_consolidado.execute("""
    INSERT INTO SALDO_MES_2 (Material, DESCRICAO_MATERIAL, TIPO_AVALIACAO, Empresa, Centro,UM)
    SELECT Material, "Descrição Material", "Tipo de Avaliação", Empresa, Centro,UM
    FROM ENTRADAS_3C
    GROUP BY Material, "Tipo de Avaliação", Centro;""")
    cursor_consolidado.execute("""CREATE TABLE SALDO_MES (
    Material TEXT,
    DESCRICAO_MATERIAL TEXT,
    TIPO_AVALIACAO TEXT,
    Empresa TEXT,
    Centro TEXT,
	UM TEXT,
    Quantidade_SI REAL,
    TOTAL_ST_SI REAL,
	Quantidade_ENTRADAS_3C REAL,
    TOTAL_ST_ENTRADAS_3C REAL,
	UNITARIO_ST_MES REAL,
	Quantidade_SAIDAS_3C REAL,
	TOTAL_ST_SAIDAS_3C REAL,
	Quantidade_DEVOLUCOES_3C REAL,
	TOTAL_ST_DEVOLUCOES_3C REAL,
	SALDO_QUANTIDADE_MES REAL,
    TOTAL_ST_MES REAL)""")
    cursor_consolidado.execute("""
    INSERT INTO SALDO_MES (Material, DESCRICAO_MATERIAL, TIPO_AVALIACAO, Empresa, Centro,UM)
    SELECT MATERIAL, DESCRICAO_MATERIAL, TIPO_AVALIACAO, EMPRESA, CENTRO,UM
    FROM SALDO_MES_2
    GROUP BY Material, TIPO_AVALIACAO, Centro;""")
    cursor_consolidado.execute("""DROP TABLE SALDO_MES_2""")
    cursor_consolidado.execute("""
        UPDATE SALDO_MES
        SET Quantidade_SI = (
            SELECT COALESCE(SUM(SI.Quantidade), 0)  -- Use COALESCE para tratar valores NULL
            FROM SALDO_INICIAL AS SI
            WHERE SALDO_MES.Material = SI.Material
                AND COALESCE(SALDO_MES.TIPO_AVALIACAO, '') = COALESCE(SI.TIPO_AVALIACAO, '')
                AND SALDO_MES.Centro = SI.Centro
        )
    """)


    cursor_consolidado.execute("""
        UPDATE SALDO_MES
        SET Quantidade_ENTRADAS_3C = (
            SELECT COALESCE(SUM(E.Quantidade), 0)
            FROM ENTRADAS_3C AS E
            WHERE SALDO_MES.Material = E.Material
                AND COALESCE(SALDO_MES.TIPO_AVALIACAO, '') = COALESCE(E."Tipo de Avaliação", '')
                AND SALDO_MES.Centro = E.Centro
        )
    """)

    cursor_consolidado.execute("""
        UPDATE SALDO_MES
        SET Quantidade_SAIDAS_3C = (
            SELECT COALESCE(SUM(S.Quantidade), 0)
            FROM SAIDAS_3C AS S
            WHERE SALDO_MES.Material = S.Material
                AND COALESCE(SALDO_MES.TIPO_AVALIACAO, '') = COALESCE(S."Tipo de Avaliação", '')
                AND SALDO_MES.Centro = S.Centro
        )
    """)


    cursor_consolidado.execute("""
        UPDATE SALDO_MES
        SET Quantidade_DEVOLUCOES_3C = (
            SELECT COALESCE(SUM(D.Quantidade), 0)
            FROM DEVOLUCOES_3C AS D
            WHERE SALDO_MES.Material = D.Material
                AND COALESCE(SALDO_MES.TIPO_AVALIACAO, '') = COALESCE(D."Tipo de Avaliação", '')
                AND SALDO_MES.Centro = D.Centro
        )
    """)
    cursor_consolidado.execute("""UPDATE SALDO_MES
    SET TOTAL_ST_SI = (
        SELECT SUM(SI.TOTAL_ST)
        FROM SALDO_INICIAL AS SI
        WHERE SALDO_MES.Material = SI.Material
            AND (
                SALDO_MES.TIPO_AVALIACAO = SI.TIPO_AVALIACAO
                OR (SALDO_MES.TIPO_AVALIACAO IS NULL AND SI.TIPO_AVALIACAO IS NULL)
                )
            AND SALDO_MES.Centro = SI.Centro
    )
    WHERE EXISTS (
        SELECT 1
        FROM SALDO_INICIAL AS SI
        WHERE SALDO_MES.Material = SI.Material
            AND (
                SALDO_MES.TIPO_AVALIACAO = SI.TIPO_AVALIACAO
                OR (SALDO_MES.TIPO_AVALIACAO IS NULL AND SI.TIPO_AVALIACAO IS NULL)
                )
            AND SALDO_MES.Centro = SI.Centro);""")
    cursor_consolidado.execute("""
        UPDATE SALDO_MES
        SET TOTAL_ST_ENTRADAS_3C = (
            SELECT COALESCE(SUM(E."Valor ICMS"), 0) + COALESCE(SUM(E."Valor ICMS ST"), 0)
            FROM ENTRADAS_3C AS E
            WHERE SALDO_MES.Material = E.Material
                AND COALESCE(SALDO_MES.TIPO_AVALIACAO, '') = COALESCE(E."Tipo de Avaliação", '')
                AND SALDO_MES.Centro = E.Centro
        )""")

    cursor_consolidado.execute("""UPDATE SALDO_MES
    SET UNITARIO_ST_MES = (COALESCE(TOTAL_ST_SI, 0) + COALESCE(TOTAL_ST_ENTRADAS_3C, 0)) / (COALESCE(Quantidade_SI, 0) + COALESCE(Quantidade_ENTRADAS_3C, 0));""")
    cursor_consolidado.execute("""UPDATE SALDO_MES
    SET SALDO_QUANTIDADE_MES = (COALESCE(Quantidade_SI, 0) + COALESCE(Quantidade_ENTRADAS_3C, 0)) + (COALESCE(Quantidade_DEVOLUCOES_3C, 0) - COALESCE(Quantidade_SAIDAS_3C, 0));""")
    cursor_consolidado.execute("""UPDATE SALDO_MES
    SET TOTAL_ST_MES = (COALESCE(SALDO_QUANTIDADE_MES, 0) * COALESCE(UNITARIO_ST_MES, 0))""")
    cursor_consolidado.execute("""UPDATE SALDO_MES
        SET TOTAL_ST_SAIDAS_3C = (COALESCE(QUANTIDADE_SAIDAS_3C, 0) * COALESCE(UNITARIO_ST_MES, 0) );""")
    cursor_consolidado.execute("""UPDATE SALDO_MES
        SET TOTAL_ST_DEVOLUCOES_3C = (COALESCE(QUANTIDADE_DEVOLUCOES_3C, 0) * COALESCE(UNITARIO_ST_MES, 0) );""")
    cursor_consolidado.execute("""UPDATE SALDO_MES
        SET UNITARIO_ST_MES = (COALESCE(TOTAL_ST_SI, 0) + COALESCE(TOTAL_ST_ENTRADAS_3C, 0)) / (COALESCE(Quantidade_SI, 0) + COALESCE(Quantidade_ENTRADAS_3C, 0));""")


    cxn_consolidado.commit()
    cxn_consolidado.close()
    try:
        writer = pd.ExcelWriter('Resumo ICMS ST a Apropriar.xlsx', engine='xlsxwriter')
        df = pd.read_sql("SELECT EMPRESA,CENTRO,MATERIAL,DESCRICAO_MATERIAL,TIPO_AVALIACAO,UNITARIO_ST_MES,SUM(QUANTIDADE_SI)"
                         "SUM(TOTAL_ST_SI) AS ST_APROPRIAR_SI, SUM(Quantidade_ENTRADAS_3C) AS QUANTIDADE_ENTRADAS,"
                         "SUM(TOTAL_ST_ENTRADAS_3C) AS ST_ENTRADAS,SUM(Quantidade_SAIDAS_3C)*-1 AS QUANTIDADE_SAIDAS,"
                         "SUM(TOTAL_ST_SAIDAS_3C) AS ST_SAIDAS, SUM(Quantidade_DEVOLUCOES_3C) AS QUANTIDADE_DEVOLUCOES,"
                         "SUM(TOTAL_ST_DEVOLUCOES_3C) AS ST_DEVOLUCOES, SUM(TOTAL_ST_MES) AS SALDO_ST_APROPRIAR FROM SALDO_MES GROUP BY EMPRESA,CENTRO", cxn_consolidado)
        df.to_excel(writer, index=False, sheet_name="RESUMO POR MATERIAL")
        df2 = pd.read_sql("SELECT EMPRESA, SUM(TOTAL_ST_SI) AS ST_APROPRIAR_SI, "
                          "SUM(TOTAL_ST_ENTRADAS_3C) AS ST_ENTRADAS,SUM(TOTAL_ST_SAIDAS_3C)*-1 AS ST_SAIDAS, "
                          "SUM(TOTAL_ST_DEVOLUCOES_3C) AS ST_DEVOLUCOES, SUM(TOTAL_ST_MES) AS ST_APROPRIAR FROM SALDO_MES GROUP BY EMPRESA",
                          cxn_consolidado)
        df2.to_excel(writer, index=False, sheet_name="RESUMO POR EMPRESA")
        df3 = pd.read_sql(
            "SELECT EMPRESA,CENTRO, SUM(TOTAL_ST_SI) AS ST_APROPRIAR_SI, "
                          "SUM(TOTAL_ST_ENTRADAS_3C) AS ST_ENTRADAS,SUM(TOTAL_ST_SAIDAS_3C)*-1 AS ST_SAIDAS, "
                          "SUM(TOTAL_ST_DEVOLUCOES_3C) AS ST_DEVOLUCOES, SUM(TOTAL_ST_MES) AS ST_APROPRIAR FROM SALDO_MES GROUP BY EMPRESA, CENTRO",
            cxn_consolidado)
        df3.to_excel(writer, index=False, sheet_name="RESUMO POR CENTRO")
        writer.save()
        mensagem_exporta_sucesso()
    except Exception as e:
        logging.error(str(e), exc_info=True)


    mensagem_processamento_sucesso()
def encerrar_mes():
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    cursor_consolidado.execute("""DELETE FROM SALDO_INICIAL;""")
    cursor_consolidado.execute("""INSERT INTO SALDO_INICIAL (MATERIAL, DESCRICAO_MATERIAL, TIPO_AVALIACAO, EMPRESA, CENTRO, QUANTIDADE, UM, TOTAL_ST)
    SELECT MATERIAL, DESCRICAO_MATERIAL, TIPO_AVALIACAO, EMPRESA, CENTRO, SALDO_QUANTIDADE_MES, UM, TOTAL_ST_MES
    FROM SALDO_MES;""")
    cursor_consolidado.execute("""UPDATE SALDO_INICIAL
    SET UNIT_ST = TOTAL_ST/Quantidade""")
    cursor_consolidado.execute("""DROP TABLE SALDO_MES""")
    cursor_consolidado.execute("""DROP TABLE ENTRADAS_3C""")
    cursor_consolidado.execute("""DROP TABLE SAIDAS_3C""")
    cursor_consolidado.execute("""DROP TABLE DEVOLUCOES_3C""")

def depuracao():
    cxn_consolidado = sqlite3.connect(
        r'\\10.85.1.22\dir_fin_control\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()
    writer = pd.ExcelWriter('Resumo ICMS ST a Apropriar.xlsx', engine='xlsxwriter')
    df = pd.read_sql("SELECT * FROM SALDO_MES",
                     cxn_consolidado)
    df.to_excel(writer, index=False, sheet_name="RESUMO POR MATERIAL")
    df2 = pd.read_sql("SELECT EMPRESA, SUM(TOTAL_ST_SI) AS ST_APROPRIAR_SI, "
                      "SUM(TOTAL_ST_ENTRADAS_3C) AS ST_ENTRADAS,SUM(TOTAL_ST_SAIDAS_3C)*-1 AS ST_SAIDAS, "
                      "SUM(TOTAL_ST_DEVOLUCOES_3C) AS ST_DEVOLUCOES, SUM(TOTAL_ST_MES) AS ST_APROPRIAR FROM SALDO_MES GROUP BY EMPRESA",
                      cxn_consolidado)
    df2.to_excel(writer, index=False, sheet_name="RESUMO POR EMPRESA")
    df3 = pd.read_sql(
        "SELECT EMPRESA,CENTRO, SUM(TOTAL_ST_SI) AS ST_APROPRIAR_SI, "
                      "SUM(TOTAL_ST_ENTRADAS_3C) AS ST_ENTRADAS,SUM(TOTAL_ST_SAIDAS_3C)*-1 AS ST_SAIDAS, "
                      "SUM(TOTAL_ST_DEVOLUCOES_3C) AS ST_DEVOLUCOES, SUM(TOTAL_ST_MES) AS ST_APROPRIAR FROM SALDO_MES GROUP BY EMPRESA, CENTRO",
        cxn_consolidado)
    df3.to_excel(writer, index=False, sheet_name="RESUMO POR CENTRO")
    writer.save()


if __name__ == "__main__":
    #trata_dados()
    #depuracao()
    #importa_saldo_inicial()
    #exporta_consolidado()
    #importa_saidas()
    pass

