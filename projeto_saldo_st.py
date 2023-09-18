import sqlite3
import tkinter as tk
import pandas as pd
from tkinter import filedialog, simpledialog,dialog
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

def transforma_data(data):
    # Verifica se a data está no formato "dd/mm/aaaa" usando uma expressão regular
    if pd.notnull(data) and re.match(r'\d{4}-\d{2}-\d{2}', str(data)):
        # Transforma a data no formato "dd.mm.aaaa"
        return pd.to_datetime(data, format='%d/%m/%Y').strftime('%d.%m.%Y')
    else:
        return data
def importa_entradas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(
        r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cxn.execute("""DROP TABLE IF EXISTS ENTRADAS_3C""")

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


            if primeiro_digito in ('01', '02') and (
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
        wb3['Data Lançamento'] = wb3['Data Lançamento'].apply(transforma_data)
        wb3.to_sql(name='ENTRADAS_3C', con=cxn_consolidado, if_exists='append', index=False)

        cria_coluna_mes_ano()
        preenche_coluna_mes_ano()

        cxn_consolidado.commit()
        cxn.commit()
        cxn.close()
        cxn_consolidado.close()
        mensagem_importa_sucesso()

    except Exception as e:
        logging.error(str(e), exc_info=True)
        mensagem_erro()
def cria_coluna_mes_ano():
    try:
        cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
        cxn_consolidado = sqlite3.connect(r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
        cursor = cxn.cursor()
        cursor_consolidado = cxn_consolidado.cursor()
        cursor.execute("""ALTER TABLE ENTRADAS_3C
                                ADD COLUMN MES_ANO TEXT; """)
        cursor_consolidado.execute("""ALTER TABLE ENTRADAS_3C
                                ADD COLUMN MES_ANO TEXT; """)
        cxn_consolidado.commit()
        cxn.commit()
    except:
        pass
def preenche_coluna_mes_ano():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(
        r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor_consolidado = cxn_consolidado.cursor()
    cursor.execute("""UPDATE entradas_3c
                    SET MES_ANO = SUBSTR("Data Lançamento", -7)""")
    cursor_consolidado.execute("""UPDATE entradas_3c
                    SET MES_ANO = SUBSTR("Data Lançamento", -7)""")
    cxn_consolidado.commit()
    cxn.commit()


## EXCLUSÃO ##

def get_entries():
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal

    dialog = tk.Toplevel()
    dialog.title("Exclusão")

    # Variáveis para armazenar os valores inseridos pelo usuário
    centro_var = tk.StringVar()
    periodo_var = tk.StringVar()

    # Label e campo de entrada para o Centro
    centro_label = tk.Label(dialog, text="Digite o Centro:")
    centro_label.pack()
    centro_entry = tk.Entry(dialog, textvariable=centro_var)
    centro_entry.pack()

    # Label e campo de entrada para o Período
    periodo_label = tk.Label(dialog, text="Digite o Período no formato MM.AAAA:")
    periodo_label.pack()
    periodo_entry = tk.Entry(dialog, textvariable=periodo_var)
    periodo_entry.pack()

    # Botão OK para confirmar a entrada
    ok_button = tk.Button(dialog, text="OK", command=dialog.destroy)
    ok_button.pack()

    dialog.wait_window(dialog)  # Esperar até que a janela seja fechada

    centro = centro_var.get()
    periodo = periodo_var.get()

    return centro, periodo
def exclui_dados_entradas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cxn_consolidado = sqlite3.connect(
        r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor_consolidado = cxn_consolidado.cursor()

    entrada_filial, entrada_periodo = get_entries()

    try:
        cursor.execute("DELETE FROM ENTRADAS_3C WHERE Centro = ? AND MES_ANO = ?", (entrada_filial, entrada_periodo))
        cursor_consolidado.execute("DELETE FROM ENTRADAS_3C WHERE Centro = ? AND MES_ANO = ?",
                                   (entrada_filial, entrada_periodo))
        cxn.commit()
        cxn_consolidado.commit()

        rows_deleted = cursor_consolidado.rowcount
        if rows_deleted == 0:
            logging.info(
                f'Nada foi excluído. Revise os dados e tente novamente')
            mensagem_erro_exclusao()
        else:
            logging.info(
                f'Arquivo de entrada da filial {entrada_filial} no período {entrada_periodo} foi excluído do sistema')
            mensagem_exclui_sucesso()
    except Exception as e:
        logging.error(str(e), exc_info=True)
        #mensagem_erro()

## EXPORTAÇÃO ##
def planilha_modelo_template_entradas():
    cxn = sqlite3.connect(r'C:\TEMP\bd_saldo_icmsst.db')
    cursor = cxn.cursor()
    cursor.execute("""DROP TABLE IF EXISTS modelo_template_entradas;""")
    try:
        try:
            cursor.execute("""CREATE table modelo_template_entradas AS SELECT "ID do Cenário" AS "CodigoCenario", "Data Lançamento" AS "Data", "Material", 
            "Tipo de Avaliação" AS "TipoAvaliacao","Docnum", "Empresa","Centro" AS "CodigoCentro","Divisão" AS "Divisao", "Valor ICMS" AS "ValorICMS",
             CASE WHEN substr("ID do Cenário",1,2) = "01" OR substr("ID do Cenário",1,2) = "02" THEN "Valor ICMS ST" AS "ValorICMSST" =""  ELSE "Valor ICMS ST" end as "ValorICMSST",
			"Valor IPI" AS "ValorIPI" FROM ENTRADAS_3C  """)
            cxn.commit()
            df = pd.read_sql("select * from modelo_template_entradas", cxn)
            df.to_excel(r"C:\TEMP\planilha_modelo_template_entradas.xlsx", index=False)
            logging.info('planilha template entradas exportada')
            mensagem_exporta_sucesso()


        except:
            df = pd.read_sql("select * from modelo_template_entradas", cxn)
            df.to_excel(r"C:\TEMP\planilha_modelo_template_entradas.xlsx", index=False)
            logging.info('planilha template entradas exportada')
            mensagem_exporta_sucesso()
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
    messagebox.showinfo("Erro!", "Ação não executada. Verifique o preenchimento da planilha e tente novamente !")
def mensagem_erro_exclusao():
    messagebox.showinfo("Erro!", "Nada foi excluído. Verifique as entradas e tente novamente")


def exporta_consolidado():
    cxn_consolidado = sqlite3.connect(
        r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
    cursor_consolidado = cxn_consolidado.cursor()

    planilha_excel = pd.read_excel(r'C:\Users\abimaelsoares\Desktop\[S_4-E001] Estrutura Organizacional Completa - Centros - Hana (1).xlsx',sheet_name='Estrutura as Unidades - DEPARA')
    planilha_excel.to_sql(name='tabela_externa_organizacao', con=cxn_consolidado, if_exists='replace', index=False)

    try:
        cursor_consolidado.execute("""ALTER TABLE ENTRADAS_3C
        ADD COLUMN IF NOT EXISTS LOCAL_NEGOCIOS text;""")
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



if __name__ == "__main__":
    pass
    importa_entradas()
    # exclui_dados_entradas()
    # planilha_modelo_template_entradas()
    # exporta_consolidado()




