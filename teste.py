import sqlite3

cxn_consolidado = sqlite3.connect(
    r'X:\CONTROLADORIA\COMPLIANCE FISCAL\APURAÇÃO & CONCILIAÇÃO FISCAL\CONTROLES\Saldos Contábeis\MR22\bd_saldo_icmsst.db')
cursor_consolidado = cxn_consolidado.cursor()

cursor_consolidado.execute("""UPDATE saldo_final
    SET "Valor ICMS" = "Valor ICMS" - COALESCE(
        (SELECT SUM(s."icms1" + s."st1")
         FROM saidas_3c s
         WHERE s."Material1" = saldo_final."Material"
           AND s."centro1" = saldo_final."centro",
        0
    );
    """)