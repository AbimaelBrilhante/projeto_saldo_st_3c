import tkinter as tk
from tkinter import ttk
import projeto_saldo_st
from tkinter import *
import tkinter as tk
from tkinter import messagebox


class Single_window(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Banco de Impostos ST')
        self.state('zoomed')

def mensagem_exporta():
    messagebox.showinfo("Exportar para Excel", "Planilhas Geradas com Sucesso !")

def mensagem_importa():
    messagebox.showinfo("Importação", "Planilha importada com Sucesso !!")

def mensagem_processamento():
    messagebox.showinfo("Processado !", "Dados processados com sucesso !!")


root = Single_window()

title_app = tk.Frame(root)
space_0 = ttk.Label(title_app,text="Banco de Impostos - ICMS ST",width=27, padding=80, font= "Arial 24 bold")
space_0.grid(row=0,column=0)
title_app.pack(side = 'top')



button_f1 = tk.Frame(root)
button_4 = Button(button_f1, text = 'Importar Saidas',bg="#d99591",width=15, pady=70,padx=50, border=2,font='arial 12')
button_4.grid(row=1, column=0, columnspan=1)
button_4["command"] = lambda:[projeto_saldo_st.importa_saidas(), mensagem_importa()]

space_1 = ttk.Panedwindow(button_f1,width=100)
space_1.grid(row=1, column=1, columnspan=1)

button_5 = Button(button_f1, text = 'Importar Entradas',bg="#84ab84",width=15, pady=70,padx=50,font='arial 12')
button_5.grid(row=1, column=2)
button_5["command"] = lambda:[projeto_saldo_st.importa_entradas(),mensagem_importa()]
button_f1.pack(side = 'top')



button_f2 = tk.Frame(root)
button_7 = Button(button_f2, text = 'Exportar Planilhas Templates',bg="#95cfe8",width=15, pady=70,padx=50,font='arial 12')
button_7.grid(row=3, column=0)
button_7["command"] = lambda:[projeto_saldo_st.planilha_modelo_template_saidas(),projeto_saldo_st.planilha_modelo_template_entradas(),mensagem_exporta()]

space_2 = ttk.Panedwindow(button_f2,width=100,height=30)
space_2.grid(row=2, column=1, columnspan=1)

button_8 = Button(button_f2, text = 'Exportar Saldo Atualizado',bg="#95cfe8",width=15, pady=70,padx=50,font='arial 12')
button_8.grid(row=3, column=2)
button_8['command'] = lambda:[projeto_saldo_st.exportar_saldo_atual(),mensagem_exporta()]
button_f2.pack(side = 'top')



button_f3 = tk.Frame(root)
button_10 = Button(button_f3, text = 'Importar Ressarcimento TIMP',bg="#ab91bd",width=15, pady=70,padx=50,font='arial 12')
button_10.grid(row=4, column=0)
button_10["command"] = mensagem_exporta

space_3 = ttk.Panedwindow(button_f3,width=100,height=30)
space_3.grid(row=3, column=1, columnspan=1)

button_11 = Button(button_f3, text = 'Importar Ressarcimento',bg="#ab91bd",width=15, pady=70,padx=50,font='arial 12')
button_11.grid(row=4, column=2)
button_f3.pack(side = 'top')


title_app10 = tk.Frame(root)
space_10 = Label(title_app10,text="",width=27, font= "Arial 24 bold")
space_10.grid(row=6,column=0)
title_app10.pack(side = 'bottom')

button_f4 = tk.Frame(root)
button_12 = Button(button_f4, text = 'Processar',bg="#a19f9f",width=15, pady=10,padx=20,font='arial 12 bold')
button_12.grid(row=7, column=0)
button_12['command'] = lambda:[projeto_saldo_st.criar_coluna_tipo_contabilizacao_saidas(), projeto_saldo_st.saldo_atual_provisorio(),
                               projeto_saldo_st.sintetiza_dados(),projeto_saldo_st.saldo_consistido(),mensagem_processamento()]
button_f4.pack(side = 'bottom')



root.mainloop()





