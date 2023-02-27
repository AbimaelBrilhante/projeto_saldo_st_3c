import tkinter as tk
from tkinter import ttk
import random
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


root = Single_window()

title_app = tk.Frame(root)
space_0 = ttk.Label(title_app,text="Banco de Impostos - ICMS ST",width=27, padding=60, font= "Arial 20 bold")
space_0.grid(row=0,column=0)
title_app.pack(side = 'top')



button_f1 = tk.Frame(root)
button_4 = ttk.Button(button_f1, text = 'Importar Saidas',width=15, padding=50)
button_4.grid(row=1, column=0, columnspan=1)
button_4["command"] = lambda:[projeto_saldo_st.importa_saidas()]

space_1 = ttk.Panedwindow(button_f1,width=100)
space_1.grid(row=1, column=1, columnspan=1)

button_5 = ttk.Button(button_f1, text = 'Importar Entradas',width=15, padding=50,)
button_5.grid(row=1, column=2)
button_5["command"] = lambda:[projeto_saldo_st.importa_entradas()]
button_f1.pack(side = 'top')



button_f2 = tk.Frame(root)
button_7 = ttk.Button(button_f2, text = 'Exportar Planilhas Templates',width=15, padding=50)
button_7.grid(row=3, column=0)
button_7["command"] = lambda:[projeto_saldo_st.planilha_modelo_template_saidas(),projeto_saldo_st.planilha_modelo_template_entradas(),mensagem_exporta()]

space_2 = ttk.Panedwindow(button_f2,width=100,height=30)
space_2.grid(row=2, column=1, columnspan=1)

button_8 = ttk.Button(button_f2, text = 'Exportar Saldo Atualizado',width=15, padding=50)
button_8.grid(row=3, column=2)
button_f2.pack(side = 'top')



button_f3 = tk.Frame(root)
button_10 = ttk.Button(button_f3, text = 'Importar Ressarcimento TIMP',width=15, padding=50)
button_10.grid(row=4, column=0)
button_10["command"] = mensagem_exporta

space_3 = ttk.Panedwindow(button_f3,width=100,height=30)
space_3.grid(row=3, column=1, columnspan=1)

button_11 = ttk.Button(button_f3, text = 'Importar Ressarcimento',width=15, padding=50)
button_11.grid(row=4, column=2)
button_f3.pack(side = 'top')

root.mainloop()





