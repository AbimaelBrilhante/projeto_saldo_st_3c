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
space_0 = ttk.Label(title_app,text="\n                                                                                  "
                                   "                                                                                    "
                                   "                                                          Menu Principal\n"
                                   " _________________________________________________________________________"
                                   "______________________________________________________________________________________________"
                                   "____________________________________________",width=6007, padding=5, font= "Arial 12", foreground="#a19f9f")
space_0.grid(row=0,column=0)
title_app.pack(side = 'top')


title_app = tk.Frame(root)
space_0 = ttk.Label(title_app,text="",width=6007, padding=15, font= "Arial 12")
space_0.grid(row=0,column=0)
title_app.pack(side = 'top')


msg1 = "Importar Relatorio Saidas".format("bold")
msg2 = "Importar Relatorio Entradas"
msg3 = "Exportar Template Entradas"
msg30 = "Exportar Template Saidas"
msg4 = "Exportar Saldo Atualizado"
msg5 = f"Importar Relatorio \n Devoluções"

button_f1 = tk.Frame(root)
button_4 = Button(button_f1, text = f"IRE \n \n {msg2}", bg="#95c45c",width=21, pady=60,padx=10, border=2,font='arial 16')
button_4.grid(row=1, column=0, columnspan=1)
button_4["command"] = lambda:[projeto_saldo_st.importa_entradas(), mensagem_importa()]

space_1 = PanedWindow(button_f1,width=100,background="#cacbd2")
space_1.grid(row=1, column=1, columnspan=1)

button_5 = Button(button_f1, text = f"IRS \n \n {msg1}", bg="#95c45c",width=21, pady=60,padx=10, border=2,font='arial 16')
button_5.grid(row=1, column=2)
button_5["command"] = lambda:[projeto_saldo_st.importa_saidas(),mensagem_importa()]
button_f1.pack(side = 'top')

space_6 = PanedWindow(button_f1,width=100,background="#cacbd2")
space_6.grid(row=1, column=3, columnspan=1)

button_20 = Button(button_f1, text = f"IRD \n \n {msg5}", bg="#95c45c",width=21, pady=48,padx=10, border=2,font='arial 16')
button_20.grid(row=1, column=4)
button_20["command"] = lambda:[projeto_saldo_st.importa_devolucoes(),mensagem_importa()]
button_f1.pack(side = 'top')



button_f2 = tk.Frame(root)
button_7 = Button(button_f2, text = f"ETE \n \n {msg3}", bg="#d99591",width=21, pady=60,padx=10, border=2,font='arial 16')
button_7.grid(row=3, column=0)
button_7["command"] = lambda:[projeto_saldo_st.planilha_modelo_template_entradas(),mensagem_exporta()]

space_2 = PanedWindow(button_f2,width=100,height=30,background="#cacbd2")
space_2.grid(row=2, column=1)

#button_f22 = tk.Frame(root)
button_7 = Button(button_f2, text = f"ETS \n \n {msg30}", bg="#d99591",width=21, pady=60,padx=10, border=2,font='arial 16')
button_7.grid(row=3, column=2)
button_7["command"] = lambda:[projeto_saldo_st.planilha_modelo_template_saidas(),mensagem_exporta()]

space_2 = PanedWindow(button_f2,width=100,height=30,background="#cacbd2")
space_2.grid(row=2, column=3)

button_8 = Button(button_f2, text = f"ESA \n \n {msg4}", bg="#d99591",width=21, pady=60,padx=10, border=2,font='arial 16')
button_8.grid(row=3, column=4)
button_8['command'] = lambda:[projeto_saldo_st.exportar_saldo_atual(),mensagem_exporta()]
button_f2.pack(side = 'top')



button_f3 = tk.Frame(root)
button_10 = Button(button_f3, text = '\n IRT \n',bg="#ab91bd",width=21, pady=60,padx=10, border=2,font='arial 16')
button_10.grid(row=4, column=0)
button_10["command"] = mensagem_exporta

space_3 = PanedWindow(button_f3,width=100,height=30,background="#cacbd2")
space_3.grid(row=3, column=1, columnspan=1)

button_11 = Button(button_f3, text = '\n IRC \n ',bg="#ab91bd",width=21, pady=60,padx=10, border=2,font='arial 16')
button_11.grid(row=4, column=2)
button_f3.pack(side = 'top')

space_3 = PanedWindow(button_f3,width=100,height=30,background="#cacbd2")
space_3.grid(row=3, column=3,)

button_11 = Button(button_f3, text = '\n IRC \n ',bg="#ab91bd",width=21, pady=60,padx=10, border=2,font='arial 16')
button_11.grid(row=4, column=4)
button_f3.pack(side = 'top')


title_app10 = tk.Frame(root)
space_10 = Label(title_app10,text="",width=2700, font= "Arial 24 bold",border=1,foreground="#cacbd2")
space_10.grid(row=6,column=0)
title_app10.pack(side = 'bottom')




button_f41 = tk.Frame(root)
button_121 = Button(button_f41, text = 'Processar',width=1500, pady=7,padx=20,font='arial 12 bold', border=10, borderwidth=0)
button_121.grid(row=8, column=0)
button_f41.pack(side = 'bottom')

button_f40 = tk.Frame(root)
button_120 = Button(button_f40, text = '',width=1500, pady=7,padx=2000,font='arial 12 bold', border=0,background="#cacbd2")
button_120.grid(row=8, column=0)
button_f40.pack(side = 'bottom')



button_f4 = tk.Frame(root)
button_12 = Button(button_f4, text = 'Processar',bg="#a19f9f",width=15, pady=10,padx=20,font='arial 12 bold')
button_12.grid(row=7, column=0)
button_12['command'] = lambda:[projeto_saldo_st.criar_coluna_tipo_contabilizacao_saidas(), projeto_saldo_st.saldo_atual_provisorio(),
                               projeto_saldo_st.sintetiza_dados(),projeto_saldo_st.sintetiza_dados_devolucoes()
                                ,projeto_saldo_st.saldo_consistido(),mensagem_processamento()]
button_f4.pack(side = 'bottom')




button_f1.configure(background="#cacbd2")
button_f2.configure(background="#cacbd2")
button_f3.configure(background="#cacbd2")



root.configure(background="#cacbd2")

root.mainloop()





