from tkinter import *
import tkinter as tk
import webbrowser
import projeto_saldo_st
from tkinter import messagebox


class Application:
    def __init__(self, master=None):

        self.fonte = ("Verdana", "8","bold")
        self.container1 = Frame(master)
        self.container1["pady"] = 100
        self.container1["padx"] = 4000
        self.container1.pack()
        self.container1.configure(bg='#333333')
        self.container8 = Frame(master)
        self.container8["padx"] = 1
        self.container8["pady"] = 1
        self.container8.pack()
        self.container8.configure(bg='#333333')
        self.container9 = Frame(master)
        self.container9["padx"] = 2
        self.container9["pady"] = 2
        self.container9.pack()
        self.container9.configure(bg='#333333')
        self.container10 = Frame(master)
        self.container10["padx"] = 20
        self.container10["pady"] = 5
        self.container10.pack()
        self.container10.configure(bg='#333333')
        self.container11 = Frame(master)
        self.container11["padx"] = 20
        self.container11["pady"] = 5
        self.container11.pack()
        self.container11.configure(bg='#333333')
        self.container12 = Frame(master)
        self.container12["padx"] = 20
        self.container12["pady"] = 5
        self.container12.pack()
        self.container12.configure(bg='#333333')
        self.container13 = Frame(master)
        self.container13["padx"] = 20
        self.container13["pady"] = 5
        self.container13.pack()
        self.container13.configure(bg='#333333')
        self.container14 = Frame(master)
        self.container14["padx"] = 20
        self.container14["pady"] = 10
        self.container14.pack()
        self.container14.configure(bg='#333333')






        self.bntConsultar = Button(self.container12, text="Importar Entradas",
        font=self.fonte, height=10, width=25,bg='#4c4c4c', foreground='white',border=3)
        self.bntConsultar["command"] = projeto_saldo_st.importa
        self.bntConsultar.pack (side=TOP)
        # self.btnConsultar.grid(row=12,column=6)

        self.bntInsert = Button(self.container12, text="Importar Saidas",
        font=self.fonte, height=10,width=25,bg='#4c4c4c', foreground='white',border=3)
        self.bntInsert["command"] = projeto_saldo_st.importa
        self.bntInsert.pack (side=TOP)

        self.bntAlterar = Button(self.container12, text="Exportar Planilhas Templates",
        font=self.fonte, height=10,width=25,bg='#4c4c4c', foreground='white',border=3)
        self.bntAlterar["command"] = lambda:[projeto_saldo_st.planilha_modelo_template_saidas(),projeto_saldo_st.planilha_modelo_template_entradas(),self.mensagem_exporta()]
        self.bntAlterar.pack (side=TOP)

        self.bntLimpar = Button(self.container12, text="Exportar Saldo atual da Conta",
        font=self.fonte, height=10,width=25,bg='#4c4c4c', foreground='white',border=3)
        self.bntLimpar["command"] = projeto_saldo_st.exportar_saldo_atual()
        self.bntLimpar.pack (side=TOP)




    def mensagem_exporta(self):
     from tkinter import messagebox
     messagebox.showinfo("PalmTax", "Planilhas Geradas com Sucesso !")




    def limpar_dados(self):
        try:
            self.label.pack_forget()
            self.container14.pack_forget()
        except:
            self.container14.pack_forget()



    def outro(self):
        try:
            self.limpar_dados()
            self.consultar_dados()
        except:
            self.consultar_dados()

    def chatestudos(self):
        webbrowser.open('https://mail.google.com/chat/u/0/?zx=vm4xeof813n7#chat/space/AAAAwoBbS_k')
        pass

if __name__ == "__main__":

    root = Tk()
    frame = Frame()
    root.title('PalmTax')
    root.configure(bg='#333333')
    frame.pack(expand=False, fill=BOTH)
    root.state('zoomed')
    Application(root)
    root.mainloop()


