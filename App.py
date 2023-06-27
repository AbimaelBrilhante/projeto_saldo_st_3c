import tkinter as tk
from tkinter import messagebox
import projeto_saldo_st

def abrir_tela_banco_impostos():
    try:
        pagina_inicial.destroy()  # Fechar a janela da página inicial
    except:
        pass

    # Criação da janela do Banco de Impostos
    root = tk.Tk()
    root.title("BANCO DE IMPOSTOS")
    root.configure(bg=cor_primaria)

    menu_principal = tk.Menu(root)
    root.config(menu=menu_principal)

    # Adicionar opções ao menu
    menu_arquivo = tk.Menu(menu_principal, tearoff=False)
    menu_principal.add_cascade(label="Arquivo", menu=menu_arquivo)
    menu_arquivo.add_command(label="Página Inicial", command=abrir_tela_banco_impostos)
    menu_arquivo.add_separator()
    menu_arquivo.add_command(label="Consultar Banco de Dados Individual", command="")
    menu_arquivo.add_separator()
    menu_arquivo.add_command(label="Consultar Banco de Dados Consolidado", command="")
    menu_arquivo.add_separator()
    menu_arquivo.add_command(label="Sair", command=root.quit)

    # Funções dos botões
    def mensagem_exporta():
        messagebox.showinfo("Exportar para Excel", "Planilhas Geradas com Sucesso !")

    def mensagem_exclui():
        messagebox.showinfo("Exclusão", "Dados Excluídos com Sucesso!")

    def mensagem_importa():
        messagebox.showinfo("Importação", "Planilha importada com Sucesso !!")

    def mensagem_processamento():
        messagebox.showinfo("Processado !", "Dados processados com sucesso !!")

    # Cabeçalho
    label_titulo = tk.Label(root, text="BANCO DE IMPOSTOS", font=("Arial", 20), bg=cor_hover, fg=cor_texto_2, pady=10, padx=300)
    label_titulo.pack(pady=0)

    # Criação dos botões com estilo personalizado
    estilo_botao = {"font": ("Arial", 12), "width": 30}

    def on_enter(event):
        event.widget.config(bg=cor_hover)

    def on_leave(event):
        event.widget.config(bg=cor_secundaria)

    btn_importar_entradas = tk.Button(root, text="Importar Entradas", command=lambda:[projeto_saldo_st.importa_entradas(), mensagem_importa()], bg=cor_secundaria, fg=cor_texto, **estilo_botao)
    btn_importar_entradas.bind("<Enter>", on_enter)
    btn_importar_entradas.bind("<Leave>", on_leave)
    btn_importar_entradas.pack(pady=10)

    btn_importar_saidas = tk.Button(root, text="Importar Saídas", command=lambda:[projeto_saldo_st.importa_saidas(),mensagem_importa()], bg=cor_secundaria, fg=cor_texto, **estilo_botao)
    btn_importar_saidas.bind("<Enter>", on_enter)
    btn_importar_saidas.bind("<Leave>", on_leave)
    btn_importar_saidas.pack(pady=10)

    btn_excluir_dados_entradas = tk.Button(root, text="Excluir Dados de Entradas", command=lambda:[projeto_saldo_st.exclui_dados_entradas(),mensagem_exclui()], bg=cor_secundaria, fg=cor_texto, **estilo_botao)
    btn_excluir_dados_entradas.bind("<Enter>", on_enter)
    btn_excluir_dados_entradas.bind("<Leave>", on_leave)
    btn_excluir_dados_entradas.pack(pady=10)

    btn_exportar_template_entradas = tk.Button(root, text="Exportar Planilha Template Entradas", command=lambda:[projeto_saldo_st.planilha_modelo_template_entradas(),mensagem_exporta()], bg=cor_secundaria, fg=cor_texto, **estilo_botao)
    btn_exportar_template_entradas.bind("<Enter>", on_enter)
    btn_exportar_template_entradas.bind("<Leave>", on_leave)
    btn_exportar_template_entradas.pack(pady=10)

    btn_exportar_template_saidas = tk.Button(root, text="Exportar Planilha Template Saídas", command=lambda:[projeto_saldo_st.planilha_modelo_template_saidas(),mensagem_exporta()], bg=cor_secundaria, fg=cor_texto, **estilo_botao)
    btn_exportar_template_saidas.bind("<Enter>", on_enter)
    btn_exportar_template_saidas.bind("<Leave>", on_leave)
    btn_exportar_template_saidas.pack(pady=10)

    btn_consistir_saldo = tk.Button(root, text="Consistir Saldo", command=lambda:[projeto_saldo_st.criar_coluna_tipo_contabilizacao_saidas(), projeto_saldo_st.saldo_atual_provisorio(),
                               projeto_saldo_st.sintetiza_dados()
                                ,projeto_saldo_st.saldo_consistido(),mensagem_processamento()], bg=cor_secundaria, fg=cor_texto, **estilo_botao)
    btn_consistir_saldo.bind("<Enter>", on_enter)
    btn_consistir_saldo.bind("<Leave>", on_leave)
    btn_consistir_saldo.pack(pady=10)

    # Rodapé
    label_rodape = tk.Label(root, text="Solucões Fiscais 3C", font=("Arial", 9), bg=cor_hover, fg=cor_texto_2, pady=10)
    label_rodape.pack(side="bottom", fill="x")

# Definindo uma paleta de cores
cor_primaria = "#222222"  # Cinza escuro
cor_secundaria = "#FFFFFF"  # Branco
cor_texto = "#000000"  # Preto
cor_texto_2 = "#383837"
cor_rodape = "#CCCCCC"  # Cinza claro
cor_hover = "#E0E0E0"  # Cinza claro (intensidade alterada)

# Criação da janela da página inicial
pagina_inicial = tk.Tk()
pagina_inicial.title("Banco de Impostos")
pagina_inicial.configure(bg=cor_primaria)

# Cabeçalho
label_titulo = tk.Label(pagina_inicial, text="BANCO DE IMPOSTOS", font=("Arial", 20), bg=cor_primaria, fg=cor_hover, pady=10)
label_titulo.pack(pady=20)

# Conteúdo da página inicial
label_informacoes = tk.Label(pagina_inicial, text="Bem-vindo ao Banco de Impostos!\n \nEste é um software para geração, gerenciamento e cálculo de impostos e saldos.\n "
                                                  "\nEle foi desenvolvido pela equipe de Soluções Fiscais, Inovação e Tecnologia da Controladoria do Grupo 3 Corações", font=("Arial", 12), bg=cor_primaria, fg=cor_hover)
label_informacoes.pack(pady=100, padx=50)

# Rodapé
label_rodape = tk.Label(pagina_inicial, text="Solucões Fiscais 3C", font=("Arial", 10), bg=cor_primaria,fg=cor_texto_2)
label_rodape.pack(side="bottom", fill="x")

# Agendar a abertura da tela do Banco de Impostos após 3 segundos
pagina_inicial.after(3000, abrir_tela_banco_impostos)

# Execução da interface gráfica
pagina_inicial.mainloop()
