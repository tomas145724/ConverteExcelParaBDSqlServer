import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox, scrolledtext
from sqlalchemy import create_engine
import os





# Configuração do tema da interface
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("dark-blue")

# Função para selecionar arquivos Excel
def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos Excel",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if arquivos:
        entry_arquivos.delete(0, ctk.END)
        entry_arquivos.insert(0, "; ".join(arquivos))

# Função para adicionar mensagens ao log
def log(mensagem):
    caixa_log.insert(ctk.END, mensagem + "\n")
    caixa_log.see(ctk.END)
    janela.update_idletasks()

# Função para processar os arquivos e importar para o banco de dados
def importar_arquivos():
    server = entry_servidor.get()
    database = entry_banco.get()
    username = entry_usuario.get()
    password = entry_senha.get()
    arquivos = entry_arquivos.get().split("; ")
    header_row = int(header_choice.get()) - 1  # Converta a escolha para índice (0 ou 1)

    if not all([server, database, username, password, arquivos]):
        messagebox.showerror("Erro", "Preencha todos os campos!")
        return

    try:
        engine = create_engine(
            f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server'
        )
        log("Conexão com o banco de dados estabelecida com sucesso.")

        total_arquivos = len(arquivos)
        progresso.set(0)
        log(f"Iniciando importação de {total_arquivos} arquivo(s)...")

        for i, arquivo in enumerate(arquivos):
            log(f"Processando arquivo: {arquivo}")
            df = pd.read_excel(arquivo, header=header_row)  # Use a linha escolhida como cabeçalho

            table_name = os.path.splitext(os.path.basename(arquivo))[0]
            log(f"Importando dados para a tabela: {table_name}")

            df.to_sql(table_name, con=engine, if_exists='replace', index=False)
            #log(f"Arquivo {arquivo} está sendo importado!")

            total_linhas = len(df)
            for linha_idx in range(total_linhas):
                progresso.set((linha_idx + 1) / total_linhas)
                janela.update_idletasks()

        log("Importação concluída com sucesso!")
        messagebox.showinfo("Sucesso", "Arquivos importados com sucesso!")
    except Exception as e:
        log(f"Erro: {str(e)}")
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
    finally:
        progresso.set(0)

# Criar a janela principal
janela = ctk.CTk()
janela.title("Excel para Banco de Dados")
janela.geometry("600x500")
janela.maxsize(width=900, height=500)
janela.resizable(width=True, height=False)

# Adicione uma variável para armazenar a escolha do cabeçalho
header_choice = ctk.StringVar(value="1")  # Valor padrão: primeira linha como cabeçalho

# Campos de entrada para as informações de conexão
ctk.CTkLabel(janela, text="Servidor:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
entry_servidor = ctk.CTkEntry(janela, width=300)
entry_servidor.insert(0, "localhost")
entry_servidor.grid(row=0, column=1, padx=10, pady=5)

ctk.CTkLabel(janela, text="Banco de Dados:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
entry_banco = ctk.CTkEntry(janela, width=300)
entry_banco.grid(row=1, column=1, padx=10, pady=5)

ctk.CTkLabel(janela, text="Usuário:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
entry_usuario = ctk.CTkEntry(janela, width=300)
entry_usuario.insert(0, "sa")
entry_usuario.grid(row=2, column=1, padx=10, pady=5)

ctk.CTkLabel(janela, text="Senha:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
entry_senha = ctk.CTkEntry(janela, width=300, show="*")
entry_senha.insert(0, "145724")
entry_senha.grid(row=3, column=1, padx=10, pady=5)

# Adicione opções para selecionar o cabeçalho
ctk.CTkLabel(janela, text="Linha do Cabeçalho:").grid(row=4, column=0, padx=10, pady=5, sticky="e")
ctk.CTkRadioButton(janela, text="Primeira Linha", variable=header_choice, value="1").grid(row=4, column=1, sticky="w")
ctk.CTkRadioButton(janela, text="Segunda Linha", variable=header_choice, value="2").grid(row=4, column=1, sticky="e")


# Frame para agrupar o campo de entrada e o botão "Selecionar"
frame_arquivos = ctk.CTkFrame(janela)
frame_arquivos.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

entry_arquivos = ctk.CTkEntry(frame_arquivos, width=200, height=50)
entry_arquivos.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

btn_selecionar = ctk.CTkButton(frame_arquivos, text="Selecionar Arquivos Excel", command=selecionar_arquivos)
btn_selecionar.grid(row=0, column=2, pady=10, sticky="nsew")

frame_arquivos.columnconfigure(1, weight=1)

# Botão para iniciar a importação
btn_importar = ctk.CTkButton(janela, text="Iniciar", fg_color="#239B56", command=importar_arquivos)
btn_importar.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Caixa de log
caixa_log = scrolledtext.ScrolledText(janela, width=60, height=10, wrap=ctk.WORD)
caixa_log.grid(row=7, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Barra de progresso
progresso = ctk.CTkProgressBar(janela, orientation="horizontal", width=400)
progresso.grid(row=8, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
progresso.set(0)

 

# Ajustar configuração de layout para a última linha
janela.rowconfigure(10, weight=1)
janela.columnconfigure(0, weight=1)
janela.columnconfigure(1, weight=1)
janela.columnconfigure(2, weight=1)
 
# Iniciar a interface gráfica
janela.mainloop()
