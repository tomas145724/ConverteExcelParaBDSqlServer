import customtkinter as ctk
from tkinter import filedialog, messagebox
from converters.excel_to_db import ExcelToDBConverter
from utils.db_connection import create_connection

def log(message):
    caixa_log.insert(ctk.END, message + "\n")
    caixa_log.see(ctk.END)
    janela.update_idletasks()

def iniciar_conversao():
    server = entry_servidor.get()
    database = entry_banco.get()
    username = entry_usuario.get()
    password = entry_senha.get()
    arquivos = entry_arquivos.get().split("; ")

    if not all([server, database, username, password, arquivos]):
        messagebox.showerror("Erro", "Preencha todos os campos!")
        return

    try:
        engine = create_connection(server, database, username, password)
        converter = ExcelToDBConverter(engine)

        for arquivo in arquivos:
            log(f"Processando arquivo: {arquivo}")
            converter.convert(arquivo)
            log(f"Arquivo {arquivo} importado com sucesso!")

        log("Importação concluída com sucesso!")
        messagebox.showinfo("Sucesso", "Arquivos importados com sucesso!")
    except Exception as e:
        log(f"Erro: {str(e)}")
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

janela = ctk.CTk()
janela.title("Excel para Banco de Dados")
janela.geometry("600x500")
janela.maxsize(width=900, height=500)
janela.resizable(width=True, height=False)

ctk.CTkLabel(janela, text="Servidor:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
entry_servidor = ctk.CTkEntry(janela, width=300)
entry_servidor.grid(row=0, column=1, padx=10, pady=5)

ctk.CTkLabel(janela, text="Banco de Dados:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
entry_banco = ctk.CTkEntry(janela, width=300)
entry_banco.grid(row=1, column=1, padx=10, pady=5)

ctk.CTkLabel(janela, text="Usuário:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
entry_usuario = ctk.CTkEntry(janela, width=300)
entry_usuario.grid(row=2, column=1, padx=10, pady=5)

ctk.CTkLabel(janela, text="Senha:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
entry_senha = ctk.CTkEntry(janela, width=300, show="*")
entry_senha.grid(row=3, column=1, padx=10, pady=5)

frame_arquivos = ctk.CTkFrame(janela)
frame_arquivos.grid(row=4, column=1, pady=5, sticky="nsew")

entry_arquivos = ctk.CTkEntry(frame_arquivos, width=200)
entry_arquivos.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

btn_selecionar = ctk.CTkButton(frame_arquivos, text="Selecionar Arquivos Excel", command=lambda: entry_arquivos.insert(0, "; ".join(filedialog.askopenfilenames(title="Selecione os arquivos Excel", filetypes=[("Arquivos Excel", "*.xlsx")]))))
btn_selecionar.grid(row=0, column=2, pady=10, sticky="nsew")

btn_importar = ctk.CTkButton(janela, text="Iniciar", fg_color="#239B56", command=iniciar_conversao)
btn_importar.grid(row=5, column=1, pady=30, sticky="nsew")

caixa_log = ctk.scrolledtext.ScrolledText(janela, width=60, height=10, wrap=ctk.WORD)
caixa_log.grid(row=6, column=1, sticky="nsew")

janela.mainloop()