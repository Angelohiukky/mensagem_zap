import time
from pathlib import Path
import pandas as pd
import pywhatkit as kit
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk

# Função que envia as mensagens
def enviar_mensagens():
    caminho_excel = Path(entry_planilha.get())
    
    try:
        df = pd.read_excel(caminho_excel, dtype=str)

        for indice, linha in df.iterrows():
            nome = linha["nome"]
            numero_telefone = linha["telefone"]
            mensagem = linha["mensagem"]

            mensagem = f"Olá {nome}, está tudo bem? Tenho uma ótima notícia!\n\n{mensagem}"

            try:
                # Envia a mensagem pelo WhatsApp
                
                kit.sendwhatmsg_instantly(
                    numero_telefone,
                    mensagem,
                    time.localtime().tm_hour,
                    time.localtime().tm_min + 1,
                )
                print(f"Mensagem enviada para o {nome}")
                label_status.config(text=f"Mensagem enviada para {nome}", fg="green")

            except Exception as e:
                print(f"Falha ao enviar a mensagem para {nome}: {str(e)}")
                label_status.config(text=f"Falha ao enviar mensagem para {nome}: Não enviado", fg="red")
                continue  # Se falhar, ele continua para o próximo contato

            # Aguardar 2 segundos antes de enviar a próxima mensagem
            time.sleep(2)

    except Exception as e:
        label_status.config(text=f"Erro ao carregar planilha: {str(e)}", fg="red")

# Função para selecionar o arquivo Excel
def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Arquivos Excel", ".xlsx"), ("Todos os arquivos", ".*")]
    )
    entry_planilha.delete(0, tk.END)
    entry_planilha.insert(0, arquivo)

janela = ctk.CTk()
janela.title("Envio de Mensagens pelo WhatsApp")
janela.geometry("380x300")

fonte_bold = ctk.CTkFont(family="Helvetica", size=16, weight="bold")
label_planilha = ctk.CTkLabel(janela, text="Convocação via Whatsapp", font=fonte_bold, compound="top")
label_planilha.pack(pady=20)


entry_planilha = ctk.CTkEntry(janela, width=300, placeholder_text="Selecione a Planilha")
entry_planilha.pack(pady=10)

btn_selecionar = ctk.CTkButton(janela, text="Selecionar", command=selecionar_arquivo)
btn_selecionar.pack(pady=10)

btn_enviar = ctk.CTkButton(janela, text="Enviar Mensagens", command=enviar_mensagens)
btn_enviar.pack(pady=20)

label_status = ctk.CTkLabel(janela, text="")
label_status.pack(pady=10)

janela.mainloop()