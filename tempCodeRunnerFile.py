import time
from datetime import datetime
from pathlib import Path
import pandas as pd
import pywhatkit as kit
import tkinter as tk
from tkinter import filedialog, Text
import customtkinter as ctk
import pyautogui as pag  # Para automação de teclado/mouse

import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import pywhatkit as kit
import tkinter as tk
from tkinter import filedialog, Text
import customtkinter as ctk

def enviar_mensagens():
    """
    Envia mensagens para os contatos listados em um arquivo Excel.
    As mensagens são enviadas pelo WhatsApp com uma pausa de 2 segundos entre cada envio.
    """
    caminho_excel = Path(entry_planilha.get())
    saudacao_formato = entry_formato.get()  # Saudação personalizada com {nome}
    mensagem_base = text_mensagem.get("1.0", tk.END).strip()  # Obter texto da mensagem base

    try:
        # Carregar dados da planilha Excel
        df = pd.read_excel(caminho_excel, dtype=str)
        
        # Obter o tempo atual
        agora = datetime.now()
        hora = agora.hour
        minuto = agora.minute

        # Iterar sobre cada linha do DataFrame
        for _, linha in df.iterrows():
            nome = linha["nome"]
            numero_telefone = linha["telefone"]

            # Substituir o placeholder {nome} pelo nome do contato
            saudacao_formatada = saudacao_formato.replace("{nome}", nome)

            # Combinar a saudacao com a mensagem base
            mensagem_completa = f"{saudacao_formatada}\n\n{mensagem_base}"

            try:
                # Enviar mensagem pelo WhatsApp
                kit.sendwhatmsg_instantly(
                    numero_telefone,
                    mensagem_completa,
                    hora,
                    minuto + 1
                )
                print(f"Mensagem enviada para {nome}")
                label_status.configure(text=f"Mensagem enviada para {nome}", text_color="green")

            except Exception as e:
                print(f"Falha ao enviar a mensagem para {nome}: {str(e)}")
                label_status.configure(text=f"Falha ao enviar mensagem para {nome}: Não enviado", text_color="red")

            # Aguardar 2 segundos antes de enviar a próxima mensagem
            time.sleep(2)

    except Exception as e:
        label_status.configure(text=f"Erro ao carregar planilha: {str(e)}", text_color="red")

def selecionar_arquivo():
    """
    Abre uma caixa de diálogo para selecionar um arquivo Excel e atualiza o campo de entrada com o caminho do arquivo.
    """
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Arquivos Excel", ".xlsx"), ("Todos os arquivos", ".*")]
    )
    
    if caminho_arquivo:
        entry_planilha.delete(0, tk.END)
        entry_planilha.insert(0, caminho_arquivo)
        btn_status.configure(text="✔ Arquivo Anexado", text_color="green")
    else:
        entry_planilha.delete(0, tk.END)
        btn_status.configure(text="✘ Nenhum Arquivo Anexado", text_color="red")

def criar_interface():
    """
    Cria e configura a interface gráfica usando customtkinter.
    """
    janela = ctk.CTk()
    janela.title("Envio de Mensagens pelo WhatsApp")
    janela.geometry("600x600")  # Definir tamanho inicial
    janela.minsize(600, 600)  # Definir tamanho mínimo

    # Configurar grid para centralizar elementos
    janela.grid_columnconfigure(0, weight=1)
    janela.grid_rowconfigure(0, weight=1)
    janela.grid_rowconfigure(1, weight=1)
    janela.grid_rowconfigure(2, weight=3)
    janela.grid_rowconfigure(3, weight=1)
    janela.grid_rowconfigure(4, weight=1)
    janela.grid_rowconfigure(5, weight=1)

    fonte_bold = ctk.CTkFont(family="Helvetica", size=16, weight="bold")

    # Saudação personalizada
    label_formato = ctk.CTkLabel(janela, text="Saudação (use {nome}):", font=fonte_bold)
    label_formato.grid(row=0, column=0, padx=20, pady=10, sticky="n")

    global entry_formato
    entry_formato = ctk.CTkEntry(janela, width=400, placeholder_text="Ex: Olá {nome}, como vai?")
    entry_formato.grid(row=1, column=0, padx=20, pady=5, sticky="n")

    # Campo de mensagem base (editor ajustável)
    label_mensagem = ctk.CTkLabel(janela, text="Escreva sua mensagem base:", font=fonte_bold)
    label_mensagem.grid(row=2, column=0, padx=20, pady=10, sticky="n")

    global text_mensagem
    text_mensagem = Text(janela, wrap=tk.WORD, bg="#2b2b2b", fg="white", height=5, width=50)  # Ajustar tamanho inicial
    text_mensagem.grid(row=3, column=0, padx=20, pady=5, sticky="nsew")

    # Lembrete de formatação de texto
    lembrete_formatacao = ctk.CTkLabel(janela, text="Como usar formatações no WhatsApp:\n- *Negrito*: Use asteriscos (*exemplo*)\n- _Itálico_: Use sublinhado (_exemplo_)\n- ~Tachado~: Use til (~exemplo~)\n- ```Monoespaçado```: Use crases (```exemplo```)", font=ctk.CTkFont(size=12), text_color="gray")
    lembrete_formatacao.grid(row=4, column=0, padx=20, pady=10, sticky="n")

    # Botão de selecionar arquivo Excel
    global entry_planilha
    entry_planilha = ctk.CTkEntry(janela, width=400, placeholder_text="Caminho da Planilha")
    entry_planilha.grid(row=5, column=0, padx=20, pady=5, sticky="n")

    btn_selecionar = ctk.CTkButton(janela, text="Selecionar Planilha", command=selecionar_arquivo)
    btn_selecionar.grid(row=6, column=0, padx=20, pady=10, sticky="n")

    # Status de anexação do arquivo
    global btn_status
    btn_status = ctk.CTkLabel(janela, text="✘ Nenhum Arquivo Anexado", text_color="red")
    btn_status.grid(row=7, column=0, padx=20, pady=10, sticky="n")

    # Botão de enviar mensagem
    btn_enviar = ctk.CTkButton(janela, text="Enviar Mensagens", command=enviar_mensagens)
    btn_enviar.grid(row=8, column=0, padx=20, pady=20, sticky="n")

    # Status de envio
    global label_status
    label_status = ctk.CTkLabel(janela, text="")
    label_status.grid(row=9, column=0, padx=20, pady=10, sticky="n")

    return janela, entry_planilha, entry_formato, text_mensagem, label_status, btn_status

# Criar a interface gráfica
janela, entry_planilha, entry_formato, text_mensagem, label_status, btn_status = criar_interface()
janela.mainloop()
