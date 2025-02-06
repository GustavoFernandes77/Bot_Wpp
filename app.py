import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
import datetime
import customtkinter as ctk
from tkinter import messagebox

'''Bibliotecas necessárias:
pip install pyautogui
pip install openpyxl
pip install customtkinter
'''

# Mensagem base para boleto não vencido
MENSAGEM_BASE = "Olá {nome}, seu boleto vence no dia {vencimento}. Favor entrar em contato conosco para atendimento!"

# Mensagem base para boleto vencido
MENSAGEM_VENCIDO = "Olá {nome}, seu boleto venceu no dia {vencimento}. Favor entrar em contato conosco para atendimento!"

def carregar_clientes():
    """Carrega os clientes da planilha CLIENTES.xlsx"""
    arquivo = os.path.join(os.path.dirname(__file__), 'CLIENTES.xlsx')

    if not os.path.exists(arquivo):
        messagebox.showerror("Erro", "Planilha CLIENTES.xlsx não encontrada!")
        return []

    workbook = openpyxl.load_workbook(arquivo)
    pagina_clientes = workbook.active

    clientes = []
    for linha in pagina_clientes.iter_rows(min_row=2):
        nome = linha[0].value
        telefone = linha[1].value
        pagamento = linha[6].value
        vencimento = linha[5].value

        if pagamento == "Pagou" or not vencimento:
            continue

        # Converte a data para o formato correto
        if isinstance(vencimento, str):
            vencimento = datetime.datetime.strptime(vencimento, '%d/%m/%Y')
        elif isinstance(vencimento, datetime.datetime):
            pass  # Já está no formato correto
        else:
            continue

        # Extrai o primeiro nome
        primeiro_nome = nome.split()[0] if nome else "Cliente"

        # Verifica se a data de vencimento já passou
        hoje = datetime.datetime.now()
        mensagem = MENSAGEM_VENCIDO if vencimento < hoje else MENSAGEM_BASE

        clientes.append({
            "nome": primeiro_nome,  # Usa apenas o primeiro nome
            "telefone": telefone,
            "vencimento": vencimento.strftime('%d/%m/%Y'),
            "mensagem": mensagem.format(nome=primeiro_nome, vencimento=vencimento.strftime('%d/%m/%Y'))
        })

    return clientes

def enviar_mensagens(clientes):
    """Envia as mensagens para os clientes"""
    for cliente in clientes:
        nome = cliente["nome"]
        telefone = cliente["telefone"]
        mensagem = cliente["mensagem"]

        print(f"Enviando para: {nome} ({telefone}) - Mensagem: {mensagem}")

        # Criar link do WhatsApp
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(15)  # Tempo para carregar a conversa

        # Tenta enviar a mensagem pressionando ENTER
        try:
            sleep(5)  # Tempo para garantir que a conversa foi carregada
            pyautogui.hotkey('enter')

            # Fecha a aba da conversa
            sleep(2)
            pyautogui.hotkey('ctrl', 'w')
            sleep(2)

        except Exception as e:
            print(f'Erro ao enviar mensagem para {nome}: {e}')
            pyautogui.hotkey('ctrl', 'w')  # Fecha a aba mesmo em caso de erro
            try:
                with open('erro.csv', 'a', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f'{nome},{telefone}{os.linesep}')
            except Exception as e:
                print(f'Erro ao escrever no arquivo: {e}')

# Interface gráfica
ctk.set_appearance_mode("dark")  # Tema escuro
ctk.set_default_color_theme("blue")  # Tema azul

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Enviar Mensagens")
        self.geometry("800x600")

        # Carrega os clientes
        self.clientes = carregar_clientes()

        if not self.clientes:
            messagebox.showinfo("Info", "Nenhum cliente encontrado para enviar mensagens.")
            self.destroy()
            return

        # Frame principal
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Lista de clientes
        self.lista_clientes = ctk.CTkScrollableFrame(self.frame)
        self.lista_clientes.pack(fill="both", expand=True, padx=10, pady=10)

        # Dicionário para armazenar as caixas de texto de mensagens
        self.mensagens_entry = {}

        # Adiciona os clientes à lista
        for i, cliente in enumerate(self.clientes):
            frame_cliente = ctk.CTkFrame(self.lista_clientes)
            frame_cliente.pack(fill="x", pady=5)

            # Nome e telefone
            label_cliente = ctk.CTkLabel(frame_cliente, text=f"{cliente['nome']} ({cliente['telefone']})")
            label_cliente.pack(side="left", padx=10)

            # Caixa de texto para personalizar a mensagem
            mensagem_entry = ctk.CTkEntry(frame_cliente, width=400)
            mensagem_entry.insert(0, cliente["mensagem"])
            mensagem_entry.pack(side="left", padx=10, fill="x", expand=True)
            self.mensagens_entry[i] = mensagem_entry

        # Botão de enviar
        botao_enviar = ctk.CTkButton(self.frame, text="Enviar Mensagens", command=self.enviar)
        botao_enviar.pack(pady=10)

    def enviar(self):
        """Atualiza as mensagens dos clientes e inicia o envio"""
        for i, cliente in enumerate(self.clientes):
            cliente["mensagem"] = self.mensagens_entry[i].get()

        enviar_mensagens(self.clientes)
        messagebox.showinfo("Sucesso", "Mensagens enviadas com sucesso!")


if __name__ == "__main__":
    app = App()
    app.mainloop()