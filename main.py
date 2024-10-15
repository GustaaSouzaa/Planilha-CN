import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os

# Função para carregar o caminho flexível
def get_flexible_path(relative_path):
    base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# Autenticação com Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(get_flexible_path('credentials.json'), scope)
client = gspread.authorize(creds)

spreadsheet = client.open("Planilha Cainiao")
sheet = spreadsheet.sheet1

# Verifica se o cabeçalho está correto
def verificar_cabecalho():
    esperado = ["Data", "Atendente", "AWB", "Nome do Buyer", "Email", "Chamado", "Observações"]
    cabecalho = sheet.row_values(1)
    if cabecalho[:7] != esperado:
        raise Exception("Cabeçalho incorreto! Verifique a planilha.")

# Função para adicionar os dados na planilha
def add_data(atendente, awb, nome_buyer, email, chamado, observacoes):
    verificar_cabecalho()
    
    awb = awb.lower()  # Ignora maiúsculas/minúsculas para verificar duplicatas
    
    # Verifica se o AWB já existe
    existing_awbs = [row[2].lower() for row in sheet.get_all_values()[1:] if len(row) > 2]
    if awb in existing_awbs:
        linha_awb = existing_awbs.index(awb) + 2  # Acha a linha correta
        dados_existentes = sheet.row_values(linha_awb)
        messagebox.showerror("Erro", f"AWB já cadastrado! Dados: {dados_existentes}")
        return False  # Indica que a inserção não foi bem-sucedida

    # Adiciona a data atual
    current_date = datetime.now().strftime("%d/%m/%Y")
    
    # Adiciona os dados à planilha
    data = [current_date, atendente, awb, nome_buyer, email, chamado, observacoes]
    sheet.append_row(data)
    return True  # Indica que a inserção foi bem-sucedida

# Função para limpar os campos após inserção
def clear_fields():
    entry_atendente.delete(0, tk.END)
    entry_awb.delete(0, tk.END)
    entry_nome_buyer.delete(0, tk.END)
    entry_email.delete(0, tk.END)
    combo_chamado.set('')
    entry_observacoes.delete(0, tk.END)

# Função para envio de dados com verificação de autocomplete e duplicata de AWB
def submit_data():
    atendente = entry_atendente.get()
    awb = entry_awb.get()
    nome_buyer = entry_nome_buyer.get()
    email = entry_email.get()
    chamado = combo_chamado.get()
    observacoes = entry_observacoes.get()

    if chamado not in ["Entrega Push", "Acareação", "Alteração de Endereço"]:
        success_message.config(text="Chamado inválido!", fg="red")
        return

    try:
        if add_data(atendente, awb, nome_buyer, email, chamado, observacoes):
            success_message.config(text="Dados inseridos com sucesso!", fg="green")
            clear_fields()
        else:
            success_message.config(text="", fg="red")  # Limpa mensagem de sucesso se falhar
    except Exception as e:
        success_message.config(text=str(e), fg="red")

# Função para filtrar opções do Combobox de Chamado
def filter_chamado(event):
    typed = combo_chamado.get()
    filtered_options = [option for option in chamado_opcoes if option.lower().startswith(typed.lower())]
    combo_chamado['values'] = filtered_options
    if filtered_options:
        combo_chamado.current(0)  # Seleciona o primeiro item da lista filtrada

# Configuração da interface gráfica
root = tk.Tk()
root.title("Planilha CN")
root.geometry("400x400")
root.iconbitmap(get_flexible_path("Logo-Cainiao.ico"))

# Estilo customizado
style = ttk.Style()
style.theme_use("clam")

# Centraliza a janela na tela
root.eval('tk::PlaceWindow . center')

# Labels e campos de entrada
label_atendente = tk.Label(root, text="Atendente:")
label_atendente.pack(pady=5)
entry_atendente = tk.Entry(root)
entry_atendente.pack()

label_awb = tk.Label(root, text="AWB:")
label_awb.pack(pady=5)
entry_awb = tk.Entry(root)
entry_awb.pack()

label_nome_buyer = tk.Label(root, text="Nome do Buyer:")
label_nome_buyer.pack(pady=5)
entry_nome_buyer = tk.Entry(root)
entry_nome_buyer.pack()

label_email = tk.Label(root, text="Email:")
label_email.pack(pady=5)
entry_email = tk.Entry(root)
entry_email.pack()

label_chamado = tk.Label(root, text="Chamado:")
label_chamado.pack(pady=5)

# Combobox para autocomplete de "Chamado"
chamado_opcoes = ["Entrega Push", "Acareação", "Alteração de Endereço"]
combo_chamado = ttk.Combobox(root, values=chamado_opcoes)
combo_chamado.pack()
combo_chamado.bind("<KeyRelease>", filter_chamado)  # Liga a função de filtro à tecla solta

label_observacoes = tk.Label(root, text="Observações:")
label_observacoes.pack(pady=5)
entry_observacoes = tk.Entry(root)
entry_observacoes.pack()

# Botão de envio
submit_button = tk.Button(root, text="Adicionar Dados", command=submit_data)
submit_button.pack(pady=20)

# Mensagem de sucesso ou erro
success_message = tk.Label(root, text="")
success_message.pack()

# Iniciar a interface gráfica
root.mainloop()
