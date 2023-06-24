import tkinter as tk
import requests
from tkinter import messagebox
from tkinter import ttk
import pandas as pd

# SCRIPT
def consultar_cnpj(cnpj):
    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"
    response = requests.get(url)
    data = response.json()
    print(response.text)
    return data

def verificar_dispensa_licenciamento():
    cnpj = cnpj_entry.get()
    data = consultar_cnpj(cnpj)

    if 'message' in data:
        messagebox.showerror("Erro", "CNPJ inválido")
        return

    cnaes = data['atividade_principal'] + data['atividades_secundarias']

    cnaes_encontrados = []
    cnpjs_licenciados = []

    with open("CNAES.txt", "r", encoding="utf-8") as file:
        cnaes_desejados = set(line.strip().split(";")[0] for line in file)

    dispensa_licenciamento = True

    for cnae_data in cnaes:
        cnae = cnae_data['code']
        nome = cnae_data['text']

        cnaes_encontrados.append(f"{cnae} - {nome}")
        
    for cnae_data in cnaes:
        if cnae_data['code'] not in cnaes_desejados:
            dispensa_licenciamento = False
            break

    if dispensa_licenciamento:
        messagebox.showinfo("Resultado", "A empresa pode ser dispensada de licenciamento")
        messagebox.showinfo("CNAEs Encontrados", "\n".join(cnaes_encontrados))

        cnpjs_licenciados.append(cnpj)

        adicionar_cnpj_planilha(cnpj, data['nome'])

    else:
        messagebox.showinfo("Resultado", "A empresa não pode ser dispensada de licenciamento")

        adicionar_cnpj_planilha1(cnpj, data['nome'])

# EXCEL
def adicionar_cnpj_planilha(cnpj, nome_empresa):
    try:
        df = pd.read_excel("cnpj_licenciados.xlsx")

        df_novos_cnpjs = pd.DataFrame({"CNPJ": [cnpj], "Nome": [nome_empresa]})
        df = pd.concat([df, df_novos_cnpjs])

        df.to_excel("cnpj_licenciados.xlsx", index=False)

        messagebox.showinfo("Planilha Atualizada", "O CNPJ foi adicionado à planilha com sucesso!")

    except FileNotFoundError:
        messagebox.showerror("Erro", "A planilha 'cnpj_licenciados.xlsx' não foi encontrada.")

def adicionar_cnpj_planilha1(cnpj, nome_empresa):
    try:
        df = pd.read_excel("cnpj_nao_licenciados.xlsx")

        df_novos_cnpjs = pd.DataFrame({"CNPJ": [cnpj], "Nome": [nome_empresa]})
        df = pd.concat([df, df_novos_cnpjs])

        df.to_excel("cnpj_nao_licenciados.xlsx", index=False)

        messagebox.showinfo("Planilha Atualizada", "O CNPJ foi adicionado à planilha com sucesso!")

    except FileNotFoundError:
        messagebox.showerror("Erro", "A planilha 'cnpj_nao_licenciados.xlsx' não foi encontrada.")

# CSS
window = tk.Tk()
window.title("Verificação de Dispensa de Licenciamento")
window.geometry("500x300")
window.configure(bg="#f2f2f2")

logo_image = tk.PhotoImage(file="color.png")

logo_label = tk.Label(window, image=logo_image)
logo_label.pack()

cnpj_label = tk.Label(window, text="Digite o CNPJ:")
cnpj_label.pack(pady=5)
cnpj_entry = ttk.Entry(window, style="Custom.TEntry")
cnpj_entry.pack(pady=5)

style = ttk.Style()
style.configure("Custom.TButton",
                background="#FF8C00",
                foreground="black",
                font=("Arial", 12),
                padding=5)

verificar_button = ttk.Button(window, text="Verificar", style="Custom.TButton", command=verificar_dispensa_licenciamento)
verificar_button.pack()

window.mainloop()