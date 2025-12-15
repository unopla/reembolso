import json
import os
from datetime import date, timedelta
from tkinter import *
from tkinter import messagebox, simpledialog
from docx import Document
import sys
import os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


CONFIG_FILE = "config.json"
HORARIO_IDA = "07:40"
HORARIO_VOLTA = "17:15"

def carregar_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return None

def salvar_config(dados):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)

def pedir_dados():
    nome = simpledialog.askstring("Dados", "Nome completo:")
    cpf = simpledialog.askstring("Dados", "CPF:")
    instituicao = simpledialog.askstring("Dados", "Instituição de ensino:")

    if not nome or not cpf or not instituicao:
        messagebox.showerror("Erro", "Todos os campos são obrigatórios.")
        return

    salvar_config({
        "nome": nome,
        "cpf": cpf,
        "instituicao": instituicao
    })

    messagebox.showinfo("OK", "Dados salvos com sucesso!")

def dias_uteis(mes, ano):
    d = date(ano, mes, 1)
    dias = []
    while d.month == mes:
        if d.weekday() < 5:
            dias.append(d)
        d += timedelta(days=1)
    return dias

def gerar_documento():
    config = carregar_config()
    if not config:
        messagebox.showerror("Erro", "Cadastre os dados pessoais primeiro.")
        return

    mes = simpledialog.askinteger("Mês", "Digite o mês (1-12):")
    ano = simpledialog.askinteger("Ano", "Digite o ano:")

    if not mes or not ano:
        return

    doc = Document(resource_path("modelo.docx"))


    texto = doc.paragraphs[0].text
    texto = texto.replace("__________________", config["nome"])
    texto = texto.replace("CPF________________________", f"CPF {config['cpf']}")
    texto = texto.replace(
        "_____________________________________",
        config["instituicao"]
    )
    doc.paragraphs[0].text = texto

    doc.paragraphs[1].text = f"Reembolso referente ao mês: {mes}/{ano}"

    tabela = doc.add_table(rows=1, cols=4)
    hdr = tabela.rows[0].cells
    hdr[0].text = "DATA"
    hdr[1].text = "HORÁRIO IDA"
    hdr[2].text = "HORÁRIO VOLTA"
    hdr[3].text = "ASSINATURA"

    for dia in dias_uteis(mes, ano):
        resp = messagebox.askyesno(
            "Presença",
            f"Você foi no dia {dia.strftime('%d/%m/%Y')}?"
        )
        if resp:
            row = tabela.add_row().cells
            row[0].text = dia.strftime("%d/%m/%Y")
            row[1].text = HORARIO_IDA
            row[2].text = HORARIO_VOLTA
            row[3].text = ""

    nome_arquivo = f"Reembolso_{mes}_{ano}.docx"
    doc.save(nome_arquivo)

    messagebox.showinfo("Sucesso", f"Arquivo criado:\n{nome_arquivo}")

# INTERFACE
root = Tk()
root.title("Gerador de Reembolso")
root.geometry("300x200")

Button(root, text="Gerar Reembolso", width=25, command=gerar_documento).pack(pady=20)
Button(root, text="Alterar dados pessoais", width=25, command=pedir_dados).pack()

if not carregar_config():
    pedir_dados()

root.mainloop()
