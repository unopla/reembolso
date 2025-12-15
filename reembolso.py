import os
import sys
import json
from datetime import date, timedelta
from tkinter import Tk, Button, messagebox, simpledialog
from docx import Document

# =====================
# CONFIGURAÇÕES
# =====================

HORARIO_IDA = "07:40"
HORARIO_VOLTA = "17:15"

def get_app_dir():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    app_dir = os.path.join(base, "Reembolso")
    os.makedirs(app_dir, exist_ok=True)
    return app_dir

CONFIG_FILE = os.path.join(get_app_dir(), "config.json")

def get_docs_dir():
    pasta = os.path.join(os.path.expanduser("~"), "Documents", "Reembolsos")
    os.makedirs(pasta, exist_ok=True)
    return pasta

def resource_path(relative):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, relative)

# =====================
# DADOS FIXOS
# =====================

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
    inst = simpledialog.askstring("Dados", "Instituição de ensino:")

    if not nome or not cpf or not inst:
        messagebox.showerror("Erro", "Todos os campos são obrigatórios.")
        return

    salvar_config({
        "nome": nome,
        "cpf": cpf,
        "instituicao": inst
    })

# =====================
# DATAS
# =====================

def dias_uteis(mes, ano):
    d = date(ano, mes, 1)
    out = []
    while d.month == mes:
        if d.weekday() < 5:
            out.append(d)
        d += timedelta(days=1)
    return out

# =====================
# DOCUMENTO
# =====================

def gerar_documento():
    cfg = carregar_config()
    if not cfg:
        pedir_dados()
        cfg = carregar_config()

    mes = simpledialog.askinteger("Mês", "Digite o mês (1-12):", minvalue=1, maxvalue=12)
    ano = simpledialog.askinteger("Ano", "Digite o ano:")

    if not mes or not ano:
        return

    doc = Document(resource_path("modelo.docx"))

    # TEXTO INICIAL
    doc.paragraphs[0].text = (
        f"Eu, {cfg['nome']}, CPF {cfg['cpf']}, declaro, para fins de recebimento "
        f"de auxílio, que frequentei presencialmente as aulas na instituição de ensino "
        f"{cfg['instituicao']} e utilizei de transporte, conforme datas e horários abaixo."
    )

    doc.paragraphs[1].text = f"Reembolso referente ao mês: {mes}/{ano}"

    # === TABELA ÚNICA DO DOCUMENTO ===
    if not doc.tables:
        messagebox.showerror("Erro", "Tabela não encontrada no documento.")
        return

    tabela = doc.tables[0]

    linha = 1  # pula cabeçalho

    for dia in dias_uteis(mes, ano):
        if messagebox.askyesno("Presença", f"Você foi no dia {dia.strftime('%d/%m/%Y')}?"):
            if linha >= len(tabela.rows):
                break  # não cria novas linhas, respeita o modelo

            cells = tabela.rows[linha].cells
            cells[0].text = dia.strftime("%d/%m/%Y")
            cells[1].text = HORARIO_IDA
            cells[2].text = HORARIO_VOLTA
            cells[3].text = ""
            linha += 1

    caminho = os.path.join(get_docs_dir(), f"Reembolso_{mes}_{ano}.docx")

    try:
        doc.save(caminho)
    except Exception as e:
        messagebox.showerror("Erro ao salvar", str(e))
        return

    messagebox.showinfo("Sucesso", f"Arquivo gerado em:\n{caminho}")

# =====================
# INTERFACE
# =====================

root = Tk()
root.title("Gerador de Reembolso")
root.geometry("300x220")
root.resizable(False, False)

Button(root, text="Gerar Reembolso", width=30, command=gerar_documento).pack(pady=20)
Button(root, text="Alterar dados pessoais", width=30, command=pedir_dados).pack()

if not carregar_config():
    pedir_dados()

root.mainloop()
