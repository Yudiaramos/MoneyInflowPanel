import tkinter as tk
from tkinter import ttk
import subprocess

def run_planilha_projetos():
    subprocess.Popen(["python", "planilha_projetos.py"])

def run_planilha_funcionarios():
    subprocess.Popen(["python", "planilha_funcionarios.py"])

# Interface de seleção inicial
root = tk.Tk()
root.title("Seleção de Planilha")

frame = ttk.Frame(root, padding="10")
frame.pack(fill='both', expand=True)

lbl = ttk.Label(frame, text="Selecione a planilha:")
lbl.pack(pady=10)

btn_projetos = ttk.Button(frame, text="Planilha de Projetos", command=run_planilha_projetos)
btn_projetos.pack(fill='x', pady=5)

btn_funcionarios = ttk.Button(frame, text="Planilha de Funcionários", command=run_planilha_funcionarios)
btn_funcionarios.pack(fill='x', pady=5)

root.mainloop()
