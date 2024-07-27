import tkinter as tk
from tkinter import ttk
import pandas as pd
import os
import tkinter.messagebox as messagebox

funcionarios_file_path = '/mnt/data/PAGAMENTO DE FUNCIONARIOS.xlsx'

# Função para adicionar os dados na planilha de Funcionários
def adicionar_dados_funcionarios():
    # Obter os dados dos campos de entrada
    nome = entry_nome.get()
    cpf = entry_cpf.get()
    cargo = entry_cargo.get()
    data_admissao = entry_data_admissao.get()
    salario = entry_salario.get()
    forma_pagto = combobox_forma_pagto_func.get()

    # Criar um DataFrame com os novos dados
    novos_dados = pd.DataFrame([[nome, cpf, cargo, data_admissao, salario, forma_pagto]],
                               columns=['NOME', 'CPF', 'CARGO', 'DATA ADMISSÃO', 'SALÁRIO', 'FORMA DE PAGAMENTO'])
    
    try:
        if os.path.exists(funcionarios_file_path):
            # Ler a planilha existente
            df = pd.read_excel(funcionarios_file_path)
            # Verificar se há uma linha vazia para adicionar os novos dados
            for i in range(len(df)):
                if pd.isna(df.iloc[i]['NOME']):
                    df.iloc[i] = novos_dados.iloc[0]
                    break
            else:
                # Se não houver linha vazia, adicionar os novos dados ao final
                df = pd.concat([df, novos_dados], ignore_index=True)
        else:
            # Criar um DataFrame vazio com os cabeçalhos apropriados
            df = pd.DataFrame(columns=['NOME', 'CPF', 'CARGO', 'DATA ADMISSÃO', 'SALÁRIO', 'FORMA DE PAGAMENTO'])
            # Adicionar os novos dados
            df = pd.concat([df, novos_dados], ignore_index=True)
        
        # Salvar a planilha
        df.to_excel(funcionarios_file_path, index=False)
        
        # Limpar os campos de entrada
        entry_nome.delete(0, tk.END)
        entry_cpf.delete(0, tk.END)
        entry_cargo.delete(0, tk.END)
        entry_data_admissao.delete(0, tk.END)
        entry_salario.delete(0, tk.END)
        combobox_forma_pagto_func.set('')
    except PermissionError:
        messagebox.showerror("Erro de Permissão", "Não foi possível acessar o arquivo. Verifique se ele está aberto em outro programa e tente novamente.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao tentar salvar os dados: {e}")

# Função para carregar os dados na Treeview para a planilha de Funcionários
def carregar_dados_funcionarios():
    for i in tree.get_children():
        tree.delete(i)
    if os.path.exists(funcionarios_file_path):
        df = pd.read_excel(funcionarios_file_path)
        for index, row in df.iterrows():
            tree.insert("", tk.END, iid=index, values=list(row))

# Função para preencher os campos de entrada ao selecionar uma linha na planilha de Funcionários
def preencher_campos_funcionarios(event):
    selected_item = tree.selection()[0]
    values = tree.item(selected_item, 'values')
    entry_nome_alt.delete(0, tk.END)
    entry_nome_alt.insert(0, values[0])
    entry_cpf_alt.delete(0, tk.END)
    entry_cpf_alt.insert(0, values[1])
    entry_cargo_alt.delete(0, tk.END)
    entry_cargo_alt.insert(0, values[2])
    entry_data_admissao_alt.delete(0, tk.END)
    entry_data_admissao_alt.insert(0, values[3])
    entry_salario_alt.delete(0, tk.END)
    entry_salario_alt.insert(0, values[4])
    combobox_forma_pagto_func_alt.set(values[5])

# Função para atualizar o dado selecionado
def atualizar_dado_funcionarios():
    selected_item = tree.selection()[0]
    df = pd.read_excel(funcionarios_file_path)
    df.at[int(selected_item), 'NOME'] = entry_nome_alt.get()
    df.at[int(selected_item), 'CPF'] = entry_cpf_alt.get()
    df.at[int(selected_item), 'CARGO'] = entry_cargo_alt.get()
    df.at[int(selected_item), 'DATA ADMISSÃO'] = entry_data_admissao_alt.get()
    df.at[int(selected_item), 'SALÁRIO'] = entry_salario_alt.get()
    df.at[int(selected_item), 'FORMA DE PAGAMENTO'] = combobox_forma_pagto_func_alt.get()
    df.to_excel(funcionarios_file_path, index=False)
    carregar_dados_funcionarios()
    messagebox.showinfo("Sucesso", "Dados atualizados com sucesso")

# Função para deletar o dado selecionado
def deletar_dado_funcionarios():
    selected_item = tree.selection()[0]
    df = pd.read_excel(funcionarios_file_path)
    df.drop(index=int(selected_item), inplace=True)
    df.to_excel(funcionarios_file_path, index=False)
    carregar_dados_funcionarios()
    messagebox.showinfo("Sucesso", "Dados deletados com sucesso")

# Função para abrir a planilha no aplicativo padrão
def abrir_planilha():
    try:
        os.startfile(funcionarios_file_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir a planilha: {e}")

# Criar a interface gráfica
root = tk.Tk()
root.title("Controle de Funcionários")

# Criação do Menu
menubar = tk.Menu(root)
root.config(menu=menubar)

# Adicionar menu 'Opções'
opcoes_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Opções", menu=opcoes_menu)
opcoes_menu.add_command(label="Abrir Planilha", command=abrir_planilha)

# Criação do Notebook para abas
notebook = ttk.Notebook(root)
notebook.pack(pady=10, expand=True)

# Frame para Adicionar Dados
frame_adicionar = ttk.Frame(notebook, width=400, height=280)
frame_adicionar.pack(fill='both', expand=True)
notebook.add(frame_adicionar, text='Adicionar Dados')

# Frame para Alterar/Remover Dados
frame_alterar = ttk.Frame(notebook, width=400, height=280)
frame_alterar.pack(fill='both', expand=True)
notebook.add(frame_alterar, text='Alterar/Remover Dados')

# --- Aba Adicionar Dados ---
# Criar campos de entrada
labels = ['NOME', 'CPF', 'CARGO', 'DATA ADMISSÃO', 'SALÁRIO', 'FORMA DE PAGAMENTO']
entries = []

# NOME
frame = ttk.Frame(frame_adicionar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[0])
lbl.pack(side='left')
entry_nome = ttk.Entry(frame)
entry_nome.pack(fill='x', expand=True)

# CPF
frame = ttk.Frame(frame_adicionar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[1])
lbl.pack(side='left')
entry_cpf = ttk.Entry(frame)
entry_cpf.pack(fill='x', expand=True)

# CARGO
frame = ttk.Frame(frame_adicionar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[2])
lbl.pack(side='left')
entry_cargo = ttk.Entry(frame)
entry_cargo.pack(fill='x', expand=True)

# DATA ADMISSÃO
frame = ttk.Frame(frame_adicionar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[3])
lbl.pack(side='left')
entry_data_admissao = ttk.Entry(frame)
entry_data_admissao.pack(fill='x', expand=True)

# SALÁRIO
frame = ttk.Frame(frame_adicionar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[4])
lbl.pack(side='left')
entry_salario = ttk.Entry(frame)
entry_salario.pack(fill='x', expand=True)

# Adicionar o Combobox para forma de pagamento
frame_pagto = ttk.Frame(frame_adicionar)
frame_pagto.pack(fill='x', padx=5, pady=5)
lbl_pagto = ttk.Label(frame_pagto, text='FORMA DE PAGAMENTO')
lbl_pagto.pack(side='left')
combobox_forma_pagto_func = ttk.Combobox(frame_pagto, values=['Cartão', 'Pix', 'Dinheiro'])
combobox_forma_pagto_func.pack(fill='x', expand=True)

# Botão para adicionar os dados
btn_adicionar = ttk.Button(frame_adicionar, text="Adicionar Dados", command=adicionar_dados_funcionarios)
btn_adicionar.pack(pady=10)

# --- Aba Alterar/Remover Dados ---
# Treeview para exibir os dados
tree = ttk.Treeview(frame_alterar, columns=labels, show='headings')
for label in labels:
    tree.heading(label, text=label)
    tree.column(label, minwidth=0, width=100)
tree.pack(fill='both', expand=True, padx=5, pady=5)

# Evento para preencher campos ao selecionar uma linha
tree.bind('<ButtonRelease-1>', preencher_campos_funcionarios)

# Botões de ação
frame_botoes = ttk.Frame(frame_alterar)
frame_botoes.pack(fill='x', padx=5, pady=5)

btn_carregar = ttk.Button(frame_botoes, text="Carregar Dados", command=carregar_dados_funcionarios)
btn_carregar.pack(side='left', padx=5, pady=5)

btn_atualizar = ttk.Button(frame_botoes, text="Atualizar Dado", command=atualizar_dado_funcionarios)
btn_atualizar.pack(side='left', padx=5, pady=5)

btn_deletar = ttk.Button(frame_botoes, text="Deletar Dado", command=deletar_dado_funcionarios)
btn_deletar.pack(side='left', padx=5, pady=5)

# Campos de entrada para edição
entries_alt = []

# NOME
frame = ttk.Frame(frame_alterar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[0])
lbl.pack(side='left')
entry_nome_alt = ttk.Entry(frame)
entry_nome_alt.pack(fill='x', expand=True)

# CPF
frame = ttk.Frame(frame_alterar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[1])
lbl.pack(side='left')
entry_cpf_alt = ttk.Entry(frame)
entry_cpf_alt.pack(fill='x', expand=True)

# CARGO
frame = ttk.Frame(frame_alterar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[2])
lbl.pack(side='left')
entry_cargo_alt = ttk.Entry(frame)
entry_cargo_alt.pack(fill='x', expand=True)

# DATA ADMISSÃO
frame = ttk.Frame(frame_alterar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[3])
lbl.pack(side='left')
entry_data_admissao_alt = ttk.Entry(frame)
entry_data_admissao_alt.pack(fill='x', expand=True)

# SALÁRIO
frame = ttk.Frame(frame_alterar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[4])
lbl.pack(side='left')
entry_salario_alt = ttk.Entry(frame)
entry_salario_alt.pack(fill='x', expand=True)

# Adicionar o Combobox para forma de pagamento
frame_pagto_alt = ttk.Frame(frame_alterar)
frame_pagto_alt.pack(fill='x', padx=5, pady=5)
lbl_pagto_alt = ttk.Label(frame_pagto_alt, text='FORMA DE PAGAMENTO')
lbl_pagto_alt.pack(side='left')
combobox_forma_pagto_func_alt = ttk.Combobox(frame_pagto_alt, values=['Cartão', 'Pix', 'Dinheiro'])
combobox_forma_pagto_func_alt.pack(fill='x', expand=True)

# Executar a aplicação
root.mainloop()
