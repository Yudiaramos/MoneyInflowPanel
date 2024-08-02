import tkinter as tk
from tkinter import ttk
import pandas as pd
import os
import tkinter.messagebox as messagebox

financeiro_file_path = 'PLANILHA DE PROJETOS E CUSTOS.xlsx'

# Função para adicionar os dados na planilha de Projetos
def adicionar_dados_projetos():
    # Obter os dados dos campos de entrada
    proprietario = entry_proprietario.get()
    projeto_simplificado = 'Sim' if var_projeto.get() == 1 else ''
    projeto_detalhado = 'Sim' if var_projeto.get() == 2 else ''
    regularizacoes = entry_regularizacoes.get()
    desdobro = 'Sim' if var_desdobro.get() == 1 else 'Não'
    proj_bombeiro = entry_proj_bombeiro.get()
    data_inicio = entry_data_inicio.get()
    data_termico = entry_data_termico.get()
    valor_total = entry_valor_total.get()
    forma_pagto = combobox_forma_pagto.get()
    valor_pago = entry_valor_pago.get()

    # Criar um DataFrame com os novos dados
    novos_dados = pd.DataFrame([[proprietario, projeto_simplificado, projeto_detalhado, regularizacoes, desdobro, proj_bombeiro, data_inicio, data_termico, valor_total, forma_pagto, valor_pago]],
                               columns=['PROPRIETÁRIO', 'PROJETO SIMPLIFICADO', 'PROJETO DETALHADO', 'REGULARIZAÇÕES', 'DESDOBRO', 'PROJ BOMBEIRO', 'DATA INICIO', 'DATA TÉRMICO', 'VALOR TOTAL', 'FORMA DE PAGTO', 'VALOR PAGO'])
    
    try:
        if os.path.exists(financeiro_file_path):
            # Ler a planilha existente
            df = pd.read_excel(financeiro_file_path)
            # Verificar se há uma linha vazia para adicionar os novos dados
            for i in range(len(df)):
                if pd.isna(df.iloc[i]['PROPRIETÁRIO']):
                    df.iloc[i] = novos_dados.iloc[0]
                    break
            else:
                # Se não houver linha vazia, adicionar os novos dados ao final
                df = pd.concat([df, novos_dados], ignore_index=True)
        else:
            # Criar um DataFrame vazio com os cabeçalhos apropriados
            df = pd.DataFrame(columns=['PROPRIETÁRIO', 'PROJETO SIMPLIFICADO', 'PROJETO DETALHADO', 'REGULARIZAÇÕES', 'DESDOBRO', 'PROJ BOMBEIRO', 'DATA INICIO', 'DATA TÉRMICO', 'VALOR TOTAL', 'FORMA DE PAGTO', 'VALOR PAGO'])
            # Adicionar os novos dados
            df = pd.concat([df, novos_dados], ignore_index=True)
        
        # Salvar a planilha
        df.to_excel(financeiro_file_path, index=False)
        
        # Limpar os campos de entrada
        entry_proprietario.delete(0, tk.END)
        var_projeto.set(0)
        entry_regularizacoes.delete(0, tk.END)
        var_desdobro.set(0)
        entry_proj_bombeiro.delete(0, tk.END)
        entry_data_inicio.delete(0, tk.END)
        entry_data_termico.delete(0, tk.END)
        entry_valor_total.delete(0, tk.END)
        combobox_forma_pagto.set('')
        entry_valor_pago.delete(0, tk.END)
    except PermissionError:
        messagebox.showerror("Erro de Permissão", "Não foi possível acessar o arquivo. Verifique se ele está aberto em outro programa e tente novamente.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao tentar salvar os dados: {e}")

# Função para carregar os dados na Treeview para a planilha de Projetos
def carregar_dados_projetos():
    for i in tree.get_children():
        tree.delete(i)
    if os.path.exists(financeiro_file_path):
        df = pd.read_excel(financeiro_file_path)
        for index, row in df.iterrows():
            tree.insert("", tk.END, iid=index, values=list(row))

# Função para preencher os campos de entrada ao selecionar uma linha na planilha de Projetos
def preencher_campos_projetos(event):
    selected_item = tree.selection()[0]
    values = tree.item(selected_item, 'values')
    entry_proprietario_alt.delete(0, tk.END)
    entry_proprietario_alt.insert(0, values[0])
    if values[1] == 'Sim':
        var_projeto_alt.set(1)
    elif values[2] == 'Sim':
        var_projeto_alt.set(2)
    entry_regularizacoes_alt.delete(0, tk.END)
    entry_regularizacoes_alt.insert(0, values[3])
    if values[4] == 'Sim':
        var_desdobro_alt.set(1)
    else:
        var_desdobro_alt.set(2)
    entry_proj_bombeiro_alt.delete(0, tk.END)
    entry_proj_bombeiro_alt.insert(0, values[5])
    entry_data_inicio_alt.delete(0, tk.END)
    entry_data_inicio_alt.insert(0, values[6])
    entry_data_termico_alt.delete(0, tk.END)
    entry_data_termico_alt.insert(0, values[7])
    entry_valor_total_alt.delete(0, tk.END)
    entry_valor_total_alt.insert(0, values[8])
    combobox_forma_pagto_alt.set(values[9])
    entry_valor_pago_alt.delete(0, tk.END)
    entry_valor_pago_alt.insert(0, values[10])

# Função para atualizar o dado selecionado
def atualizar_dado_projetos():
    selected_item = tree.selection()[0]
    df = pd.read_excel(financeiro_file_path)
    df.at[int(selected_item), 'PROPRIETÁRIO'] = entry_proprietario_alt.get()
    df.at[int(selected_item), 'PROJETO SIMPLIFICADO'] = 'Sim' if var_projeto_alt.get() == 1 else ''
    df.at[int(selected_item), 'PROJETO DETALHADO'] = 'Sim' if var_projeto_alt.get() == 2 else ''
    df.at[int(selected_item), 'REGULARIZAÇÕES'] = entry_regularizacoes_alt.get()
    df.at[int(selected_item), 'DESDOBRO'] = 'Sim' if var_desdobro_alt.get() == 1 else 'Não'
    df.at[int(selected_item), 'PROJ BOMBEIRO'] = entry_proj_bombeiro_alt.get()
    df.at[int(selected_item), 'DATA INICIO'] = entry_data_inicio_alt.get()
    df.at[int(selected_item), 'DATA TÉRMICO'] = entry_data_termico_alt.get()
    df.at[int(selected_item), 'VALOR TOTAL'] = entry_valor_total_alt.get()
    df.at[int(selected_item), 'FORMA DE PAGTO'] = combobox_forma_pagto_alt.get()
    df.at[int(selected_item), 'VALOR PAGO'] = entry_valor_pago_alt.get()
    df.to_excel(financeiro_file_path, index=False)
    carregar_dados_projetos()
    messagebox.showinfo("Sucesso", "Dados atualizados com sucesso")

# Função para deletar o dado selecionado
def deletar_dado_projetos():
    selected_item = tree.selection()[0]
    df = pd.read_excel(financeiro_file_path)
    df.drop(index=int(selected_item), inplace=True)
    df.to_excel(financeiro_file_path, index=False)
    carregar_dados_projetos()
    messagebox.showinfo("Sucesso", "Dados deletados com sucesso")

# Função para abrir a planilha no aplicativo padrão
def abrir_planilha():
    try:
        os.startfile(financeiro_file_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir a planilha: {e}")

# Criar a interface gráfica
root = tk.Tk()
root.title("Controle de Projetos")

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
labels = ['PROPRIETÁRIO', 'PROJETO SIMPLIFICADO', 'PROJETO DETALHADO', 'REGULARIZAÇÕES', 'DESDOBRO', 'PROJ BOMBEIRO', 'DATA INICIO', 'DATA TÉRMICO', 'VALOR TOTAL', 'FORMA DE PAGTO', 'VALOR PAGO']
entries = []

# PROPRIETÁRIO
frame = ttk.Frame(frame_adicionar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[0])
lbl.pack(side='left')
entry_proprietario = ttk.Entry(frame)
entry_proprietario.pack(fill='x', expand=True)

# PROJETO SIMPLIFICADO e PROJETO DETALHADO
var_projeto = tk.IntVar()
frame_projeto = ttk.Frame(frame_adicionar)
frame_projeto.pack(fill='x', padx=5, pady=5)
lbl_projeto = ttk.Label(frame_projeto, text="Tipo de Projeto")
lbl_projeto.pack(side='left')
rb_simplificado = ttk.Radiobutton(frame_projeto, text='Simplificado', variable=var_projeto, value=1)
rb_simplificado.pack(side='left')
rb_detalhado = ttk.Radiobutton(frame_projeto, text='Detalhado', variable=var_projeto, value=2)
rb_detalhado.pack(side='left')

# DESDOBRO
var_desdobro = tk.IntVar()
frame_desdobro = ttk.Frame(frame_adicionar)
frame_desdobro.pack(fill='x', padx=5, pady=5)
lbl_desdobro = ttk.Label(frame_desdobro, text="Desdobro")
lbl_desdobro.pack(side='left')
rb_desdobro_sim = ttk.Radiobutton(frame_desdobro, text='Sim', variable=var_desdobro, value=1)
rb_desdobro_sim.pack(side='left')
rb_desdobro_nao = ttk.Radiobutton(frame_desdobro, text='Não', variable=var_desdobro, value=2)
rb_desdobro_nao.pack(side='left')

# Outros campos
for label in labels[3:-2]:
    if label == 'DESDOBRO':
        continue
    frame = ttk.Frame(frame_adicionar)
    frame.pack(fill='x', padx=5, pady=5)
    lbl = ttk.Label(frame, text=label)
    lbl.pack(side='left')
    entry = ttk.Entry(frame)
    entry.pack(fill='x', expand=True)
    entries.append(entry)

entry_regularizacoes, entry_proj_bombeiro, entry_data_inicio, entry_data_termico, entry_valor_total = entries

# Adicionar o Combobox para forma de pagamento
frame_pagto = ttk.Frame(frame_adicionar)
frame_pagto.pack(fill='x', padx=5, pady=5)
lbl_pagto = ttk.Label(frame_pagto, text='FORMA DE PAGTO')
lbl_pagto.pack(side='left')
combobox_forma_pagto = ttk.Combobox(frame_pagto, values=['Cartão', 'Pix', 'Dinheiro'])
combobox_forma_pagto.pack(fill='x', expand=True)

# Adicionar campo para valor pago
frame_valor_pago = ttk.Frame(frame_adicionar)
frame_valor_pago.pack(fill='x', padx=5, pady=5)
lbl_valor_pago = ttk.Label(frame_valor_pago, text='VALOR PAGO')
lbl_valor_pago.pack(side='left')
entry_valor_pago = ttk.Entry(frame_valor_pago)
entry_valor_pago.pack(fill='x', expand=True)

# Botão para adicionar os dados
btn_adicionar = ttk.Button(frame_adicionar, text="Adicionar Dados", command=adicionar_dados_projetos)
btn_adicionar.pack(pady=10)

# --- Aba Alterar/Remover Dados ---
# Treeview para exibir os dados
tree = ttk.Treeview(frame_alterar, columns=labels, show='headings')
for label in labels:
    tree.heading(label, text=label)
    tree.column(label, minwidth=0, width=100)
tree.pack(fill='both', expand=True, padx=5, pady=5)

# Evento para preencher campos ao selecionar uma linha
tree.bind('<ButtonRelease-1>', preencher_campos_projetos)

# Botões de ação
frame_botoes = ttk.Frame(frame_alterar)
frame_botoes.pack(fill='x', padx=5, pady=5)

btn_carregar = ttk.Button(frame_botoes, text="Carregar Dados", command=carregar_dados_projetos)
btn_carregar.pack(side='left', padx=5, pady=5)

btn_atualizar = ttk.Button(frame_botoes, text="Atualizar Dado", command=atualizar_dado_projetos)
btn_atualizar.pack(side='left', padx=5, pady=5)

btn_deletar = ttk.Button(frame_botoes, text="Deletar Dado", command=deletar_dado_projetos)
btn_deletar.pack(side='left', padx=5, pady=5)

# Campos de entrada para edição
entries_alt = []

# PROPRIETÁRIO
frame = ttk.Frame(frame_alterar)
frame.pack(fill='x', padx=5, pady=5)
lbl = ttk.Label(frame, text=labels[0])
lbl.pack(side='left')
entry_proprietario_alt = ttk.Entry(frame)
entry_proprietario_alt.pack(fill='x', expand=True)

# PROJETO SIMPLIFICADO e PROJETO DETALHADO
var_projeto_alt = tk.IntVar()
frame_projeto_alt = ttk.Frame(frame_alterar)
frame_projeto_alt.pack(fill='x', padx=5, pady=5)
lbl_projeto_alt = ttk.Label(frame_projeto_alt, text="Tipo de Projeto")
lbl_projeto_alt.pack(side='left')
rb_simplificado_alt = ttk.Radiobutton(frame_projeto_alt, text='Simplificado', variable=var_projeto_alt, value=1)
rb_simplificado_alt.pack(side='left')
rb_detalhado_alt = ttk.Radiobutton(frame_projeto_alt, text='Detalhado', variable=var_projeto_alt, value=2)
rb_detalhado_alt.pack(side='left')

# DESDOBRO
var_desdobro_alt = tk.IntVar()
frame_desdobro_alt = ttk.Frame(frame_alterar)
frame_desdobro_alt.pack(fill='x', padx=5, pady=5)
lbl_desdobro_alt = ttk.Label(frame_desdobro_alt, text="Desdobro")
lbl_desdobro_alt.pack(side='left')
rb_desdobro_sim_alt = ttk.Radiobutton(frame_desdobro_alt, text='Sim', variable=var_desdobro_alt, value=1)
rb_desdobro_sim_alt.pack(side='left')
rb_desdobro_nao_alt = ttk.Radiobutton(frame_desdobro_alt, text='Não', variable=var_desdobro_alt, value=2)
rb_desdobro_nao_alt.pack(side='left')

# Outros campos
for label in labels[3:-2]:
    if label == 'DESDOBRO':
        continue
    frame = ttk.Frame(frame_alterar)
    frame.pack(fill='x', padx=5, pady=5)
    lbl = ttk.Label(frame, text=label)
    lbl.pack(side='left')
    entry = ttk.Entry(frame)
    entry.pack(fill='x', expand=True)
    entries_alt.append(entry)

entry_regularizacoes_alt, entry_proj_bombeiro_alt, entry_data_inicio_alt, entry_data_termico_alt, entry_valor_total_alt = entries_alt

# Adicionar o Combobox para forma de pagamento
frame_pagto_alt = ttk.Frame(frame_alterar)
frame_pagto_alt.pack(fill='x', padx=5, pady=5)
lbl_pagto_alt = ttk.Label(frame_pagto_alt, text='FORMA DE PAGTO')
lbl_pagto_alt.pack(side='left')
combobox_forma_pagto_alt = ttk.Combobox(frame_pagto_alt, values=['Cartão', 'Pix', 'Dinheiro'])
combobox_forma_pagto_alt.pack(fill='x', expand=True)

# Adicionar campo para valor pago
frame_valor_pago_alt = ttk.Frame(frame_alterar)
frame_valor_pago_alt.pack(fill='x', padx=5, pady=5)
lbl_valor_pago_alt = ttk.Label(frame_valor_pago_alt, text='VALOR PAGO')
lbl_valor_pago_alt.pack(side='left')
entry_valor_pago_alt = ttk.Entry(frame_valor_pago_alt)
entry_valor_pago_alt.pack(fill='x', expand=True)

# Executar a aplicação
root.mainloop()
