import tkinter as tk
from tkinter import ttk
import pandas as pd
import os
import tkinter.messagebox as messagebox

# Função para adicionar os dados na planilha
def adicionar_dados():
    # Obter os dados dos campos de entrada
    proprietario = entry_proprietario.get()
    projeto_simplificado = entry_projeto_simplificado.get()
    projeto_detalhado = entry_projeto_detalhado.get()
    regularizacoes = entry_regularizacoes.get()
    desdobro = entry_desdobro.get()
    proj_bombeiro = entry_proj_bombeiro.get()
    data_inicio = entry_data_inicio.get()
    data_termico = entry_data_termico.get()
    valor_total = entry_valor_total.get()
    forma_pagto = combobox_forma_pagto.get()

    # Criar um DataFrame com os novos dados
    novos_dados = pd.DataFrame([[proprietario, projeto_simplificado, projeto_detalhado, regularizacoes, desdobro, proj_bombeiro, data_inicio, data_termico, valor_total, forma_pagto]],
                               columns=['PROPRIETÁRIO', 'PROJETO SIMPLIFICADO', 'PROJETO DETALHADO', 'REGULARIZAÇOES', 'DESDOBRO', 'PROJ BOMBEIRO', 'DATA INICIO', 'DATA TÉRMICO', 'VALOR TOTAL', 'FORMA DE PAGTO'])
    
    # Checar se o arquivo existe
    file_path = 'PLANILHA DE PROJETOS E CUSTOS.xlsx'
    try:
        if os.path.exists(file_path):
            try:
                # Ler a planilha existente sem cabeçalhos para detectar a estrutura
                df = pd.read_excel(file_path, header=None)
                # Encontrar a linha de cabeçalhos
                header_row = df[df.iloc[:, 0] == 'PROPRIETÁRIO'].index[0]
                df.columns = df.iloc[header_row]
                df = df[header_row + 1:].reset_index(drop=True)
                # Verificar se há uma linha vazia para adicionar os novos dados
                for i in range(len(df)):
                    if pd.isna(df.iloc[i]['PROPRIETÁRIO']):
                        df.iloc[i] = novos_dados.iloc[0]
                        break
                else:
                    # Se não houver linha vazia, adicionar os novos dados ao final
                    df = pd.concat([df, novos_dados], ignore_index=True)
            except Exception as e:
                print(f"Erro ao ler a planilha existente: {e}")
                # Criar um DataFrame vazio com os cabeçalhos apropriados se houver erro
                df = pd.DataFrame(columns=['PROPRIETÁRIO', 'PROJETO SIMPLIFICADO', 'PROJETO DETALHADO', 'REGULARIZAÇOES', 'DESDOBRO', 'PROJ BOMBEIRO', 'DATA INICIO', 'DATA TÉRMICO', 'VALOR TOTAL', 'FORMA DE PAGTO'])
                # Adicionar os novos dados
                df = pd.concat([df, novos_dados], ignore_index=True)
        else:
            # Criar um DataFrame vazio com os cabeçalhos apropriados
            df = pd.DataFrame(columns=['PROPRIETÁRIO', 'PROJETO SIMPLIFICADO', 'PROJETO DETALHADO', 'REGULARIZAÇOES', 'DESDOBRO', 'PROJ BOMBEIRO', 'DATA INICIO', 'DATA TÉRMICO', 'VALOR TOTAL', 'FORMA DE PAGTO'])
            # Adicionar os novos dados
            df = pd.concat([df, novos_dados], ignore_index=True)
        
        # Salvar a planilha
        df.to_excel(file_path, index=False)
        
        # Limpar os campos de entrada
        entry_proprietario.delete(0, tk.END)
        entry_projeto_simplificado.delete(0, tk.END)
        entry_projeto_detalhado.delete(0, tk.END)
        entry_regularizacoes.delete(0, tk.END)
        entry_desdobro.delete(0, tk.END)
        entry_proj_bombeiro.delete(0, tk.END)
        entry_data_inicio.delete(0, tk.END)
        entry_data_termico.delete(0, tk.END)
        entry_valor_total.delete(0, tk.END)
        combobox_forma_pagto.set('')
    except PermissionError:
        messagebox.showerror("Erro de Permissão", "Não foi possível acessar o arquivo. Verifique se ele está aberto em outro programa e tente novamente.")

# Criar a interface gráfica
root = tk.Tk()
root.title("Controle de Projetos")

# Criar campos de entrada
labels = ['PROPRIETÁRIO', 'PROJETO SIMPLIFICADO', 'PROJETO DETALHADO', 'REGULARIZAÇOES', 'DESDOBRO', 'PROJ BOMBEIRO', 'DATA INICIO', 'DATA TÉRMICO', 'VALOR TOTAL', 'FORMA DE PAGTO']
entries = []
for label in labels[:-1]:
    frame = ttk.Frame(root)
    frame.pack(fill='x')
    lbl = ttk.Label(frame, text=label)
    lbl.pack(side='left')
    entry = ttk.Entry(frame)
    entry.pack(fill='x', expand=True)
    entries.append(entry)

entry_proprietario, entry_projeto_simplificado, entry_projeto_detalhado, entry_regularizacoes, entry_desdobro, entry_proj_bombeiro, entry_data_inicio, entry_data_termico, entry_valor_total = entries

# Adicionar o Combobox para forma de pagamento
frame_pagto = ttk.Frame(root)
frame_pagto.pack(fill='x')
lbl_pagto = ttk.Label(frame_pagto, text='FORMA DE PAGTO')
lbl_pagto.pack(side='left')
combobox_forma_pagto = ttk.Combobox(frame_pagto, values=['Cartão', 'Pix', 'Dinheiro'])
combobox_forma_pagto.pack(fill='x', expand=True)

# Botão para adicionar os dados
btn_adicionar = ttk.Button(root, text="Adicionar Dados", command=adicionar_dados)
btn_adicionar.pack()

root.mainloop()
