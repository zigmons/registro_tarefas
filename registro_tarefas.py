import tkinter as tk
from tkinter import ttk
import time
import pandas as pd
import os
import json
import subprocess

# Variáveis globais
start_time = None
end_time = None
filename = 'Registro.xlsx'
config_filename = 'config.json'

# Carrega o caminho do arquivo a partir do arquivo de configuração


def load_file_path():
    try:
        with open(config_filename, 'r') as f:
            config = json.load(f)
        return config.get('file_path')
    except FileNotFoundError:
        return None

# Salva o caminho do arquivo no arquivo de configuração


def save_file_path(file_path):
    with open(config_filename, 'w') as f:
        json.dump({'file_path': file_path}, f)


def atualizar_cronometro():
    if start_time is not None:
        elapsed_time = time.time() - start_time
        label_cronometro['text'] = time.strftime(
            '%H:%M:%S', time.gmtime(elapsed_time))
        root.after(1000, atualizar_cronometro)


def iniciar_tarefa():
    global start_time
    start_time = time.time()
    # Esconde os widgets
    label_nome.pack_forget()
    entry_nome.pack_forget()
    label_tipo.pack_forget()
    combo_tipo.pack_forget()
    button_iniciar.pack_forget()
    button_abrir_pasta.pack_forget()
    # Mostra o cronômetro e o botão "Parar"
    label_cronometro.pack()
    button_parar.pack()
    atualizar_cronometro()


def parar_tarefa():
    global start_time, end_time
    end_time = time.time()
    elapsed_time = end_time - start_time

    # Cria um DataFrame com os dados da tarefa
    data = {
        'Nome': [entry_nome.get()],
        'Tipo': [combo_tipo.get()],
        'Data': [time.strftime('%d/%m/%Y')],
        'Hora Início': [time.strftime('%H:%M:%S', time.gmtime(start_time))],
        'Hora Final': [time.strftime('%H:%M:%S', time.gmtime(end_time))],
        'Duração': [time.strftime('%H:%M:%S', time.gmtime(elapsed_time))]
    }
    df_tarefa = pd.DataFrame(data)

    # Se for a primeira tarefa, define o local do arquivo como a raiz do diretório C:
    filename = load_file_path()
    if filename is None:
        filename = os.path.join('C:\\', 'Registro_Tarefas', 'Registro.xlsx')
        save_file_path(filename)  # Salva o caminho do arquivo

    # Se o arquivo já existir, carrega o DataFrame existente
    if os.path.exists(filename):
        df = pd.read_excel(filename)
        df = pd.concat([df, df_tarefa])
    else:
        df = df_tarefa

    # Salva o DataFrame em um arquivo do Excel
    df.to_excel(filename, index=False)

    # Esconde o cronômetro e o botão "Parar" e mostra os widgets
    label_cronometro.pack_forget()
    button_parar.pack_forget()
    label_nome.pack()
    entry_nome.pack()
    label_tipo.pack()
    combo_tipo.pack()
    button_iniciar.pack()
    button_abrir_pasta.pack()


def abrir_pasta():
    filename = load_file_path()
    if filename is not None:
        subprocess.Popen(f'explorer /select,"{filename}"')


root = tk.Tk()
root.title('Registro de Tarefas')
root.geometry('300x200')

# Campo para o nome da tarefa
label_nome = tk.Label(root, text='Nome da tarefa:')
label_nome.pack()
entry_nome = tk.Entry(root)
entry_nome.pack()

# Lista de tipos de tarefa
label_tipo = tk.Label(root, text='Tipo de tarefa:')
label_tipo.pack()
combo_tipo = ttk.Combobox(root, values=['Comercial', 'Projetos', 'Tech'])
combo_tipo.pack()

# Botões de iniciar e parar
button_iniciar = tk.Button(root, text='Iniciar', command=iniciar_tarefa)
button_iniciar.pack()
button_parar = tk.Button(root, text='Parar', command=parar_tarefa)

# Botão para abrir a pasta
button_abrir_pasta = tk.Button(root, text='Abrir Pasta', command=abrir_pasta)
button_abrir_pasta.pack()

# Cronômetro
label_cronometro = tk.Label(root)

# Créditos
label_creditos = tk.Label(root, text='Criado por Rafael Sousa', anchor='sw')
label_creditos.pack(side=tk.BOTTOM, fill=tk.X)

# Cria o diretório "Registro_Tarefas" se ele não existir
os.makedirs('C:\\Registro_Tarefas', exist_ok=True)

root.mainloop()
