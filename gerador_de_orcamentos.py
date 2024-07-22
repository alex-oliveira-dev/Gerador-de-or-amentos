import platform
from cx_Freeze import setup, Executable
import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from docx import Document
import os
from docx2pdf import convert
import threading


# Função para atualizar a barra de progresso
def update_progress(progress):
    try:
        progressbar["value"] = progress
        root.update_idletasks()
    except Exception as e:
        messagebox.showerror("Erro ao atualizar progresso", f"Ocorreu um erro ao atualizar a barra de progresso: {e}")

# Função para atualizar a mensagem de status
def update_status(message):
    try:
        status_label.config(text=message)
    except Exception as e:
        messagebox.showerror("Erro ao atualizar status", f"Ocorreu um erro ao atualizar o status: {e}")

# Função para resetar a interface
def reset_ui():
    try:
        progressbar["value"] = 0
        update_status("")
        root.update_idletasks()
    except Exception as e:
        messagebox.showerror("Erro ao resetar interface", f"Ocorreu um erro ao resetar a interface: {e}")

# Função para gerar o orçamento
def generate_quote():
    try:
        # Obter os valores dos campos value1 a value5
        values = [
            value1_entry.get(),
            value2_entry.get(),
            value3_entry.get(),
            value4_entry.get(),
            value5_entry.get(),
            nome_entry.get(),
        ]

        # Verificar se algum campo está vazio
        if not all(values) or not nome_entry.get() or not prefix_entry.get() or not plate_entry.get() or not data_entry.get() or not description1_entry.get():
            messagebox.showwarning("Aviso", "Por favor, preencha todos os campos!")
            return

        # Exibir mensagem de "AGUARDE, GERANDO ORÇAMENTO"
        update_status("AGUARDE, GERANDO ORÇAMENTO...")

        # Converter os valores para inteiros e somá-los
        total = sum(int(value) for value in values if value.isdigit())

        # Atualizar a barra de progresso para 10%
        update_progress(10)

        # Obtenha os valores de entrada
        nome = nome_entry.get()
        prefix = prefix_entry.get()
        plate = plate_entry.get()
        data = data_entry.get()
        description1 = description1_entry.get()
        description2 = description2_entry.get()
        description3 = description3_entry.get()
        description4 = description4_entry.get()
        description5 = description5_entry.get()

        value1 = value1_entry.get()
        value2 = value2_entry.get()
        value3 = value3_entry.get()
        value4 = value4_entry.get()
        value5 = value5_entry.get()

        # Atualizar a barra de progresso para 30%
        update_progress(30)

        # Atualizar a mensagem de status
        update_status("AGUARDE, GERANDO ORÇAMENTO...")

        # Gerar o arquivo apenas se todos os campos estiverem preenchidos
        template_path = "templates/orcamento.docx"
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Arquivo de template não encontrado: {template_path}")
        
        doc = Document(template_path)
        for table in doc.tables:
            # Percorrer todas as linhas e células da tabela
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text
                    # Substituir os marcadores de posição pelos dados fornecidos
                    cell_text = cell_text.replace("[NOME]", nome)
                    cell_text = cell_text.replace("[PREFIX]", prefix)
                    cell_text = cell_text.replace("[PLATE]", plate)
                    cell_text = cell_text.replace("[DATA]", data)
                    cell_text = cell_text.replace("[DESCRIPTION1]", description1)
                    cell_text = cell_text.replace("[DESCRIPTION2]", description2)
                    cell_text = cell_text.replace("[DESCRIPTION3]", description3)
                    cell_text = cell_text.replace("[DESCRIPTION4]", description4)
                    cell_text = cell_text.replace("[DESCRIPTION5]", description5)
                    cell_text = cell_text.replace("[VALUE1]", value1)
                    cell_text = cell_text.replace("[VALUE2]", value2)
                    cell_text = cell_text.replace("[VALUE3]", value3)
                    cell_text = cell_text.replace("[VALUE4]", value4)
                    cell_text = cell_text.replace("[VALUE5]", value5)
                    cell_text = cell_text.replace("[VALUET]", str(total))
                    # Atribuir o texto modificado à célula
                    cell.text = cell_text

        # Atualizar a barra de progresso para 50%
        update_progress(50)

        # Define o nome do arquivo DOCX
        file_name = f"Orçamento_{prefix}_{plate}_{data}.docx"
        # Define o nome da pasta dentro do diretório raiz do projeto onde os arquivos serão salvos
        output_folder = "orçamentos"  # A pasta 'orçamentos' deve estar no diretório raiz do seu projeto

        # Caminho completo para a pasta de destino
        output_path = os.path.join(os.getcwd(), output_folder)

        # Verifica se a pasta de destino existe, caso contrário, cria a pasta
        if not os.path.exists(output_path):
            os.makedirs(output_path)

        # Define o caminho completo para o arquivo incluindo o nome da pasta de destino
        file_path = os.path.join(output_path, file_name)

        # Salva o arquivo DOCX
        doc.save(file_path)

        # Atualizar a barra de progresso para 70%
        update_progress(70)

        # Converte para PDF e salva na mesma pasta
        pdf_output_path = file_path.replace(".docx", ".pdf")
        try:
            convert(file_path, pdf_output_path)
        except Exception as e:
            raise RuntimeError(f"Erro ao converter DOCX para PDF: {e}")

        # Atualizar a barra de progresso para 90%
        update_progress(90)

        # Remove o arquivo DOCX após a conversão
        try:
            os.remove(file_path)
        except Exception as e:
            raise RuntimeError(f"Erro ao remover o arquivo DOCX: {e}")

        # Atualizar a barra de progresso para 100%
        update_progress(100)

        # Atualizar a mensagem de status para "CONCLUÍDO"
        update_status("CONCLUÍDO")

        # Limpa os campos de entrada
        nome_entry.delete(0, "end")
        prefix_entry.delete(0, "end")
        plate_entry.delete(0, "end")
        data_entry.delete(0, "end")
        description1_entry.delete(0, "end")
        description2_entry.delete(0, "end")
        description3_entry.delete(0, "end")
        description4_entry.delete(0, "end")
        description5_entry.delete(0, "end")
        value1_entry.delete(0, "end")
        value2_entry.delete(0, "end")
        value3_entry.delete(0, "end")
        value4_entry.delete(0, "end")
        value5_entry.delete(0, "end")

        # Exibe uma mensagem de sucesso
        message = f"Orçamento gerado com sucesso! Salvo como '{pdf_output_path}'"
        if messagebox.askyesno(
            "Abrir arquivo", message + "\n\nDeseja abrir o arquivo gerado?"
        ):
            try:
                if platform.system() == "Windows":  # Verifica o sistema operacional
                    os.startfile(pdf_output_path)
                else:
                    subprocess.call(["xdg-open", pdf_output_path])  # Para sistemas Linux
            except Exception as e:
                messagebox.showerror("Erro ao abrir o arquivo", f"Erro ao abrir o arquivo: {e}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

    # Limpar a interface para permitir a criação de outro orçamento
    reset_ui()


# Create the main window
root = tk.Tk()

# Definir o título da janela
root.wm_title("GERADOR DE ORÇAMENTOS HSO - versão 1.5")

# Define o tamanho da janela interativa
root.minsize(width=600, height=550)
root.maxsize(width=600, height=550)

# Define a cor da tela do usuário
root.configure(bg="#20B2AA")

# Create the input fields
# itens da coluna da esquerda
nome_label = ttk.Label(root, text="NOME DO CLIENTE:")
nome_entry = ttk.Entry(root, width=40)

prefix_label = ttk.Label(root, text="PREFIXO DO VEICULO:")
prefix_entry = ttk.Entry(root, width=40)

plate_label = ttk.Label(root, text="PLACA DO VEICULO:")
plate_entry = ttk.Entry(root, width=40)

data_label = ttk.Label(root, text="DATA DE SOLICITAÇÃO:")
data_entry = ttk.Entry(root, width=40)

description1_label = ttk.Label(root, text="DESCRIÇÃO DO SERVIÇO:")
description1_entry = ttk.Entry(root, width=40)

description2_label = ttk.Label(root, text="DESCRIÇÃO DO SERVIÇO:")
description2_entry = ttk.Entry(root, width=40)

description3_label = ttk.Label(root, text="DESCRIÇÃO DO SERVIÇO:")
description3_entry = ttk.Entry(root, width=40)

description4_label = ttk.Label(root, text="DESCRIÇÃO DO SERVIÇO:")
description4_entry = ttk.Entry(root, width=40)

description5_label = ttk.Label(root, text="DESCRIÇÃO DO SERVIÇO:")
description5_entry = ttk.Entry(root, width=40)
# itens da coluna da esquerda

# itens da coluna da direita

value1_label = ttk.Label(root, text="VALOR:")
value1_entry = ttk.Entry(root, width=10)

value2_label = ttk.Label(root, text="VALOR:")
value2_entry = ttk.Entry(root, width=10)

value3_label = ttk.Label(root, text="VALOR:")
value3_entry = ttk.Entry(root, width=10)

value4_label = ttk.Label(root, text="VALOR:")
value4_entry = ttk.Entry(root, width=10)

value5_label = ttk.Label(root, text="VALOR:")
value5_entry = ttk.Entry(root, width=10)

# Create the generate quote button
generate_button = ttk.Button(
    root, text="GERAR ORÇAMENTO", command=generate_quote, width=30
)

# Create progress bar
progressbar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")

# Create status label
status_label = ttk.Label(root, text="", foreground="blue")

# Empacote os campos de entrada, botões e rótulos
nome_label.grid(row=0, column=0, padx=(40, 10), pady=(10, 10), sticky="e")
nome_entry.grid(row=0, column=1, padx=5, pady=(10, 10))

prefix_label.grid(row=1, column=0, padx=(40, 10), pady=(10, 10), sticky="e")
prefix_entry.grid(row=1, column=1, padx=5, pady=(10, 10))

plate_label.grid(row=2, column=0, padx=(40, 10), pady=(10, 10), sticky="e")
plate_entry.grid(row=2, column=1, padx=5, pady=(10, 10))

data_label.grid(row=3, column=0, padx=(40, 10), pady=(10, 10), sticky="e")
data_entry.grid(row=3, column=1, padx=5, pady=(10, 10))

description1_label.grid(row=4, column=0, padx=(40, 10), pady=(10, 10), sticky="e")
description1_entry.grid(row=4, column=1, padx=5, pady=(10, 10))

description2_label.grid(row=5, column=0, padx=(40, 10), pady=(10, 10), sticky="e")
description2_entry.grid(row=5, column=1, padx=5, pady=(10, 10))

description3_label.grid(row=6, column=0, padx=(40, 10), pady=(10, 10), sticky="e")
description3_entry.grid(row=6, column=1, padx=5, pady=(10, 10))

description4_label.grid(row=7, column=0, padx=(40, 10), pady=(10, 10), sticky="e")
description4_entry.grid(row=7, column=1, padx=5, pady=(10, 10))

description5_label.grid(row=8, column=0, padx=(40, 10), pady=(10, 10), sticky="e")
description5_entry.grid(row=8, column=1, padx=5, pady=(10, 10))

value1_label.grid(row=4, column=2, padx=(10, 10), pady=(10, 10), sticky="")
value1_entry.grid(row=4, column=3, padx=(10, 10), pady=(10, 10))

value2_label.grid(row=5, column=2, padx=(10, 10), pady=(10, 10), sticky="")
value2_entry.grid(row=5, column=3, padx=(10, 10), pady=(10, 10))

value3_label.grid(row=6, column=2, padx=(10, 10), pady=(10, 10), sticky="")
value3_entry.grid(row=6, column=3, padx=(10, 10), pady=(10, 10))

value4_label.grid(row=7, column=2, padx=(10, 10), pady=(10, 10), sticky="")
value4_entry.grid(row=7, column=3, padx=(10, 10), pady=(10, 10))

value5_label.grid(row=8, column=2, padx=(10, 10), pady=(10, 10), sticky="")
value5_entry.grid(row=8, column=3, padx=(10, 10), pady=(10, 10))

generate_button.grid(
    row=13, column=0, columnspan=4, padx=100, pady=(10, 10), sticky="ew"
)
progressbar.grid(row=14, column=0, columnspan=4, padx=10, pady=(10, 10))
status_label.grid(row=15, column=0, columnspan=4, padx=10, pady=(0, 10))

# Start the mainloop
root.mainloop()
  