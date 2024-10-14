import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import PyPDF2
import os
import openpyxl
import threading

def split_pdf(input_pdf, output_dir, lote_number, excel_name):
    """Dividir o PDF em arquivos de 3 páginas e salvar detalhes em um arquivo Excel."""
    try:
        with open(input_pdf, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            num_pages = len(reader.pages)

            # Cria o arquivo Excel
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Detalhes Segurados"
            sheet.append(["Arquivo"])

            count_associado = 1  # Iniciar a contagem a partir de 1

            for start_page in range(0, num_pages, 3):
                writer = PyPDF2.PdfWriter()

                # Adiciona 3 páginas ao novo arquivo PDF
                for i in range(3):
                    if start_page + i < num_pages:
                        writer.add_page(reader.pages[start_page + i])

                # Nome do arquivo será AP_MAPFRE + número do lote + número sequencial
                output_pdf = os.path.join(output_dir, f'AP_MAPFRE_{lote_number}_{count_associado:04d}.pdf')

                with open(output_pdf, 'wb') as output_file:
                    writer.write(output_file)

                # Adiciona informações ao registro do Excel
                sheet.append([output_pdf])

                count_associado += 1  # Incrementa o contador para o próximo arquivo
            
            # Salva o arquivo Excel
            excel_output = os.path.join(output_dir, f'{excel_name}.xlsx')
            workbook.save(excel_output)

            return True, output_dir  # Retorna sucesso e o diretório de saída
            
    except FileNotFoundError:
        return False, "Arquivo não encontrado. Verifique o caminho do arquivo PDF."
    except PyPDF2.errors.PdfReadError:
        return False, "Erro ao ler o arquivo PDF. Verifique se o arquivo está corrompido."
    except Exception as e:
        return False, f"Ocorreu um erro: {str(e)}"

def drop(event):
    """Função para lidar com o evento de arrastar e soltar um arquivo."""    
    file_path = event.data.strip('"')  # Remove aspas se estiverem presentes
    choose_output_directory(file_path)

def choose_output_directory(input_pdf):
    """Escolhe o diretório de saída e inicia o processamento do PDF."""
    output_dir = filedialog.askdirectory(title="Escolha o diretório de saída")
    if output_dir:
        lote_number = lote_entry.get().strip()  # Obtém o número do lote da entrada
        excel_name = excel_name_entry.get().strip()  # Obtém o nome do arquivo Excel

        # Validação do número do lote
        if not lote_number.isdigit() or not lote_number:  
            messagebox.showwarning("Aviso", "Por favor, insira um número de lote válido.")
            return

        if not excel_name:
            messagebox.showwarning("Aviso", "Por favor, insira um nome para o arquivo Excel.")
            return

        # Desativa os botões e a área de arrasto durante o processamento
        btn_upload.config(state=tk.DISABLED)
        drop_area.config(state=tk.DISABLED)

        # Atualiza a mensagem de status e inicia a barra de progresso
        status_label.config(text="Processando, aguarde...")
        progress_bar.start()
        root.update()  # Atualiza a interface

        # Executa o processamento em uma thread separada
        threading.Thread(target=process_pdf, args=(input_pdf, output_dir, lote_number, excel_name)).start()

def process_pdf(input_pdf, output_dir, lote_number, excel_name):
    """Função para processar o PDF e dividir em arquivos menores."""    
    success, message = split_pdf(input_pdf, output_dir, lote_number, excel_name)
    
    # Finaliza a barra de progresso
    progress_bar.stop()

    if success:
        messagebox.showinfo("Sucesso", f"Arquivos salvos com sucesso em {output_dir}!\nDetalhes salvos em {excel_name}.xlsx.")
    else:
        messagebox.showerror("Erro", message)

    # Limpa as entradas após a conclusão
    clear_inputs()

def upload_file():
    """Permite ao usuário selecionar um arquivo PDF."""
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        choose_output_directory(file_path)

def clear_inputs():
    """Limpa as entradas da interface."""
    lote_entry.delete(0, tk.END)  # Limpa a entrada do número do lote
    excel_name_entry.delete(0, tk.END)  # Limpa a entrada do nome do Excel

def create_help_menu():
    """Cria um menu de ajuda."""
    help_menu = tk.Menu(menu)
    help_menu.add_command(label="Instruções", command=lambda: messagebox.showinfo("Instruções", 
    "1. Arraste e solte um arquivo PDF ou use o botão 'Selecionar PDF'.\n"
    "2. Insira o número do lote.\n"
    "3. Insira o nome do arquivo Excel.\n"
    "4. Clique em 'Processar'.\n"
    "5. Os arquivos gerados serão salvos no diretório selecionado."))
    menu.add_cascade(label="Ajuda", menu=help_menu)

# Cria a janela principal
root = TkinterDnD.Tk()
root.title("Divisor de PDF")
root.geometry("600x450")
root.configure(bg="#f7f7f7")  # Cor de fundo da janela
root.resizable(True, True)  # Permite redimensionar a janela

# Cria um menu
menu = tk.Menu(root)
root.config(menu=menu)
create_help_menu()

# Cria um frame para a entrada do lote
input_frame = tk.Frame(root, bg="#f7f7f7")
input_frame.pack(pady=10)

# Cria um label e uma entrada para o número do lote
lote_label = tk.Label(input_frame, text="Número do Lote:", bg="#f7f7f7", font=("Arial", 12))
lote_label.pack(side=tk.LEFT)

lote_entry = tk.Entry(input_frame, width=10, font=("Arial", 12), bd=2, relief="solid")
lote_entry.pack(side=tk.LEFT, padx=5)

# Cria um label e uma entrada para o nome do arquivo Excel
excel_name_label = tk.Label(input_frame, text="Nome do Arquivo Excel:", bg="#f7f7f7", font=("Arial", 12))
excel_name_label.pack(side=tk.LEFT)

excel_name_entry = tk.Entry(input_frame, width=20, font=("Arial", 12), bd=2, relief="solid")
excel_name_entry.pack(side=tk.LEFT, padx=5)

# Cria um frame para a área de arrasto e o botão
action_frame = tk.Frame(root, bg="#f7f7f7")
action_frame.pack(pady=20)

# Define a área para arrastar e soltar
drop_area = tk.Label(action_frame, text="Arraste e solte seu PDF aqui", padx=10, pady=10, relief="groove", bg="#e6e6e6", font=("Arial", 12))
drop_area.pack(padx=20, pady=10, expand=True, fill=tk.BOTH)

# Configura a área para aceitar arquivos
drop_area.drop_target_register(DND_FILES)
drop_area.dnd_bind('<<Drop>>', drop)

# Adiciona botão de upload
btn_upload = tk.Button(action_frame, text="Selecionar PDF", command=upload_file, bg="#4CAF50", fg="white", font=("Arial", 12), activebackground="#45a049")
btn_upload.pack(pady=10)

# Adiciona um label para status
status_label = tk.Label(root, text="Pronto!", bg="#f7f7f7", font=("Arial", 12))
status_label.pack(pady=5)

# Adiciona uma barra de progresso
progress_bar = ttk.Progressbar(root, orient='horizontal', mode='indeterminate')
progress_bar.pack(pady=10, fill=tk.X, padx=20)

# Adiciona o rodapé
footer_label = tk.Label(root, text="Desenvolvido por: L7ps", bg="#f7f7f7", font=("Arial", 10))
footer_label.pack(side=tk.BOTTOM, pady=10)

# Inicia a janela
root.mainloop()
