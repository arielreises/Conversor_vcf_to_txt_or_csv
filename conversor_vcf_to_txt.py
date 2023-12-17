# Programa Python - Desenvolvido por Ariel Reises (@arielreises)
# GitHub: https://github.com/arielreises

# Destinado ao cliente Cauhe Filipin.

import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import platform

def clear_terminal():
    if platform.system() == 'Windows':
        os.system('cls')
    else:
        os.system('clear')
        
# Função para instalar uma biblioteca usando o pip
def install_package(package):
    subprocess.check_call(["python", "-m", "pip", "install", package])

# Verifica e instala a biblioteca vobject se necessário
try:
    import vobject
except ImportError:
    print("A biblioteca vobject não está instalada. Instalando agora...")
    install_package("vobject")
    import vobject

# Função para limpar os números de telefone
def clean_phone(phone):
    cleaned_phone = phone.replace('+55', '').replace('-', '').replace(' ', '')
    
    # Verifica se o primeiro dígito é zero e remove se necessário
    if cleaned_phone.startswith('0'):
        cleaned_phone = cleaned_phone[1:]
    
    if len(cleaned_phone) == 11:  # Verifica se é um número com 11 dígitos (incluindo o DDD)
        return f"{cleaned_phone}"
    elif len(cleaned_phone) == 12:  # Verifica se é um número com 10 dígitos (sem o DDD)
        return f"{cleaned_phone[1:]}"
    elif len(cleaned_phone) == 9:  # Verifica se consta o DDD
        return f"{cleaned_phone} - Telefone sem DDD"
    elif len(cleaned_phone) == 10:  # Verifica se tem todos os digitos
        return f"{cleaned_phone} - Telefone sem o 9 na frente"

    return cleaned_phone

# Função para selecionar um arquivo VCF usando a caixa de diálogo
def select_vcf_file():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal do tkinter

    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo VCF",
        filetypes=[("Arquivos VCF", "*.vcf"), ("Todos os arquivos", "*.*")]
    )

    return file_path

# Função chamada ao clicar no botão "Salvar" ou "Abrir arquivo"
def save_or_open_file(save_and_open=False):
    global vcf_file_path  # Adicionando global para a variável vcf_file_path
    client_name = client_name_entry.get()
    output_format = output_format_var.get()
    export_names_only = content_choice_var.get()

    output_folder = os.path.dirname(vcf_file_path)  # Usa o diretório do arquivo selecionado como destino
    os.makedirs(output_folder, exist_ok=True)
    output_file = f"contatos_{client_name}{output_format.lower()}"

    try:
        output_path = os.path.join(output_folder, output_file)
        with open(vcf_file_path, 'r', encoding='utf-8') as file:
            vcard_contents = file.read().split("BEGIN:VCARD")[1:]

        for vcard_content in vcard_contents:
            try:
                vcard = vobject.readOne("BEGIN:VCARD" + vcard_content)
                export_contacts(vcard, output_format, client_name, output_path, export_names_only)
            except vobject.base.ParseError as e:
                print(f"Ignorando linha com erro: {e}")

        # Remover zeros do início dos números de telefone
        with open(output_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        with open(output_path, 'w', encoding='utf-8') as file:
            for line in lines:
                file.write(clean_phone(line))

        output_path_label.config(text=f"Salvando em: {output_folder}")  # Atualizando o label de output_path
        success_message = f"O arquivo {output_file} foi salvo com sucesso na pasta de origem!"
        messagebox.showinfo("Sucesso", success_message)

        if save_and_open:
            os.startfile(output_path)
        else:
            select_another_file_button.config(state=tk.NORMAL)  # Habilita o botão "Nova conversão"
            # Exibir caixa de diálogo com opções "Nova conversão" e "Fechar"
            choice = messagebox.askquestion("Nova Conversão", "Arquivo salvo na pasta de origem\n \nDeseja realizar uma nova conversão?")
            if choice == "no":
                root.destroy()  # Fecha a aplicação
                clear_terminal()
    except Exception as e:
        error_message = f"Erro durante a leitura do arquivo VCF.\n O arquivo foi gerado, mas pode conter erros.\n"
        choice = messagebox.askquestion("Nova Conversão", "Arquivo salvo na pasta de origem\n \nDeseja realizar uma nova conversão?")
        if choice == "no":
            root.destroy()  # Fecha a aplicação
            clear_terminal()

# Função chamada ao clicar no botão "Nova conversão"
def new_conversion():
    global vcf_file_path
    vcf_file_path = select_vcf_file()
    if vcf_file_path:
        print(f"Novo arquivo selecionado: {vcf_file_path}")
        select_another_file_button.config(state=tk.DISABLED)  # Desabilita o botão "Nova conversão" até a próxima exportação
        output_path_label.config(text="Salvando em:")  # Reinicia o label de output_path
    else:
        print("Nenhum arquivo selecionado.")

# Função para exportar os contatos
def export_contacts(vcard, output_format, client_name, output_file, export_phones_only=False):
    name = vcard.n.value if vcard.n else ""
    phone = clean_phone(vcard.tel.value) if vcard.tel else ""

    with open(output_file, 'a', encoding='utf-8') as file:
        if export_phones_only:
            if phone:
                file.write(f'{phone}\n')
        else:
            file.write(f'{name},{phone}\n')

# Cria a janela principal do tkinter
root = tk.Tk()
root.title("Conversor VCF para TXT/CSV")

# Variáveis para armazenar as escolhas do usuário
output_format_var = tk.StringVar(value=".TXT")
content_choice_var = tk.StringVar(value="Somente Telefones")  # Inicializa com "Somente Telefones"
save_and_open_var = tk.BooleanVar()

# Botão para selecionar o arquivo VCF
select_file_button = tk.Button(root, text="Selecionar arquivo", command=new_conversion)
select_file_button.grid(row=0, column=0, columnspan=2, sticky="w")

# Seção para inserir o nome do cliente
tk.Label(root, text="Insira o nome do cliente:").grid(row=1, column=0, sticky="w")
client_name_entry = tk.Entry(root)
client_name_entry.grid(row=1, column=1, sticky="w")

# Seção para escolher o formato de saída
tk.Label(root, text="Formato de Saída:").grid(row=2, column=0, sticky="w")
output_format_menu = tk.OptionMenu(root, output_format_var, ".TXT", ".CSV")
output_format_menu.grid(row=2, column=1, sticky="w")

# Seção para escolher o conteúdo da exportação
tk.Label(root, text="Conteúdo da exportação:").grid(row=3, column=0, sticky="w")
content_choice_menu = tk.OptionMenu(root, content_choice_var, "Somente Telefones")
content_choice_menu.grid(row=3, column=1, sticky="w")

# Botão para Salvar
tk.Button(root, text="Salvar arquivo", command=lambda: save_or_open_file(False)).grid(row=4, column=0, columnspan=2, sticky="w")

# Inicia o loop principal do tkinter
root.mainloop()
