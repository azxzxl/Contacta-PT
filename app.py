import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import main_logic
from PIL import Image, ImageTk
import base64
import io


def abrir_arquivo():
    global caminho_do_template
    caminho_do_template = filedialog.askopenfilename(
        filetypes=[("Word documents", "*.docx")])
    if caminho_do_template:
        entry_caminho.delete(0, tk.END)
        entry_caminho.insert(0, caminho_do_template)


def preencher_tabela_controle():
    global dados_para_tabela_controle
    dados_para_tabela_controle = []
    num_linhas = int(entry_num_linhas_controle.get())
    for i in range(num_linhas):
        linha = simpledialog.askstring(
            "Preencher Tabela de Controle", f"Informe os dados para a linha {i+1} (separados por vírgula):")
        if linha:
            dados_para_tabela_controle.append(tuple(linha.split(',')))
    # Atualiza a entrada de texto com os dados da tabela de controle
    if dados_para_tabela_controle:
        entry_tabela_controle.delete("1.0", tk.END)
        for linha in dados_para_tabela_controle:
            entry_tabela_controle.insert(tk.END, ','.join(linha) + '\n')


def preencher_tabela_produtos():
    global dados_para_tabela_produtos
    dados_para_tabela_produtos = []
    num_linhas = int(entry_num_linhas_produtos.get())
    for i in range(num_linhas):
        linha = simpledialog.askstring(
            "Preencher Tabela de Produtos", f"Informe os dados para a linha {i+1} (separados por vírgula):")
        if linha:
            dados_para_tabela_produtos.append(tuple(linha.split(',')))
    # Atualiza a entrada de texto com os dados da tabela de produtos
    if dados_para_tabela_produtos:
        entry_tabela_produtos.delete("1.0", tk.END)
        for linha in dados_para_tabela_produtos:
            entry_tabela_produtos.insert(tk.END, ','.join(linha) + '\n')


def preencher_tabela_servicos():
    global dados_para_tabela_servicos
    dados_para_tabela_servicos = []
    num_linhas = int(entry_num_linhas_servicos.get())
    for i in range(num_linhas):
        linha = simpledialog.askstring(
            "Preencher Tabela de Serviços", f"Informe os dados para a linha {i+1} (separados por vírgula):")
        if linha:
            dados_para_tabela_servicos.append(tuple(linha.split(',')))
    # Atualiza a entrada de texto com os dados da tabela de serviços
    if dados_para_tabela_servicos:
        entry_tabela_servicos.delete("1.0", tk.END)
        for linha in dados_para_tabela_servicos:
            entry_tabela_servicos.insert(tk.END, ','.join(linha) + '\n')


def processar_documento():
    if caminho_do_template:
        try:
            # Solicita ao usuário as substituições
            substituicoes = {
                "[fabricante]": entry_fabricante.get(),
                "[solução]": entry_solucao.get(),
                "[documento]": entry_documento.get(),
                "[serviços]": entry_servicos.get(),
                "[Produto]": entry_produto.get(),
                # Converte para minúsculo
                "[produto]": entry_produto.get().lower(),
                "[Texto descritivo]": entry_descritivo.get()
            }

            # Chama a função do módulo main_logic com os argumentos corretos
            main_logic.substituir_texto_no_documento_e_preencher_tabelas(
                caminho_do_template,
                substituicoes,
                dados_para_tabela_controle,
                dados_para_tabela_produtos,
                dados_para_tabela_servicos
            )
            messagebox.showinfo(
                "Sucesso!", "Sua tabela foi gerada no mesmo local do arquivo de origem :)")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
    else:
        messagebox.showwarning("Aviso", "Por favor, selecione um arquivo.")


root = tk.Tk()
root.title("Automatização - Proposta tecnica :)")

# Criação do Notebook
notebook = ttk.Notebook(root)

# Frame para a primeira aba
frame_aba1 = tk.Frame(notebook)
label_caminho = tk.Label(frame_aba1, text="Caminho do Arquivo:")
label_caminho.grid(row=0, column=0, padx=5, pady=5)
entry_caminho = tk.Entry(frame_aba1, width=50)
entry_caminho.grid(row=0, column=1, padx=5, pady=5)
btn_abrir_arquivo = tk.Button(
    frame_aba1, text="Abrir Arquivo", command=abrir_arquivo)
btn_abrir_arquivo.grid(row=1, column=0, columnspan=2, pady=5)

# Imagem incorporada em base64
image_base64 = "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"

# Decodifica a imagem
image_bytes = base64.b64decode(image_base64)

# Cria uma imagem a partir dos bytes decodificados
image = Image.open(io.BytesIO(image_bytes))

# Redimensiona a imagem para o tamanho desejado (por exemplo, 200x200)
resized_image = image.resize((80, 80))

# Converte a imagem redimensionada em um objeto PhotoImage
logo_image = ImageTk.PhotoImage(resized_image)

# Crie um label para exibir a logo
label_logo = tk.Label(frame_aba1, image=logo_image)
label_logo.grid(row=3, column=0, columnspan=3, padx=2, pady=2, sticky="nsew")

# Definir uma fonte personalizada
fonte_instrucoes = ("Helvetica", 12, "bold")
fonte_detalhes = ("Helvetica", 10)

# Criar label para as instruções
instrucoes_label = tk.Label(
    frame_aba1, text="#AlwaysGetBetter", font=fonte_instrucoes)
instrucoes_label.grid(row=2, column=0, columnspan=2, pady=240)

# Criar label para as instruções detalhadas
instrucoes_detalhadas = tk.Label(
    frame_aba1, text="Contacta", justify="left", font=fonte_detalhes)
instrucoes_detalhadas.grid(row=4, column=0, columnspan=2, pady=10)

# Adiciona a primeira aba ao notebook
notebook.add(frame_aba1, text="Seleção de Arquivo")

# Frame para a segunda aba
frame_aba2 = tk.Frame(notebook)

# Campos para preenchimento das substituições
label_substituicoes = tk.Label(frame_aba2, text="Substituições:")
label_substituicoes.grid(row=0, column=0, columnspan=3, padx=5, pady=5)

label_fabricante = tk.Label(frame_aba2, text="Fabricante:")
label_fabricante.grid(row=1, column=0, padx=5, pady=5)
entry_fabricante = tk.Entry(frame_aba2)
entry_fabricante.grid(row=1, column=1, padx=5, pady=5)

label_solucao = tk.Label(frame_aba2, text="Solução:")
label_solucao.grid(row=2, column=0, padx=5, pady=5)
entry_solucao = tk.Entry(frame_aba2)
entry_solucao.grid(row=2, column=1, padx=5, pady=5)

label_documento = tk.Label(frame_aba2, text="Documento:")
label_documento.grid(row=3, column=0, padx=5, pady=5)
entry_documento = tk.Entry(frame_aba2)
entry_documento.grid(row=3, column=1, padx=5, pady=5)

label_servicos = tk.Label(frame_aba2, text="Serviços:")
label_servicos.grid(row=4, column=0, padx=5, pady=5)
entry_servicos = tk.Entry(frame_aba2)
entry_servicos.grid(row=4, column=1, padx=5, pady=5)

label_produto = tk.Label(frame_aba2, text="Produto:")
label_produto.grid(row=5, column=0, padx=5, pady=5)
entry_produto = tk.Entry(frame_aba2)
entry_produto.grid(row=5, column=1, padx=5, pady=5)

label_descritivo = tk.Label(frame_aba2, text="Texto Descritivo:")
label_descritivo.grid(row=6, column=0, padx=5, pady=5)
entry_descritivo = tk.Entry(frame_aba2)
entry_descritivo.grid(row=6, column=1, padx=5, pady=5)

# Campos para preenchimento das tabelas
label_controle = tk.Label(
    frame_aba2, text="Tabela de Controle (adicionar nº de linhas):")
label_controle.grid(row=7, column=0, padx=5, pady=5)
entry_num_linhas_controle = tk.Entry(frame_aba2, width=10)
entry_num_linhas_controle.grid(row=7, column=1, padx=5, pady=5)
btn_preencher_controle = tk.Button(
    frame_aba2, text="Preencher", command=preencher_tabela_controle)
btn_preencher_controle.grid(row=7, column=2, padx=5, pady=5)
entry_tabela_controle = tk.Text(
    frame_aba2, height=4, width=50)
entry_tabela_controle.grid(row=8, column=0, columnspan=3, padx=5, pady=5)

label_produtos = tk.Label(
    frame_aba2, text="Tabela de Produtos (adicionar nº de linhas):")
label_produtos.grid(row=9, column=0, padx=5, pady=5)
entry_num_linhas_produtos = tk.Entry(frame_aba2, width=10)
entry_num_linhas_produtos.grid(row=9, column=1, padx=5, pady=5)
btn_preencher_produtos = tk.Button(
    frame_aba2, text="Preencher", command=preencher_tabela_produtos)
btn_preencher_produtos.grid(row=9, column=2, padx=5, pady=5)
entry_tabela_produtos = tk.Text(
    frame_aba2, height=4, width=50)
entry_tabela_produtos.grid(row=10, column=0, columnspan=3, padx=5, pady=5)

label_servicos = tk.Label(
    frame_aba2, text="Tabela de Serviços (adicionar nº de linhas):")
label_servicos.grid(row=11, column=0, padx=5, pady=5)
entry_num_linhas_servicos = tk.Entry(frame_aba2, width=10)
entry_num_linhas_servicos.grid(row=11, column=1, padx=5, pady=5)
btn_preencher_servicos = tk.Button(
    frame_aba2, text="Preencher", command=preencher_tabela_servicos)
btn_preencher_servicos.grid(row=11, column=2, padx=5, pady=5)
entry_tabela_servicos = tk.Text(
    frame_aba2, height=4, width=50)
entry_tabela_servicos.grid(row=12, column=0, columnspan=3, padx=5, pady=5)

# Botão para processar o documento
btn_processar = tk.Button(
    frame_aba2, text="Clique aqui para gerar o novo documento", command=processar_documento)
btn_processar.grid(row=13, column=0, columnspan=3, pady=10)

# Adiciona a segunda aba ao notebook
notebook.add(frame_aba2, text="Edição de Documento")

# Layout do notebook
notebook.pack(expand=True, fill="both")

# Rodapé
rodape = tk.Label(root, text="v0.2 - Desenvolvido por Lucas R.",
                  bg="#6a5acd", fg="white")
rodape.pack(side="bottom", fill="x")

root.mainloop()
