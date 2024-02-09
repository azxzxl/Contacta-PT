import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import main_logic


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
                "Sucesso", "O documento foi processado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
    else:
        messagebox.showwarning("Aviso", "Por favor, selecione um arquivo.")


root = tk.Tk()
root.title("Processador de Documentos")

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

# Adiciona a primeira aba ao notebook
notebook.add(frame_aba1, text="Seleção de Arquivo")

# Frame para a segunda aba
frame_aba2 = tk.Frame(notebook)

# Campos para preenchimento das substituições
label_substituicoes = tk.Label(frame_aba2, text="Substituições:")
label_substituicoes.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

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
label_controle = tk.Label(frame_aba2, text="Tabela de Controle:")
label_controle.grid(row=7, column=0, padx=5, pady=5)
entry_num_linhas_controle = tk.Entry(frame_aba2, width=10)
entry_num_linhas_controle.grid(row=7, column=1, padx=5, pady=5)
btn_preencher_controle = tk.Button(
    frame_aba2, text="Preencher Tabela", command=preencher_tabela_controle)
btn_preencher_controle.grid(row=7, column=2, padx=5, pady=5)
entry_tabela_controle = tk.Text(
    frame_aba2, height=4, width=50)
entry_tabela_controle.grid(row=8, column=0, columnspan=3, padx=5, pady=5)

label_produtos = tk.Label(frame_aba2, text="Tabela de Produtos:")
label_produtos.grid(row=9, column=0, padx=5, pady=5)
entry_num_linhas_produtos = tk.Entry(frame_aba2, width=10)
entry_num_linhas_produtos.grid(row=9, column=1, padx=5, pady=5)
btn_preencher_produtos = tk.Button(
    frame_aba2, text="Preencher Tabela", command=preencher_tabela_produtos)
btn_preencher_produtos.grid(row=9, column=2, padx=5, pady=5)
entry_tabela_produtos = tk.Text(
    frame_aba2, height=4, width=50)
entry_tabela_produtos.grid(row=10, column=0, columnspan=3, padx=5, pady=5)

label_servicos = tk.Label(frame_aba2, text="Tabela de Serviços:")
label_servicos.grid(row=11, column=0, padx=5, pady=5)
entry_num_linhas_servicos = tk.Entry(frame_aba2, width=10)
entry_num_linhas_servicos.grid(row=11, column=1, padx=5, pady=5)
btn_preencher_servicos = tk.Button(
    frame_aba2, text="Preencher Tabela", command=preencher_tabela_servicos)
btn_preencher_servicos.grid(row=11, column=2, padx=5, pady=5)
entry_tabela_servicos = tk.Text(
    frame_aba2, height=4, width=50)
entry_tabela_servicos.grid(row=12, column=0, columnspan=3, padx=5, pady=5)

# Botão para processar o documento
btn_processar = tk.Button(
    frame_aba2, text="Processar Documento", command=processar_documento)
btn_processar.grid(row=13, column=0, columnspan=3, pady=10)

# Adiciona a segunda aba ao notebook
notebook.add(frame_aba2, text="Edição de Documento")

# Layout do notebook
notebook.pack(expand=True, fill="both")

# Rodapé
rodape = tk.Label(root, text="Versão 0.1 - Desenvolvido por Lucas R.")
rodape.pack(side="bottom", pady=5)

root.mainloop()
