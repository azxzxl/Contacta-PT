import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import main_logic


def abrir_arquivo():
    global caminho_do_template
    caminho_do_template = filedialog.askopenfilename(
        filetypes=[("Word documents", "*.docx")])
    if caminho_do_template:
        entry_caminho.delete(0, tk.END)
        entry_caminho.insert(0, caminho_do_template)


def processar_documento():
    if caminho_do_template:
        try:
            # Solicita ao usuário as substituições
            substituicoes = {
                "[fabricante]": simpledialog.askstring("Input", "Fabricante desejado:", parent=root),
                "[solução]": simpledialog.askstring("Input", "Solução desejada:", parent=root),
                "[documento]": simpledialog.askstring("Input", "Documento desejado:", parent=root),
                "[serviços]": simpledialog.askstring("Input", "Serviços desejados:", parent=root),
                "[Produto] [produto] ": simpledialog.askstring("Input", "Produto desejado:", parent=root),
                "[Texto descritivo]": simpledialog.askstring("Input", "Texto descritivo desejado:", parent=root)
            }

            # Solicita ao usuário os dados para preencher a tabela de controle
            dados_para_tabela_controle = []
            num_linhas = simpledialog.askinteger(
                "Input", "Qntas linhas para a tabela de controle? ", parent=root)
            for _ in range(num_linhas):
                linha = simpledialog.askstring(
                    "TABELA DE CONTROLE", "Insira por ordem, separados por vírgula e sem espaço EX: versao,data,comentarios", parent=root)
                dados_para_tabela_controle.append(tuple(linha.split(',')))

            # Solicita ao usuário os dados para preencher a tabela de produtos
            dados_para_tabela_produtos = []
            num_linhas = simpledialog.askinteger(
                "Input", "Quantas linhas para a tabela de produtos?", parent=root)
            for _ in range(num_linhas):
                linha = simpledialog.askstring(
                    "TABELA DE PRODUTOS", "Insira os dados da linha para a tabela de produtos (separados por vírgula e sem espaço) e sempre contendo o N da linha antes. Ex: 1,partnumber,descriçao,qntd:", parent=root)
                dados_para_tabela_produtos.append(tuple(linha.split(',')))

            # Solicita ao usuário os dados para preencher a tabela de serviços
            dados_para_tabela_servicos = []
            num_linhas = simpledialog.askinteger(
                "Input", "Quantas linhas para a tabela de serviços?", parent=root)
            for _ in range(num_linhas):
                linha = simpledialog.askstring(
                    "TABELA DE SERVIÇO", "Insira os dados da linha para a tabela de serviços (separados por vírgula e sem espaço) e sempre contendo o N da linha antes. Ex: 1,descriçao,partnumber,qntd:", parent=root)
                dados_para_tabela_servicos.append(tuple(linha.split(',')))

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

entry_caminho = tk.Entry(root, width=50)
entry_caminho.pack(pady=5)

btn_abrir_arquivo = tk.Button(
    root, text="Abrir Arquivo", command=abrir_arquivo)
btn_abrir_arquivo.pack(pady=5)

btn_processar = tk.Button(
    root, text="Processar Documento", command=processar_documento)
btn_processar.pack(pady=5)

root.mainloop()
