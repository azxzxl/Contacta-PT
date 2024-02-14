# main_logic.py

from docx import Document


def substituir_texto_no_documento_e_preencher_tabelas(caminho_do_template, substituicoes, dados_para_tabela_controle, dados_para_tabela_produtos, dados_para_tabela_servicos):
    doc = Document(caminho_do_template)

    # Substituição de texto nos parágrafos
    for p in doc.paragraphs:
        for key, value in substituicoes.items():
            if key in p.text:
                inline = p.runs
                for i in inline:
                    if key in i.text:
                        i.text = i.text.replace(key, value)

    # Preenchimento da tabela de controle de documento
    tabela_controle = doc.tables[0]  # A primeira tabela é a de controle
    # Inicia no 1 para pular o cabeçalho
    for i, row_data in enumerate(dados_para_tabela_controle, start=1):
        for j, cell_data in enumerate(row_data):
            tabela_controle.cell(i, j).text = str(cell_data)

    # Preenchimento da tabela de produtos
    tabela_produtos = doc.tables[1]  # A segunda tabela é a de produtos
    # Inicia no 1 para pular o cabeçalho
    for i, row_data in enumerate(dados_para_tabela_produtos, start=1):
        for j, cell_data in enumerate(row_data):
            # Insere os dados apenas se a célula estiver vazia
            if tabela_produtos.cell(i, j).text.strip() == '':
                tabela_produtos.cell(i, j).text = str(cell_data)

    # Preenchimento da tabela de serviços
    tabela_servicos = doc.tables[2]  # A terceira tabela é a de serviços
    # Inicia no 1 para pular o cabeçalho
    for i, row_data in enumerate(dados_para_tabela_servicos, start=1):
        for j, cell_data in enumerate(row_data):
            tabela_servicos.cell(i, j).text = str(cell_data)

    doc.save(caminho_do_template.replace('.docx', '_modificado.docx'))
