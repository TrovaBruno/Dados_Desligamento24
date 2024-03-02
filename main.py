import re
import openpyxl

# Função para extrair os dados do texto usando expressões regulares
def extrair_dados(texto):
    padrao_nome = r"Nome Completo:\s*(.+)"
    padrao_gestor = r"Gestor:\s*(.+)"
    padrao_telefone = r"Celular:\s*(\d+-\d+)"
    padrao_data_desligamento = r"Data do desligamento:\s*(\d+/\d+/\d+)"
    padrao_area = r"Diretoria \/ Area:\s*(.+)"

    nome = re.search(padrao_nome, texto).group(1)
    gestor = re.search(padrao_gestor, texto).group(1)
    telefone = re.search(padrao_telefone, texto).group(1)
    data_desligamento = re.search(padrao_data_desligamento, texto).group(1)
    area = re.search(padrao_area, texto).group(1)

    return nome, gestor, telefone, data_desligamento, area

# Função para criar ou abrir uma planilha Excel e inserir os dados
def inserir_dados_excel(dados):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Adicionando cabeçalhos
    sheet.append(["Nome", "Gestor", "Telefone", "Data do Desligamento", "Área"])

    # Inserindo os dados
    sheet.append(dados)

    # Salvando a planilha
    workbook.save("dados_desligamento.xlsx")
    print("Dados inseridos com sucesso no arquivo 'dados_desligamento.xlsx'.")

# Texto fornecido
texto_comunicado = """
Olá,

Comunicamos o desligamento de ISIS CRISTINA DA SILVA ALVES, conforme dados detalhados abaixo. Por favor, prossigam com o fluxo interno para:

• Bloqueio dos acessos;
• Agendamento de exame demissional;
• Confirmação Auxílio Educação/Idiomas para desconto em folha;
• Confirmação de mobiliários para serem devolvidos;
• Agendamento com a transportadora para retirada de equipamentos para os casos de trabalho remoto ( e enquanto durar a pandemia todos estão home office) ou agendamento para devolução dos equipamentos para os que residem em São Paulo.


Nome Completo: ISIS CRISTINA DA SILVA ALVES



Empresa: LOCAWEB SERVICOS DE INTERNET S/A

Diretoria / Area: ALL IN CUSTOMER SUCCESS


Gestor: FLAVIO DOS SANTOS DE NIJS

Data do desligamento: 20/02/2024

CPF: 43459800895-


Endereço: Rua Francisco Pessoa , 690

Celular: 1198947-7479

Equipamentos: Deverão ser retirados na residência
"""

# Extraindo os dados do texto
dados = extrair_dados(texto_comunicado)

# Inserindo os dados na planilha Excel
inserir_dados_excel(dados)
