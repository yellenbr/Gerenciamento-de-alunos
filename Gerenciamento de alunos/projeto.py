import pandas as pd

# Define o caminho do arquivo Excel
FILE_PATH = "alunos.xlsx"

# Função para carregar dados do Excel para um DataFrame
def load_data():
    try:
        return pd.read_excel(FILE_PATH)
    except FileNotFoundError:
        # Se o arquivo não existir, cria um DataFrame vazio
        return pd.DataFrame(columns=["Nome", "Idade", "Curso"])

# Função para salvar o DataFrame no Excel
def save_data(df):
    df.to_excel(FILE_PATH, index=False)

# Função para cadastrar um novo aluno
def cadastrar_aluno(nome, idade, curso):
    df = load_data()
    novo_aluno = pd.DataFrame({"Nome": [nome], "Idade": [idade], "Curso": [curso]})
    df = pd.concat([df, novo_aluno], ignore_index=True)
    save_data(df)
    print(f"Aluno {nome} cadastrado com sucesso.")

# Função para listar todos os alunos
def listar_alunos():
    df = load_data()
    if df.empty:
        print("Nenhum aluno cadastrado.")
    else:
        print(df)

# Função para atualizar dados de um aluno
def atualizar_aluno(nome, novo_nome=None, nova_idade=None, novo_curso=None):
    df = load_data()
    if nome in df["Nome"].values:
        # Atualiza os dados conforme informado
        if novo_nome:
            df.loc[df["Nome"] == nome, "Nome"] = novo_nome
        if nova_idade:
            df.loc[df["Nome"] == nome, "Idade"] = nova_idade
        if novo_curso:
            df.loc[df["Nome"] == nome, "Curso"] = novo_curso
        save_data(df)
        print(f"Dados do aluno {nome} atualizados com sucesso.")
    else:
        print(f"Aluno {nome} não encontrado.")

# Função para excluir um aluno
def excluir_aluno(nome):
    df = load_data()
    if nome in df["Nome"].values:
        df = df[df["Nome"] != nome]
        save_data(df)
        print(f"Aluno {nome} excluído com sucesso.")
    else:
        print(f"Aluno {nome} não encontrado.")

# Função principal para interação com o usuário
def menu():
    while True:
        print("\nSistema de Gerenciamento de Alunos")
        print("1. Cadastrar Aluno")
        print("2. Listar Alunos")
        print("3. Atualizar Aluno")
        print("4. Excluir Aluno")
        print("5. Sair")
        
        opcao = input("Escolha uma opção: ")
        
        if opcao == "1":
            nome = input("Nome do Aluno: ")
            idade = int(input("Idade do Aluno: "))
            curso = input("Curso do Aluno: ")
            cadastrar_aluno(nome, idade, curso)
        
        elif opcao == "2":
            listar_alunos()
        
        elif opcao == "3":
            nome = input("Nome do Aluno a atualizar: ")
            novo_nome = input("Novo Nome (ou pressione Enter para não alterar): ")
            nova_idade = input("Nova Idade (ou pressione Enter para não alterar): ")
            novo_curso = input("Novo Curso (ou pressione Enter para não alterar): ")
            
            # Converte entrada de idade, se fornecida
            nova_idade = int(nova_idade) if nova_idade else None
            
            atualizar_aluno(nome, novo_nome or None, nova_idade, novo_curso or None)
        
        elif opcao == "4":
            nome = input("Nome do Aluno a excluir: ")
            excluir_aluno(nome)
        
        elif opcao == "5":
            print("Encerrando o sistema.")
            break
        else:
            print("Opção inválida. Tente novamente.")

# Executa o menu
menu()
