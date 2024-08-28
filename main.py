import pandas as pd

# Carrega o arquivo Excel
file_path = 'Carteira_CELI_LEGACY_FIA_IE_26_08_2024.xlsx'  # Substitua pelo caminho do seu arquivo
sheet_name = 'Planilha1'  # Substitua pelo nome da sua planilha

def Automatizacao(file_path,sheet_name):

    # Lê a planilha em um DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Adiciona uma nova coluna categorizando como "Ações" ou "Renda Fixa"
    df['Categoria'] = df['Unnamed: 1'].apply(lambda x: 'Ações' if 'FIA' in x or 'FICFIA' in x or 'FCFA' in x else 'Renda Fixa')

    # Separa os fundos em dois DataFrames diferentes
    acoes_df = df[df['Categoria'] == 'Ações']
    renda_fixa_df = df[df['Categoria'] == 'Renda Fixa']

    # Cria um novo DataFrame com os títulos e os dados organizados
    final_df = pd.DataFrame()

    # Adiciona a seção de Fundos de Ações
    final_df = pd.concat([final_df, pd.DataFrame({'Unnamed: 1': ['Fundos de Ações']})], ignore_index=True)
    final_df = pd.concat([final_df, acoes_df], ignore_index=True)

    # Adiciona a seção de Fundos de Renda Fixa
    final_df = pd.concat([final_df, pd.DataFrame({'Unnamed: 1': ['Fundos de Renda Fixa']})], ignore_index=True)
    final_df = pd.concat([final_df, renda_fixa_df], ignore_index=True)

    # Salva o DataFrame final em um novo arquivo Excel
    output_file = r'C:\Users\gabriel.martins\Documents\fundos_organizados.xlsx'
    final_df.to_excel(output_file, index=False, sheet_name='Fundos Organizados')

    print(f"Arquivo salvo como {output_file}")
