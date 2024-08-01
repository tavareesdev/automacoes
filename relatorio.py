import pandas as pd
from datetime import datetime

# Lê a planilha existente para preservar a situação dos alunos
try:
    df_existente = pd.read_excel('C:\\Users\\gtava\\OneDrive\\Documentos\\Relatórios\\Relatório Coordenação Junho.xlsx')
except FileNotFoundError:
    df_existente = pd.DataFrame(columns=['Aluno', 'Data do Último Acesso', 'Data de Primeiro Acesso', 'Dias desde primeiro acesso', 'Dias desde o último acesso', 'Situação'])

# Lê a nova planilha de agendamentos
df = pd.read_excel('C:\\Users\\gtava\\Downloads\\2251-AgendamentosAluno-2c3714067b7241b7bdb2040f5a963256.xlsx', header=1, names=['Data', 'Hora Início', 'Hora Fim', 'Aluno', 'Telefone', 'Situacao', 'Agendamento'])

# Converte a coluna de datas para o formato datetime
df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')

# Encontra a data mais recente de acesso de cada aluno
ultimas_datas_acesso = df.groupby('Aluno')['Data'].max().reset_index()

# Calcula os dias desde o último acesso e renomeia a coluna
ultimas_datas_acesso['Dias desde o último acesso'] = (datetime.now() - ultimas_datas_acesso['Data']).dt.days

# Define a situação com base nos dias desde o último acesso
ultimas_datas_acesso['Situação'] = 'ATIVO'
ultimas_datas_acesso.loc[ultimas_datas_acesso['Dias desde o último acesso'] >= 21, 'Situação'] = 'DESAPARECIDO'

# Formata a data mais recente para string e renomeia a coluna
ultimas_datas_acesso['Data do Último Acesso'] = ultimas_datas_acesso['Data'].dt.strftime('%d/%m/%Y')

# Encontra a data mais antiga de acesso de cada aluno
primeiras_datas_acesso = df.groupby('Aluno')['Data'].min().reset_index()

# Calcula a diferença em dias da data mais antiga para o dia de hoje
primeiras_datas_acesso['Dias desde primeiro acesso'] = (datetime.now() - primeiras_datas_acesso['Data']).dt.days

# Formata a data mais antiga para string e renomeia a coluna
primeiras_datas_acesso['Data de Primeiro Acesso'] = primeiras_datas_acesso['Data'].dt.strftime('%d/%m/%Y')

# Calcula a diferença em meses da data mais antiga para o dia de hoje
primeiras_datas_acesso['Meses desde primeiro acesso'] = primeiras_datas_acesso['Data'].apply(lambda x: (datetime.now().year - x.year) * 12 + datetime.now().month - x.month)

# Junta os dados de última e primeira data de acesso
datas_e_situacao = pd.merge(ultimas_datas_acesso, primeiras_datas_acesso[['Aluno', 'Data de Primeiro Acesso', 'Dias desde primeiro acesso', 'Meses desde primeiro acesso']], on='Aluno', how='left')

# Integra as informações da planilha existente
datas_e_situacao = pd.merge(datas_e_situacao, df_existente[['Aluno', 'Situação']], on='Aluno', how='left', suffixes=('', '_existente'))

# Mantém a situação 'CANCELADO', 'BLOQUEADO' ou 'FORMADO' da planilha existente
datas_e_situacao['Situação'] = datas_e_situacao.apply(
    lambda row: row['Situação_existente'] if row['Situação_existente'] in ['CANCELADO', 'BLOQUEADO', 'FORMADO'] else row['Situação'],
    axis=1
)

# Remove a coluna auxiliar 'Situação_existente'
datas_e_situacao = datas_e_situacao.drop(columns=['Situação_existente'])

# Adiciona os alunos com situação 'CANCELADO', 'BLOQUEADO' ou 'FORMADO' que não estão na nova planilha
alunos_especiais = df_existente[df_existente['Situação'].isin(['CANCELADO', 'BLOQUEADO', 'FORMADO'])]
datas_e_situacao = pd.concat([datas_e_situacao, alunos_especiais], ignore_index=True).drop_duplicates(subset=['Aluno'], keep='last')

# Reordena as colunas
datas_e_situacao = datas_e_situacao[['Aluno', 'Data do Último Acesso', 'Data de Primeiro Acesso', 'Dias desde primeiro acesso', 'Meses desde primeiro acesso', 'Dias desde o último acesso', 'Situação']]

# Salva o resultado no arquivo Excel
datas_e_situacao.to_excel('C:\\Users\\gtava\\OneDrive\\Documentos\\Relatórios\\Relatório Coordenação TESTE.xlsx', index=False)
