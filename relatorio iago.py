import pandas as pd
from datetime import datetime, timedelta

df = pd.read_excel('C:\\Users\\Ped\\Downloads\\2251-AgendamentosAluno-28781c48766c45d8a926735667082f83.xlsx', header=1, names=['Data', 'Hora Início', 'Hora Fim', 'Aluno', 'Telefone', 'Situacao', 'Agendamento'])

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

# Junta os dados de última e primeira data de acesso
datas_e_situacao = pd.merge(ultimas_datas_acesso, primeiras_datas_acesso[['Aluno', 'Data de Primeiro Acesso', 'Dias desde primeiro acesso']], on='Aluno', how='left')

# Filtra os alunos que começaram há menos de 30 dias
datas_e_situacao = datas_e_situacao[datas_e_situacao['Dias desde primeiro acesso'] < 31]

# Reordena as colunas e remove a coluna 'Data'
datas_e_situacao = datas_e_situacao[['Aluno', 'Data do Último Acesso', 'Data de Primeiro Acesso', 'Dias desde primeiro acesso', 'Dias desde o último acesso', 'Situação']]

datas_e_situacao.to_excel('C:\\Users\\Ped\\\Documents\\Relatórios\\Relatório Novos Alunos.xlsx', index=False)
