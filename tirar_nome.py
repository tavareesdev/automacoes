import pandas as pd

# Lista de nomes dos alunos a serem procurados
nomes_alunos = [
    "Vitória Batista de Souza"
]

# Lista de arquivos Excel para verificar
arquivos_excel = ["mensagens_segunda - Copia.xlsx", "mensagens_terca - Copia.xlsx", "mensagens_quarta - Copia.xlsx", "mensagens_quinta - Copia.xlsx", "mensagens_sexta - Copia.xlsx", "mensagens_sabado - Copia.xlsx"]

def remover_nomes(nomes, arquivo):
    # Carrega todas as planilhas do arquivo Excel em um dicionário de DataFrames
    xls = pd.ExcelFile(arquivo)
    planilhas = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    
    # Verifica e remove os nomes dos alunos de cada planilha
    alterado = False
    for sheet, df in planilhas.items():
        for nome in nomes:
            # Cria uma máscara para identificar linhas que contêm o nome do aluno
            mask = df.apply(lambda row: row.astype(str).str.contains(nome, regex=True).any(), axis=1)
            if mask.any():
                # Remove as linhas onde o nome aparece
                df = df[~mask]
                planilhas[sheet] = df  # Atualiza o DataFrame no dicionário
                alterado = True
    
    # Se alguma planilha foi alterada, salva o arquivo atualizado
    if alterado:
        with pd.ExcelWriter(arquivo) as writer:
            for sheet, df in planilhas.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        print(f"Os nomes {nomes} foram removidos do arquivo '{arquivo}'.")

for arquivo in arquivos_excel:
    remover_nomes(nomes_alunos, arquivo)
