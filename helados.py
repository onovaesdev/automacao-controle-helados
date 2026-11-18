import shutil
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.protection import SheetProtection

ARQUIVO_MODELO = "Controle de Helados - Modelo Padrao.xlsx"        
ARQUIVO_LOJAS  = "exemplo_lista_de_lojas.xlsx"
PASTA_SAIDA    = "exemplo_mes_do_controle"  

COL_NOME = "Nome da loja"   
COL_COD  = "Codigo Linx"    

def ler_lista_lojas(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(caminho, sep=None, engine="python", dtype=str)
    else:
        df = pd.read_excel(caminho, dtype=str)
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=[COL_NOME, COL_COD])
    df[COL_NOME] = df[COL_NOME].str.strip()
    df[COL_COD]  = df[COL_COD].str.strip()
    return df


def nome_arquivo_seguro(nome):
    for c in r'\/:*?"<>|':
        nome = nome.replace(c, "-")
    return nome.strip()


def gerar_planilha(linha, pasta_saida):
    nome_loja = linha[COL_NOME]
    codigo    = int(linha[COL_COD])

    # Copia o modelo
    nome_arquivo = nome_arquivo_seguro("Controle de Helados - " + nome_loja) + ".xlsx"
    destino = os.path.join(pasta_saida, nome_arquivo)
    shutil.copy2(ARQUIVO_MODELO, destino)

    wb = load_workbook(destino)

    # Aba '01': nome da loja em D4
    aba01 = wb["01"]
    aba01["D4"] = nome_loja
    aba01["D4"].number_format = "@"

    # Aba 'CONSOLIDADO': código da loja - A7:A22
    consolidado = wb["CONSOLIDADO"]

    # Desprotege temporariamente
    consolidado.protection = SheetProtection()
    consolidado.protection.sheet = False

    for linha_num in range(7, 24):  # A7 até A23 (inclusive)
        consolidado.cell(row=linha_num, column=1).value = codigo

    # Reprotege (sem senha, igual ao original)
    consolidado.protection.sheet = True

    wb.save(destino)
    return nome_arquivo


def main():
    for arq in [ARQUIVO_MODELO, ARQUIVO_LOJAS]:
        if not os.path.exists(arq):
            print(f"[ERRO] Arquivo não encontrado: {arq}")
            return

    os.makedirs(PASTA_SAIDA, exist_ok=True)

    df = ler_lista_lojas(ARQUIVO_LOJAS)
    total = len(df)
    print(f"\n{total} lojas encontradas. Gerando planilhas...\n")

    erros = []
    for i, (_, linha) in enumerate(df.iterrows(), 1):
        try:
            arquivo = gerar_planilha(linha, PASTA_SAIDA)
            print(f"  [{i:02d}/{total}] ✓ {arquivo}")
        except Exception as e:
            nome = linha.get(COL_NOME, f"linha {i}")
            print(f"  [{i:02d}/{total}] ✗ ERRO em '{nome}': {e}")
            erros.append((nome, str(e)))

    print(f"\n{'='*50}")
    print(f"Concluído! {total - len(erros)}/{total} planilhas geradas em: {PASTA_SAIDA}/")
    if erros:
        print(f"\nErros ({len(erros)}):")
        for nome, erro in erros:
            print(f"  - {nome}: {erro}")

if __name__ == "__main__":
    main()