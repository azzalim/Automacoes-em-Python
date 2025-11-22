from pathlib import Path
import pandas as pd
import sys

def empilhar_planilhas(base_dir: Path, subpastas=("Ativos", "Excluídos"), arquivo_saida="empilhado.xlsx"):
    arquivos = []
    for sub in subpastas:
        pasta = base_dir / sub
        if pasta.exists() and pasta.is_dir():
            # procura recursivamente por .xls
            arquivos.extend(sorted(pasta.rglob("*.xls")))
        else:
            print(f"Atenção: pasta não encontrada ou não é diretório: {pasta}")

    if not arquivos:
        print("Nenhum arquivo .xls encontrado nas pastas especificadas. Finalizando script.")
        return

    dfs = []
    for caminho in arquivos:
        try:
            print(f"Lendo: {caminho}")
            # usa engine "xlrd" para arquivos .xls (Excel 97-2003)
            df = pd.read_excel(caminho, engine="xlrd", header=0)
        except Exception as e:
            print(f"Erro ao ler {caminho}: {e}. Pulando arquivo.")
            continue

        if not df.empty:
            primeira_linha = df.iloc[0].astype(str).tolist()
            nomes_colunas = df.columns.astype(str).tolist()
            if primeira_linha == nomes_colunas:
                df = df.iloc[1:].reset_index(drop=True)

        dfs.append(df)

    if not dfs:
        print("Nenhum DataFrame válido para concatenar. Saindo.")
        return

    df_final = pd.concat(dfs, ignore_index=True, sort=False)

    caminho_saida = base_dir / arquivo_saida
    try:
        df_final.to_excel(caminho_saida, index=False, engine="openpyxl")
        print(f"\nArquivo final salvo em: {caminho_saida}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo final: {e}")

if __name__ == "__main__":
    try:
        base = Path(__file__).parent.resolve()
    except NameError:
        base = Path.cwd().resolve()

    if len(sys.argv) > 1:
        base = Path(sys.argv[1]).resolve()

    empilhar_planilhas(base)
