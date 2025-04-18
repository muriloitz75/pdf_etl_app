import os
import sys
import pandas as pd

# Adicionar o diretório src ao caminho de busca do Python
src_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src')
sys.path.append(src_path)

# Importar os módulos do projeto
from extraction import extract_tables_from_pdf
from transformation import clean_dataframe
from loading_csv import load_to_csv

def test_csv_export():
    """Testa a exportação para CSV"""
    # Definir caminhos
    pdf_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
    csv_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\saida_test.csv"

    print(f"Testando exportação para CSV com arquivo: {pdf_path}")
    print(f"Saída será salva em: {csv_path}")

    try:
        # Extração
        print("\n1. EXTRAÇÃO")
        print("Extraindo tabelas do PDF...")
        tables = extract_tables_from_pdf(pdf_path)
        print(f"Extraídas {len(tables)} tabelas")

        if not tables:
            print("ERRO: Nenhuma tabela encontrada no PDF")
            return False

        # Transformação
        print("\n2. TRANSFORMAÇÃO")
        print("Limpando e transformando dados...")
        cleaned_tables = [clean_dataframe(df) for df in tables]
        print(f"Transformadas {len(cleaned_tables)} tabelas")

        # Carregamento
        print("\n3. CARREGAMENTO")
        print("Salvando dados no CSV...")

        # Testar com cabeçalho
        csv_with_header = csv_path.replace('.csv', '_with_header.csv')
        load_to_csv(cleaned_tables[0], csv_with_header, include_header=True)
        print(f"Dados salvos com cabeçalho em: {csv_with_header}")

        # Testar sem cabeçalho
        csv_without_header = csv_path.replace('.csv', '_without_header.csv')
        load_to_csv(cleaned_tables[0], csv_without_header, include_header=False)
        print(f"Dados salvos sem cabeçalho em: {csv_without_header}")

        # Verificar se os arquivos CSV foram criados
        if os.path.exists(csv_with_header) and os.path.exists(csv_without_header):
            print("\nTESTE CONCLUÍDO COM SUCESSO!")
            print(f"Arquivo CSV com cabeçalho: {csv_with_header}")
            print(f"Tamanho: {os.path.getsize(csv_with_header)} bytes")
            print(f"Arquivo CSV sem cabeçalho: {csv_without_header}")
            print(f"Tamanho: {os.path.getsize(csv_without_header)} bytes")
            return True
        else:
            print("ERRO: Arquivos CSV não foram criados")
            return False

    except Exception as e:
        print(f"\nERRO durante o teste: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_csv_export()
