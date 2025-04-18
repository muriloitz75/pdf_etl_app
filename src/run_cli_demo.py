import os
import sys

# Adicionar o diretório atual ao caminho de busca do Python
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Importar os módulos do projeto
from extraction import extract_tables_from_pdf
from transformation import clean_dataframe
from loading import load_to_excel
from loading_csv import load_to_csv

def run_demo():
    """Executa uma demonstração do pipeline ETL em modo linha de comando"""
    # Definir caminhos
    pdf_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
    excel_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\demo_excel.xlsx"
    csv_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\demo_csv.csv"

    print("=" * 80)
    print("DEMONSTRAÇÃO DO PIPELINE ETL")
    print("=" * 80)
    print(f"Arquivo PDF: {pdf_path}")
    print(f"Saída Excel: {excel_path}")
    print(f"Saída CSV: {csv_path}")
    print("=" * 80)

    try:
        # Extração
        print("\n1. EXTRAÇÃO")
        print("-" * 40)
        print("Extraindo tabelas do PDF...")
        tables = extract_tables_from_pdf(pdf_path, method='auto')
        print(f"Extraídas {len(tables)} tabelas")

        if not tables:
            print("ERRO: Nenhuma tabela encontrada no PDF")
            return False

        # Mostrar exemplo da primeira tabela
        print("\nExemplo da primeira tabela (primeiras 5 linhas):")
        print(tables[0].head())

        # Transformação
        print("\n2. TRANSFORMAÇÃO")
        print("-" * 40)
        print("Limpando e transformando dados...")

        # Opções de transformação
        transform_options = {
            'remove_empty_rows': True,
            'remove_empty_cols': True,
            'convert_dates': False,  # Desativar conversão de datas para evitar conflitos
            'convert_money': True
        }

        cleaned_tables = [clean_dataframe(df, **transform_options) for df in tables]
        print(f"Transformadas {len(cleaned_tables)} tabelas")

        # Mostrar exemplo da primeira tabela limpa
        print("\nExemplo da primeira tabela limpa (primeiras 5 linhas):")
        print(cleaned_tables[0].head())

        # Combinar todas as tabelas em uma única
        print("\nCombinando todas as tabelas...")
        import pandas as pd
        combined_df = pd.concat(cleaned_tables, ignore_index=True)
        print(f"Tabela combinada: {len(combined_df)} linhas, {len(combined_df.columns)} colunas")

        # Formatar datas corretamente antes de exportar
        print("\nFormatando datas...")
        # Extrair as datas originais do PDF
        import re
        for i, table in enumerate(tables):
            for col in table.columns:
                col_lower = str(col).lower()
                if any(date_term in col_lower for date_term in ['data', 'dt', 'date', 'emissão', 'emissao']):
                    # Copiar as datas originais para o DataFrame limpo
                    date_values = []
                    for idx, row in table.iterrows():
                        cell_value = str(row[col])
                        match = re.search(r'(\d{2}/\d{2}/\d{4})', cell_value)
                        if match:
                            date_values.append(match.group(1))
                        else:
                            date_values.append('')

                    # Atualizar as datas no DataFrame limpo
                    if len(date_values) == len(cleaned_tables[i]):
                        cleaned_tables[i][col] = date_values

        # Recombinar as tabelas com as datas corrigidas
        combined_df = pd.concat(cleaned_tables, ignore_index=True)

        # Carregamento para Excel
        print("\n3. CARREGAMENTO PARA EXCEL")
        print("-" * 40)
        print("Salvando dados no Excel...")
        excel_file = load_to_excel(combined_df, excel_path)
        print(f"Dados salvos em: {excel_file}")

        # Carregamento para CSV
        print("\n4. CARREGAMENTO PARA CSV")
        print("-" * 40)
        print("Salvando dados no CSV...")
        csv_file = load_to_csv(combined_df, csv_path, include_header=True)
        print(f"Dados salvos em: {csv_file}")

        # Verificar se os arquivos foram criados
        if os.path.exists(excel_path) and os.path.exists(csv_path):
            print("\nDEMONSTRAÇÃO CONCLUÍDA COM SUCESSO!")
            print("=" * 80)
            print(f"Arquivo Excel: {excel_path}")
            print(f"Tamanho: {os.path.getsize(excel_path)} bytes")
            print(f"Arquivo CSV: {csv_path}")
            print(f"Tamanho: {os.path.getsize(csv_path)} bytes")
            print("=" * 80)
            return True
        else:
            print("ERRO: Arquivos não foram criados corretamente")
            return False

    except Exception as e:
        print(f"\nERRO durante a demonstração: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    run_demo()
