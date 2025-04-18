import os
import sys
import traceback

# Adicionar o diretório atual ao caminho de busca do Python
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Importar os módulos do projeto
from extraction import extract_tables_from_pdf
from transformation import clean_dataframe
from loading import load_to_excel
from loading_csv import load_to_csv

def run_etl_cli(pdf_path, output_path, format_type='excel', include_header=True, apply_formatting=True):
    """Executa o pipeline ETL em modo linha de comando"""
    print(f"Iniciando processamento do arquivo: {pdf_path}")
    print(f"Saída será salva em: {output_path}")
    print(f"Formato de saída: {format_type.upper()}")

    # Verificar se o arquivo PDF existe
    if not os.path.exists(pdf_path):
        print(f"ERRO: Arquivo PDF não encontrado: {pdf_path}")
        return False

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

        # Combinar todas as tabelas em uma única
        print("\nCombinando todas as tabelas...")
        import pandas as pd
        combined_df = pd.concat(cleaned_tables, ignore_index=True)
        print(f"Tabela combinada: {len(combined_df)} linhas, {len(combined_df.columns)} colunas")

        # Carregamento
        print("\n3. CARREGAMENTO")
        if format_type.lower() == 'csv':
            print("Salvando dados no CSV...")
            output_file = load_to_csv(combined_df, output_path, include_header=include_header)
            print(f"Dados salvos em: {output_file}")
        else:  # Excel é o padrão
            print("Salvando dados no Excel...")
            output_file = load_to_excel(combined_df, output_path)
            print(f"Dados salvos em: {output_file}")

        # Verificar se o arquivo de saída foi criado
        if os.path.exists(output_path):
            print("\nPROCESSAMENTO CONCLUÍDO COM SUCESSO!")
            print(f"Arquivo criado: {output_path}")
            print(f"Tamanho do arquivo: {os.path.getsize(output_path)} bytes")
            return True
        else:
            print(f"ERRO: Arquivo não foi criado: {output_path}")
            return False

    except Exception as e:
        print(f"\nERRO durante o processamento: {str(e)}")
        traceback.print_exc()
        return False

def main():
    # Tentar usar a interface gráfica
    try:
        # Importar o módulo gui
        from gui import create_gui
        print("Iniciando interface gráfica...")
        create_gui()
    except Exception as e:
        print(f"Erro ao iniciar a interface gráfica: {str(e)}")
        print("Executando em modo linha de comando...")

        # Definir caminhos e opções
        pdf_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
        output_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\saida_cli.xlsx"
        format_type = 'excel'  # ou 'csv'
        include_header = True

        # Executar o pipeline ETL
        run_etl_cli(pdf_path, output_path, format_type, include_header)

if __name__ == '__main__':
    main()
