import os
import sys

# Adicionar o diretório atual ao caminho de busca do Python
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Importar os módulos do projeto
from extraction import extract_tables_from_pdf
from transformation import clean_dataframe
from loading import load_to_excel

def run_etl_cli(pdf_path, excel_path):
    """Executa o pipeline ETL em modo linha de comando"""
    print(f"Iniciando processamento do arquivo: {pdf_path}")
    print(f"Saída será salva em: {excel_path}")
    
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
        
        # Carregamento
        print("\n3. CARREGAMENTO")
        print("Salvando dados no Excel...")
        excel_file = load_to_excel(cleaned_tables[0], excel_path)
        print(f"Dados salvos em: {excel_file}")
        
        # Verificar se o arquivo Excel foi criado
        if os.path.exists(excel_path):
            print("\nPROCESSAMENTO CONCLUÍDO COM SUCESSO!")
            print(f"Arquivo Excel criado: {excel_path}")
            print(f"Tamanho do arquivo: {os.path.getsize(excel_path)} bytes")
            return True
        else:
            print(f"ERRO: Arquivo Excel não foi criado: {excel_path}")
            return False
            
    except Exception as e:
        print(f"\nERRO durante o processamento: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    # Definir caminhos
    pdf_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
    excel_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\saida_cli.xlsx"
    
    # Executar o pipeline ETL
    run_etl_cli(pdf_path, excel_path)

if __name__ == '__main__':
    main()
