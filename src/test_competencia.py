import os
import sys
import pandas as pd

# Adicionar o diretório src ao caminho de busca do Python
src_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(src_path)

# Importar os módulos do projeto
from extraction import extract_tables_from_pdf
from transformation import clean_dataframe
from loading import load_to_excel

def test_competencia():
    """Testa a funcionalidade de adicionar a coluna de competência"""
    # Definir caminhos
    pdf_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
    excel_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\saida_competencia_test.xlsx"
    
    print(f"Testando funcionalidade de competência com arquivo: {pdf_path}")
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
        
        # Verificar se a coluna de competência foi adicionada
        for i, df in enumerate(cleaned_tables):
            if 'competência' in df.columns:
                print(f"\nTabela {i+1}: Coluna de competência adicionada com sucesso!")
                print("Primeiras 5 linhas:")
                # Mostrar apenas as colunas de data e competência
                date_col = [col for col in df.columns if any(date_term in str(col).lower() for date_term in ['data', 'dt', 'date', 'emissão', 'emissao'])][0]
                print(df[[date_col, 'competência']].head())
            else:
                print(f"\nERRO: Tabela {i+1}: Coluna de competência não foi adicionada!")
        
        # Carregamento
        print("\n3. CARREGAMENTO")
        print("Salvando dados no Excel...")
        excel_file = load_to_excel(cleaned_tables[0], excel_path)
        print(f"Dados salvos em: {excel_file}")
        
        # Verificar se o arquivo Excel foi criado
        if os.path.exists(excel_path):
            print("\nTESTE CONCLUÍDO COM SUCESSO!")
            print(f"Arquivo Excel criado: {excel_path}")
            print(f"Tamanho do arquivo: {os.path.getsize(excel_path)} bytes")
            return True
        else:
            print(f"ERRO: Arquivo Excel não foi criado: {excel_path}")
            return False
            
    except Exception as e:
        print(f"\nERRO durante o teste: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_competencia()
