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

def test_resumo():
    """Testa a funcionalidade de criar uma aba de resumo"""
    # Definir caminhos
    pdf_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
    excel_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\saida_resumo_test.xlsx"
    
    print(f"Testando funcionalidade de resumo com arquivo: {pdf_path}")
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
        
        # Combinar todas as tabelas em uma única
        print("\nCombinando todas as tabelas...")
        combined_df = pd.concat(cleaned_tables, ignore_index=True)
        print(f"Tabela combinada: {len(combined_df)} linhas, {len(combined_df.columns)} colunas")
        
        # Adicionar coluna de situação para teste
        if 'situação' not in combined_df.columns:
            combined_df['situação'] = 'EM ABERTO'
            # Marcar algumas notas como canceladas para teste
            if len(combined_df) > 20:
                combined_df.loc[0:10, 'situação'] = 'CANCELADA'
        
        # Adicionar colunas de valores para teste
        if 'base de cálculo' not in combined_df.columns:
            combined_df['base de cálculo'] = 1000.0
        if 'iss próprio' not in combined_df.columns:
            combined_df['iss próprio'] = 50.0
        if 'iss retido' not in combined_df.columns:
            combined_df['iss retido'] = 25.0
        
        # Carregamento
        print("\n3. CARREGAMENTO")
        print("Salvando dados no Excel com aba de resumo...")
        excel_file = load_to_excel(combined_df, excel_path)
        print(f"Dados salvos em: {excel_file}")
        
        # Verificar se o arquivo Excel foi criado
        if os.path.exists(excel_path):
            print("\nTESTE CONCLUÍDO COM SUCESSO!")
            print(f"Arquivo Excel criado: {excel_path}")
            print(f"Tamanho do arquivo: {os.path.getsize(excel_path)} bytes")
            print("\nVerifique se a aba 'Resumo' foi criada corretamente no arquivo Excel.")
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
    test_resumo()
