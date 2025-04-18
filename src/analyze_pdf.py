import camelot
import pdfplumber
import pandas as pd
import os

def analyze_pdf(pdf_path):
    """Analisa a estrutura de um arquivo PDF e mostra informações detalhadas sobre as tabelas"""
    print(f"Analisando PDF: {pdf_path}")
    
    # Verificar se o arquivo existe
    if not os.path.exists(pdf_path):
        print(f"ERRO: Arquivo não encontrado: {pdf_path}")
        return
    
    # Extrair tabelas com camelot
    print("\n=== ANÁLISE COM CAMELOT ===")
    try:
        tables = camelot.read_pdf(pdf_path, pages='1-5', flavor='stream')
        print(f"Camelot encontrou {len(tables)} tabelas nas primeiras 5 páginas")
        
        # Analisar as primeiras 3 tabelas
        for i, table in enumerate(tables[:3]):
            print(f"\nTabela {i+1}:")
            print(f"Dimensões: {table.df.shape}")
            print(f"Colunas: {table.df.columns.tolist()}")
            print("Primeiras 5 linhas:")
            print(table.df.head())
            
            # Salvar tabela para análise
            output_file = f"tabela_camelot_{i+1}.csv"
            table.df.to_csv(output_file, index=False)
            print(f"Tabela salva em: {output_file}")
    
    except Exception as e:
        print(f"Erro ao analisar com camelot: {str(e)}")
    
    # Extrair tabelas com pdfplumber
    print("\n=== ANÁLISE COM PDFPLUMBER ===")
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"Total de páginas: {len(pdf.pages)}")
            
            # Analisar as primeiras 5 páginas
            for page_num in range(min(5, len(pdf.pages))):
                page = pdf.pages[page_num]
                tables = page.extract_tables()
                print(f"\nPágina {page_num+1}: encontradas {len(tables)} tabelas")
                
                # Analisar as tabelas da página
                for i, table in enumerate(tables):
                    if table:
                        print(f"  Tabela {i+1}:")
                        print(f"  Dimensões: {len(table)}x{len(table[0]) if table else 0}")
                        print(f"  Cabeçalho: {table[0]}")
                        print("  Primeiras 2 linhas de dados:")
                        for row in table[1:3]:
                            print(f"    {row}")
                        
                        # Salvar tabela para análise
                        df = pd.DataFrame(table[1:], columns=table[0])
                        output_file = f"tabela_pdfplumber_p{page_num+1}_t{i+1}.csv"
                        df.to_csv(output_file, index=False)
                        print(f"  Tabela salva em: {output_file}")
    
    except Exception as e:
        print(f"Erro ao analisar com pdfplumber: {str(e)}")
    
    print("\nAnálise concluída!")

if __name__ == "__main__":
    pdf_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
    analyze_pdf(pdf_path)
