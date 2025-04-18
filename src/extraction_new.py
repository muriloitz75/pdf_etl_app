import camelot
import pdfplumber
import pandas as pd

def extract_tables_from_pdf(pdf_path):
    """Extrai tabelas de um arquivo PDF e retorna lista de DataFrames"""
    print(f"Tentando extrair tabelas de: {pdf_path}")
    
    try:
        # Primeira tentativa usando camelot
        print("Tentando com camelot...")
        tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
        if len(tables) > 0:
            print(f"Camelot encontrou {len(tables)} tabelas")
            return [table.df for table in tables]
        
        # Segunda tentativa usando pdfplumber
        print("Tentando com pdfplumber...")
        with pdfplumber.open(pdf_path) as pdf:
            dfs = []
            for page_num, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables()
                print(f"Página {page_num}: encontradas {len(tables)} tabelas")
                for table in tables:
                    if table:  # verifica se a tabela não está vazia
                        df = pd.DataFrame(table[1:], columns=table[0])
                        dfs.append(df)
            return dfs
            
    except Exception as e:
        print(f"Erro durante a extração: {str(e)}")
        raise Exception(f"Falha ao extrair tabelas do PDF: {str(e)}")
