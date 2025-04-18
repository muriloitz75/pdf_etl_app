import camelot
import pdfplumber
import pandas as pd
import sys

def debug_extraction(pdf_path):
    """Depura a extração de tabelas de um arquivo PDF"""
    print(f"Tentando extrair tabelas de: {pdf_path}")
    
    try:
        # Primeira tentativa usando camelot
        print("Tentando com camelot...")
        print(f"Versão do camelot: {camelot.__version__}")
        tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
        print(f"Camelot encontrou {len(tables)} tabelas")
        
        # Segunda tentativa usando pdfplumber
        print("Tentando com pdfplumber...")
        with pdfplumber.open(pdf_path) as pdf:
            print(f"Número de páginas: {len(pdf.pages)}")
            for page_num, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables()
                print(f"Página {page_num}: encontradas {len(tables)} tabelas")
                
    except Exception as e:
        print(f"Erro durante a depuração: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    pdf_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
    debug_extraction(pdf_path)
