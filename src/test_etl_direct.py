from extraction import extract_tables_from_pdf
from transformation import clean_dataframe
from loading import load_to_excel

def run_etl(pdf_path, excel_output):
    print(f"Iniciando processamento do arquivo: {pdf_path}")

    # Extração
    print("Extraindo tabelas...")
    tables = extract_tables_from_pdf(pdf_path)
    print(f"Encontradas {len(tables)} tabelas")

    # Transformação
    print("Limpando dados...")
    cleaned_tables = [clean_dataframe(df) for df in tables]

    # Loading
    if cleaned_tables:
        print("Salvando no Excel...")
        load_to_excel(cleaned_tables[0], excel_output)
        print(f"Arquivo Excel salvo em: {excel_output}")
    else:
        print("Nenhuma tabela encontrada para processar")

if __name__ == '__main__':
    # Caminhos configurados para seus arquivos
    PDF_PATH = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
    EXCEL_OUTPUT = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\saida_2022.xlsx"

    run_etl(PDF_PATH, EXCEL_OUTPUT)

