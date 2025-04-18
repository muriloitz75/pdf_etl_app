import os
import sys
import threading
import time
import PySimpleGUI as sg

# Garantir que os módulos possam ser importados
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Importar os módulos do projeto
from extraction import extract_tables_from_pdf
from transformation import clean_dataframe
from loading import load_to_excel
from loading_csv import load_to_csv

def process_etl(pdf_path, output_path, window, values):
    """Processa o ETL em uma thread separada e atualiza a interface"""
    try:
        # Obter opções de processamento
        format_type = 'csv' if values['format_csv'] else 'excel'
        include_header = values['include_header']
        apply_formatting = values['apply_formatting']

        # Configurar método de extração
        extraction_method = 'auto'
        if values['camelot_extract']:
            extraction_method = 'camelot'
        elif values['pdfplumber_extract']:
            extraction_method = 'pdfplumber'

        # Etapa 1: Extração
        window.write_event_value('-UPDATE-', {'status': 'Extraindo tabelas do PDF...', 'progress': 25})
        tables = extract_tables_from_pdf(pdf_path, method=extraction_method)

        if not tables:
            window.write_event_value('-ERROR-', 'Nenhuma tabela encontrada no PDF.')
            return

        # Etapa 2: Transformação
        window.write_event_value('-UPDATE-', {'status': 'Limpando e transformando dados...', 'progress': 50})

        # Aplicar opções de transformação
        transform_options = {
            'remove_empty_rows': values['remove_empty_rows'],
            'remove_empty_cols': values['remove_empty_cols'],
            'convert_dates': values['convert_dates'],
            'convert_money': values['convert_money']
        }

        cleaned = [clean_dataframe(df, **transform_options) for df in tables]

        if not cleaned:
            window.write_event_value('-ERROR-', 'Erro ao transformar os dados.')
            return

        # Combinar todas as tabelas em uma única
        window.write_event_value('-UPDATE-', {'status': 'Combinando todas as tabelas...', 'progress': 60})
        import pandas as pd
        combined_df = pd.concat(cleaned, ignore_index=True)

        # Etapa 3: Carregamento
        if format_type == 'csv':
            window.write_event_value('-UPDATE-', {'status': 'Salvando no CSV...', 'progress': 75})
            load_to_csv(combined_df, output_path, include_header=include_header)
        else:  # Excel é o padrão
            window.write_event_value('-UPDATE-', {'status': 'Salvando no Excel...', 'progress': 75})
            load_to_excel(combined_df, output_path)

        # Concluído
        window.write_event_value('-UPDATE-', {'status': 'Processamento concluído!', 'progress': 100})
        window.write_event_value('-DONE-', output_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.write_event_value('-ERROR-', str(e))

def create_gui():
    sg.theme('DefaultNoMoreNagging')  # Define o tema da interface

    # Definição das abas
    tab1_layout = [
        [sg.Text('Arquivo PDF:', size=(12, 1)),
         sg.Input(key='pdf', size=(50, 1)),
         sg.FileBrowse('Procurar', file_types=(('PDF Files', '*.pdf'),))],
        [sg.Text('Output Excel:', size=(12, 1)),
         sg.Input(key='excel', size=(50, 1)),
         sg.FileSaveAs('Salvar como', file_types=(('Excel Files', '*.xlsx'),))],
        [sg.Text('Formato de saída:'),
         sg.Radio('Excel', 'FORMAT', key='format_excel', default=True),
         sg.Radio('CSV', 'FORMAT', key='format_csv')],
        [sg.Checkbox('Incluir cabeçalho com informações do documento', key='include_header', default=True)],
        [sg.Checkbox('Aplicar formatação avançada (cores, bordas, etc.)', key='apply_formatting', default=True)],
        [sg.Text('_' * 80)],
        [sg.Text('Status:', size=(10, 1)), sg.Text('Aguardando...', key='status', size=(50, 1))],
        [sg.ProgressBar(100, orientation='h', size=(50, 20), key='progress')],
        [sg.Button('Executar ETL', size=(15, 1), button_color=('white', '#007BFF')),
         sg.Button('Cancelar', size=(15, 1))]
    ]

    tab2_layout = [
        [sg.Text('Configurações de Extração', font=('Arial', 12, 'bold'))],
        [sg.Text('Método de extração:')],
        [sg.Radio('Automático (tenta todos os métodos)', 'EXTRACTION', key='auto_extract', default=True)],
        [sg.Radio('Camelot (melhor para tabelas com linhas)', 'EXTRACTION', key='camelot_extract')],
        [sg.Radio('PDFPlumber (melhor para tabelas sem linhas)', 'EXTRACTION', key='pdfplumber_extract')],
        [sg.Text('_' * 80)],
        [sg.Text('Configurações de Transformação', font=('Arial', 12, 'bold'))],
        [sg.Checkbox('Remover linhas vazias', key='remove_empty_rows', default=True)],
        [sg.Checkbox('Remover colunas vazias', key='remove_empty_cols', default=True)],
        [sg.Checkbox('Converter datas (DD/MM/YYYY)', key='convert_dates', default=True)],
        [sg.Checkbox('Converter valores monetários', key='convert_money', default=True)],
    ]

    tab3_layout = [
        [sg.Text('Sobre o Aplicativo', font=('Arial', 12, 'bold'))],
        [sg.Text('ETL PDF para Excel - Versão 1.1')],
        [sg.Text('Este aplicativo extrai tabelas de arquivos PDF e as exporta para Excel.')],
        [sg.Text('Desenvolvido com Python, PySimpleGUI, Camelot e PDFPlumber.')],
        [sg.Text('_' * 80)],
        [sg.Text('Instruções de Uso:', font=('Arial', 12, 'bold'))],
        [sg.Text('1. Selecione o arquivo PDF de entrada')],
        [sg.Text('2. Escolha o local para salvar o arquivo Excel')],
        [sg.Text('3. Configure as opções de extração e transformação (opcional)')],
        [sg.Text('4. Clique em "Executar ETL" para iniciar o processamento')],
    ]

    layout = [
        [sg.Text('ETL PDF para Excel', size=(30, 1), font=('Arial', 16, 'bold'), justification='center')],
        [sg.TabGroup([[sg.Tab('Principal', tab1_layout),
                       sg.Tab('Configurações', tab2_layout),
                       sg.Tab('Ajuda', tab3_layout)]])],
    ]

    window = sg.Window('PDF ETL App', layout, resizable=True, finalize=True)

    # Inicializar a barra de progresso
    progress_bar = window['progress']
    progress_bar.update(0)

    # Loop principal
    processing = False
    while True:
        event, values = window.read(timeout=100 if processing else None)

        if event == sg.WIN_CLOSED or event == 'Cancelar':
            break

        if event == 'Executar ETL' and not processing:
            if not values['pdf'] or not values['excel']:
                sg.popup_error('Por favor, selecione o arquivo PDF e o caminho do Excel!')
                continue

            # Ajustar a extensão do arquivo de saída conforme o formato selecionado
            output_path = values['excel']
            if values['format_csv'] and not output_path.lower().endswith('.csv'):
                output_path = output_path.rsplit('.', 1)[0] + '.csv'
                window['excel'].update(output_path)
            elif values['format_excel'] and not output_path.lower().endswith(('.xlsx', '.xls')):
                output_path = output_path.rsplit('.', 1)[0] + '.xlsx'
                window['excel'].update(output_path)

            # Iniciar processamento
            processing = True
            window['status'].update('Iniciando processamento...')
            progress_bar.update(0)

            # Iniciar thread de processamento
            thread = threading.Thread(target=process_etl, args=(values['pdf'], output_path, window, values), daemon=True)
            thread.start()

        # Atualizações da thread de processamento
        elif event == '-UPDATE-':
            window['status'].update(values[event]['status'])
            progress_bar.update(values[event]['progress'])

        elif event == '-ERROR-':
            processing = False
            window['status'].update('Erro!')
            sg.popup_error(f'Erro durante o processamento:\n{values[event]}')

        elif event == '-DONE-':
            processing = False
            window['status'].update('Concluído!')
            sg.popup('ETL concluído com sucesso!', f'Arquivo salvo em:\n{values[event]}')

    window.close()

if __name__ == '__main__':
    create_gui()

