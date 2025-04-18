import os
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def load_to_excel(df, excel_path):
    """Exporta DataFrame para arquivo Excel com formatação básica"""
    # Criar diretório de saída se não existir
    output_dir = os.path.dirname(excel_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Exportar para Excel
    df.to_excel(excel_path, index=False, sheet_name='Dados')

    # Aplicar formatação
    wb = load_workbook(excel_path)
    ws = wb['Dados']

    # Definir estilos
    header_font = Font(bold=True, size=12, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Formatar cabeçalho
    for col_num, _ in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # Formatar células de dados
    for row in range(2, len(df) + 2):  # +2 porque Excel é 1-indexado e temos cabeçalho
        for col in range(1, len(df.columns) + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border

            # Alinhar números à direita
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal='right')

            # Formatar datas
            column_name = df.columns[col-1].lower()
            if any(date_term in column_name for date_term in ['data', 'dt', 'date', 'emissão', 'emissao']):
                # Verificar se o valor da célula é uma data no formato DD/MM/YYYY
                if cell.value and re.match(r'\d{2}/\d{2}/\d{4}', str(cell.value)):
                    # Formatar como texto para garantir que seja exibido corretamente
                    cell.number_format = '@'
                    # Garantir que o valor seja uma string
                    cell.value = str(cell.value)

            # Formatar competência
            if 'competência' in column_name:
                # Verificar se o valor da célula é uma competência no formato MM/AAAA
                if cell.value and re.match(r'\d{2}/\d{4}', str(cell.value)):
                    # Formatar como texto para garantir que seja exibido corretamente
                    cell.number_format = '@'
                    # Garantir que o valor seja uma string
                    cell.value = str(cell.value)
                    # Centralizar o texto
                    cell.alignment = Alignment(horizontal='center')

            # Formatar valores monetários
            if any(value_term in column_name for value_term in ['valor', 'vlr', 'preço', 'preco', 'total']):
                cell.number_format = 'R$ #,##0.00'

    # Ajustar largura das colunas
    for col_num, _ in enumerate(df.columns, 1):
        col_letter = get_column_letter(col_num)
        # Definir largura mínima e máxima
        max_length = 0
        for row_num in range(1, len(df) + 2):
            cell_value = str(ws.cell(row=row_num, column=col_num).value or '')
            max_length = max(max_length, len(cell_value))
        adjusted_width = min(max(max_length + 2, 10), 50)  # Entre 10 e 50 caracteres
        ws.column_dimensions[col_letter].width = adjusted_width

    # Congelar painel no cabeçalho
    ws.freeze_panes = 'A2'

    # Salvar o arquivo formatado
    wb.save(excel_path)

    return excel_path
