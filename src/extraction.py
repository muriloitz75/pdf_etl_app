import camelot
import pdfplumber
import pandas as pd

def extract_tables_from_pdf(pdf_path, method='auto'):
    """Extrai tabelas de um arquivo PDF e retorna lista de DataFrames

    Args:
        pdf_path: Caminho para o arquivo PDF
        method: Método de extração ('auto', 'camelot' ou 'pdfplumber')
    """
    print(f"Tentando extrair tabelas de: {pdf_path}")
    print(f"Método de extração: {method}")

    try:
        # Usar Camelot se o método for 'auto' ou 'camelot'
        if method in ['auto', 'camelot']:
            print("Tentando com camelot...")
            tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
            if len(tables) > 0:
                print(f"Camelot encontrou {len(tables)} tabelas")

                # Filtrar tabelas relevantes (com dados de serviços)
                filtered_tables = []
                for table in tables:
                    df = table.df
                    # Verificar se a tabela tem pelo menos 5 linhas e contém dados de serviços
                    if len(df) >= 5 and df.iloc[:, 0].astype(str).str.contains('2022000000').any():
                        # Processar a tabela para extrair os dados corretamente
                        processed_df = process_service_table(df)
                        if not processed_df.empty:
                            filtered_tables.append(processed_df)

                if filtered_tables:
                    print(f"Encontradas {len(filtered_tables)} tabelas relevantes")
                    return filtered_tables
                elif method == 'camelot':
                    print("Nenhuma tabela relevante encontrada com Camelot")
                    return []

        # Usar PDFPlumber se o método for 'auto' ou 'pdfplumber'
        if method in ['auto', 'pdfplumber']:
            print("Tentando com pdfplumber...")
            with pdfplumber.open(pdf_path) as pdf:
                dfs = []
                for page_num, page in enumerate(pdf.pages, 1):
                    tables = page.extract_tables()
                    print(f"Página {page_num}: encontradas {len(tables)} tabelas")
                    for table in tables:
                        if table and len(table) > 1:  # verifica se a tabela não está vazia
                            # Verificar se a tabela contém dados de serviços
                            table_text = ' '.join([' '.join(str(cell) for cell in row if cell is not None) for row in table])
                            if '2022000000' in table_text:
                                # Processar a tabela para extrair os dados corretamente
                                df = process_pdfplumber_table(table)
                                if not df.empty:
                                    dfs.append(df)

                if dfs:
                    print(f"PDFPlumber encontrou {len(dfs)} tabelas relevantes")
                    return dfs

        print("Nenhuma tabela relevante encontrada")
        return []
    except Exception as e:
        print(f"Erro durante a extração: {str(e)}")
        raise Exception(f"Falha ao extrair tabelas do PDF: {str(e)}")

def process_service_table(df):
    """Processa uma tabela de serviços extraída pelo Camelot"""
    try:
        # Encontrar a linha que contém os cabeçalhos reais
        header_row = None
        for i, row in df.iterrows():
            if any(col.lower() == 'n° nota' for col in row if isinstance(col, str)):
                header_row = i
                break

        if header_row is not None:
            # Usar a linha do cabeçalho como colunas
            headers = df.iloc[header_row].tolist()
            # Pegar apenas as linhas após o cabeçalho
            data = df.iloc[header_row+1:].reset_index(drop=True)
            # Definir os novos cabeçalhos
            data.columns = headers

            # Filtrar apenas as linhas que contém números de nota fiscal
            data = data[data.iloc[:, 0].astype(str).str.contains('2022000000', na=False)]

            # Manter as datas originais do PDF
            # Sem processamento para preservar o formato original

            return data
        else:
            # Se não encontrou o cabeçalho, verificar se a tabela já contém dados de serviços
            if any(df.iloc[:, 0].astype(str).str.contains('2022000000', na=False)):
                # Criar cabeçalhos padrão se não encontrou
                headers = ['N° Nota', 'Dt,. Emissão', 'CPF/CNPJ Tomador', 'Tomador do Serviço',
                          'Serviço', 'Vrl. Serviço', 'Base de Cálculo', 'Aliq.',
                          'ISS Próprio', 'ISS Retido', 'Nat. da Operação', 'Incidência', 'Situação']

                # Filtrar apenas as linhas que contém números de nota fiscal
                data = df[df.iloc[:, 0].astype(str).str.contains('2022000000', na=False)].reset_index(drop=True)

                # Processar as datas para garantir que estejam no formato correto
                if len(data.columns) > 1:  # Verificar se há pelo menos duas colunas (nota e data)
                    # Extrair as datas usando regex da segunda coluna (geralmente a data)
                    import re
                    # Primeiro, extrair as datas no formato DD/MM/YYYY
                    data[1] = data[1].astype(str).apply(lambda x: re.search(r'(\d{2}/\d{2}/\d{4})', x).group(1) if re.search(r'(\d{2}/\d{2}/\d{4})', x) else '')

                # Ajustar o número de colunas se necessário
                if len(data.columns) < len(headers):
                    for i in range(len(data.columns), len(headers)):
                        data[i] = ''
                elif len(data.columns) > len(headers):
                    data = data.iloc[:, :len(headers)]

                data.columns = headers
                return data

            return pd.DataFrame()  # Retorna DataFrame vazio se não encontrou dados relevantes
    except Exception as e:
        print(f"Erro ao processar tabela: {str(e)}")
        return pd.DataFrame()

def process_pdfplumber_table(table):
    """Processa uma tabela extraída pelo PDFPlumber"""
    try:
        # Verificar se a tabela contém dados de serviços
        has_service_data = False
        for row in table:
            row_text = ' '.join([str(cell) for cell in row if cell is not None])
            if '2022000000' in row_text:
                has_service_data = True
                break

        if has_service_data:
            # Criar DataFrame a partir da tabela
            df = pd.DataFrame(table)

            # Encontrar a linha que contém os cabeçalhos
            header_row = None
            for i, row in df.iterrows():
                row_text = ' '.join([str(cell) for cell in row if cell is not None])
                if 'N° Nota' in row_text or 'Nota' in row_text and 'Emissão' in row_text:
                    header_row = i
                    break

            if header_row is not None:
                # Usar a linha do cabeçalho como colunas
                headers = df.iloc[header_row].tolist()
                # Pegar apenas as linhas após o cabeçalho
                data = df.iloc[header_row+1:].reset_index(drop=True)
                # Definir os novos cabeçalhos
                data.columns = headers

                # Processar as datas para garantir que estejam no formato correto
                for col in data.columns:
                    if 'data' in str(col).lower() or 'dt' in str(col).lower() or 'emissão' in str(col).lower():
                        # Extrair as datas usando regex
                        import re
                        # Primeiro, extrair as datas no formato DD/MM/YYYY
                        data[col] = data[col].astype(str).apply(lambda x: re.search(r'(\d{2}/\d{2}/\d{4})', x).group(1) if re.search(r'(\d{2}/\d{2}/\d{4})', x) else '')
            else:
                # Se não encontrou o cabeçalho, criar cabeçalhos padrão
                headers = ['N° Nota', 'Dt,. Emissão', 'CPF/CNPJ Tomador', 'Tomador do Serviço',
                          'Serviço', 'Vrl. Serviço', 'Base de Cálculo', 'Aliq.',
                          'ISS Próprio', 'ISS Retido', 'Nat. da Operação', 'Incidência', 'Situação']

                # Filtrar apenas as linhas que contém dados de serviços
                data = pd.DataFrame()
                for row in table:
                    row_text = ' '.join([str(cell) for cell in row if cell is not None])
                    if '2022000000' in row_text:
                        # Processar a linha para extrair os campos
                        processed_row = process_service_line(row_text)
                        if processed_row:
                            data = pd.concat([data, pd.DataFrame([processed_row])], ignore_index=True)

                if not data.empty:
                    # Ajustar o número de colunas se necessário
                    for i in range(len(headers)):
                        if i >= len(data.columns):
                            data[i] = ''

                    data.columns = headers

            return data

        return pd.DataFrame()  # Retorna DataFrame vazio se não encontrou dados relevantes
    except Exception as e:
        print(f"Erro ao processar tabela PDFPlumber: {str(e)}")
        return pd.DataFrame()

def process_service_line(line_text):
    """Processa uma linha de texto para extrair os campos de serviço"""
    try:
        # Extrair o número da nota (começa com 2022000000)
        import re
        nota_match = re.search(r'(2022000000\d+)', line_text)
        if not nota_match:
            return None

        nota = nota_match.group(1)

        # Extrair a data (formato DD/MM/YYYY)
        data_match = re.search(r'(\d{2}/\d{2}/\d{4})', line_text)
        data = data_match.group(1) if data_match else ''

        # Extrair CPF/CNPJ (formato XX.XXX.XXX/XXXX-XX ou XXX.XXX.XXX-XX)
        cpf_cnpj_match = re.search(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})', line_text)
        cpf_cnpj = cpf_cnpj_match.group(1) if cpf_cnpj_match else ''

        # Extrair valores monetários (formato X.XXX,XX)
        valores = re.findall(r'(\d+\.\d+,\d{2})', line_text)

        # Criar um dicionário com os campos extraídos
        row = {
            'N° Nota': nota,
            'Data Emissão': data,
            'CPF/CNPJ Tomador': cpf_cnpj,
            'Tomador do Serviço': '',  # Difícil extrair com precisão
            'Valor Serviço': valores[0] if valores else '',
            'Base de Cálculo': valores[1] if len(valores) > 1 else '',
            'ISS Próprio': valores[3] if len(valores) > 3 else '',
            'ISS Retido': valores[4] if len(valores) > 4 else ''
        }

        return row
    except Exception as e:
        print(f"Erro ao processar linha: {str(e)}")
        return None
