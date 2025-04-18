import pandas as pd
import re

def clean_dataframe(df: pd.DataFrame, remove_empty_rows=True, remove_empty_cols=True,
                  convert_dates=True, convert_money=True) -> pd.DataFrame:
    """Limpa e padroniza colunas, tipos de dados e trata valores faltantes

    Args:
        df: DataFrame a ser limpo
        remove_empty_rows: Se True, remove linhas completamente vazias
        remove_empty_cols: Se True, remove colunas completamente vazias
        convert_dates: Se True, tenta converter colunas de data
        convert_money: Se True, tenta converter colunas de valores monetários
    """
    # Cria uma cópia para não modificar o original
    df = df.copy()

    # Remover linhas completamente vazias
    if remove_empty_rows:
        df = df.dropna(how='all')

    # Remover colunas completamente vazias
    if remove_empty_cols:
        df = df.dropna(axis=1, how='all')

    # Normalizar nomes das colunas (remover espaços extras, converter para minúsculas)
    df.columns = [str(col).strip().lower() for col in df.columns]

    # Tentar identificar e converter colunas de data
    if convert_dates:
        date_columns = [col for col in df.columns if any(date_term in str(col).lower() for date_term in ['data', 'dt', 'date', 'emissão', 'emissao'])]
        for col in date_columns:
            if col in df.columns:
                try:
                    # Converter para string e limpar
                    df[col] = df[col].astype(str).str.strip()

                    # Aplicar regex para extrair datas no formato DD/MM/YYYY
                    df[col] = df[col].astype(str).apply(lambda x: re.search(r'(\d{2}/\d{2}/\d{4})', x).group(1) if re.search(r'(\d{2}/\d{2}/\d{4})', x) else '')

                    # Substituir valores NaN ou 'nan' por string vazia
                    df[col] = df[col].replace(['nan', 'NaN', 'NaT', 'None', 'none'], '')
                except Exception as e:
                    print(f"Erro ao processar coluna {col}: {str(e)}")
                    pass  # Se não conseguir processar, mantém como está

    # Tentar identificar e converter colunas de valores monetários
    if convert_money:
        value_columns = [col for col in df.columns if any(value_term in str(col).lower() for value_term in ['valor', 'vlr', 'preço', 'preco', 'total', 'base', 'iss', 'alíquota'])]
        for col in value_columns:
            if col in df.columns:
                try:
                    # Verificar se a coluna não é 'tomador do serviço' ou similar
                    if 'tomador' in str(col).lower() or 'serviço' in str(col).lower() or 'servico' in str(col).lower():
                        continue

                    # Converter valores monetários (substituir pontos e vírgulas)
                    df[col] = df[col].astype(str).str.replace('R$', '', regex=False).str.strip()
                    df[col] = df[col].apply(lambda x: re.sub(r'[^\d,.]', '', str(x)))
                    # Substituir strings vazias por '0'
                    df[col] = df[col].replace('', '0')
                    df[col] = df[col].str.replace('.', '').str.replace(',', '.').astype(float)
                except Exception as e:
                    print(f"Erro ao converter coluna {col} para valor monetário: {str(e)}")
                    pass  # Se não conseguir converter, mantém como está

    # Remover linhas que parecem ser cabeçalhos duplicados ou rodapés
    # (geralmente contêm palavras como 'total', 'página', etc.)
    try:
        header_footer_terms = ['total', 'página', 'pagina', 'subtotal', 'sub-total']
        mask = ~df.astype(str).apply(lambda x: x.str.lower().str.contains('|'.join(header_footer_terms), na=False)).any(axis=1)
        df = df[mask]
    except Exception as e:
        print(f"Erro ao remover cabeçalhos/rodapés: {str(e)}")

    return df
