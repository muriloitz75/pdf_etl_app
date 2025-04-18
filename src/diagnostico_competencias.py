import os
import sys
import pandas as pd

# Adicionar o diretório src ao caminho de busca do Python
src_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(src_path)

# Importar os módulos do projeto
from extraction import extract_tables_from_pdf
from transformation import clean_dataframe

def diagnosticar_competencias(pdf_path, competencias_alvo=['07/2022', '08/2022', '10/2022']):
    """Analisa detalhadamente as competências específicas para identificar problemas"""
    print(f"Analisando competências problemáticas no arquivo: {pdf_path}")
    
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
        
        # Analisar cada competência alvo
        for competencia in competencias_alvo:
            print(f"\n{'='*80}")
            print(f"ANÁLISE DETALHADA DA COMPETÊNCIA: {competencia}")
            print(f"{'='*80}")
            
            # Filtrar dados da competência
            comp_data = combined_df[combined_df['competência'] == competencia]
            
            if len(comp_data) == 0:
                print(f"Nenhum dado encontrado para a competência {competencia}")
                continue
            
            # Estatísticas gerais
            print(f"\nEstatísticas gerais:")
            print(f"Total de registros: {len(comp_data)}")
            
            # Verificar situações
            if 'situação' in comp_data.columns:
                situacoes = comp_data['situação'].value_counts()
                print(f"\nDistribuição de situações:")
                for situacao, count in situacoes.items():
                    print(f"  {situacao}: {count} registros")
                
                # Contar notas canceladas
                notas_canceladas = comp_data[comp_data['situação'].str.contains('CANCELAD', case=False, na=False)]
                print(f"\nNotas canceladas: {len(notas_canceladas)} de {len(comp_data)}")
                
                # Filtrar notas válidas (não canceladas)
                notas_validas = comp_data[~comp_data['situação'].str.contains('CANCELAD', case=False, na=False)]
                print(f"Notas válidas: {len(notas_validas)} de {len(comp_data)}")
            
            # Analisar valores financeiros
            print(f"\nAnálise de valores financeiros:")
            
            # Base de cálculo
            if 'base de cálculo' in comp_data.columns:
                print(f"\nBase de Cálculo:")
                # Converter para numérico
                comp_data['base_calc_num'] = pd.to_numeric(
                    comp_data['base de cálculo'].astype(str).str.replace(',', '.').str.replace('R$', '').str.strip(), 
                    errors='coerce'
                )
                
                # Estatísticas gerais
                print(f"  Soma total: {comp_data['base_calc_num'].sum():.2f}")
                print(f"  Média: {comp_data['base_calc_num'].mean():.2f}")
                print(f"  Mínimo: {comp_data['base_calc_num'].min():.2f}")
                print(f"  Máximo: {comp_data['base_calc_num'].max():.2f}")
                
                # Analisar por situação
                if 'situação' in comp_data.columns:
                    print(f"\n  Análise por situação:")
                    for situacao in comp_data['situação'].unique():
                        situacao_data = comp_data[comp_data['situação'] == situacao]
                        soma = situacao_data['base_calc_num'].sum()
                        print(f"    {situacao}: {soma:.2f}")
                
                # Verificar valores extremos
                print(f"\n  Valores extremos (top 5):")
                extremos = comp_data.nlargest(5, 'base_calc_num')[['base_calc_num', 'situação']]
                for i, (_, row) in enumerate(extremos.iterrows(), 1):
                    print(f"    {i}. {row['base_calc_num']:.2f} - Situação: {row['situação']}")
            
            # ISS Próprio
            if 'iss próprio' in comp_data.columns:
                print(f"\nISS Próprio:")
                # Converter para numérico
                comp_data['iss_proprio_num'] = pd.to_numeric(
                    comp_data['iss próprio'].astype(str).str.replace(',', '.').str.replace('R$', '').str.strip(), 
                    errors='coerce'
                )
                
                # Estatísticas gerais
                print(f"  Soma total: {comp_data['iss_proprio_num'].sum():.2f}")
                print(f"  Média: {comp_data['iss_proprio_num'].mean():.2f}")
                
                # Analisar por situação
                if 'situação' in comp_data.columns:
                    print(f"\n  Análise por situação:")
                    for situacao in comp_data['situação'].unique():
                        situacao_data = comp_data[comp_data['situação'] == situacao]
                        soma = situacao_data['iss_proprio_num'].sum()
                        print(f"    {situacao}: {soma:.2f}")
            
            # ISS Retido
            if 'iss retido' in comp_data.columns:
                print(f"\nISS Retido:")
                # Converter para numérico
                comp_data['iss_retido_num'] = pd.to_numeric(
                    comp_data['iss retido'].astype(str).str.replace(',', '.').str.replace('R$', '').str.strip(), 
                    errors='coerce'
                )
                
                # Estatísticas gerais
                print(f"  Soma total: {comp_data['iss_retido_num'].sum():.2f}")
                print(f"  Média: {comp_data['iss_retido_num'].mean():.2f}")
                
                # Analisar por situação
                if 'situação' in comp_data.columns:
                    print(f"\n  Análise por situação:")
                    for situacao in comp_data['situação'].unique():
                        situacao_data = comp_data[comp_data['situação'] == situacao]
                        soma = situacao_data['iss_retido_num'].sum()
                        print(f"    {situacao}: {soma:.2f}")
            
            # Verificar se há inconsistências nos dados
            print(f"\nVerificação de inconsistências:")
            
            # Verificar se há valores nulos ou inválidos
            for col in ['base de cálculo', 'iss próprio', 'iss retido']:
                if col in comp_data.columns:
                    nulos = comp_data[col].isna().sum()
                    print(f"  Valores nulos em '{col}': {nulos}")
            
            # Verificar se há registros duplicados
            duplicados = comp_data.duplicated().sum()
            print(f"  Registros duplicados: {duplicados}")
            
            # Verificar se há inconsistências entre base de cálculo e ISS
            if 'base_calc_num' in comp_data.columns and 'iss_proprio_num' in comp_data.columns:
                # Verificar se há registros com base de cálculo mas sem ISS
                inconsistentes = comp_data[(comp_data['base_calc_num'] > 0) & (comp_data['iss_proprio_num'] == 0)].shape[0]
                print(f"  Registros com base de cálculo > 0 mas ISS próprio = 0: {inconsistentes}")
                
                # Verificar se há registros com ISS mas sem base de cálculo
                inconsistentes = comp_data[(comp_data['base_calc_num'] == 0) & (comp_data['iss_proprio_num'] > 0)].shape[0]
                print(f"  Registros com base de cálculo = 0 mas ISS próprio > 0: {inconsistentes}")
            
            # Mostrar alguns exemplos de registros
            print(f"\nExemplos de registros (primeiros 5):")
            colunas_exemplo = [col for col in ['competência', 'situação', 'base de cálculo', 'iss próprio', 'iss retido'] if col in comp_data.columns]
            print(comp_data[colunas_exemplo].head())
        
        return True
            
    except Exception as e:
        print(f"\nERRO durante a análise: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    pdf_path = r"C:\Users\Murilo\Desktop\pdf_etl_app\Arquivo\2022.pdf"
    diagnosticar_competencias(pdf_path)
