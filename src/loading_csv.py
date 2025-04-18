import os
import pandas as pd

def load_to_csv(df, csv_path, include_header=True):
    """Exporta DataFrame para arquivo CSV com opções de formatação"""
    # Criar diretório de saída se não existir
    output_dir = os.path.dirname(csv_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Adicionar cabeçalho com informações do documento se solicitado
    if include_header:
        # Criar um novo DataFrame para o cabeçalho
        header_df = pd.DataFrame([
            ["PREFEITURA DE IMPERATRIZ"],
            ["SECRETARIA DE FAZENDA E GESTÃO ORÇAMENTARIA"],
            ["SEFAZGO"],
            ["CNPJ: 06.158.455/0001-16"],
            ["Rua Godofredo Viana 722/738, Centro CEP: 65901-480 - Imperatriz-MA"],
            [""],  # Linha em branco
            ["RELATÓRIO DE SERVIÇOS PRESTADOS"],
            ["Gerado em: " + pd.Timestamp.now().strftime("%d/%m/%Y %H:%M:%S")],
            [""],  # Linha em branco
        ])

        # Salvar o cabeçalho em um arquivo temporário
        temp_header_path = csv_path + ".header.tmp"
        header_df.to_csv(temp_header_path, index=False, header=False, encoding='utf-8-sig', sep=';')

        # Salvar os dados em um arquivo temporário
        temp_data_path = csv_path + ".data.tmp"
        df.to_csv(temp_data_path, index=False, encoding='utf-8-sig', sep=';')

        # Combinar os arquivos
        with open(csv_path, 'w', encoding='utf-8-sig') as outfile:
            # Adicionar o cabeçalho
            with open(temp_header_path, 'r', encoding='utf-8-sig') as header_file:
                outfile.write(header_file.read())

            # Adicionar os dados
            with open(temp_data_path, 'r', encoding='utf-8-sig') as data_file:
                outfile.write(data_file.read())

        # Remover arquivos temporários
        os.remove(temp_header_path)
        os.remove(temp_data_path)
    else:
        # Salvar diretamente sem cabeçalho
        df.to_csv(csv_path, index=False, encoding='utf-8-sig', sep=';')

    return csv_path
