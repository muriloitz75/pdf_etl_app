# ETL de PDF para Excel com Interface

Este projeto implementa um pipeline ETL (Extract, Transform, Load) em Python que:
1. **Extrai** tabelas de arquivos PDF (Livro de Serviços Prestados) usando Camelot e PDFPlumber.
2. **Transforma** e higieniza dados (remoção de cabeçalhos duplicados, normalização de formatos, conversão de tipos).
3. **Carrega** a saída em planilhas Excel formatadas profissionalmente.
4. Disponibiliza uma interface gráfica intuitiva para seleção de arquivos e geração de relatórios.

## Estrutura de pastas

- **src/**: Código-fonte dividido em módulos:
  - `extraction.py`: Rotinas de extração usando Camelot e PDFPlumber para detectar e extrair tabelas de PDFs.
  - `transformation.py`: Funções de limpeza e padronização com pandas (normalização de colunas, conversão de tipos).
  - `loading.py`: Exportação para Excel com openpyxl, incluindo formatação profissional.
  - `gui.py`: Interface gráfica intuitiva desenvolvida com PySimpleGUI.
  - `main.py`: Ponto de entrada da aplicação.
  - `test_pipeline.py`: Script para testar o pipeline ETL completo.
- **Arquivo/**: Pasta com arquivos de exemplo e saídas geradas.
- **requirements.txt**: Dependências do projeto.
- **build_exe.py**: Script para gerar o executável.
- **dist/**: Saída do build do executável (gerado via PyInstaller).
- **README.md**: Documentação do projeto.

## Pré-requisitos

- Python 3.8+
- Bibliotecas listadas em `requirements.txt`

## Instalação

```bash
# Clonar o repositório
git clone <repositório>
cd pdf_etl_app

# Instalar dependências
pip install -r requirements.txt
```

## Uso

### Usando o executável

1. Execute o arquivo `dist/pdf_etl_app.exe`
2. Na interface, selecione o arquivo PDF de entrada
3. Escolha o local onde deseja salvar o arquivo Excel de saída
4. Clique em "Executar ETL" e aguarde a confirmação

### Em modo desenvolvimento

```bash
python src/main.py
```

## Testes

Para testar o pipeline ETL completo:

```bash
python src/test_pipeline.py
```

## Build do executável

O projeto já inclui um executável pré-compilado em `dist/pdf_etl_app.exe`. Para gerar um novo executável:

```bash
python build_exe.py
```

Ou manualmente com PyInstaller:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name pdf_etl_app src/main.py
```

## Funcionalidades

- **Extração inteligente**: Detecta automaticamente tabelas em PDFs usando múltiplos métodos
- **Transformação robusta**: Limpa e padroniza dados, detectando automaticamente tipos de colunas
- **Formatação profissional**: Gera planilhas Excel com formatação profissional (cabeçalhos, cores, etc.)
- **Interface amigável**: Seleção de arquivos intuitiva e feedback visual do processo

## Observações

- O aplicativo funciona melhor com PDFs que contêm tabelas bem estruturadas
- Para PDFs com layouts complexos, pode ser necessário ajustar parâmetros de extração

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias.

## Licença

Este projeto está licenciado sob a licença MIT - veja o arquivo LICENSE para detalhes.
