import os
import subprocess
import shutil
import glob

def build_executable():
    """Constrói o executável usando PyInstaller"""
    print("Iniciando build do executável...")

    # Remover arquivos temporários do Excel
    print("Removendo arquivos temporários do Excel...")
    for temp_file in glob.glob("Arquivo/~$*.xlsx") + glob.glob("Arquivo\\~$*.xlsx"):
        try:
            print(f"Removendo {temp_file}...")
            os.remove(temp_file)
        except Exception as e:
            print(f"Erro ao remover {temp_file}: {str(e)}")

    # Verificar se o diretório dist existe e removê-lo
    if os.path.exists("dist"):
        print("Removendo diretório dist existente...")
        shutil.rmtree("dist")

    # Verificar se o diretório build existe e removê-lo
    if os.path.exists("build"):
        print("Removendo diretório build existente...")
        shutil.rmtree("build")

    # Verificar se o arquivo spec existe e removê-lo
    if os.path.exists("pdf_etl_app.spec"):
        print("Removendo arquivo spec existente...")
        os.remove("pdf_etl_app.spec")

    # Construir o executável
    print("Executando PyInstaller...")
    cmd = ["pyinstaller",
           "--onefile",  # Criar um único arquivo executável
           "--console",  # Mostrar console ao executar (para depuração)
           "--name", "pdf_etl_app",  # Nome do executável
           "--hidden-import", "PySimpleGUI",  # Garantir que o PySimpleGUI seja incluído
           "--hidden-import", "camelot",  # Garantir que o camelot seja incluído
           "--hidden-import", "pdfplumber",  # Garantir que o pdfplumber seja incluído
           "--hidden-import", "pandas",  # Garantir que o pandas seja incluído
           "--hidden-import", "openpyxl",  # Garantir que o openpyxl seja incluído
           "--hidden-import", "re",  # Garantir que o re seja incluído
           "--hidden-import", "threading",  # Garantir que o threading seja incluído
           "--hidden-import", "time",  # Garantir que o time seja incluído
           "src/main.py"]  # Script principal (versão com interface gráfica)

    result = subprocess.run(cmd, capture_output=True, text=True)

    # Verificar se o build foi bem-sucedido
    if result.returncode == 0:
        print("Build concluído com sucesso!")
        print(f"Executável gerado em: {os.path.abspath('dist/pdf_etl_app.exe')}")
        return True
    else:
        print("Erro durante o build:")
        print(result.stdout)
        print(result.stderr)
        return False

if __name__ == "__main__":
    build_executable()
