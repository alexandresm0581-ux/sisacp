@echo off
REM Atalho para abrir o SISACP. Rode com duplo clique.
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo Criando ambiente virtual e instalando bibliotecas ^(so na primeira vez^)...
    python -m venv .venv
    call ".venv\Scripts\activate.bat"
    pip install -r requirements.txt
) else (
    call ".venv\Scripts\activate.bat"
)

if not exist "controle_peso.db" (
    echo.
    echo Banco de dados nao encontrado.
    echo Copie o Excel para esta pasta e rode:
    echo     python importar_excel.py "01 - Controle de Peso_v3.xlsm"
    pause
    exit /b 1
)

streamlit run app.py
