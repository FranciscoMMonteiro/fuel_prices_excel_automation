@echo off

REM Save the current directory
set CURRENT_DIR=%~dp0

REM Activate the Conda environment
call activate pyQuant_3_11

REM Run the Python script
python "%CURRENT_DIR%\atualiza_preco_prod.py"


