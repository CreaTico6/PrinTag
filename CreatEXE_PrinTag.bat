@echo off
title Gerador de Executavel PrinTag v1.4.5
echo -------------------------------------------------------
echo A verificar instalacao do PyInstaller...
python -m pip install pyinstaller --quiet

echo -------------------------------------------------------
echo A criar o executavel... (Isto pode demorar um pouco)
echo -------------------------------------------------------

:: Usamos "python -m" para evitar erros de PATH
python -m PyInstaller --noconsole --onefile --clean PrinTag.pyw

echo -------------------------------------------------------
echo Processo concluido! 
echo O teu executavel esta na pasta "dist".
echo -------------------------------------------------------
pause