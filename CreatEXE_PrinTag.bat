@echo off
title Gerador de Executavel PrinTag v1.6.4

:: O truque de mestre para lidar com pastas de rede (UNC)
pushd "%~dp0"

echo -------------------------------------------------------
echo A verificar dependencias (PyInstaller, Pandas, Openpyxl)...
python -m pip install pyinstaller pandas openpyxl --quiet

echo -------------------------------------------------------
echo A forjar o executavel... (Isto pode demorar um pouco)
echo -------------------------------------------------------

:: Compilacao com supressao de consola e ficheiro unico
python -m PyInstaller --noconsole --onefile --clean PrinTag.pyw

echo -------------------------------------------------------
echo A extrair o executavel e a obliterar os residuos temporarios...
echo -------------------------------------------------------

:: Resgatar o executavel para a pasta raiz
move /y "dist\PrinTag.exe" . >nul

:: Erradicar as provas do crime (pastas e ficheiros de configuracao da build)
rmdir /s /q build
rmdir /s /q dist
del /q PrinTag.spec

echo -------------------------------------------------------
echo Processo concluido!
echo O teu executavel PrinTag.exe aguarda-te.
echo -------------------------------------------------------

:: Remove o disfarce da unidade de rede
popd
pause