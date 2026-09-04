@echo off
REM Script de configuração - Instala as dependências necessárias

echo.
echo ============================================================
echo Instalando dependencias do projeto...
echo ============================================================
echo.

python -m pip install --upgrade pip --user
python -m pip install -r requirements.txt --user

echo.
echo ============================================================
echo Instalacao concluida!
echo.
echo Agora voce pode executar o script com:
echo   run.bat
echo ============================================================
echo.
pause
