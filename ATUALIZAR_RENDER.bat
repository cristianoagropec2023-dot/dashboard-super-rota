@echo off
setlocal
cd /d "%~dp0"

title SUPERROTA - ATUALIZAR RENDER

echo.
echo ================================================
echo       SUPERROTA - PUBLICAR ATUALIZACAO
echo ================================================
echo.

git --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Git nao foi encontrado no computador.
    pause
    exit /b 1
)

echo [1/4] Verificando alteracoes...
echo.
git status --short
echo.

git status --porcelain > "%TEMP%\superrota_status.txt"
for %%A in ("%TEMP%\superrota_status.txt") do if %%~zA==0 (
    echo Nenhuma alteracao encontrada para publicar.
    echo.
    pause
    exit /b 0
)

echo ================================================
echo ARQUIVOS QUE SERAO PUBLICADOS
echo ================================================
type "%TEMP%\superrota_status.txt"
echo.
echo ================================================
echo.

choice /C SN /N /M "Deseja publicar estas alteracoes? [S/N]: "
if errorlevel 2 (
    echo.
    echo Publicacao cancelada.
    del "%TEMP%\superrota_status.txt" >nul 2>&1
    pause
    exit /b 0
)

echo.
echo [2/4] Preparando arquivos...
git add .
if errorlevel 1 (
    echo.
    echo [ERRO] Nao foi possivel preparar os arquivos.
    pause
    exit /b 1
)

echo.
echo [3/4] Criando commit...
set "COMMIT_MSG=Atualizacao SuperRota %date% %time%"
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo.
    echo [ERRO] Nao foi possivel criar o commit.
    echo Verifique se existem alteracoes validas para publicar.
    pause
    exit /b 1
)

echo.
echo [4/4] Enviando para o GitHub...
git push origin main
if errorlevel 1 (
    echo.
    echo [ERRO] O envio para o GitHub falhou.
    echo.
    echo Nao tente fazer outro push automaticamente.
    echo Verifique a mensagem acima e entre em contato para resolver.
    pause
    exit /b 1
)

echo.
echo ================================================
echo       PUBLICACAO ENVIADA COM SUCESSO!
echo ================================================
echo.
echo O GitHub recebeu a nova versao.
echo O Render fara o deploy automaticamente.
echo.
echo Aguarde o Render concluir o deploy antes de conferir
echo a versao publicada do Dashboard.
echo.

del "%TEMP%\superrota_status.txt" >nul 2>&1
pause
exit /b 0
