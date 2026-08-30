@echo off
setlocal
cd /d "%~dp0"

title SUPERROTA - ATUALIZAR RENDER

echo.
echo ================================================
echo        SUPERROTA - PUBLICAR ATUALIZACAO
echo ================================================
echo.

git --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Git nao foi encontrado no computador.
    echo Instale o Git e tente novamente.
    pause
    exit /b 1
)

echo [1/3] Verificando alteracoes...
echo.
git status --short

git diff --quiet
set DIFF_EXIT=%errorlevel%
git diff --cached --quiet
set CACHED_EXIT=%errorlevel%

if "%DIFF_EXIT%"=="0" if "%CACHED_EXIT%"=="0" (
    echo.
    echo Nenhuma alteracao encontrada para publicar.
    echo O Render continua com a versao atual.
    echo.
    pause
    exit /b 0
)

echo.
echo [2/3] Enviando alteracoes para o GitHub...
git add .
if errorlevel 1 (
    echo.
    echo [ERRO] Nao foi possivel preparar os arquivos.
    pause
    exit /b 1
)

set "COMMIT_MSG=Atualizacao SuperRota %date% %time%"
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo.
    echo [ERRO] Nao foi possivel criar o commit.
    pause
    exit /b 1
)

git push origin main
if errorlevel 1 (
    echo.
    echo [ERRO] O envio para o GitHub falhou.
    echo Verifique sua conexao e o acesso ao repositorio.
    pause
    exit /b 1
)

echo.
echo [3/3] PUBLICACAO ENVIADA!
echo.
echo O GitHub recebeu a nova versao.
echo Se o Render estiver conectado ao repositorio, ele iniciara o deploy automaticamente.
echo.
echo Aguarde alguns instantes e abra o Dashboard no Render.
echo.
pause
exit /b 0
