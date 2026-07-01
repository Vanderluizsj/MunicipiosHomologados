@echo off
:: Define a codificação do terminal para UTF-8 (corrige acentuações)
chcp 65001 > nul

title Publicador de Executável Único e Compacto - .NET

echo =======================================================================
echo        INICIANDO PUBLICAÇÃO DO EXECUTÁVEL COMPACTO (SINGLE FILE)
echo =======================================================================
echo.

:: 1. Define o caminho onde o executável final será gerado
set OUT_DIR=.\bin\Release\Publish_Compacto

:: 2. Executa o comando de publicação avançada do .NET
echo ⚙ Compilando e gerando arquivo único...
echo.

dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=true -o %OUT_DIR%

if %errorlevel% neq 0 (
    echo.
    echo ❌ Falha crítica ao publicar o projeto.
    goto ERRO
)

echo.
echo =======================================================================
echo 🎉 EXECUTÁVEL GERADO COM SUCESSO!
echo =======================================================================
echo.
echo 📁 O seu executável está na pasta: %OUT_DIR%
echo 📌 Nome do arquivo: MunicipiosHomologados.ConsoleApp.exe
echo.
echo O Windows Explorer vai abrir a pasta automaticamente agora...
echo.

:: Abre a pasta do executável final no Windows
explorer %OUT_DIR%
goto FIM

:ERRO
echo.
echo ⚠ O processo de publicação foi abortado devido a erros.
echo.

:FIM
pause