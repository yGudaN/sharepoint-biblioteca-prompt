@echo off
REM =========================================================================
REM Iniciador do Assistente de Implantacao da Biblioteca de Prompts
REM Duplo-clique para abrir o wizard grafico.
REM =========================================================================

setlocal
cd /d "%~dp0"

REM Tenta o PowerShell 7 (pwsh). Se nao existir, avisa o usuario.
where pwsh >nul 2>nul
if errorlevel 1 (
    echo.
    echo [ERRO] PowerShell 7 nao encontrado.
    echo.
    echo Instale em: https://aka.ms/powershell-release?tag=stable
    echo.
    echo Depois execute este arquivo novamente.
    echo.
    pause
    exit /b 1
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0Assistente.ps1"
if errorlevel 1 (
    echo.
    echo O assistente encerrou com erro. Pressione qualquer tecla para fechar.
    pause >nul
)
endlocal
