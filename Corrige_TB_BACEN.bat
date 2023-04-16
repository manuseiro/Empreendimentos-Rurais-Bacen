@echo off

REM Muda para o diretório onde o script está localizado
pushd "%~dp0"

REM Define o caminho da pasta onde estão os arquivos a serem copiados
set current_directory=%cd%\PreencheEmpreendimento

REM Define o caminho da pasta de destino para onde os arquivos serão copiados
set destination_folder=C:\SISTEMAS\COMPONENTES\PreencheEmpreendimento

REM Verifica se o Excel está em execução e pergunta ao usuário se ele deseja fechá-lo
tasklist /fi "imagename eq excel.exe" 2>NUL | find /I /N "excel.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo Excel está em execução.
    set /p close_excel=Você deseja fechar o Excel antes de continuar? (s/n)
    if /I "%close_excel%"=="s" (
        taskkill /f /im excel.exe >NUL
        echo Excel foi fechado.
    ) else (
        echo O Excel nao foi fechado. O script pode não funcionar corretamente.
    )
)

REM Copia os arquivos da pasta de origem para a pasta de destino
xcopy "%current_directory%\*" "%destination_folder%\" /E /Y

REM Registra os arquivos DLL
regsvr32 /s "%destination_folder%\SPR32X30.OCX"
regsvr32 /s "%destination_folder%\PreencheEmpreendimento.dll"

REM Exibe mensagem de sucesso
echo Arquivos copiados com sucesso!

REM Exibediretorios
echo O diretório atual é: %current_directory%
echo O diretório atual é: %destination_folder%

REM Pausa a execução do script para que o usuário possa visualizar a mensagem de sucesso
pause