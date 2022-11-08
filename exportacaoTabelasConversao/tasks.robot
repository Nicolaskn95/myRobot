
*** Settings ***
DOCUMENTATION     Exporta tabelas no processo de conversão, mais precisamente na 
...               na complementar. TOTAL: 22 tabelas => start  at 9:00 01/11
...                                                     finish at

Library           RPA.Excel.Application
Library           RPA.Excel.Files
Library           RPA.Desktop
Library           RPA.Dialogs
Library           RPA.Windows
Library           RPA.Browser
Library           RPA.Tables
Library           String
Library           RPA.Desktop.Windows
Library           RPA.HTTP
Library           RPA.Excel.Files 
Library           OperatingSystem
Resource          keyword.robot

*** Tasks ***
exportTable
    # ${pesquisa}    Set Variable    Exportação/Importação{SPACE}de{SPACE}Dados
    # Iniciar    ${pesquisa}
    # Limpar_Pesquisa
    # exportacao_importacao
chooseProcess
    estruturas_input
    processo_especifico_input
processDownload
        