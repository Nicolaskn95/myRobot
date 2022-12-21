*** Settings ***
DOCUMENTATION    Criar NF e enviar

Library           Collections
#Library           RPA.Browser.Selenium
Library           RPA.Excel.Files
Library           RPA.Dialogs
Library           RPA.Robocloud.Secrets
Library           RPA.Desktop
Library           RPA.Desktop.Windows
Library           RPA.JSON
Library           RPA.FileSystem
Library           RPA.Tables
Library           RPA.Windows
Library           RPA.Browser
Library           RPA.Robocorp.WorkItems
Library           RPA.Excel.Application
Library           RPA.PDF
Library           RPA.FileSystem
Library           String
Library           RPA.HTTP
Library           DateTime
Resource          keywords.robot
# Library           test_executable.py
*** Tasks ***
criar_nf
    ${arquivo}=    Coleta_Nome_do_Arquivo_Excel  
    sleep  1s
    go_to_invoice
    add_invoice    ${arquivo}
    go_to_itens_nf
    add_itens_of_nf    ${arquivo}
    go_to_obs_NF
    add_obs_NF    ${arquivo}
    

    

