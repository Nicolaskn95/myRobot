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

*** Tasks ***
criar_nf
    # ${arquivo}=    Coleta_Nome_do_Arquivo_Excel  
    # sleep  1s
    # entrar_Notas_fiscais
    add_invoice
    go_to_itens_nf
    add_itens_of_nf
    # aba_notas_fiscais    ${arquivo}
    # aba_notas_fiscais2

    

