*** Settings ***
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

*** Keywords ***
Notas_fiscais
    # Send Keys To Input    {VK_MENU}    FALSE    
    Send Keys To Input    {VK_MENU}    FALSE   0.0    0.0   
    Send Keys To Input    {VK_RIGHT}    FALSE  0.0    0.0   
    Send Keys To Input    {VK_RIGHT}    FALSE  0.0    0.0   
    Send Keys To Input    {VK_RIGHT}    FALSE  0.0    0.0   
    Send Keys To Input    {VK_RIGHT}    FALSE  0.0    0.0  
    Repeat Keyword    14x    loops_for_VK_DOWN
    Send Keys To Input    {ENTER}    FALSE
    RPA.Desktop.Wait For Element    alias:add    
    ${region}=    RPA.Desktop.Find Element    alias:add
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click  

loops_for_VK_DOWN
    Send Keys To Input    {VK_DOWN}    FALSE  0.0  0.0
loops_for_VK_TAB
    Send Keys To Input    {VK_TAB}    FALSE    0.0  0.0

Coleta_Nome_do_Arquivo_Excel
    Add heading    Upload Excel File
    Add file input
    ...    label=Upload the Excel file
    ...    name=fileupload
    ...    file_type=Excel files (*.xls;*.xlsx)
    ...    destination=${OUTPUT_DIR}
    ${result}=    Run dialog
    [Return]    ${result.fileupload}[0]

Procurar_e_digitar_as_NFs
    [Arguments]   ${arquivo}
    #Botao.Click    alias:Id.Global    ESQUERDO
    # ${arquivo}=   ${CURDIR}${/}output${/}Faturamento_2022_.xlsx
    RPA.Excel.Files.Open Workbook     ${arquivo}
    ${sheets}=    List Worksheets
    FOR    ${sheet}    IN    @{sheets}
        ${tabela}=    Read Worksheet As Table    name=${sheet}    header=True
        RPA.Excel.Files.Set Active Worksheet    ${sheet}
        ${linhas2}    ${colunas2}=    Get Table Dimensions    ${tabela}
        ${linhas}=    Set Variable    ${${linhas2}+${2}}
        ${colunas}=    Set Variable    ${${colunas2}+${1}}
        FOR    ${linha}    IN RANGE    2    ${linhas}
            FOR    ${coluna}    IN RANGE    1    ${colunas}
                Set Cell Format    ${linha}    ${coluna}    fmt=@
            END
                ${nomes}=    RPA.Tables.Get Table Row    ${tabela}    ${${linha}-${2}}
                ${cont}=    Set Variable    1
                FOR    ${nome}    IN    @{nomes}
                    IF    '${nome}' == 'razao_social'
                        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        ${conteudo}=    Convert To String    ${conteudo}
                        digitar_numero_nf    ${conteudo}
                        Repeat Keyword    5x    loops_for_VK_TAB
                    END
                        IF    '${nome}' == 'cond_cobranca'
                        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        ${conteudo}=    Convert To String    ${conteudo}
                        digitar_numero_nf    ${conteudo}
                        Repeat Keyword    2x    loops_for_VK_TAB
                    END
                        IF    '${nome}' == 'centro_custo_emit'
                        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        ${conteudo}=    Convert To String    ${conteudo}
                        digitar_numero_nf    ${conteudo}
                        Repeat Keyword    2x    loops_for_VK_TAB
                    END
                        IF    '${nome}' == 'vendedor'
                        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        ${conteudo}=    Convert To String    ${conteudo}
                        digitar_numero_nf    ${conteudo}
                        Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
                    END
                        IF    '${nome}' == 'cod_subgrupo'
                        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        ${conteudo}=    Convert To String    ${conteudo}
                        digitar_numero_nf    ${conteudo}
                        Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
                    END                    
              ${cont}=    Set Variable    ${${cont}+${1}}
            END
        END
    END
    Save Workbook
    Close Workbook
    Add heading    Rotina Finalizada!
    Run dialog

digitar_numero_nf
    [Arguments]  ${conteudo}   
    RPA.Desktop.Type Text  ${conteudo}
    sleep  1s
    
