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


*** Variables ***  
${contador_linha}=    0
${linha_acima}=    1
${index}=    0
${cont}=    0
${contador_coluna}=    0
*** Keywords ***

entrar_Notas_fiscais
    # Send Keys To Input    {VK_MENU}    FALSE    
    Send Keys To Input    {VK_MENU}    FALSE   0.0    0.0   
    Send Keys To Input    {VK_RIGHT}    FALSE  0.0    0.0   
    Send Keys To Input    {VK_RIGHT}    FALSE  0.0    0.0   
    Send Keys To Input    {VK_RIGHT}    FALSE  0.0    0.0   
    Send Keys To Input    {VK_RIGHT}    FALSE  0.0    0.0  
    Repeat Keyword    16x    loops_for_VK_DOWN
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

                                

add_nota_fiscal
    ${arquivo}=  Set Variable  C:${/}Users${/}nicolas${/}robots${/}CriarNFSatisFaturamento${/}DataInput.xlsx
    RPA.Excel.Files.Open Workbook    ${arquivo}
    ${table}=    Read Worksheet As Table    header=True
    ${testess}=    RPA.Tables.Get Table Row    ${table}    0
    ${sheets}=    List Worksheets
    FOR    ${sheet}    IN    @{sheets}
        ${tabela}=    Read Worksheet As Table    name=${sheet}    header=True
        RPA.Excel.Files.Set Active Worksheet    ${sheet}
        ${linhas2}    ${colunas2}=    Get Table Dimensions    ${tabela}
        ${linhas}=    Set Variable    ${${linhas2}+${2}}
        ${colunas}=    Set Variable    ${${colunas2}+${1}}
        FOR    ${linha}    IN RANGE    2    ${linhas}  #ELE IRÁ PERCORRER AS LINHAS 
            FOR    ${coluna}    IN RANGE    1    ${colunas}  # NÃO IRA PERCORRER A COLUNA, O ${CONT} QUE IRÁ PERCORRER, ELE SÓ PEGA O TOTAL DE COLUNAS             
                 Set Cell Format    ${linha}    ${coluna}    fmt=@
            END
                ${nomes}=    RPA.Tables.Get Table Row    ${tabela}    ${${linha}-${2}}
                ${cont}=    Set Variable    1
                ${nf_compare}=    RPA.Excel.Files.Get Cell Value    ${${linha}-${linha_acima}}    ${1}
                ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}
                #  ${n_itens}=    Set Global Variable   ${n_itens}
                IF    '${conteudo}' != '${nf_compare}'
                    FOR    ${nome}    IN    @{nomes}               
                        input_nf    ${nome}    ${linha}
                    END
                ELSE
                 ${linha}=    Set Variable    ${${linha}+${1}}    
                END                       
        END  #final_for_nome     
    END  #final_for_linha
input_nf
    [Arguments]    ${nome}    ${linha}
        # ${n_itens}=    Convert To Integer    ${n_itens}
        ${n_itens}=  RPA.EXCEL.FILES.Get Cell Value    ${linha}    B  #pega valor de numero de itens     
        IF    '${nome}' == 'cnpj'  # IRÁ SER O POSICIONADOR DA COLUNA
            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}
            ${conteudo}=    Convert To String    ${conteudo}
            ${conteudo}=    Remove String    ${conteudo}    /
            ...                                             .
            ...                                             -
            # click_on_add_itens
            # input_cnpj
            digitar_numero_nf    ${conteudo}
            # press_selecionar
            # Repeat Keyword    5x    loops_for_VK_TAB
        END  #final_cnpj E FAZ UM CONT++ PARA PASSAR P/ PROX. COLUNA
        # IF    '${nome}' == 'cond_cobranca'
        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
        #     ${conteudo}=    Convert To String    ${conteudo}
        #     digitar_numero_nf    ${conteudo}
        #     Repeat Keyword    2x    loops_for_VK_TAB
        # END
        # IF    '${nome}' == 'centro_custo_emit'
        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
        #     ${conteudo}=    Convert To String    ${conteudo}
        #     digitar_numero_nf    ${conteudo}
        #     Repeat Keyword    2x    loops_for_VK_TAB
        # END  
        # # IF    '${nome}' == 'vendedor'
        # #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${${linha}+${n_itens}}    ${cont}  
        # #     ${conteudo}=    Convert To String    ${conteudo}
        # #     digitar_numero_nf    ${conteudo}
        # Repeat Keyword    1x    loops_for_VK_TAB
        # IF  '${nome}' == 'cod_subgrupo'
        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
        #     ${conteudo}=    Convert To String    ${conteudo}
        #     digitar_numero_nf    ${conteudo}
        #     Repeat Keyword    1x    loops_for_VK_TAB
        # END
        # IF    '${nome}' == 'cfop'
        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
        #     ${conteudo}=    Convert To String    ${conteudo}
        #     digitar_numero_nf    ${conteudo}
        # END
        # salvar
        # dados_adicionais_NF
     ${cont}=    Set Variable    ${${cont}+${1}}    
criar_itens_da_nf
    [Arguments]    ${n_itens}  ${nome}  ${linha}   
        ${table}=    Read Worksheet As Table    header=True
        ${testess}=    RPA.Tables.Get Table Row    ${table}    0
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
                  #final_coluna
                # IF    $result.Caracteres
                ${nomes}=    RPA.Tables.Get Table Row    ${tabela}    ${${linha}-${2}}
                ${contador}=    Set Variable    1
                 # ${n_itens}=    Set Global Variable   ${n_itens}
                #${n_itens}=  RPA.EXCEL.FILES.Get Cell Value    ${linha}    B  #pega valor de numero de itens
                Log    ${n_itens}
                FOR    ${n_itens}    IN RANGE    0    ${n_itens}
                    FOR    ${nome}    IN    @{nomes}
                        IF    '${nome}' == 'cod_servico'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${n_itens}    ${contador} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero_nf    ${conteudo}
                            # Repeat Keyword    5x    loops_for_VK_TAB  #CERTO              
                            Repeat Keyword    1x    loops_for_VK_TAB  #ERRADO {PARA TESTES}                                        
                          #cod_servico
                        IF    '${nome}' == 'quantidade'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${n_itens}    ${contador} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero_nf    ${conteudo}
                            Repeat Keyword    2x    loops_for_VK_TAB
                        END
                        ${contador}=    Set Variable    ${${contador}+${1}}
                    END  #final_nome_in_nome
                END  #for_colunas
            END  #for_linhas
        END  #for_sheets
    END 

input_cnpj
    RPA.Desktop.Wait For Element    alias:razao_social    
    ${region}=    RPA.Desktop.Find Element    alias:razao_social
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    {VK_TAB}    FALSE    0.0  0.0
    Send Keys To Input    {VK_UP}    FALSE    0.0   0.0
press_selecionar
    RPA.Desktop.Wait For Element    alias:selecionar    
    ${region}=    RPA.Desktop.Find Element    alias:selecionar
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
click_on_add_itens
    RPA.Desktop.Wait For Element    alias:add    
    ${region}=    RPA.Desktop.Find Element    alias:add
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click 
aba_itens_nf
    RPA.Desktop.Wait For Element    alias:aba_itens_nf    
    ${region}=    RPA.Desktop.Find Element    alias:aba_itens_nf
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click 
digitar_numero_nf
    [Arguments]  ${conteudo}  
    RPA.Desktop.Type Text  ${conteudo}
    sleep  1s
salvar
    RPA.Desktop.Wait For Element    alias:salvar  
    ${region}=    RPA.Desktop.Find Element    alias:salvar
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
dados_adicionais_NF
    RPA.Desktop.Wait For Element    alias:dados_adicionais_nf  
    ${region}=    RPA.Desktop.Find Element    alias:dados_adicionais_nf
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    {VK_DELETE}    FALSE    0.0   0.2
    Send Keys To Input    {VK_TAB}    FALSE    0.0   0.2
    Send Keys To Input    {VK_ENTER}    FALSE    0.0   0.2
        


    
