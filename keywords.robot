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
@{numero_itens}=    5    8    2    4    
@{quantidade_de_itens}=    1    2    3    4    5    1	2	3	4	5	6	7	8    1    2    1    2    3    4    
${contador_linha}=    0
${cont}=    0
${index}=    0
# @{test_dictionary}=    example= 1    example1= 2    example= 3




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

aba_notas_fiscais2
    
    FOR    ${element}    IN    @{numero_itens}
        Log    ${element}
        Repeat Keyword    ${element}    repeat_itens  ${element}    ${cont}
            ${cont}=    Set Variable    ${${cont}+${1}}
    END                                            
repeat_itens
    [Arguments]   ${element}    ${cont}
    ${total_element}=  Get Length    ${quantidade_de_itens}
    
    # FOR    ${counter}    IN RANGE    0   ${total_element}
        # IF    ${quantidade_de_itens} >= ${0}
            # Call Keyword
        # ELSE
            
        # END
        # Log    ${counter}
        # ${quantidade_de_itens}=   Set Variable   ${${index}+${element}}
    # END
    # FOR    ${counter}    IN RANGE    0    {END}

    #     Log    @{LIST_2}
        
    # END

teste_notas_nf
    # [Arguments]    ${arquivo}
    ${arquivo}=    ${CURDIR}${/}DataInput.xlsx
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
        FOR    ${linha}    IN RANGE    2    ${linhas}
            FOR    ${coluna}    IN RANGE    1    ${colunas}
                Set Cell Format    ${linha}    ${coluna}    fmt=@
            END
            # IF    $result.Caracteres
                ${nomes}=    RPA.Tables.Get Table Row    ${tabela}    ${${linha}-${2}}
                ${cont}=    Set Variable    1
                ${n_itens}=   Set Variable    1
                 # ${n_itens}=    Set Global Variable   ${n_itens}
                FOR    ${nome}    IN    @{nomes}
                # ${n_itens}=    Convert To Integer    ${n_itens}  # .
                    ${n_itens}=  RPA.EXCEL.FILES.Get Cell Value    ${linha}    B  #pega valor de numero de itens                    
                    Log    ${n_itens}
                    IF    '${nome}' == 'cnpj'
                        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}
                        ${conteudo}=    Convert To String    ${conteudo}
                        ${conteudo}=    Remove String    ${conteudo}    /
                        ...                                             .
                        ...                                             -
                        # input_cnpj
                        digitar_numero_nf    ${conteudo}
                        Repeat Keyword    2x    loops_for_VK_TAB
                    END  #final_cnpj
                        
                    IF    '${nome}' == 'cond_cobranca'
                        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        ${conteudo}=    Convert To String    ${conteudo}
                        digitar_numero_nf    ${conteudo}
                        Repeat Keyword    2x    loops_for_VK_TAB
                    END  #final_cond_cobranca
                    # salvar
                    # add_itens_nf
                    # click_on_add_itens
                    log  ${n_itens}  
                    Log    ${nome}  #
                    Log    ${linha}
                    log  ${cont}  #09171140000105        506              
                    Repeat Keyword    ${n_itens}    itens_da_nf  ${n_itens}  ${nome}  ${linha}  ${cont}

                ${cont}=    Set Variable    ${${cont}+${1}}
                END  #final_for_nome     
        END  #final_for_linha  
    END  #final_for_sheet


itens_da_nf
    [Arguments]    ${n_itens}  ${nome}  ${linha}  ${cont}
        Log    ${n_itens}
        IF    '${nome}' == 'cod_servico'
        ##Alterar variável linha ########
        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${${linha}+${contador_linha}}    ${cont}  
        ${conteudo}=    Convert To String    ${conteudo}
        digitar_numero_nf    ${conteudo}
        Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
        END  #final_cod_sevico
    
        IF    '${nome}' == 'quantidade'
        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${${linha}+${contador_linha}}    ${cont}  
        ${conteudo}=    Convert To String    ${conteudo}
        digitar_numero_nf    ${conteudo}
        Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
        END  #final_quantidade
        ${contador_linha}  Set Variable  ${${contador_linha}+${1}}    
        ${linha}=    Set Variable    ${${linha}+${1}} 
input_cnpj
    RPA.Desktop.Wait For Element    alias:razao_social    
    ${region}=    RPA.Desktop.Find Element    alias:razao_social
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    {VK_TAB}    FALSE    0.0  0.0
    Send Keys To Input    {VK_UP}    FALSE    0.0   0.0

            
aba_notas_fiscais
    [Arguments]    ${arquivo}

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
        FOR    ${linha}    IN RANGE    2    ${linhas}
            FOR    ${coluna}    IN RANGE    1    ${colunas}
                Set Cell Format    ${linha}    ${coluna}    fmt=@
            END
            # IF    $result.Caracteres
                ${nomes}=    RPA.Tables.Get Table Row    ${tabela}    ${${linha}-${2}}
                ${cont}=    Set Variable    1
                FOR    ${nome}    IN    @{nomes}
                    ${n_itens}=  RPA.EXCEL.FILES.Get Cell Value    ${${cont}+${1}}    B  #pega valor de numero de itens
                    IF    '${nome}' == 'cnpj'
                        ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}
                        ${conteudo}=    Convert To String    ${conteudo}
                        ${conteudo}=    Remove String    ${conteudo}    /
                        ...                                             .
                        ...                                             -
                        RPA.Excel.Files.Set Cell Value    ${linha}    ${cont}    ${conteudo}
                    END
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
                    # IF    '${nome}' == 'vendedor'
                    #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                    #     ${conteudo}=    Convert To String    ${conteudo}
                    #     digitar_numero_nf    ${conteudo}
                    #     Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
                    # END
                    # IF    '${nome}' == 'cod_subgrupo'
                    #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                    #     ${conteudo}=    Convert To String    ${conteudo}
                    #     digitar_numero_nf    ${conteudo}
                    #     Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
                    # END
                    # IF    '${nome}' == 'cfop'
                    #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                    #     ${conteudo}=    Convert To String    ${conteudo}
                    #     digitar_numero_nf    ${conteudo}
                    # END
                    # salvar
                    # add_itens_nf
                    # click_on_add_itens
                    
                    # ITENS_DA_NOTA_FISCAL
                    # FOR    ${element}    IN    @{LIST}
                    #     IF    ${var1} == ${var1}
                            
                    #     END
                    #     Log    ${element}
                        
                    # END
                    # IF    ${var1} == ${var1}
                    #     Call Keyword
                    # ELSE
                    # FOR    ${cont_linha}    IN RANGE    1    ${n_itens}                      
                    #     IF    '${nome}' == 'cod_servico'
                    #     ##Alterar variável linha ########
                    #         ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                    #         ${conteudo}=    Convert To String    ${conteudo}
                    #         digitar_numero_nf    ${conteudo}
                    #         Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
                    #     END
                    #     IF    '${nome}' == 'quantidade'
                    #         ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                    #         ${conteudo}=    Convert To String    ${conteudo}
                    #         digitar_numero_nf    ${conteudo}
                    #         Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
                    #     END
                        # IF    '${nome}' == 'valor_unit_moeda'
                        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        #     ${conteudo}=    Convert To String    ${conteudo}
                        #     digitar_numero_nf    ${conteudo}
                        #     Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
                        # END
                        # IF    '${nome}' == 'valor_merc_terce'
                        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        #     ${conteudo}=    Convert To String    ${conteudo}
                        #     digitar_numero_nf    ${conteudo}
                        #     Send Keys To Input    {VK_TAB}    FALSE  0.0  0.0
                        # END 
                        # IF    '${nome}' == '%_desc_acresc'
                        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        #     ${conteudo}=    Convert To String    ${conteudo}
                        #     digitar_numero_nf    ${conteudo}
                        #     Repeat Keyword    2x    loops_for_VK_TAB
                        # END 
                        # IF    '${nome}' == 'valor_base_ret_inss'
                        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        #     ${conteudo}=    Convert To String    ${conteudo}
                        #     digitar_numero_nf    ${conteudo}
                        #     Repeat Keyword    2x    loops_for_VK_TAB
                        # END 
                        # IF    '${nome}' == 'val_merc_propria'
                        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        #     ${conteudo}=    Convert To String    ${conteudo}
                        #     digitar_numero_nf    ${conteudo}
                        #     Repeat Keyword    2x    loops_for_VK_TAB
                        # END
                        # IF    '${nome}' == 'cod_tribut'
                        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        #     ${conteudo}=    Convert To String    ${conteudo}
                        #     digitar_numero_nf    ${conteudo}
                        #     Send Keys To Input    {VK_TAB}    FALSE    0.0  0.0
                        # END
                        # IF    '${nome}' == 'mun_prest_serv' 
                        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        #     ${conteudo}=    Convert To String    ${conteudo}
                        #     digitar_numero_nf    ${conteudo}
                        #     Send Keys To Input    {VK_TAB}    FALSE    0.0  0.0
                        # END    
                        # IF    '${nome}' == 'tipo_serv'
                        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        #     ${conteudo}=    Convert To String    ${conteudo}
                        #     digitar_numero_nf    ${conteudo}
                        #     Repeat Keyword    2x    loops_for_VK_TAB
                        # END    
                        # IF    '${nome}' == 'centro_custo'
                        #     ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                        #     ${conteudo}=    Convert To String    ${conteudo}
                        #     digitar_numero_nf    ${conteudo}
                        # END       
                                                                                                                                                                                                                         
                # ${cont_linha}=    Set Variable    ${${cont}+${1}}

                END
             ${cont}=    Set Variable    ${${cont}+${1}} 
        END
    END      
    # Save Workbook
    # Close Workbook
    Add heading    Rotina Finalizada!
    Run dialog

click_on_add_itens
    RPA.Desktop.Wait For Element    alias:add    
    ${region}=    RPA.Desktop.Find Element    alias:add
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click 
add_itens_nf
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
    
