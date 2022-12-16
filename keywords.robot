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
${linha_abaixo}=    1
${PRINT}=    V_ 
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

add_invoice  #IT WORKS!!! --TEST PASS--
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
                        Log    ${nome}
                        IF    '${nome}' == 'cnpj'  # IRÁ SER O POSICIONADOR DA COLUNA
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}
                            ${conteudo}=    Convert To String    ${conteudo}
                            ${conteudo}=    Remove String    ${conteudo}    /
                            ...                                             .
                            ...                                             -
                        click_on_add_itens
                        input_cnpj   
                        sleep  1s                           
                        digitar_numero    ${conteudo}     
                        Send Keys To Input    {ENTER}    FALSE    0.2  0.0
                        press_selecionar
                        Repeat Keyword    5x    loops_for_VK_TAB                        
                        END
                        IF    '${nome}' == 'cond_cobranca'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero    ${conteudo}
                            Sleep    0.5s
                            Repeat Keyword    2x    loops_for_VK_TAB
                        END
                        IF    '${nome}' == 'centro_custo_emit'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero    ${conteudo}
                            Repeat Keyword    3x    loops_for_VK_TAB
                        END  
                        # #  IF    '${nome}' == 'vendedor'
                        # #      ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${${linha}+${n_itens}}    ${cont}  
                        # #      ${conteudo}=    Convert To String    ${conteudo}
                        # #      digitar_numero    ${conteudo}
                        # # Repeat Keyword    3x    loops_for_VK_TAB
                        IF  '${nome}' == 'cod_subgrupo'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero    ${conteudo}
                            Repeat Keyword    1x    loops_for_VK_TAB
                        END
                        IF    '${nome}' == 'cfop'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}  
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero    ${conteudo}
                         salvar
                         dados_adicionais_NF
                        END
                         ${cont}=    Set Variable    ${${cont}+${1}}
                    END                             
                ELSE
                ${linha}=    Set Variable    ${${linha}+${1}}                      
                END  #final_if_nome
        END             
    END  #final_for_linha

add_itens_of_nf  #IT WORKS!!!
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
        FOR    ${linha}    IN RANGE    2    ${linhas}
            FOR    ${coluna}    IN RANGE    1    ${colunas}
                Set Cell Format    ${linha}    ${coluna}    fmt=@
            END
              #final_coluna
            # IF    $result.Caracteres
            ${nomes}=    RPA.Tables.Get Table Row    ${tabela}    ${${linha}-${2}}
            ${cont}=    Set Variable    1
            ${nf_compare}=    RPA.Excel.Files.Get Cell Value    ${${linha}+${linha_abaixo}}    ${1}
            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont}                    
             # ${n_itens}=    Set Global Variable   ${n_itens}
             #${n_itens}=  RPA.EXCEL.FILES.Get Cell Value    ${linha}    B  #pega valor de numero de itens
                IF    '${conteudo}' != '${nf_compare}'
                    FOR    ${nome}    IN    @{nomes}
                        IF    '${nome}' == 'cod_servico'
                            click_on_add_itens                        
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero    ${conteudo}
                            # Repeat Keyword    5x    loops_for_VK_TAB  #CERTO              
                            Repeat Keyword    1x    loops_for_VK_TAB  #ERRADO {PARA TESTES} 
                            Send Keys To Input    ${PRINT}    FALSE    0.2  0.2
                        END
                        IF    '${nome}' == 'quantidade'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero    ${conteudo}
                            Repeat Keyword    1x    loops_for_VK_TAB
                            Send Keys To Input    ${PRINT}    FALSE    0.2  0.2                            
                        END
                        IF    '${nome}' == 'valor_unit_moeda'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}            
                            Repeat Keyword    13x    loops_for_VK_TAB                                 
                        END                                           
                        IF    '${nome}' == 'cod_tibut_servico'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}            
                            Repeat Keyword    1x    loops_for_VK_TAB                                 
                        END
                        IF    '${nome}' == 'mun_prest_serv'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}          
                            Repeat Keyword    1x    loops_for_VK_TAB                                 
                        END
                        IF    '${nome}' == 'tipo_serv'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}            
                            Repeat Keyword    2x    loops_for_VK_TAB                                 
                        END                        
                        IF    '${nome}' == 'centro_custo'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}            
                            # Repeat Keyword    3x    loops_for_VK_TAB                                             
                         salvar
                         Sleep    1s
                         sair
                         wait_until_appear_itens
                         sleep  1s
                         conferir
                         sleep  1s
                         go_to_NF
                         sleep  1s
                         Send Keys To Input    {VK_SHIFT}+{TAB}    FALSE    0.2  0.0
                         ${NF_dos_itens_linhaAbaixo}=    RPA.Excel.Files.Get Cell Value    ${${linha}+${linha_abaixo}}    A
                         sleep  1s
                         digitar_numero    ${NF_dos_itens_linhaAbaixo}
                         Send Keys To Input    {ENTER}    FALSE    0.2  0.0
                         go_to_itens_nf
                        END                                                      
                         ${cont}=    Set Variable    ${${cont}+${1}}
                    END
                ELSE
                    FOR    ${nome}    IN    @{nomes}     
                        # sleep    5s
                        IF    '${nome}' == 'cod_servico'
                         click_on_add_itens
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}
                            # Type Text    ${conteudo}
                            Repeat Keyword    1x    loops_for_VK_TAB
                        END    
                        IF    '${nome}' == 'quantidade'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero    ${conteudo}
                            Repeat Keyword    1x    loops_for_VK_TAB
                        END                
                        IF    '${nome}' == 'valor_unit_moeda'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}            
                            Repeat Keyword    13x    loops_for_VK_TAB                                 
                        END                                           
                        IF    '${nome}' == 'cod_tibut_servico'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}            
                            Repeat Keyword    1x    loops_for_VK_TAB                                 
                        END
                        IF    '${nome}' == 'mun_prest_serv'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}          
                            Repeat Keyword    1x    loops_for_VK_TAB                                 
                        END
                        IF    '${nome}' == 'tipo_serv'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}            
                            Repeat Keyword    2x    loops_for_VK_TAB                                 
                        END                        
                        IF    '${nome}' == 'centro_custo'
                            ${conteudo}=    RPA.Excel.Files.Get Cell Value    ${linha}    ${cont} 
                            ${conteudo}=    Convert To String    ${conteudo}
                            digitar_numero   ${conteudo}            
                            # Repeat Keyword    3x    loops_for_VK_TAB                                 
                         salvar               
                        END                                                                      
                        ${cont}=    Set Variable    ${${cont}+${1}}
                    END
                END       
        END  #for_colunas
    END  #for_linhas
add_obs_NF
    
conferir
    RPA.Desktop.Wait For Element    alias:conferir_nf    
    ${region}=    RPA.Desktop.Find Element    alias:conferir_nf
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
wait_until_appear_itens
    RPA.Desktop.Wait For Element    alias:item
go_to_NF
    RPA.Desktop.Wait For Element    alias:nota_fiscal    
    ${region}=    RPA.Desktop.Find Element    alias:nota_fiscal
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    RPA.Desktop.Click
    RPA.Desktop.Click
sair
    RPA.Desktop.Wait For Element    alias:sair    
    ${region}=    RPA.Desktop.Find Element    alias:sair
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
go_to_obs_NF
    RPA.Desktop.Wait For Element    alias:obs_nf    
    ${region}=    RPA.Desktop.Find Element    alias:obs_nf
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
go_to_itens_nf
    RPA.Desktop.Wait For Element    alias:aba_itens_nf    
    ${region}=    RPA.Desktop.Find Element    alias:aba_itens_nf
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    click_on_add_itens
input_cnpj
    RPA.Desktop.Wait For Element    alias:razao_social    
    ${region}=    RPA.Desktop.Find Element    alias:razao_social
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    {VK_TAB}    FALSE    0.2  0.0
    Send Keys To Input    {VK_UP}    FALSE    0.2   0.0
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
digitar_numero
    [Arguments]  ${conteudo}  
    sleep  0.5s
    RPA.Desktop.Type Text  ${conteudo}
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
    Send Keys To Input    {VK_DELETE}    FALSE    0.2   0.2
    Send Keys To Input    {VK_TAB}    FALSE    0.2   0.2
    Send Keys To Input    {ENTER}    FALSE    0.2   0.2
        


    
