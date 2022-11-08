*** Settings ***
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
Library    Collections

*** Variables ***
@{especificTable}    Valores{SPACE}da{SPACE}Pensão{SPACE}Alimentícia  Usuários{SPACE}da{SPACE}Linha{SPACE}de{SPACE}Transporte  Pessoa{SPACE}Física{SPACE}-{SPACE}Estrutura  Períodos{SPACE}de{SPACE}Férias
...                  Registro{SPACE}de{SPACE}Emprego{SPACE}-{SPACE}Estrutura

@{array_nameTable}    agban  cargo  funcao  Centro{SPACE}de{SPACE}Resultado{SPACE}-{SPACE}Estrutura  Dependentes{SPACE}-{SPACE}Estrutura
...        Histórico{SPACE}Salarial{SPACE}-{SPACE}Estrutura  Histórico{SPACE}de{SPACE}Funções{SPACE}-{SPACE}Estrutura
...        Histórico{SPACE}de{SPACE}Lotação{SPACE}-{SPACE}Estrutura  Afastamentos{SPACE}Férias  
...        Afastamento{SPACE}Geral  Afastamentos{SPACE}Rescisão{SPACE}-{SPACE}Estrutura  Associado{SPACE}ao{SPACE}Plano
...        Conveniados,{SPACE}Segurados,{SPACE}Usuários  Pensão  Histórico{SPACE}de{SPACE}Ocorrências{SPACE}do{SPACE}SEFIP  Conta{SPACE}Vinculada{SPACE}do{SPACE}    Fornecedor   
# 0_Agban
# 1_CARGO
# 2_FUNCAO
# 3_UNIORG
# 4_Pfisica
# 5_Regem
# 6_DEPENDENTES
# 7_Historico_Salarial
# 8_Historico_de_FUNCAO
# 9_Historico_de_Lotação
# 10_Afastamento_de_férias
# 11_Período_de_férias
# 12_Afastamento_geral
# 13_Afastamento_rescisão	
# 14_Associado_ao_Plano_Previdência_Privada
# 15_Conveniados_Segurados_Usuários
# 16_Pensão
# 17_Hist_Sefip
# 18_Conta_FGTS
# 19_Fornecedor

*** Keywords ***
Limpar_Pesquisa
    Send Keys To Input    ^a
    Send Keys To Input    {VK_BACK}    false
    #FOR    ${counter}    IN RANGE    1    24    1
    #    Send Keys To Input    {VK_BACK}    false
    #END
    RPA.Desktop.Wait For Element    alias:PesqLimpa
Botao_Click
    [Arguments]    ${alias}    ${botao}
    RPA.Desktop.Wait For Element    ${alias}    timeout=30
    ${region}=    RPA.Desktop.Find Element    ${alias}
    #RPA.Desktop.Move Mouse    ${region}
    IF    "${botao}" == "ESQUERDO"
        RPA.Desktop.Click    ${region}
    ELSE
        RPA.Desktop.Click    ${region}    right click
    END

estruturas_input
    # Send Keys To Input    {VK_TAB}    false    0.0  0.0

    FOR    ${counter}    IN RANGE    0    21
        RPA.Desktop.Wait For Element    alias:startIndex_denominacao
        ${region}=    RPA.Desktop.Find Element    alias:startIndex_denominacao
        RPA.Desktop.Move Mouse    ${region}
        RPA.Desktop.Click
        Send Keys To Input    {VK_TAB}    false    0.0  0.0
        Send Keys To Input    ${array_nameTable[${counter}]}    false
        Send Keys To Input    {VK_TAB}    false   0.3    0.0
        Send Keys To Input    {ENTER}     false   0.3    0.0
        RPA.Desktop.Wait For Element    alias:startIndex_denominacao
        ${region}=    RPA.Desktop.Find Element    alias:startIndex_denominacao
        RPA.Desktop.Move Mouse    ${region}
        RPA.Desktop.Click
        Send Keys To Input    {VK_TAB}    false   0.3    0.0        
        Send Keys To Input    {VK_TAB}    false   0.3    0.0
        Send Keys To Input    {VK_TAB}    false   0.3    0.0
        Send Keys To Input    {ENTER}    false   0.3    0.0
        Sleep    1.5s
        concluir_estrutura
        Sleep    1s
        volta_para_estrutura
    END
loops_VK_DOWN
    Send Keys To Input    {VK_DOWN}  FALSE    0.0  0.0
processo_especifico_input
            RPA.Desktop.Wait For Element    alias:startIndex_denominacao
            ${region}=    RPA.Desktop.Find Element    alias:startIndex_denominacao
            RPA.Desktop.Move Mouse    ${region}
            RPA.Desktop.Click
            valores_alimenticia  #TABLE#
            concluir_estrutura
            volta_para_estrutura
            usuarios_transporte  #TABLE#
            concluir_estrutura
            volta_para_estrutura
            periodo_ferias       #TABLE#
            concluir_estrutura
            volta_para_estrutura
            pessoa_fisica        #TABLE# 
            concluir_estrutura
            volta_para_estrutura
            registro_emprego     #TABLE#
            concluir_estrutura



registro_emprego
    RPA.Desktop.Wait For Element    alias:startIndex_denominacao
    ${region}=    RPA.Desktop.Find Element    alias:startIndex_denominacao
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    {VK_TAB}    false    0.0  0.0
    Send Keys To Input    ${especificTable[${4}]}    false
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}     false   0.3    0.0
    control_+_F
    Send Keys To Input    {ENTER}    FALSE    0.0  0.0
    Send Keys To Input    ${especificTable[${4}]}   FALSE
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}    FALSE   0.3    0.0
    RPA.Desktop.Wait For Element    alias:registro_emprego
    ${region}=    RPA.Desktop.Find Element    alias:registro_emprego
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Sleep    1s

pessoa_fisica
    RPA.Desktop.Wait For Element    alias:startIndex_denominacao
    ${region}=    RPA.Desktop.Find Element    alias:startIndex_denominacao
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    {VK_TAB}    false    0.0  0.0
    Send Keys To Input    ${especificTable[${2}]}    false
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}     false   0.3    0.0
    control_+_F
    Send Keys To Input    {ENTER}    FALSE    0.0  0.0
    Send Keys To Input    ${especificTable[${2}]}   FALSE
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}    FALSE   0.3    0.0
    RPA.Desktop.Wait For Element    alias:pessoa_fisica
    ${region}=    RPA.Desktop.Find Element    alias:pessoa_fisica
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    RPA.Desktop.Click
    Sleep    1s
periodo_ferias
    RPA.Desktop.Wait For Element    alias:startIndex_denominacao
    ${region}=    RPA.Desktop.Find Element    alias:startIndex_denominacao
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    {VK_TAB}    false    0.0  0.0
    Send Keys To Input    ${especificTable[${3}]}    false
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}     false   0.3    0.0
    control_+_F
    Send Keys To Input    {ENTER}    FALSE    0.0  0.0
    Send Keys To Input    ${especificTable[${3}]}   FALSE
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}    FALSE   0.3    0.0
    RPA.Desktop.Wait For Element    alias:periodo_ferias
    ${region}=    RPA.Desktop.Find Element    alias:periodo_ferias
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Sleep    1s

volta_para_estrutura
    Send Keys To Input    {ESC}    FALSE  0.0  0.0 
    RPA.Desktop.Wait For Element    alias:export_tables
    ${region}=    RPA.Desktop.Find Element    alias:export_tables
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
usuarios_transporte
    RPA.Desktop.Wait For Element    alias:startIndex_denominacao
    ${region}=    RPA.Desktop.Find Element    alias:startIndex_denominacao
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    {VK_TAB}    false    0.0  0.0
    Send Keys To Input    ${especificTable[${1}]}    false
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}     false   0.3    0.0
    control_+_F
    Send Keys To Input    {ENTER}    FALSE    0.0  0.0
    Send Keys To Input    ${especificTable[${1}]}   FALSE
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}    FALSE   0.3    0.0
    RPA.Desktop.Wait For Element    alias:usuarios_transporte
    ${region}=    RPA.Desktop.Find Element    alias:usuarios_transporte
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Sleep    1s
concluir_estrutura
    Send Keys To Input    {VK_TAB}    FALSE    0.0    0.0
    Send Keys To Input    {VK_TAB}    FALSE    0.0    0.0
    Send Keys To Input    {ENTER}    FALSE    0.0    0.0
    Sleep    2.5s
    Send Keys To Input    {VK_TAB}    FALSE    0.0    0.0
    Send Keys To Input    {ENTER}    FALSE    0.0    0.0   
valores_alimenticia
    Send Keys To Input    {VK_TAB}    false    0.0  0.0
    Send Keys To Input    ${especificTable[${0}]}    false
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}     false   0.3    0.0
    control_+_F
    Send Keys To Input    {ENTER}    FALSE    0.0  0.0
    Send Keys To Input    ${especificTable[${0}]}   FALSE
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {VK_TAB}    false   0.3    0.0
    Send Keys To Input    {ENTER}    FALSE   0.3    0.0
    RPA.Desktop.Wait For Element    alias:valores_alimenticia
    ${region}=    RPA.Desktop.Find Element    alias:valores_alimenticia
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Sleep    1.0s
control_+_F
     RPA.Desktop.Wait For Element    alias:config_chrome
     ${region}=    RPA.Desktop.Find Element    alias:config_chrome
     RPA.Desktop.Move Mouse    ${region}
     RPA.Desktop.Click         
     Repeat Keyword    12x    loops_VK_DOWN            
Iniciar
    [Arguments]    ${pesquisa}
    RPA.Desktop.Wait For Element    alias:PesqLocal
    ${region}=    RPA.Desktop.Find Element    alias:PesqLocal
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    ${pesquisa}    true
exportacao_importacao
    RPA.Desktop.Wait For Element    alias:exportacao_importacao
    ${region}=    RPA.Desktop.Find Element    alias:exportacao_importacao
    RPA.Desktop.Move Mouse    ${region}
    RPA.Desktop.Click
    Send Keys To Input    {VK_TAB}  TRUE  0.0  0.0
Processamento_Concluido

    IF    ${var1} == ${var1}
        Call Keyword
    ELSE
        
    END   

    FOR    ${counter}    IN RANGE    1    24000    1
        sleep    0.5s
        ${elements}=    RPA.Desktop.Find Elements    alias:Processamento.Concluido
        ${sair}    Set Variable    0
        FOR    ${element}    IN    @{elements}
            ${sair}    Set Variable    Sair
        END
        IF    len("${sair}") > 1
            Exit For Loop
        ELSE
            Pesquisar
        END
    END
    IF    len("${sair}") > 1
        Acoes
    ELSE
    
    END

# 0_Agban
# 1_CARGO
# 2_FUNCAO
# 3_UNIORG
# 4_Pfisica
# 5_Regem
# 6_DEPENDENTES
# 7_Historico_Salarial
# 8_Historico_de_FUNCAO
# 9_Historico_de_Lotação
# 10_Afastamento_de_férias
# 11_Período_de_férias
# 12_Afastamento_geral
# 13_Afastamento_rescisão	
# 14_Associado_ao_Plano_Previdência_Privada
# 15_Conveniados_Segurados_Usuários
# 16_Pensão
# 17_Usuários_da_Linha_de_Transporte
# 18_Valores_de_Pensão_Alimentícia
# 19_Hist_Sefip
# 20_Conta_FGTS
# 21_Fornecedor

