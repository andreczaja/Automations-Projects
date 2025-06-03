Attribute VB_Name = "PASSO0_GERAL_EXTRACAO_CARTOLAS"
Option Explicit

    Public driver As New EdgeDriver
    Public elementoInput As WebElement
    Public elementosClasse As WebElements
    Public linha, linha_fim, i, i2, bci_buscas_realizadas_contas_clp As Integer
    Public banco, banco_anterior, cuenta, usuario, senha, sociedad, ByType As String
    
    ' declaração de strings por elemento BCI
    Public bci_elemento_banco_en_linea, bci_elemento_login_usuario, bci_elemento_login_senha, bci_elemento_botao_login, bci_elemento_box_grupo_electrolux, _
    bci_elemento_opcoes_menu_geral, bci_elemento_sessao_cuentas, bci_elemento_sessao_cuentas_corrientes, bci_elemento_opcoes_abertas_menu_geral, bci_elemento_lista_de_sociedades_banco, _
    bci_elemento_conta_ativa, bci_elemento_numero_cartola_atual, bci_elemento_ultima_data_encontrada_extrato, _
    bci_elemento_botao_download, bci_elemento_download_excel, bci_elemento_consulta_de_movimientos, bci_elemento_download_excel_aba_movimientos_anterior, _
    bci_elemento_conta_ativa_aba_movimientos_anterior, bci_elemento_lista_de_sociedades_banco_aba_movimientos_anterior As String
    
    ' declaração de strings por elemento SANTANDER
    Public santander_elemento_login_usuario, santander_elemento_login_senha, santander_elemento_botao_login, santander_elemento_cuentas_corrientes, santander_elemento_cartola_historica, _
    santander_elemento_list_box_contas, santander_elemento_list_box_meses, santander_elemento_list_box_anos, santander_elemento_buscar_cartola, santander_elemento_numero_cartola_atual, santander_elemento_ultima_data_encontrada_cartola, _
    santander_elemento_botao_download, santander_elemento_download_em_excel, santander_elemento_mensagem_sem_movimentos, santander_elemento_aceptar_mensagem_sem_movimento, _
    bci_elemento_ultima_data_encontrada_extrato_aba_movimientos_anterior As String
    
    ' declaração de strings por elemento SCOTIABANK
    Public scotiabank_elemento_login_rut, scotiabank_elemento_login_usuario, scotiabank_elemento_login_senha, scotiabank_elemento_botao_login, scotiabank_elemento_cuentas, _
    scotiabank_elemento_cartolas, scotiabank_elemento_ultima_data_encontrada_cartola, scotiabank_elemento_download_excel As String
    
    ' declaração de strings por elemento BANCO DE CHILE
    Public banco_chile_elemento_login_usuario, banco_chile_elemento_login_senha, banco_chile_elemento_botao_login, banco_chile_elemento_atalho_saldos_movimentos, banco_chile_elemento_ultimo_movimento_aba_saldos_movimentos, _
    banco_chile_elemento_cartola_historica, banco_chile_conta_ativa, elemento_abrir_list_box_contas, elemento_busca_contas, elemento_primeira_conta_da_busca, banco_chile_elemento_ultima_data_encontrada_cartola, _
    banco_chile_elemento_botao_download, banco_chile_elemento_download_em_excel, banco_chile_elemento_download_em_pdf As String
    
    Public bci_selecionada_aba_iquique_importadora, arquivos_extraidos As Boolean
    Public fecha_pagos As Date
    Public by As New by
    Public aba_acesso_bancos, aba_contas, aba_saldos, aba_pagamentos, aba_numero_cartola_banco_chile As Worksheet
    Public tabela_acesso_bancos As ListObject
    Public tabela_contas As ListObject
    Public tabela_saldos As ListObject
    Public tabela_pagamentos As ListObject
    Public tabela_numero_cartola_banco_chile As ListObject

Sub extracao_cartolas_()


' ELEMENTOS PORTAL BCI

bci_elemento_banco_en_linea = "/html/body/header/nav/div/button"
bci_elemento_login_usuario = "rut_aux"
bci_elemento_login_senha = "clave_aux"
bci_elemento_botao_login = "/html/body/section/div/div/div/div[1]/div[2]/div[2]/form/div[4]/div/button"
bci_elemento_box_grupo_electrolux = "box-grupo"
bci_elemento_opcoes_menu_geral = "/html/body/app-root/app-private-container/mat-toolbar/div/div[1]/div[1]/i"
bci_elemento_sessao_cuentas = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav/app-menu/mat-list[1]/mat-list-item[2]/div"
bci_elemento_sessao_cuentas_corrientes = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav/app-menu/mat-list[1]/mat-list[1]/div[1]/mat-list-item/div"
bci_elemento_opcoes_abertas_menu_geral = "mat-list-item"
bci_elemento_lista_de_sociedades_banco = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-historical-balances/div/div[1]/div[2]/bci-select-search/div/mat-form-field"
bci_elemento_lista_de_sociedades_banco_aba_movimientos_anterior = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-current-account/div/mat-sidenav-container/mat-sidenav-content/div/app-balances-rut-account/div/div[2]/bci-select-search/div/mat-form-field/div/div[1]/div"
bci_elemento_conta_ativa = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-historical-balances/div/div[1]/div[3]/bci-select-search-cuentas/div/mat-form-field/div/div[1]/div/mat-select"
bci_elemento_numero_cartola_atual = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-historical-balances/div/section[2]/div/mat-table/mat-row[1]/mat-cell[1]"
bci_elemento_ultima_data_encontrada_extrato = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-historical-balances/div/section[2]/div/mat-table/mat-row[1]/mat-cell[2]/span"
bci_elemento_botao_download = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-historical-balances/div/section[2]/div/mat-table/mat-row[1]/mat-cell[3]/span/i"
bci_elemento_download_excel = "/html/body/div[3]/div[2]/div/div/button[2]"
bci_elemento_consulta_de_movimientos = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-current-account/div/mat-sidenav-container/mat-sidenav-content/div/div[2]/mat-tab-group/mat-tab-header/div[2]/div/div/div[2]"
bci_elemento_download_excel_aba_movimientos_anterior = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-current-account/div/mat-sidenav-container/mat-sidenav-content/div/div[1]/button[2]"
bci_elemento_ultima_data_encontrada_extrato_aba_movimientos_anterior = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-current-account/div/mat-sidenav-container/mat-sidenav-content/div/div[2]/mat-tab-group/div/mat-tab-body[2]/div/app-account-transactions-date/div/div/div[2]/mat-table/mat-row[1]/mat-cell[1]/span"
bci_elemento_conta_ativa_aba_movimientos_anterior = "/html/body/app-root/app-private-container/mat-sidenav-container/mat-sidenav-content/app-current-account/div/mat-sidenav-container/mat-sidenav-content/div/app-balances-rut-account/div/div[3]/bci-select-search/div/mat-form-field/div/div[1]/div/mat-select/div/div[1]/span/span"

' ELEMENTOS PORTAL SANTANDER

santander_elemento_login_usuario = "userInput"
santander_elemento_login_senha = "userCodeInput"
santander_elemento_botao_login = "/html/body/form/div[3]/button"
santander_elemento_cuentas_corrientes = "/html/body/app-root/app-layout-perfilado/div/app-menu-perfilado/aside/div/ul/app-nav-item/li/a"
santander_elemento_cartola_historica = "/html/body/app-root/app-layout-perfilado/div/app-menu-perfilado/app-standalone-submenu/div/div[2]/div/ul/li[3]/app-office-banking-link/a"
santander_elemento_list_box_contas = "/html/body/main/div/div[1]/div/div[1]/div[1]/span/div/button"
santander_elemento_list_box_meses = "/html/body/main/div/div[3]/div/div/div[1]/div[2]/div/button/span[1]"
santander_elemento_list_box_anos = "/html/body/main/div/div[3]/div/div/div[1]/div[3]/div/button"
santander_elemento_buscar_cartola = "/html/body/main/div/div[4]/button"
santander_elemento_numero_cartola_atual = "/html/body/main/div[1]/div[1]/div/div/div[1]/div[2]/div/div/button/span[1]"
santander_elemento_ultima_data_encontrada_cartola = "/html/body/main/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[1]/p/span"
santander_elemento_botao_download = "/html/body/main/div[2]/div[2]/div[1]/div[1]/div/a"
santander_elemento_download_em_excel = "/html/body/main/div[2]/div[2]/div[1]/div[1]/div/div/ul/li[1]/a"
santander_elemento_mensagem_sem_movimentos = "modal_message_v1"
santander_elemento_aceptar_mensagem_sem_movimento = "/html/body/div[2]/div/div/div[3]/button"


' ELEMENTOS PORTAL SCOTIABANK
scotiabank_elemento_login_rut = "/html/body/div[1]/div/div[2]/div[1]/div[2]/form/div/div[1]/div/div[2]/div/label/input"
scotiabank_elemento_login_usuario = "/html/body/div[1]/div/div[2]/div[1]/div[2]/form/div/div[2]/div/div[2]/div/label/input"
scotiabank_elemento_login_senha = "/html/body/div[1]/div/div[2]/div[1]/div[2]/form/div/div[3]/div/div[3]/div/label/input"
scotiabank_elemento_botao_login = "/html/body/div[1]/div/div[2]/div[1]/div[2]/form/div/div[4]/button"
scotiabank_elemento_cuentas = "/html/body/div[1]/div/div[1]/div[4]/div/ul/li[2]"
scotiabank_elemento_cartolas = "/html/body/div[1]/div/div[2]/section/div/div/div[3]/div[1]/div[6]/div/div[1]/div[2]"
scotiabank_elemento_ultima_data_encontrada_cartola = "/html/body/div[1]/div/div[2]/section/div/div/div[3]/div[1]/div[6]/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div[1]/div[1]/table/tbody/tr[1]/td[1]/div/div/p"
scotiabank_elemento_download_excel = "/html/body/div[1]/div/div[2]/section/div/div/div[3]/div[1]/div[6]/div/div[2]/div/div/div[2]/div/div/div/div[1]/div/div[3]/div/div[2]/div[3]/div/button"

' ELEMENTO PORTAL BANCO DE CHILE
banco_chile_elemento_login_usuario = "/html/body/div[2]/div/div/article/form/div[1]/input[2]"
banco_chile_elemento_login_senha = "/html/body/div[2]/div/div/article/form/div[2]/input"
banco_chile_elemento_botao_login = "/html/body/div[2]/div/div/article/form/div[3]/div/button"
banco_chile_elemento_atalho_saldos_movimentos = "/html/body/div[2]/hydra-mf-pemp-home-root/div/div/hydra-main/main/article/div/section[1]/hydra-saldos-movimientos-mf/div/div[2]/hydra-movimientos/section/div[1]/div/a/bch-button/div/button/span[1]/span"
banco_chile_elemento_ultimo_movimento_aba_saldos_movimentos = "/html/body/div[2]/hydra-mf-pemp-prd-cta-movimientos/div/div/hydra-main/hydra-saldosmovimientos/section/div/div/section/div[1]/div/bch-interactive-table/div/div/table/tbody/tr[1]/td[2]"
banco_chile_elemento_cartola_historica = "/html/body/div[2]/hydra-mf-pemp-prd-cta-movimientos/div/div/hydra-main/div/div/div[3]/div/section/bch-tabs/div/nav/div[2]/div/div/a[2]"
banco_chile_conta_ativa = "/html/body/div[2]/hydra-mf-pemp-prd-cta-movimientos/div/div/hydra-main/div/div/div[1]/div[2]/section/hydra-selector-producto-saldos/div/div/p[1]/b"
banco_chile_elemento_ultima_data_encontrada_cartola = "/html/body/div[2]/hydra-mf-pemp-prd-cta-movimientos/div/div/hydra-main/hydra-cartolahistorica/div[2]/div[2]/div[2]/div/bch-card-download[1]/div/div[1]/div/div[2]/div[2]/p/span"
banco_chile_elemento_botao_download = "/html/body/div[2]/hydra-mf-pemp-prd-cta-movimientos/div/div/hydra-main/hydra-cartolahistorica/div[2]/div[2]/div[2]/div/bch-card-download[1]/div/div[2]/div/bch-button/div/button"
elemento_abrir_list_box_contas = "/html/body/div[2]/hydra-mf-pemp-prd-cta-movimientos/div/div/hydra-main/div/div/div[1]/div[2]/section/hydra-selector-producto-saldos/div/i"

inicio:
    
    fecha_pagos = CDate(frm_extracao.txtbox_date)
    If fecha_pagos > Date Or Len(fecha_pagos) <> 8 Or fecha_pagos < (Date - 5) Then
        MsgBox "Digite uma data anterior a data de hoje!", vbOKOnly
        GoTo inicio
    ElseIf CStr(fecha_pagos) = "" Then
        MsgBox "Você não digitou uma data válida", vbOKOnly
    End If
    
    Set aba_acesso_bancos = ThisWorkbook.Sheets("Acessos Bancos")
    Set tabela_acesso_bancos = aba_acesso_bancos.ListObjects("Tabela_Acesso_Bancos")
    
    Set aba_contas = ThisWorkbook.Sheets("Contas")
    Set tabela_contas = aba_contas.ListObjects("Tabela_Contas")
    
    Set aba_saldos = ThisWorkbook.Sheets("Consolidado - Saldos")
    Set tabela_saldos = aba_saldos.ListObjects("Tabela_Consolidado_Saldos")
    
    Set aba_pagamentos = ThisWorkbook.Sheets("Consolidado - Pagamentos")
    Set tabela_pagamentos = aba_pagamentos.ListObjects("Tabela_Consolidado_Pagamentos")
    
    Set aba_numero_cartola_banco_chile = ThisWorkbook.Sheets("Número Cartola Banco de Chile")
    Set tabela_numero_cartola_banco_chile = aba_numero_cartola_banco_chile.ListObjects("Tabela_Número_Cartola_Banco_de_Chile")


    ' Abrir o navegador Edge
    
    linha_fim = aba_contas.Range("A200").End(xlUp).Row
    If frm_extracao.opt_extrair_todos = True Then
        aba_contas.Range("E3:F" & linha_fim).ClearContents
    End If

    
    ' setando o banco anterior para já ser o primeiro que aparece na lista para não
    ' prejudicar etapa de verificação do checkbox_executar_next (etapas subsequentes)
    banco_anterior = ""
    
    For linha = 3 To linha_fim
        
        banco = aba_contas.Range("A" & linha).Value
        sociedad = aba_contas.Range("B" & linha).Value
        cuenta = aba_contas.Range("C" & linha).Value
        usuario = Application.WorksheetFunction.VLookup(banco, aba_acesso_bancos.Columns("A:B"), 2, False)
        senha = Application.WorksheetFunction.VLookup(banco, aba_acesso_bancos.Columns("A:C"), 3, False)

        If banco = "BCI" Then
            If frm_extracao.opt_extrair_a_partir_bco_chile Or frm_extracao.opt_extrair_a_partir_santander Or frm_extracao.opt_extrair_a_partir_scotiabank Then
                GoTo proxima_linha
            End If
            Call extracao_bci_
            arquivos_extraidos = True
        ElseIf banco = "BANCO DE CHILE" Then
            If frm_extracao.opt_extrair_a_partir_santander Or frm_extracao.opt_extrair_a_partir_scotiabank Then
                GoTo proxima_linha
            End If
            Call extracao_banco_chile_
            arquivos_extraidos = True
        ElseIf banco = "SANTANDER" Then
            If frm_extracao.opt_extrair_a_partir_scotiabank Then
                GoTo proxima_linha
            End If
            Call extracao_santander_
            arquivos_extraidos = True
        ElseIf banco = "SCOTIABANK" Then
            Call extracao_scotiabank_
            arquivos_extraidos = True
        End If
        
proxima_linha:
        
    Next linha
    
    driver.Quit
    
    RenomearArquivos
    
    ThisWorkbook.RefreshAll
    
    tabela_saldos.QueryTable.Refresh False
    tabela_numero_cartola_banco_chile.QueryTable.Refresh False
    
    Application.Wait (Now + TimeValue("00:00:40"))

    
    ' preenchendo a data da cartola na aba de saldos
    linha_fim = aba_saldos.Range("C200").End(xlUp).Row
    
    For linha = 2 To linha_fim
        aba_saldos.Range("E" & linha).Value = aba_pagamentos.Range("A2").Value
    Next linha
    
    MsgBox ("Aguarde a atualização terminar para continuar a análise e contabilização dos pagamentos."), vbOKOnly
    
    
End Sub

Public Function EsperarElementoEnabled(driver As EdgeDriver, ByType As String, ByVal chave_id_classe_xpath As String)

Dim contador_de_segundos As Integer
Dim elemento_a_verificar As WebElement

    contador_de_segundos = 1
    Application.Wait (Now + TimeValue("00:00:02"))

    If UCase(ByType) = "ID" Then
        Set elemento_a_verificar = driver.FindElementById(chave_id_classe_xpath)
    ElseIf UCase(ByType) = "XPATH" Then
        Set elemento_a_verificar = driver.FindElementByXPath(chave_id_classe_xpath)
    ElseIf UCase(ByType) = "CLASS" Then
        Set elemento_a_verificar = driver.FindElementByClass(chave_id_classe_xpath)
    End If
 
    If Not elemento_a_verificar.IsEnabled Then
        Do Until elemento_a_verificar.IsEnabled
            If contador_de_segundos < 30 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                EsperarElementoEnabled = False
                Exit Function
            End If
        Loop
    End If

    EsperarElementoEnabled = True

End Function


Public Function EsperarElementoPresent(driver As EdgeDriver, ByType As String, ByVal chave_id_classe_xpath As String)

Dim contador_de_segundos As Integer
Dim elemento_a_verificar As WebElement

    contador_de_segundos = 1
    Application.Wait (Now + TimeValue("00:00:02"))

   
    If UCase(ByType) = "XPATH" Then
        Do Until driver.IsElementPresent(by.XPath(chave_id_classe_xpath))
            If contador_de_segundos < 15 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                EsperarElementoPresent = False
                Exit Function
            End If
        Loop
    ElseIf UCase(ByType) = "ID" Then
        Do Until driver.IsElementPresent(by.ID(chave_id_classe_xpath))
            If contador_de_segundos < 15 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                EsperarElementoPresent = False
                Exit Function
            End If
        Loop
    ElseIf UCase(ByType) = "CLASS" Then
        Do Until driver.IsElementPresent(by.Class(chave_id_classe_xpath))
            If contador_de_segundos < 15 Then
                Application.Wait (Now + TimeValue("00:00:01"))
                contador_de_segundos = contador_de_segundos + 1
            Else
                EsperarElementoPresent = False
                Exit Function
            End If
        Loop
    End If
    

    EsperarElementoPresent = True
    

End Function


