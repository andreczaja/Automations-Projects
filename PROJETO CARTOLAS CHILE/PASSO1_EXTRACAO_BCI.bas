Attribute VB_Name = "PASSO1_EXTRACAO_BCI"
Sub extracao_bci_()

    Dim array_clp, array_monedas_estranjeras, array_contas_iquique_importadora_electrolux
    Dim conta_clp, conta_moneda_estranjeras, contas_iquique_importadora_electrolux  As Boolean
    
    array_clp = Array(10107258, 10652680, 52022382, 10652931)
    array_monedas_estranjeras = Array(11079673, 11209658, 18537405, 19735367, 18530940, 18579574)
    array_contas_iquique_importadora_electrolux = Array(10652931, 18530940, 18579574)
    conta_clp = False
    conta_moneda_estranjeras = False
    
    ' verificacao de conta clp ou de moneda estranjera, se for clp, extra� a cartola da sess�o cartola historica,
    ' se n�o extra� da sess�o movimientos anterior


    For i = LBound(array_clp) To UBound(array_clp)
        If array_clp(i) = cuenta Then
            conta_clp = True
            conta_moneda_estranjeras = False
            bci_buscas_realizadas_contas_clp = bci_buscas_realizadas_contas_clp + 1
            Exit For
        End If
    Next i

    If Not conta_clp Then
        For i = LBound(array_monedas_estranjeras) To UBound(array_monedas_estranjeras)
            If array_monedas_estranjeras(i) = cuenta Then
                conta_clp = False
                conta_moneda_estranjeras = True
                Exit For
            End If
        Next i
    End If
    
    ' verifica��es para encaminhar c�digo para devidas sess�es dependendo do cen�rio da conta atual e de buscas anteriores
    
    If banco_anterior = banco And conta_clp = True Then
        GoTo verificacao_movimentos_conta_clp
    ElseIf banco_anterior = banco And conta_moneda_estranjeras = True Then
        GoTo verificacao_movimentos_conta_moneda_estranjeras
    End If
            
    banco_anterior = banco

    driver.Get "https://www.bci.cl/empresas"
    driver.Window.Maximize

    ' elemento bot�o "BANCO EN L�NEA"
    driver.FindElementByXPath(bci_elemento_banco_en_linea).Click
    
    Application.Wait (Now + TimeValue("00:00:01"))
    
    ' inserindo usu�rio
    If Not EsperarElementoEnabled(driver, "ID", bci_elemento_login_usuario) Then
        GoTo erro_carregamento
    End If
    
    Set elementoInput = driver.FindElementById(bci_elemento_login_usuario)
    elementoInput.Click
    elementoInput.SendKeys usuario
    'inserindo senha
    Set elementoInput = driver.FindElementById(bci_elemento_login_senha)
    elementoInput.Click
    elementoInput.SendKeys senha
    'clicando no bot�o para logar
    driver.FindElementByXPath(bci_elemento_botao_login).Click
    
    ' inserindo usu�rio
    If Not EsperarElementoEnabled(driver, "Class", bci_elemento_box_grupo_electrolux) Then
        GoTo erro_carregamento
    End If
    ' elemento do grupo electrolux, ap�s o clique entra-se na tela do portal
    driver.FindElementByClass(bci_elemento_box_grupo_electrolux).Click
    
    
    ' aguardando elemento que abre todas as op��es estar enable
    If Not EsperarElementoEnabled(driver, "XPATH", bci_elemento_opcoes_menu_geral) Then
        GoTo erro_carregamento
    End If
    Application.Wait (Now + TimeValue("00:00:02"))
    ' elemento para abrir todas as op��es, dentre eles a de extra��o de cartolas
    On Error Resume Next
    driver.FindElementByXPath(bci_elemento_opcoes_menu_geral).Click
    
    
    ' aguardando elemento que abre sess�o "cuentas"
    If Not EsperarElementoEnabled(driver, "XPATH", bci_elemento_sessao_cuentas) Then
        GoTo erro_carregamento
    End If
    ' elemento que ir� abrir a sess�o cuentas para escolher a op��o "Cuentas corrientes"
    driver.FindElementByXPath(bci_elemento_sessao_cuentas).Click
    
    ' aguardando elemento que abre sess�o "cuentas corrientes"
    If Not EsperarElementoEnabled(driver, "XPATH", bci_elemento_sessao_cuentas_corrientes) Then
        GoTo erro_carregamento
    End If
    ' elemento que ir� abrir a sess�o cuentas corrientes para escolher a op��o "Cartola Hist�rica"
    driver.FindElementByXPath(bci_elemento_sessao_cuentas_corrientes).Click
    
    If conta_clp Then
        Application.Wait (Now + TimeValue("00:00:02"))
        Set elementosClasse = driver.FindElementsByClass(bci_elemento_opcoes_abertas_menu_geral)
        For Each elementoInput In elementosClasse
            If elementoInput.text = "Cartola Hist�rica" Then
                elementoInput.Click
                GoTo verificacao_movimentos_conta_clp
            End If
        Next elementoInput
    Else
        Application.Wait (Now + TimeValue("00:00:02"))
        Set elementosClasse = driver.FindElementsByClass(bci_elemento_opcoes_abertas_menu_geral)
        For Each elementoInput In elementosClasse
            If elementoInput.text = "Movimientos (anterior)" Then
                elementoInput.Click
                GoTo verificacao_movimentos_conta_moneda_estranjeras
            End If
        Next elementoInput
    End If
            
            
verificacao_movimentos_conta_clp:

    If Not EsperarElementoEnabled(driver, "XPATH", bci_elemento_lista_de_sociedades_banco) Then
        GoTo erro_carregamento
    End If
    driver.ExecuteScript "window.scrollTo(0, 0);"
    
    ' verifica se � uma conta de iquique, se sim, faz as altera��es de busca necess�rias
    
    For i = LBound(array_contas_iquique_importadora_electrolux) To UBound(array_contas_iquique_importadora_electrolux)
        If array_contas_iquique_importadora_electrolux(i) = cuenta Then
            driver.FindElementByXPath(bci_elemento_lista_de_sociedades_banco).Click
            contas_iquique_importadora_electrolux = True
            bci_selecionada_aba_iquique_importadora = True
            Application.Wait (Now + TimeValue("00:00:01"))
            For i2 = 2 To 4
                If driver.IsElementPresent(by.XPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[3]")) Then
                    driver.FindElementByXPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[3]").Click
                    Exit For
                End If
            Next i2
            Exit For
        End If
    Next i

    ' elemento que verifica se a conta correta � a que est� selecionada, se n�o, ir� selecionar
    'conta que est� selecionada no listbox

    If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_conta_ativa) Then
        GoTo erro_carregamento
    End If
    
        ' VERIFICA��O DE SELECIONAR A SOCIEDADE CORRETA DATA A CONTA ATUAL
    If Not bci_selecionada_aba_iquique_importadora And sociedad = "TC08" Then
        For i2 = 2 To 4
            If driver.IsElementPresent(by.XPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[3]")) Then
                driver.FindElementByXPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[3]").Click
                Exit For
            End If
        Next i2
    End If
    
    
    Debug.Print driver.FindElementByXPath(bci_elemento_conta_ativa).text
    If driver.FindElementByXPath(bci_elemento_conta_ativa).text <> "Cuenta corriente (CLP) - N� " & cuenta Then
        If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_conta_ativa) Then
            GoTo erro_carregamento
        End If
        driver.FindElementByXPath(bci_elemento_conta_ativa).Click
        For i = 2 To 20
            If driver.FindElementByXPath(bci_elemento_conta_ativa).text = "Cuenta corriente (CLP) - N� " & cuenta Then
                Exit For
            End If
            ' elemento das contas disponiveis no listbox diretas da electrolux (impossivel jogar numa variavel pois tem manipula��o da xpath)
            For i2 = 2 To 4
                If driver.IsElementPresent(by.XPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[" & i & "]/span/p[2]")) Then
                    If driver.FindElementByXPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[" & i & "]/span/p[2]").text = "N� " & cuenta Then
                        driver.FindElementByXPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[" & i & "]/span/p[2]").Click
                        Exit For
                    End If
                End If
            Next i2
        Next i
    End If
    If Not EsperarElementoEnabled(driver, "XPATH", bci_elemento_conta_ativa) Then
        GoTo erro_carregamento
    End If
    ' verifica se o elemento_ultima_data_encontrada_extrato est� presente
    Application.Wait (Now + TimeValue("00:00:02"))
    If Not driver.IsElementPresent(by.XPath(bci_elemento_ultima_data_encontrada_extrato)) Then
        aba_contas.Range("E" & linha).Value = "Sem Movimentos"
        GoTo fim
    End If

    Debug.Print driver.FindElementByXPath(bci_elemento_ultima_data_encontrada_extrato).text
    If driver.FindElementByXPath(bci_elemento_ultima_data_encontrada_extrato).text = fecha_pagos Then
        If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_botao_download) Then
            GoTo erro_carregamento
        End If

        driver.FindElementByXPath(bci_elemento_botao_download).Click
        Application.Wait (Now + TimeValue("00:00:01"))
        For i = 2 To 6
            If driver.IsElementPresent(by.XPath("/html/body/div[3]/div[" & i & "]/div/div/button[2]")) Then
                If driver.FindElementByXPath("/html/body/div[3]/div[" & i & "]/div/div/button[2]").text Like "*Descarga cartola Excel" Then
                driver.FindElementByXPath("/html/body/div[3]/div[" & i & "]/div/div/button[2]").Click
                Exit For
                End If
            End If
        Next i
        aba_contas.Range("E" & linha).Value = "OK"
        aba_contas.Range("F" & linha).Value = driver.FindElementByXPath(bci_elemento_numero_cartola_atual).text
        Application.Wait (Now + TimeValue("00:00:02"))
    Else
        aba_contas.Range("E" & linha).Value = "Sem Movimentos"
    End If
    
    ' ESTRUTURA QUE IR� VERIFICAR SE TODAS AS CONTAS CLP J� FORAM VERIFICADAS, SE SIM, IR� PARA A SESS�O DE MOVIMIMENTOS(ANTERIOR)
    ' PARA BAIXAR CARTOLAS DESSAS CONTAS
    
    If bci_buscas_realizadas_contas_clp = UBound(array_clp) + 1 Then
            bci_buscas_realizadas_contas_clp = bci_buscas_realizadas_contas_clp + 1
               ' aguardando elemento que abre todas as op��es estar enable
        If Not EsperarElementoEnabled(driver, "XPATH", bci_elemento_opcoes_menu_geral) Then
            GoTo erro_carregamento
        End If
        Application.Wait (Now + TimeValue("00:00:02"))
        ' elemento para abrir todas as op��es, dentre eles a de extra��o de cartolas
        On Error Resume Next
        driver.FindElementByXPath(bci_elemento_opcoes_menu_geral).Click
        
        If Not driver.FindElementByXPath(bci_elemento_sessao_cuentas_corrientes).IsDisplayed Then
            ' aguardando elemento que abre sess�o "cuentas"
            If Not EsperarElementoEnabled(driver, "XPATH", bci_elemento_sessao_cuentas) Then
                GoTo erro_carregamento
            End If
            ' elemento que ir� abrir a sess�o cuentas para escolher a op��o "Cuentas corrientes"
            driver.FindElementByXPath(bci_elemento_sessao_cuentas).Click
            ' aguardando elemento que abre sess�o "cuentas corrientes"
            If Not EsperarElementoEnabled(driver, "XPATH", bci_elemento_sessao_cuentas_corrientes) Then
                GoTo erro_carregamento
            End If
            ' elemento que ir� abrir a sess�o cuentas corrientes para escolher a op��o "Movimientos (anterior)"
            driver.FindElementByXPath(bci_elemento_sessao_cuentas_corrientes).Click
        End If

        
        Application.Wait (Now + TimeValue("00:00:02"))
        Set elementosClasse = driver.FindElementsByClass(bci_elemento_opcoes_abertas_menu_geral)
        For Each elementoInput In elementosClasse
            If elementoInput.text = "Movimientos (anterior)" Then
                elementoInput.Click
                Exit For
            End If
        Next elementoInput
    
    End If

GoTo fim

verificacao_movimentos_conta_moneda_estranjeras:
    Application.Wait (Now + TimeValue("00:00:02"))
    If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_conta_ativa_aba_movimientos_anterior) Then
        GoTo erro_carregamento
    End If
    driver.ExecuteScript "window.scrollTo(0, 0);"
    
    ' verifica se � uma conta de iquique
    
        For i = LBound(array_contas_iquique_importadora_electrolux) To UBound(array_contas_iquique_importadora_electrolux)
            Debug.Print array_contas_iquique_importadora_electrolux(i)
            If array_contas_iquique_importadora_electrolux(i) = cuenta Then
                contas_iquique_importadora_electrolux = True
                If bci_selecionada_aba_iquique_importadora Then
                    Exit For
                End If
                driver.FindElementByXPath(bci_elemento_lista_de_sociedades_banco_aba_movimientos_anterior).Click
                bci_selecionada_aba_iquique_importadora = True
                Application.Wait (Now + TimeValue("00:00:01"))
                For i2 = 2 To 4
                    If driver.IsElementPresent(by.XPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[3]")) Then
                        driver.FindElementByXPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[3]").Click
                        Exit For
                    End If
                Next i2
                Exit For
            End If
        Next i

    ' elemento que verifica se a sociedade correta � a que est� selecionada, se n�o, ir� selecionar
    'conta que est� selecionada no listbox

    If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_lista_de_sociedades_banco_aba_movimientos_anterior) Then
        GoTo erro_carregamento
    End If
    
    ' VERIFICA��O DE SELECIONAR A SOCIEDADE CORRETA DATA A CONTA ATUAL
    If bci_selecionada_aba_iquique_importadora And sociedad = "TC04" Then
        driver.FindElementByXPath(bci_elemento_lista_de_sociedades_banco_aba_movimientos_anterior).Click
        Application.Wait (Now + TimeValue("00:00:01"))
        For i2 = 2 To 6
            If driver.IsElementPresent(by.XPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[2]")) Then
                driver.FindElementByXPath("/html/body/div[3]/div[" & i2 & "]/div/div/mat-option[2]").Click
                bci_selecionada_aba_iquique_importadora = False
                Exit For
            End If
        Next i2
    End If
    
    If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_conta_ativa_aba_movimientos_anterior) Then
        GoTo erro_carregamento
    End If
    
    
    If driver.FindElementByXPath(bci_elemento_conta_ativa_aba_movimientos_anterior).text <> "Cuenta Corriente " & cuenta Then
        driver.FindElementByXPath(bci_elemento_conta_ativa_aba_movimientos_anterior).Click
        Application.Wait (Now + TimeValue("00:00:01"))
        For i = 2 To 20
            ' elemento das contas disponiveis no listbox de contas de importadora (impossivel jogar numa variavel pois tem manipula��o da xpath)
            If driver.FindElementByXPath("/html/body/div[3]/div[3]/div/div/mat-option[" & i & "]/span/p[2]").text = cuenta Then
                driver.FindElementByXPath("/html/body/div[3]/div[3]/div/div/mat-option[" & i & "]/span/p[2]").Click
                Exit For
            End If
        Next i
    End If
    ' espera o bot�o de consulta de movimientos estar ativo
    If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_consulta_de_movimientos) Then
        GoTo erro_carregamento
    End If
    
retorno_erro_botao_consulta_movimientos:

    Application.Wait (Now + TimeValue("00:00:01"))
    On Error GoTo erro_botao_consulta_movimientos
    driver.FindElementByXPath(bci_elemento_consulta_de_movimientos).Click
    
    ' verifica se o elemento que respons�vel por fazer o download � ativo e depois
    ' se � igual a data da extra��o
    If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_download_excel_aba_movimientos_anterior) Then
        GoTo erro_carregamento
    End If
    ' se n�o for ativo, e a cuenta selecionada � a mesma que est�
    ' sendo verificada coloca Sem Movimientos
    If Not driver.IsElementPresent(by.XPath(bci_elemento_ultima_data_encontrada_extrato_aba_movimientos_anterior)) And _
        driver.FindElementByXPath(bci_elemento_conta_ativa_aba_movimientos_anterior).text = "Cuenta Corriente " & cuenta Then

        aba_contas.Range("E" & linha).Value = "Sem Movimentos"
        GoTo fim
    ' se n�o for ativo e a cuenta selecionada � diferente da que est� sendo verificada,
    ' repete o processo para trazer assertividade na analise
    ElseIf Not driver.IsElementPresent(by.XPath(bci_elemento_ultima_data_encontrada_extrato_aba_movimientos_anterior)) And _
        driver.FindElementByXPath(bci_elemento_conta_ativa_aba_movimientos_anterior).text <> "Cuenta Corriente " & cuenta Then
        Debug.Print driver.FindElementByXPath(bci_elemento_conta_ativa_aba_movimientos_anterior).text
        GoTo verificacao_movimentos_conta_moneda_estranjeras
    End If
    Application.Wait (Now + TimeValue("00:00:02"))
    driver.ExecuteScript "window.scrollTo(0, 0);"
    driver.ExecuteScript "window.scrollTo(0, 400);"
    ' verifica se a ultima data encontrada na cartola � igual a da data buscada pelo usuario,
    ' se sim, conta OK, se n�o, conta Sem Movimentos
    If CDate(driver.FindElementByXPath(bci_elemento_ultima_data_encontrada_extrato_aba_movimientos_anterior).text) = fecha_pagos Then
        aba_contas.Range("E" & linha).Value = "OK"
        ' clica no icone para fazer o download da cartola
        If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_download_excel_aba_movimientos_anterior) Then
            GoTo erro_carregamento
        End If
        driver.FindElementByXPath(bci_elemento_download_excel_aba_movimientos_anterior).Click
        Application.Wait (Now + TimeValue("00:00:02"))
        ' esperando o bot�o de consulta de movimientos ficar enabled para seguir
        If Not EsperarElementoEnabled(driver, "xpath", bci_elemento_consulta_de_movimientos) Then
            GoTo erro_carregamento
        End If
    Else
        aba_contas.Range("E" & linha).Value = "Sem Movimentos"
    End If

GoTo fim
    
erro_botao_consulta_movimientos:
    Application.Wait (Now + TimeValue("00:00:02"))
    driver.FindElementByXPath(bci_elemento_consulta_de_movimientos).Click
    GoTo retorno_erro_botao_consulta_movimientos
    
erro_carregamento:
    MsgBox "A p�gina do banco " & UCase(banco) & " n�o carregou. Por favor, apague os arquivos e rode novamente.", vbOKOnly
    End
    
fim:


End Sub

