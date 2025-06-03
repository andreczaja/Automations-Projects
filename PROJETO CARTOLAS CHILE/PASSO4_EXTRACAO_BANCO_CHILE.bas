Attribute VB_Name = "PASSO4_EXTRACAO_BANCO_CHILE"
Sub extracao_banco_chile_()

    Dim movimentos_dia_selecionado As Boolean
    
    movimentos_dia_selecionado = False

    If banco_anterior = banco Then
        GoTo verificacao_movimentos
    End If
            
    banco_anterior = banco

    driver.Get "https://login.portalempresas.bancochile.cl/bancochile-web/empresa/login/index.html#/login"
    driver.Window.Maximize

    driver.FindElementByXPath(banco_chile_elemento_login_usuario).Click
    driver.FindElementByXPath(banco_chile_elemento_login_usuario).SendKeys usuario
    
    driver.FindElementByXPath(banco_chile_elemento_login_senha).Click
    driver.FindElementByXPath(banco_chile_elemento_login_senha).SendKeys senha
    
    driver.FindElementByXPath(banco_chile_elemento_botao_login).Click
    
    If Not EsperarElementoPresent(driver, "XPATH", banco_chile_elemento_atalho_saldos_movimentos) Then
        GoTo erro_carregamento
    End If

    ' botao elemento productos
    driver.FindElementByXPath(banco_chile_elemento_atalho_saldos_movimentos).Click
    

    
verificacao_movimentos:

    If Not EsperarElementoPresent(driver, "XPATH", banco_chile_conta_ativa) Then
        GoTo erro_carregamento
    End If
    If Not EsperarElementoEnabled(driver, "XPATH", banco_chile_conta_ativa) Then
        GoTo erro_carregamento
    End If
    driver.Get "https://portalempresas.bancochile.cl/mibancochile-web/front/empresa/index.html#/movimientos-cuentas/movimientos"
    driver.ExecuteScript "window.scrollTo(0, 0);"
    If driver.FindElementByXPath(banco_chile_conta_ativa).text <> cuenta Then
        ' elemento_abrir_list_box_contas
        driver.FindElementByXPath(elemento_abrir_list_box_contas).Click
        ' elemento_busca_contas
        For i = 7 To 12
            If driver.IsElementPresent(by.XPath("/html/body/div[" & i & "]/div[2]/div/mat-dialog-container/hydra-modal/div/section/div[1]/div[1]/div[2]/mat-form-field/div/div[1]/div[3]/input")) Then
                Application.Wait (Now + TimeValue("00:00:02"))
                driver.FindElementByXPath("/html/body/div[" & i & "]/div[2]/div/mat-dialog-container/hydra-modal/div/section/div[1]/div[1]/div[2]/mat-form-field/div/div[1]/div[3]/input").Click
                Application.Wait (Now + TimeValue("00:00:02"))
                driver.FindElementByXPath("/html/body/div[" & i & "]/div[2]/div/mat-dialog-container/hydra-modal/div/section/div[1]/div[1]/div[2]/mat-form-field/div/div[1]/div[3]/input").SendKeys cuenta
                driver.FindElementByXPath("/html/body/div[" & i & "]/div[2]/div/mat-dialog-container/hydra-modal/div/section/div[1]/div[2]/ul/li/div[1]/div/mat-radio-button/label/span[1]").Click
                Application.Wait (Now + TimeValue("00:00:02"))
                driver.FindElementByXPath("/html/body/div[" & i & "]/div[2]/div/mat-dialog-container/hydra-modal/div/div/bch-button[2]/div/button").ClickDouble
                Application.Wait (Now + TimeValue("00:00:04"))
                Exit For
            End If
        Next i
    End If
        driver.ExecuteScript "window.scrollTo(0, 200);"
        
        If Not driver.IsElementPresent(by.XPath(banco_chile_elemento_ultimo_movimento_aba_saldos_movimentos)) Then
            aba_contas.Range("E" & linha).Value = "Sem Movimentos"
            GoTo fim
        End If
        
        i = 10
        Do Until i = 0
            If driver.IsElementPresent(by.XPath("/html/body/div[2]/hydra-mf-pemp-prd-cta-movimientos/div/div/hydra-main/hydra-saldosmovimientos/section/div/div/section/div[1]/div/bch-interactive-table/div/div/table/tbody/tr[" & i & "]/td[2]")) Then
                 If driver.FindElementByXPath("/html/body/div[2]/hydra-mf-pemp-prd-cta-movimientos/div/div/hydra-main/hydra-saldosmovimientos/section/div/div/section/div[1]/div/bch-interactive-table/div/div/table/tbody/tr[" & i & "]/td[2]").text = fecha_pagos Then
                    Debug.Print driver.FindElementByXPath("/html/body/div[2]/hydra-mf-pemp-prd-cta-movimientos/div/div/hydra-main/hydra-saldosmovimientos/section/div/div/section/div[1]/div/bch-interactive-table/div/div/table/tbody/tr[" & i & "]/td[2]").text
                    movimentos_dia_selecionado = True
                    aba_contas.Range("E" & linha).Value = "OK"
                    driver.FindElementByXPath(banco_chile_elemento_cartola_historica).Click
                    driver.FindElementByXPath(banco_chile_elemento_botao_download).Click
                    Application.Wait (Now + TimeValue("00:00:01"))
                    For i2 = 7 To 12
                        If driver.IsElementPresent(by.XPath("/html/body/div[" & i2 & "]/div[2]/div/div/div/button[1]")) Then
                            driver.FindElementByXPath("/html/body/div[" & i2 & "]/div[2]/div/div/div/button[1]").Click
                            Application.Wait (Now + TimeValue("00:00:01"))
                            Exit For
                        End If
                    Next i2
                    GoTo fim
                End If
            End If
            i = i - 1
        Loop
        If movimentos_dia_selecionado = False Then
            aba_contas.Range("E" & linha).Value = "Sem Movimentos"
            GoTo fim
        End If
    
    driver.ExecuteScript "window.scrollTo(0, 0);"
    
    GoTo fim

erro_carregamento:
    MsgBox "A página do banco " & UCase(banco) & " não carregou. Por favor, verifique.", vbOKOnly
    End
    
fim:

driver.ExecuteScript "window.scrollTo(0, 0);"
    
End Sub
