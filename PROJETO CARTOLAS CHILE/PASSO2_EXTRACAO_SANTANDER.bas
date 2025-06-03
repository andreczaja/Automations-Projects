Attribute VB_Name = "PASSO2_EXTRACAO_SANTANDER"
Sub extracao_santander_()

'PASSAR VARIAVEIS PARA O MODULO PRINCIPAL

Dim mes, ano As String

    mes = Mid(fecha_pagos, 4, 2)
    ano = Right(fecha_pagos, 4)

    Select Case mes
         Case "01": mes = "Enero"
         Case "02": mes = "Febrero"
         Case "03": mes = "Marzo"
         Case "04": mes = "Abril"
         Case "05": mes = "Mayo"
         Case "06": mes = "Junio"
         Case "07": mes = "Julio"
         Case "08": mes = "Agosto"
         Case "09": mes = "Septiembre"
         Case "10": mes = "Octubre"
         Case "11": mes = "Noviembre"
         Case "12": mes = "Diciembre"
     End Select

    If banco_anterior = banco Then
        GoTo verificacao_movimentos
    End If
            
    banco_anterior = banco
inicio:
    driver.Get "https://wslogin.officebanking.cl/?theme=ob"
    driver.Window.Maximize
    ' inserindo usuário
    If Not EsperarElementoEnabled(driver, "ID", santander_elemento_login_usuario) Then
        GoTo erro_carregamento
    End If
    
    Set elementoInput = driver.FindElementById(santander_elemento_login_usuario)
    elementoInput.Click
    elementoInput.SendKeys usuario
    'inserindo senha
    Set elementoInput = driver.FindElementById(santander_elemento_login_senha)
    elementoInput.Click
    elementoInput.SendKeys senha
    'clicando no botão para logar
    
    If Not EsperarElementoEnabled(driver, "XPATH", santander_elemento_botao_login) Then
        GoTo erro_carregamento
    End If
    driver.FindElementByXPath(santander_elemento_botao_login).Click
    
    Application.Wait (Now + TimeValue("00:00:02"))
    If driver.IsElementPresent(by.XPath("/html/body/app-root/app-error-seguridad/section/div[3]/div[2]/div/button")) Then
        GoTo inicio
    End If
    
    ' esperando o elemento de cuentas corrientes estar presente e enabled
    If Not EsperarElementoPresent(driver, "XPATH", santander_elemento_cuentas_corrientes) Then
        GoTo erro_carregamento
    End If
    If Not EsperarElementoEnabled(driver, "XPATH", santander_elemento_cuentas_corrientes) Then
        GoTo erro_carregamento
    End If
    driver.FindElementByXPath(santander_elemento_cuentas_corrientes).Click
    
    
    If Not EsperarElementoEnabled(driver, "XPATH", santander_elemento_cartola_historica) Then
        GoTo erro_carregamento
    End If
    driver.FindElementByXPath(santander_elemento_cartola_historica).Click
    
    
    ' esperando o elemento de list box de contas estar presente e enabled
verificacao_movimentos:
    driver.Get "https://eob.officebanking.cl/CTA.UI.Web/CartolaHistoricaCtaCte/?servicioId=TRNCNA_CTLHTRCA_ULT"
    
    Application.Wait (Now + TimeValue("00:00:02"))
    If driver.FindElementByXPath("/html/body/div/div/div[2]/div/h3/span").IsDisplayed Then
        MsgBox "O sistema do Santander está apresentando instabilidade, verifique.", vbOKOnly
        End
    End If
    If Not EsperarElementoPresent(driver, "XPATH", santander_elemento_list_box_contas) Then
        GoTo inicio
    End If
    If Not EsperarElementoEnabled(driver, "XPATH", santander_elemento_list_box_contas) Then
        GoTo inicio
    End If
    driver.FindElementByXPath(santander_elemento_list_box_contas).Click
    
    
    If Not EsperarElementoEnabled(driver, "XPATH", santander_elemento_list_box_contas) Then
        GoTo inicio
    End If
    
    ' seleciona a cuenta correta
    For i = 2 To 4
        Set elementoInput = driver.FindElementByXPath("/html/body/main/div/div[1]/div/div[1]/div[1]/span/div/div/ul/li[" & i & "]")
        If Left(elementoInput.text, 15) = cuenta Then
            elementoInput.Click
            Exit For
        End If
    Next i
    
    ' espera o elemento list_box_meses estar enable
    If Not EsperarElementoEnabled(driver, "XPATH", santander_elemento_list_box_meses) Then
        GoTo erro_carregamento
    End If
    ' verifica seleção do mes correto caso o selecionado seja diferente da data que o usuario escolheu para extrair as cartolas
    If driver.FindElementByXPath(santander_elemento_list_box_meses).text <> mes Then
        driver.FindElementByXPath(santander_elemento_list_box_meses).Click
        For i2 = 1 To 12
            Set elementoInput = driver.FindElementByXPath("/html/body/main/div/div[3]/div/div/div[1]/div[2]/div/div/ul/li[" & i2 & "]")
            If elementoInput.text = mes Then
               elementoInput.Click
                Exit For
            End If
        Next i2
    End If
    ' verifica seleção do mes correto caso o selecionado seja diferente da data que o usuario escolheu para extrair as cartolas
    If Left(driver.FindElementByXPath(santander_elemento_list_box_anos).text, 4) <> ano Then
        driver.FindElementByXPath(santander_elemento_list_box_anos).Click
        For i2 = 1 To 7
            Set elementoInput = driver.FindElementByXPath("/html/body/main/div/div[3]/div/div/div[1]/div[3]/div/div/ul/li[" & i2 & "]")
            If elementoInput.text = ano Then
                elementoInput.Click
                Exit For
            End If
        Next i2
    End If
    driver.FindElementByXPath(santander_elemento_buscar_cartola).Click
    
    Application.Wait (Now + TimeValue("00:00:04"))
    If driver.IsElementPresent(by.ID(santander_elemento_mensagem_sem_movimentos)) Then
        If driver.FindElementById(santander_elemento_mensagem_sem_movimentos).IsDisplayed Then
            ' botão de aceptar que não existe movimento na conta escolhida
            driver.FindElementByXPath(santander_elemento_aceptar_mensagem_sem_movimento).Click
            aba_contas.Range("E" & linha).Value = "Sem Movimentos"
            GoTo fim
        End If
    End If
    
    If Not EsperarElementoEnabled(driver, "XPATH", santander_elemento_botao_download) Then
        GoTo erro_carregamento
    End If
    
    If Not driver.IsElementPresent(by.XPath(santander_elemento_ultima_data_encontrada_cartola)) Then
        Application.Wait (Now + TimeValue("00:00:02"))
    End If
    
    
    If Not driver.IsElementPresent(by.XPath(santander_elemento_ultima_data_encontrada_cartola)) Then
        aba_contas.Range("E" & linha).Value = "Sem Movimentos"
        GoTo fim
    ElseIf driver.FindElementByXPath(santander_elemento_ultima_data_encontrada_cartola).text <> fecha_pagos Then
        ' aqui vai para a proxima conta caso não encontre movimentos do dia selecionado na cartola
        aba_contas.Range("E" & linha).Value = "Sem Movimentos"
        GoTo fim
    ElseIf driver.FindElementByXPath(santander_elemento_ultima_data_encontrada_cartola).text = fecha_pagos Then
        driver.FindElementByXPath(santander_elemento_botao_download).Click
        Application.Wait (Now + TimeValue("00:00:02"))
        driver.FindElementByXPath(santander_elemento_download_em_excel).Click
        aba_contas.Range("E" & linha).Value = "OK"
        For i = 2 To 10
            If driver.IsElementPresent(by.XPath("/html/body/main/div[1]/div[1]/div/div/div[1]/div[" & i & "]/div/span")) Then
                aba_contas.Range("F" & linha).Value = Left(driver.FindElementByXPath("/html/body/main/div[1]/div[1]/div/div/div[1]/div[" & i & "]/div/span").text, 3)
                Exit For
            ElseIf driver.IsElementPresent(by.XPath("/html/body/main/div[1]/div[1]/div/div/div[1]/div[" & i & "]/div/div/button/span[1]")) Then
                aba_contas.Range("F" & linha).Value = Left(driver.FindElementByXPath("/html/body/main/div[1]/div[1]/div/div/div[1]/div[" & i & "]/div/div/button/span[1]").text, 3)
                Exit For
            End If
        Next i
        Application.Wait (Now + TimeValue("00:00:02"))
    End If
    
    GoTo fim
    
erro_carregamento:
    MsgBox "A página do banco " & UCase(banco) & " não carregou. Por favor, verifique.", vbOKOnly
    End
    
fim:
    
End Sub

