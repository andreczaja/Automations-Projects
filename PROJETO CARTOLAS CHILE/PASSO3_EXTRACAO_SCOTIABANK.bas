Attribute VB_Name = "PASSO3_EXTRACAO_SCOTIABANK"
Sub extracao_scotiabank_()


Dim dia, mes, ano As String

    dia = Left(fecha_pagos, 2)
    mes = Mid(fecha_pagos, 4, 2)
    ano = Right(fecha_pagos, 4)

    Select Case mes
         Case "01": mes = "Ene"
         Case "02": mes = "Feb"
         Case "03": mes = "Mar"
         Case "04": mes = "Abr"
         Case "05": mes = "May"
         Case "06": mes = "Jun"
         Case "07": mes = "Jul"
         Case "08": mes = "Ago"
         Case "09": mes = "Sep"
         Case "10": mes = "Oct"
         Case "11": mes = "Nov"
         Case "12": mes = "Dic"
     End Select

'    If banco_anterior = banco Then
'        GoTo verificacao_movimentos
'    End If
'
'    banco_anterior = banco
inicio:
    driver.Get "https://appservtrx.scotiabank.cl/portalempresas/"
    driver.Window.Maximize
    
    driver.FindElementByXPath(scotiabank_elemento_login_rut).Click
    driver.FindElementByXPath(scotiabank_elemento_login_rut).SendKeys "76163495K"
    
    driver.FindElementByXPath(scotiabank_elemento_login_usuario).Click
    driver.FindElementByXPath(scotiabank_elemento_login_usuario).SendKeys usuario
    
    driver.FindElementByXPath(scotiabank_elemento_login_senha).Click
    driver.FindElementByXPath(scotiabank_elemento_login_senha).SendKeys senha
    
    driver.FindElementByXPath(scotiabank_elemento_botao_login).Click
    
    Do Until driver.IsElementPresent(by.XPath(scotiabank_elemento_cuentas))
        Application.Wait (Now + TimeValue("00:00:01"))
    Loop
    
    driver.FindElementByXPath(scotiabank_elemento_cuentas).Click
    
    Do Until driver.IsElementPresent(by.XPath(scotiabank_elemento_cartolas))
        Application.Wait (Now + TimeValue("00:00:01"))
    Loop
    
    
    If Not EsperarElementoEnabled(driver, "XPATH", scotiabank_elemento_cartolas) Then
        GoTo erro_carregamento
    End If
    
    driver.FindElementByXPath(scotiabank_elemento_cartolas).Click
    
    If Not EsperarElementoPresent(driver, "XPATH", scotiabank_elemento_ultima_data_encontrada_cartola) Then
        GoTo erro_carregamento
    End If
    
    If driver.FindElementByXPath(scotiabank_elemento_ultima_data_encontrada_cartola).text = dia & " " & mes & ", " & ano Then
        driver.FindElementByXPath(scotiabank_elemento_download_excel).Click
        Application.Wait (Now + TimeValue("00:00:02"))
        aba_contas.Range("E" & linha).Value = "OK"
    Else
        aba_contas.Range("E" & linha).Value = "Sem movimentos"
    End If
    
GoTo fim

erro_carregamento:
    MsgBox "A página do banco " & UCase(banco) & " não carregou. Por favor, verifique.", vbOKOnly
    End
    
fim:
End Sub
