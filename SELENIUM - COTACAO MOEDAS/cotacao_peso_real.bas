Attribute VB_Name = "cotacao_peso_real"
Option Explicit

Sub TESTE_cotacao_peso_real_()
    
    Dim driver As New Selenium.EdgeDriver
    Dim elemento_Mes_Escolhido As WebElements
    Dim elemento_Dia_Escolhido As WebElement
    Dim elemento_Ano_Escolhido As WebElement
    Dim elemento_flecha_Ano As WebElement
    Dim elemento_peso_cotado As WebElement
    Dim elemento_Calendario As WebElement
    Dim elemento_Calendario_esquerdo As WebElement
    Dim elemento_Dolar As WebElement
    Dim elemento_Euro As WebElement
    Dim elemento_Yen As WebElement
    Dim ano, mes, dia, data_completa, mes_ano As String
    Dim data_menos_10_anos As Date
    Dim by As New by
    Dim i, dia_mais_proximo_com_cotacao, dia_int, mes_int As Integer
    Dim cotacao_data_nao_encontrada As Boolean
    
    
    
    cotacao_data_nao_encontrada = False
    data_completa = frm_cotacao.txt_box_date
    dia = Left(data_completa, 2)
    mes = Mid(data_completa, 4, 2)
    mes_int = CInt(mes)
    ano = Right(data_completa, 4)
    
    'data - 10 anos para impedir que o usuário insira uma data anterior ao que o site comporta
    data_menos_10_anos = DateSerial(Year(Date) - 10, 1, 1)
    If CDate(data_completa) > Date Then
        MsgBox "Por favor insira uma data anterior a data atual", vbOKOnly
        End
        
    ElseIf CDate(data_completa) < data_menos_10_anos Then
        MsgBox "Por favor insira uma data entre " & Year(Date) & " e " & Year(Date) - 10 & ".", vbOKOnly
        End
    ElseIf CInt(dia) < 1 Or CInt(dia) > 31 Or CInt(mes) < 1 Or CInt(mes) > 12 Or CInt(ano) < 1 Or CInt(ano) > CInt(Year(Date)) Or Len(frm_cotacao.txt_box_date.text) <> 10 Then
        MsgBox "Por favor insira uma data válida", vbOKOnly
        End
    End If
    Set driver = CreateObject("Selenium.EdgeDriver")
    driver.AddArgument "--headless=new"
    driver.AddArgument "--disable-gpu"
    driver.AddArgument "--window-size=1920,1080"
    driver.get "https://si3.bcentral.cl/indicadoressiete/secure/IndicadoresDiarios.aspx"
    
' em caso sem nenhuma variação cambial na data encontrada, o sistema retorna para cá com uma data = date-1
date_menos_1:
    
     Select Case mes
        Case "01": mes = "enero"
        Case "02": mes = "febrero"
        Case "03": mes = "marzo"
        Case "04": mes = "abril"
        Case "05": mes = "mayo"
        Case "06": mes = "junio"
        Case "07": mes = "julio"
        Case "08": mes = "agosto"
        Case "09": mes = "septiembre"
        Case "10": mes = "octubre"
        Case "11": mes = "noviembre"
        Case "12": mes = "diciembre"
        Case Else
            MsgBox "Por favor, digite uma data válida!", vbOKOnly
        Exit Sub
    End Select
    
        mes_ano = mes & " de " & ano
    

        
        If CDate(data_completa) = Date Then
            GoTo fim
        End If
    
    
        Set elemento_Calendario = driver.FindElementById("_calendarioButton")
            elemento_Calendario.Click
    
            
            ' buscando o ano correto
            If driver.FindElementById("calendario_YearSelectorTitle").text <> ano Then
                
                Set elemento_Ano_Escolhido = driver.FindElementById("calendario_YearSelectorTitle")
                elemento_Ano_Escolhido.Click
                
                If driver.FindElementById("calendario_Year" & ano).IsDisplayed Then
                    driver.FindElementById("calendario_Year" & ano).Click
                Else
                    Do Until driver.FindElementById("calendario_Year" & ano).IsDisplayed
                        Set elemento_flecha_Ano = driver.FindElementById("calendario_YearSelectorMoveUp")
                        elemento_flecha_Ano.ClickAndHold
                    Loop
                    driver.FindElementById("calendario_Year" & ano).Click
                End If
            End If
                
            ' buscando o mes correto
            Set elemento_Calendario_esquerdo = driver.FindElementByXPath("/html/body/div/div/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td")
            
            Do Until elemento_Calendario_esquerdo.text = mes_ano
'                Debug.Print elemento_Calendario_esquerdo.text
'                Debug.Print mes_ano
'                Debug.Print CInt(Month(Date))
'                Debug.Print mes_int
                If CInt(Month(Date)) >= mes_int Then
                    driver.FindElementByClass("calendarArrowLeft").Click
                    Set elemento_Calendario_esquerdo = driver.FindElementByXPath("/html/body/div/div/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td")
                ElseIf CInt(Month(Date)) < mes_int Then
                    driver.FindElementByClass("calendarArrowRight").Click
                    Set elemento_Calendario_esquerdo = driver.FindElementByXPath("/html/body/div/div/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td")
                End If
            Loop
            
            'buscando o dia correto
            If elemento_Calendario_esquerdo.text = mes_ano Then
                Set elemento_Mes_Escolhido = driver.FindElementsByClass("calendarDay")
            
                For Each elemento_Dia_Escolhido In elemento_Mes_Escolhido
                    Debug.Print elemento_Dia_Escolhido.text
                    If elemento_Dia_Escolhido.text = CInt(dia) Then
                        elemento_Dia_Escolhido.Click
                        Exit For
                    End If
                Next elemento_Dia_Escolhido
            End If
            
            Application.Wait (Now + TimeValue("00:00:01"))
            'voltando o circuito para date - 1 caso a conversão para outros moedas seja igual a ND e seguindo caso não seja
            If driver.FindElementById("lblValor1_3").text = "ND" And driver.FindElementById("lblValor1_5").text = "ND" And _
                driver.FindElementById("lblValor1_10").text = "ND" Then
                data_completa = CDate(data_completa) - 1
                dia = Left(data_completa, 2)
                mes = Mid(data_completa, 4, 2)
                mes_int = CInt(mes)
                ano = Right(data_completa, 4)
                cotacao_data_nao_encontrada = True
                GoTo date_menos_1
            End If
            
            
fim:
            Set elemento_Dolar = driver.FindElementById("lblValor1_3")
            Set elemento_Euro = driver.FindElementById("lblValor1_5")
            Set elemento_Yen = driver.FindElementById("lblValor1_10")
            
            If cotacao_data_nao_encontrada Then
                frm_cotacao.lbl_dia_anterior.Caption = "OBS: NÃO ENCONTRADA A COTACÃO NO DIA SELECIONADO, O DIA MAIS PRÓXIMO FOI " & data_completa
                frm_cotacao.Height = 383
                frm_cotacao.Width = 250
                
                
                ' ajustado largura e altura label "VALOR DO PESO EM:"
                frm_cotacao.lbl_valor_das_moedas.Height = 18
                frm_cotacao.lbl_valor_das_moedas.Width = 102
                
                ' ajustado largura e altura labels dos valores das moedas
                frm_cotacao.lbl_valor_peso_euro_frm.Height = 30
                frm_cotacao.lbl_valor_peso_dolar_frm.Height = 30
                frm_cotacao.lbl_valor_peso_yen_frm.Height = 30
                
                frm_cotacao.lbl_valor_peso_euro_frm.Width = 102
                frm_cotacao.lbl_valor_peso_dolar_frm.Width = 102
                frm_cotacao.lbl_valor_peso_yen_frm.Width = 102
                
                
                ' ajustando largura e altura labels texto de moedas
                frm_cotacao.lbl_dolar.Height = 23
                frm_cotacao.lbl_euro.Height = 23
                frm_cotacao.lbl_yen.Height = 23
                
                frm_cotacao.lbl_dolar.Width = 35
                frm_cotacao.lbl_euro.Width = 35
                frm_cotacao.lbl_yen.Width = 35
                
                 ' ajustando largura e altura botões copiar
                frm_cotacao.copiar_dolar.Height = 18
                frm_cotacao.copiar_euro.Height = 18
                frm_cotacao.copiar_yen.Height = 18
                
                frm_cotacao.copiar_dolar.Width = 54
                frm_cotacao.copiar_euro.Width = 54
                frm_cotacao.copiar_yen.Width = 54
                
                frm_cotacao.lbl_dia_anterior.Height = 60
                frm_cotacao.lbl_dia_anterior.Width = 192
                

                frm_cotacao.lbl_valor_peso_dolar_frm.Caption = elemento_Dolar.text
                frm_cotacao.lbl_valor_peso_euro_frm.Caption = elemento_Euro.text
                frm_cotacao.lbl_valor_peso_yen_frm.Caption = Format(CDbl(elemento_Yen.text), "0.00")
                
                frm_cotacao.lbl_valor_peso_dolar_frm.Visible = True
                frm_cotacao.lbl_valor_peso_euro_frm.Visible = True
                frm_cotacao.lbl_valor_peso_yen_frm.Visible = True
                
                frm_cotacao.lbl_euro.Visible = True
                frm_cotacao.lbl_dolar.Visible = True
                frm_cotacao.lbl_yen.Visible = True
                
                frm_cotacao.copiar_euro.Visible = True
                frm_cotacao.copiar_dolar.Visible = True
                frm_cotacao.copiar_yen.Visible = True

                frm_cotacao.lbl_valor_das_moedas.Visible = True
                frm_cotacao.lbl_dia_anterior.Visible = True
            Else
                frm_cotacao.Height = 297
                frm_cotacao.Width = 250
                
                ' ajustado largura e altura label "VALOR DO PESO EM:"
                frm_cotacao.lbl_valor_das_moedas.Height = 18
                frm_cotacao.lbl_valor_das_moedas.Width = 102
                
                ' ajustado largura e altura labels dos valores das moedas
                frm_cotacao.lbl_valor_peso_euro_frm.Height = 24
                frm_cotacao.lbl_valor_peso_dolar_frm.Height = 24
                frm_cotacao.lbl_valor_peso_yen_frm.Height = 24
                
                frm_cotacao.lbl_valor_peso_euro_frm.Width = 102
                frm_cotacao.lbl_valor_peso_dolar_frm.Width = 102
                frm_cotacao.lbl_valor_peso_yen_frm.Width = 102
                
                
                ' ajustando largura e altura labels texto de moedas
                frm_cotacao.lbl_dolar.Height = 18
                frm_cotacao.lbl_euro.Height = 18
                frm_cotacao.lbl_yen.Height = 18
                
                frm_cotacao.lbl_dolar.Width = 50
                frm_cotacao.lbl_euro.Width = 50
                frm_cotacao.lbl_yen.Width = 50
                
                 ' ajustando largura e altura botões copiar
                frm_cotacao.copiar_dolar.Height = 18
                frm_cotacao.copiar_euro.Height = 18
                frm_cotacao.copiar_yen.Height = 18
                
                frm_cotacao.copiar_dolar.Width = 54
                frm_cotacao.copiar_euro.Width = 54
                frm_cotacao.copiar_yen.Width = 54
                

                frm_cotacao.lbl_valor_peso_dolar_frm.Caption = elemento_Dolar.text
                frm_cotacao.lbl_valor_peso_euro_frm.Caption = elemento_Euro.text
                frm_cotacao.lbl_valor_peso_yen_frm.Caption = Format(CDbl(elemento_Yen.text), "0.00")
                
                frm_cotacao.lbl_valor_peso_dolar_frm.Visible = True
                frm_cotacao.lbl_valor_peso_euro_frm.Visible = True
                frm_cotacao.lbl_valor_peso_yen_frm.Visible = True
                
                frm_cotacao.lbl_euro.Visible = True
                frm_cotacao.lbl_dolar.Visible = True
                frm_cotacao.lbl_yen.Visible = True
                
                frm_cotacao.copiar_euro.Visible = True
                frm_cotacao.copiar_dolar.Visible = True
                frm_cotacao.copiar_yen.Visible = True

                frm_cotacao.lbl_valor_das_moedas.Visible = True
                frm_cotacao.lbl_dia_anterior.Visible = False
            End If

        
        driver.Quit
            
            
End Sub

