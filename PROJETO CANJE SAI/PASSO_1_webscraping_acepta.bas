Attribute VB_Name = "PASSO_1_webscraping_acepta"
Option Explicit

Sub verificacao_notas_acepta_()
    
    Dim driver As New EdgeDriver
    Dim elementoInput, alert, elemento_data_inicial, elemento_data_final, elemento_botao_mais, elemento_estado_documento, elemento_botao_salvar As WebElement
    Dim dia_min_mes_anterior, mes_anterior, ano_mes_anterior As String
    Dim i As Integer
    Dim by As New by
    
    If frmDate.lbl_data_inicio = "" And frmDate.lbl_data_final = "" Then
        frmDate.Show
    End If
  
    dia_min_mes_anterior = Left(frmDate.lbl_data_inicio, 2)
    If Month(Date - 20) < 10 Then
        mes_anterior = 0 & (Month(Date - 20))
    Else
        mes_anterior = (Month(Date - 20))
    End If
    ano_mes_anterior = Year(Date - 20)
    

    
    
    ' Abrir o navegador Edge
    driver.Get "https://escritorio.acepta.com/ext.php?r=https://escritorio.acepta.com/dashboard%3Fapp_dinamica=dashboard%26session_id=008db7194a7a6f55f74429d4fe8ed001a5381531552%26rutUsuario=300001734%26rutCliente=76163495%26aplicacion=DTE%26%26inicio=%26mensaje_pantalla="
    driver.SwitchToAlert.Accept
    driver.Window.Maximize
    
    
    
    ' setando o elemento para o botao de documentos recibidos
    On Error GoTo Erro_Login_Acepta
    Set elementoInput = driver.FindElementById("badge_2150")
0:
    

    ' clicando no botao de documentos emitidos
        Set elementoInput = driver.FindElementById("badge_2150")
        elementoInput.Click
    
        Application.Wait (Now + TimeValue("00:00:05"))
    ' setando o elemento para o botao de busqueda avanzada
    Set elementoInput = driver.FindElementByXPath("/html/body/div[9]/div[1]/section/div[2]/div/div/div[2]/div[1]/form/div[7]/a")
    
    ' clicando no botao de busqueda avanzada
        elementoInput.Click
        
    Set elementoInput = driver.FindElementByXPath("/html/body/div[9]/div[1]/section/div[2]/div/div/div[2]/div[2]/form/div[1]/div/div/select")
      elementoInput.Click
    ' setando o elemento para a opcao de data de emissao
    Set elementoInput = driver.FindElementByXPath("/html/body/div[9]/div[1]/section/div[2]/div/div/div[2]/div[2]/form/div[1]/div/div/select/option[2]")
    
    ' clicando no botao de buscar
    elementoInput.Click

    ' Data inicial de emissao
            Set elemento_data_inicial = driver.FindElementByXPath("/html/body/div[9]/div[1]/section/div[2]/div/div/div[2]/div[2]/form/div[3]/div/input")
        
            With elemento_data_inicial
                .Click
                .Clear
                .SendKeys (dia_min_mes_anterior & mes_anterior & ano_mes_anterior)
            End With
                
    ' data final de emissao
            Set elemento_data_final = driver.FindElementByXPath("/html/body/div[9]/div[1]/section/div[2]/div/div/div[2]/div[2]/form/div[5]/div/input")
        
            With elemento_data_final
                .Click
                .Clear
                .SendKeys (Format(Date, "dd.mm.yyyy"))
            End With
            
    ' selecionando tipo de documento: factura eletronica
    Set elementoInput = driver.FindElementByXPath("/html/body/div[9]/div[1]/section/div[2]/div/div/div[2]/div[2]/form/div[15]/div/div/select")
        elementoInput.Click
        
    ' selecionando tipo de documento: factura eletronica
    Set elementoInput = driver.FindElementByXPath("/html/body/div[9]/div[1]/section/div[2]/div/div/div[2]/div[2]/form/div[15]/div/div/select/option[5]")
        elementoInput.Click
    
            
    ' clicando no botao de buscar
    Set elementoInput = driver.FindElementByXPath("/html/body/div[9]/div[1]/section/div[2]/div/div/div[2]/div[2]/form/div[25]/div/input")
        elementoInput.Click
        
        
        
    ' clicando no botao de exportar relatorio
    Application.Wait (Now + TimeValue("00:00:20"))
    Set elementoInput = driver.FindElementByXPath("/html/body/div[9]/div[1]/section/div[2]/div/div/div[2]/div[3]/div/div[1]/div[3]/input")
        elementoInput.Click
        
    driver.SwitchToAlert.Accept
        
        
    ' clicando na sessao de Reportes para baixar o relatorio
    
    Set elementoInput = driver.FindElementById("badge_2109")
        elementoInput.Click
        
          'setando e clicando no ultimo reporte gerado
        Do Until driver.IsElementPresent(by.ID("1Descargargrilla_reportesNEW"), 1000) = True
            driver.Refresh
        Loop
            For i = 1 To 10
                If driver.IsElementPresent(by.XPath("/html/body/div[9]/div/section/div[2]/div/div[1]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[1]/td[" & i & "]/div/button")) Then
                        'salvar documento
                      driver.FindElementByXPath("/html/body/div[9]/div/section/div[2]/div/div[1]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[1]/td[" & i & "]/div/button").Click
                      Application.Wait (Now + TimeValue("00:00:02"))
                      driver.FindElementByXPath("/html/body/div[9]/div/section/div[2]/div/div[1]/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[1]/td[" & i & "]/div/ul/li/a[1]").Click
                      Exit For
                End If
            Next i
        
        
            
        ' Aguarde um tempo suficiente para o download ser iniciado
    Application.Wait (Now + TimeValue("00:00:10"))
    
    driver.Close
    
    MoverUltimoArquivoBaixado
    
    End
    
Erro_Login_Acepta:

    MsgBox "Por favor, faça login no Acepta no navegador. Após isso clique em OK.", vbOKOnly
    
    GoTo 0

End Sub

Sub MoverUltimoArquivoBaixado()
    Dim sourcePath As String
    Dim fileName As String
    Dim lastModified As Date
    Dim latestFile, novo_nome_arquivo As String
    Dim fileDate As Date
    
    ' Caminho da pasta onde o Chrome salva os downloads
    sourcePath = Environ("USERPROFILE") & "\Downloads\"

    CaminhoPasta = "C:\Users\CardoAnd03\OneDrive - Electrolux\CANJE - CLIENTES SAI\EXPORTS SAP - CANJE\"
    ' Inicializa a variável para armazenar a última data de modificação
    lastModified = 0

    ' Obtém o primeiro arquivo da pasta
    fileName = Dir(sourcePath & "*.*")

    ' Loop através de todos os arquivos na pasta
    Do While fileName <> ""
        ' Obtém a data de modificação do arquivo
        fileDate = FileDateTime(sourcePath & fileName)

        ' Verifica se a data de modificação é maior que a última registrada
        If fileDate > lastModified Then
            lastModified = fileDate
            latestFile = fileName
        End If

        ' Obtém o próximo arquivo da pasta
        fileName = Dir
    Loop

    ' Verifica se encontrou algum arquivo na pasta
    If latestFile <> "" Then
        ' Combina o caminho completo do arquivo mais recente
        Dim fullpath As String
        fullpath = sourcePath & latestFile

        ' Move o arquivo para a pasta desejada
        Name fullpath As CaminhoPasta & latestFile
        fullpath = CaminhoPasta & latestFile
        On Error Resume Next
        Kill CaminhoPasta & "Reporte Acepta.zip"
        Name fullpath As CaminhoPasta & "Reporte Acepta.zip"
        
        
    Else
        MsgBox "Nenhum arquivo encontrado na pasta de downloads."
    End If
    
    MsgBox "Os Relatórios do Acepta e do SAP foram salvos na pasta designada."
End Sub
