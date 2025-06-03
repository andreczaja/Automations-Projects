Attribute VB_Name = "PASSO1_PORT_DEV"
Private contador As Integer
Sub extracao_relatorio()
    Dim driver As New EdgeDriver
    Dim data_inicial, elemento_home, elemento_home_barra_lateral, elemento_pesquisar, elemento_login, elemento_data_inicial, elemento_data_final, _
    elemento_fechar_barra_lateral, elemento_status_relatorio, elemento_gerar_relatorio, elemento_pesquisar2, elemento_download As String
    Dim by As New by


    Set aba_relatorio_portal_devolucoes = ThisWorkbook.Sheets("Relatório Portal de Devoluções")
    Set tabela_aba_relatorio_portal_devolucoes = aba_relatorio_portal_devolucoes.ListObjects("Tabela_Relatório_Portal_de_Devoluções")
    
    elemento_login = "/html/body/app-root/div[1]/div/div/div/app-login/div/div/div/div/div/div/div[2]/div/div/div[1]/button"
                    ' /html/body/app-root/div/div/div/div/app-login/div/div/div/div/div/div/div[2]/div/div/div[1]/button
    elemento_home = "/html/body/app-root/div/div/div[2]/div[1]/app-nav-bar/nav/div/div[1]/button[1]/i"
    elemento_home_barra_lateral = "/html/body/app-root/div/div/div[2]/app-side-bar-mobile/div[1]/app-side-bar-menu/ul/div[1]/li/a"
    elemento_pesquisar = "/html/body/app-root/div/div/div[1]/app-side-bar-desktop/nav/app-side-bar-menu/ul/div[5]/li/a"
    elemento_fechar_barra_lateral = "/html/body/app-root/div[2]/div/button"
    elemento_data_inicial = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[1]/div/div/div/div[1]/div[3]/div/div/input[1]"
    elemento_data_final = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[1]/div/div/div/div[1]/div[3]/div/div/input[2]"
    elemento_gerar_relatorio = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[1]/div/div/div/div[7]/div/button[2]"
    elemento_status_relatorio = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[2]/div/div/div[2]/div/div/table/tbody/tr[1]/td[4]"
    elemento_download = "/html/body/app-root/div/div/div[2]/div[2]/app-search-occurrence/app-listing-occurences/div[3]/div[2]/div/div/div[2]/div/div/table/tbody/tr[1]/td[5]/button"
    
    driver.Get "https://portaldevolucoes.electrolux.com.br/login"
    driver.Window.Maximize
    
    driver.FindElementByXPath(elemento_login).Click
    
    
    ' verificacao simples se a pagina está carregada
    Do Until driver.IsElementPresent(by.XPath(elemento_home_barra_lateral)) Or contador = 10
        If driver.IsElementPresent(by.XPath(elemento_login)) Then
            driver.FindElementByXPath(elemento_login).Click
        End If
        Application.Wait (Now + TimeValue("00:00:01"))
        contador = contador + 1
    Loop
    

    driver.Get "https://portaldevolucoes.electrolux.com.br/search_occurrence/default"
   ' começo do preenchimento do formulario de datas e solicitacao de geracao de relatorio para extrair
    On Error Resume Next
    Application.Wait (Now + TimeValue("00:00:02"))
    driver.FindElementByXPath(elemento_data_inicial).Click
    driver.FindElementByXPath(elemento_data_inicial).SendKeys Format(Date - 90, "dd/mm/yyyy")
    driver.FindElementByXPath(elemento_data_final).Click
    driver.FindElementByXPath(elemento_data_final).SendKeys Format(Date, "dd/mm/yyyy")
    driver.FindElementByXPath(elemento_gerar_relatorio).Click
    Application.Wait (Now + TimeValue("00:00:10"))
    On Error GoTo 0
    ' espera até que o relatorio solicitado esteja com status de "Concluído" quando estiver, clica no elemento_download
    Do Until driver.FindElementByXPath(elemento_status_relatorio).text = "Concluído"
        driver.Refresh
        Application.Wait (Now + TimeValue("00:00:02"))
        If driver.FindElementByXPath(elemento_status_relatorio).text <> "Concluído" And _
        driver.FindElementByXPath(elemento_status_relatorio).text <> "Em processamento" Then
        
            MsgBox "Erro ao baixar o relatório, por favor verifique!", vbOKOnly
            End
        End If
    Loop
    
    driver.FindElementByXPath(elemento_download).Click
    Application.Wait (Now + TimeValue("00:00:15"))
    
    driver.Quit

    ' chama a funçao que tira o arquivo da pasta downloads e coloca na pasta correta, nesse caso,
    ' é a pasta atalho do sharepoint CobranaRegionais presente no caminho > Automações > Macro de Cobrança > Arquivo TXT SAP
    Call manipular_arquivo

End Sub

Sub manipular_arquivo()
 
    Dim novo_caminho, nomeArquivoNovo, primeira_ocorrencia_lista As String
    Dim objFSO As Object
    Dim PastaDownloads As Object
    Dim OBJarquivo As Object, OBJarquivo_mais_recente As Object
    Dim dataModificacaoMaisRecente As Date, dataModificacao As Date
    
    ' Criação do objeto FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Seleção da pasta de destino
    Folder = BuscarPasta("", False)
    
    ' Definição da pasta Downloads
    Set PastaDownloads = objFSO.GetFolder(VBA.Environ("USERPROFILE") & "\Downloads")
    dataModificacaoMaisRecente = #1/1/1900# ' Inicializa com uma data antiga
    
    ' Loop para identificar o arquivo mais recente
    For Each OBJarquivo In PastaDownloads.Files
        On Error Resume Next ' Ignora erros ao acessar arquivos problemáticos
        dataModificacao = FileDateTime(OBJarquivo)
        On Error GoTo 0 ' Retorna ao tratamento normal de erros
        
        If dataModificacao > dataModificacaoMaisRecente Then
            dataModificacaoMaisRecente = dataModificacao
            Set OBJarquivo_mais_recente = OBJarquivo
        End If
    Next OBJarquivo
    
    ' Define o novo nome do arquivo
    nomeArquivoNovo = "Relatório Portal Devoluções." & objFSO.GetExtensionName(OBJarquivo_mais_recente)
    novo_caminho = Folder & "\" & nomeArquivoNovo
    
    
    ' exclui o arquivo antigo, se existir
    On Error Resume Next
    Kill novo_caminho
    On Error GoTo 0
    
    ' Renomeia e move o arquivo para a pasta de destino
    objFSO.MoveFile OBJarquivo_mais_recente.Path, novo_caminho
    
    
    primeira_ocorrencia_lista = aba_relatorio_portal_devolucoes.Range("A2").Value
    ' Atualiza a tabela
    tabela_aba_relatorio_portal_devolucoes.QueryTable.Refresh False
    
    contador = 1
    Do Until contador = 5
        If primeira_ocorrencia_lista = aba_relatorio_portal_devolucoes.Range("A2").Value Then
            tabela_aba_relatorio_portal_devolucoes.QueryTable.BackgroundQuery = False
            tabela_aba_relatorio_portal_devolucoes.QueryTable.Refresh False
        End If
        contador = contador + 1
    Loop
    
    MsgBox "Relatório extraído e transferido para a pasta correta. Por favor verifique se a base na aba Relatório Portal de " & _
    "Devoluções foi atualizada, caso contrário, clique no botão Atualizar", vbOKOnly

End Sub

