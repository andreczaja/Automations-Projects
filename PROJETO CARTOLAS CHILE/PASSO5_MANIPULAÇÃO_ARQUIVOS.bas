Attribute VB_Name = "PASSO5_MANIPULA��O_ARQUIVOS"
Sub RenomearArquivos()
    Dim linha As Long
    Dim linha_fim As Long
    Dim pastaDownloads As String
    Dim pastaDestino As String
    Dim nomeArquivoNovo As String
    Dim fs As Object
    Dim arquivos As Object
    Dim arquivoMaisRecente As Object
    Dim Arquivo As Object
    Dim extensao As String
    Dim dataModificacaoMaisRecente As Date

    ' Defina a pasta de downloads onde os arquivos foram salvos
    pastaDownloads = Environ("USERPROFILE") & "\Downloads\"
    ' Defina a pasta de destino para onde os arquivos renomeados ser�o movidos
    pastaDestino = Environ("USERPROFILE") & "\OneDrive - Electrolux\Projetos de Automatiza��o\CARTOLAS DIARIAS - PROJETO CONTABILIDADE\Cartolas Renomeadas\"


    ' Crie um objeto FileSystemObject para manipular arquivos
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set arquivos = fs.GetFolder(pastaDownloads).Files
verificar_pasta_novamente:
    ' Verifique se a pasta de destino existe, caso contr�rio, crie-a
    If Not fs.FolderExists(pastaDestino) Then
        MsgBox ("N�o foi encontrada a pasta atalho do Sharepoint dos documentos extra�dos no caminho" & pastaDestino & _
        ". Por favor, crie, e SOMENTE AP�S ISSO, clique em OK."), vbOKOnly
        GoTo verificar_pasta_novamente
    End If


    Set aba_acesso_bancos = ThisWorkbook.Sheets("Acessos Bancos")
    Set tabela_acesso_bancos = aba_acesso_bancos.ListObjects("Tabela_Acesso_Bancos")
    Set aba_contas = ThisWorkbook.Sheets("Contas")
    Set tabela_contas = aba_contas.ListObjects("Tabela_Contas")
    
    ' Obtenha a �ltima linha com dados na tabela_contas
    linha_fim = aba_contas.Range("A999").End(xlUp).Row

    ' Percorra as linhas de baixo para cima
    For linha = linha_fim To 2 Step -1
        ' Verifique se a conta tem movimentos na coluna E
        If aba_contas.Range("E" & linha).Value = "OK" Then
            ' Obtenha o banco e o n�mero da conta
            banco = aba_contas.Range("A" & linha).Value
            cuenta = aba_contas.Range("C" & linha).Value

            ' Inicialize a vari�vel de data de modifica��o mais recente
            dataModificacaoMaisRecente = DateSerial(1900, 1, 1)
            Set arquivoMaisRecente = Nothing

            ' Procure o arquivo mais recente na pasta de downloads com a extens�o especificada
            For Each Arquivo In arquivos
                If Arquivo.DateLastModified > dataModificacaoMaisRecente Then
                    dataModificacaoMaisRecente = Arquivo.DateLastModified
                    Set arquivoMaisRecente = Arquivo
                End If
            Next Arquivo

            ' Verifique se encontramos um arquivo correspondente
            If Not arquivoMaisRecente Is Nothing Then
                ' Defina o novo nome do arquivo
                nomeArquivoNovo = banco & " - " & cuenta & "." & fs.GetExtensionName(arquivoMaisRecente.Name)

                ' Renomeie o arquivo e mova-o para a pasta de destino
                On Error Resume Next
                Kill (pastaDestino & nomeArquivoNovo)
                fs.MoveFile arquivoMaisRecente.Path, pastaDestino & nomeArquivoNovo
                Debug.Print nomeArquivoNovo
            Else
                MsgBox "Nenhum arquivo encontrado para a conta " & cuenta & " do banco " & banco, vbExclamation
            End If
        End If

ProximaLinha:
    Next linha

    ' Limpeza de objetos
    Set fs = Nothing
End Sub

