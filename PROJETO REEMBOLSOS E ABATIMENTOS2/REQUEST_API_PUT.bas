Attribute VB_Name = "REQUEST_API_PUT"
Public Sub APITrocaResponsavelChamado(responsavel As Integer)

    Const CodigoReferenciaAutor As String = "CONTASARECEBERELECT-638618399315" ' Código de referência do autor do chamado (constante)
    Dim CodigoReferenciaResponsavel As String ' Código de referência do responsável pelo chamado
    Dim url As String ' URL da API para trocar o responsável
    Dim jsonPayload As String ' Payload JSON a ser enviado na requisição PUT
    Dim jsonParts As String ' Variável auxiliar para construir o payload JSON


    ' Obtém o token da API da planilha "API KEY"
    API_token = ThisWorkbook.Sheets("API KEY").Range("A1").Value
    ' Obtém o código de referência do responsável pela planilha "API KEY"
    CodigoReferenciaResponsavel = ThisWorkbook.Sheets("API KEY").Range("A" & responsavel).Value
    ' Monta a URL da API para trocar o responsável do chamado
    url = "https://electrolux.ellevo.com/api/v1/ticket/ticket-involved/" & chamado

    ' Cria o objeto HTTP para realizar a requisição
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Inicia a construção do payload JSON
    jsonParts = "{"

    ' Adiciona o código de referência do autor ao payload, se existir
    If CodigoReferenciaAutor <> "" Then
        jsonParts = jsonParts & """authorReferenceCode"": """ & CodigoReferenciaAutor & """"
    End If

    ' Adiciona o código de referência do responsável ao payload, se existir
    If CodigoReferenciaResponsavel <> "" Then
        If jsonParts <> "{" Then jsonParts = jsonParts & ", " ' Adiciona vírgula se já houver outra parte no payload
        jsonParts = jsonParts & """responsibleReferenceCode"": """ & CodigoReferenciaResponsavel & """"
    End If
    
    jsonParts = jsonParts & ", " & """status"": ""1"""

    ' Fecha o payload JSON
    jsonParts = jsonParts & "}"
    jsonPayload = jsonParts ' Atribui o payload completo à variável jsonPayload

    ' Configura e envia a requisição PUT para a API
    http.Open "PUT", url, False ' Abre a conexão PUT para a URL
    http.setRequestHeader "Content-Type", "application/json" ' Define o tipo de conteúdo como JSON
    http.setRequestHeader "Authorization", "Bearer " & API_token ' Define o header de autenticação com o token
    http.Send jsonPayload ' Envia o payload JSON no corpo da requisição
    
    If responsavel = 2 Then
        ' Verifica se a quantidade de NFDs é maior que 1 e define uma condição
        If qtde_NFD_OC_chamado = "Acima de 01" Then
            condicao_chamado_supplier = True
        End If
        
        texto_tramite = "Chamado atrelado à crédito Supplier Card. Número do documento: " & num_doc_supplier & ". Favor responsável dar continuidade"
    ElseIf responsavel = 3 Then
        texto_tramite = "Chamado com nova inconsistência nos dados informados. Favor responsável dar continuidade"
    End If
    
    apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/" & chamado & "/proceeding"
    requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                        """private"":false," & _
                        """status"":1," & _
                        """description"":""" & texto_tramite & """}"

    ' Envia a requisição POST para a API
    Call EnviarRequisicao("POST", http, apiUrl, requestBody)
    
    linha_fim_aba_historico_chamados_pendentes = aba_historico_chamados_pendentes.Range("A1048576").End(xlUp).Row
    For linha = linha_fim_aba_historico_chamados_pendentes To 2 Step -1
        If aba_historico_chamados_pendentes.Range("A" & linha).Value = CLng(chamado) Then
            aba_historico_chamados_pendentes.Rows(linha).Delete
        End If
    Next linha

    ' Libera a memória do objeto HTTP
    Set http = Nothing

End Sub

