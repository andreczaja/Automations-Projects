Attribute VB_Name = "REQUEST_API_PUT"
Public Sub APITrocaResponsavelChamado(responsavel As Integer)

    Const CodigoReferenciaAutor As String = "CONTASARECEBERELECT-638618399315" ' C�digo de refer�ncia do autor do chamado (constante)
    Dim CodigoReferenciaResponsavel As String ' C�digo de refer�ncia do respons�vel pelo chamado
    Dim url As String ' URL da API para trocar o respons�vel
    Dim jsonPayload As String ' Payload JSON a ser enviado na requisi��o PUT
    Dim jsonParts As String ' Vari�vel auxiliar para construir o payload JSON


    ' Obt�m o token da API da planilha "API KEY"
    API_token = ThisWorkbook.Sheets("API KEY").Range("A1").Value
    ' Obt�m o c�digo de refer�ncia do respons�vel pela planilha "API KEY"
    CodigoReferenciaResponsavel = ThisWorkbook.Sheets("API KEY").Range("A" & responsavel).Value
    ' Monta a URL da API para trocar o respons�vel do chamado
    url = "https://electrolux.ellevo.com/api/v1/ticket/ticket-involved/" & chamado

    ' Cria o objeto HTTP para realizar a requisi��o
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Inicia a constru��o do payload JSON
    jsonParts = "{"

    ' Adiciona o c�digo de refer�ncia do autor ao payload, se existir
    If CodigoReferenciaAutor <> "" Then
        jsonParts = jsonParts & """authorReferenceCode"": """ & CodigoReferenciaAutor & """"
    End If

    ' Adiciona o c�digo de refer�ncia do respons�vel ao payload, se existir
    If CodigoReferenciaResponsavel <> "" Then
        If jsonParts <> "{" Then jsonParts = jsonParts & ", " ' Adiciona v�rgula se j� houver outra parte no payload
        jsonParts = jsonParts & """responsibleReferenceCode"": """ & CodigoReferenciaResponsavel & """"
    End If
    
    jsonParts = jsonParts & ", " & """status"": ""1"""

    ' Fecha o payload JSON
    jsonParts = jsonParts & "}"
    jsonPayload = jsonParts ' Atribui o payload completo � vari�vel jsonPayload

    ' Configura e envia a requisi��o PUT para a API
    http.Open "PUT", url, False ' Abre a conex�o PUT para a URL
    http.setRequestHeader "Content-Type", "application/json" ' Define o tipo de conte�do como JSON
    http.setRequestHeader "Authorization", "Bearer " & API_token ' Define o header de autentica��o com o token
    http.Send jsonPayload ' Envia o payload JSON no corpo da requisi��o
    
    If responsavel = 2 Then
        ' Verifica se a quantidade de NFDs � maior que 1 e define uma condi��o
        If qtde_NFD_OC_chamado = "Acima de 01" Then
            condicao_chamado_supplier = True
        End If
        
        texto_tramite = "Chamado atrelado � cr�dito Supplier Card. N�mero do documento: " & num_doc_supplier & ". Favor respons�vel dar continuidade"
    ElseIf responsavel = 3 Then
        texto_tramite = "Chamado com nova inconsist�ncia nos dados informados. Favor respons�vel dar continuidade"
    End If
    
    apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/" & chamado & "/proceeding"
    requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                        """private"":false," & _
                        """status"":1," & _
                        """description"":""" & texto_tramite & """}"

    ' Envia a requisi��o POST para a API
    Call EnviarRequisicao("POST", http, apiUrl, requestBody)
    
    linha_fim_aba_historico_chamados_pendentes = aba_historico_chamados_pendentes.Range("A1048576").End(xlUp).Row
    For linha = linha_fim_aba_historico_chamados_pendentes To 2 Step -1
        If aba_historico_chamados_pendentes.Range("A" & linha).Value = CLng(chamado) Then
            aba_historico_chamados_pendentes.Rows(linha).Delete
        End If
    Next linha

    ' Libera a mem�ria do objeto HTTP
    Set http = Nothing

End Sub

