Attribute VB_Name = "REQUEST_GEMINI"
Public resumo As String

Function ResumirArquivoGemini() As String
    Dim http As Object
    Dim JSON As String
    Dim apiKey As String
    Dim url As String
    Dim responseText As String
    Dim result As Object
    
    ' Definição da API Key
    apiKey = "apikey"
    
    ' URL da API corrigida para v1beta
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" & apiKey
    
    ' Criando JSON para requisição
    JSON = "{""contents"": [{""parts"": [{""text"": """ & Replace(resumo, """", "'") & """}]}]}"
    
    ' Criando objeto HTTP para requisição
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Configurando e enviando requisição POST
    With http
        .Open "POST", url, False
        .SetRequestHeader "Content-Type", "application/json"
        .Send JSON
    End With
    
    ' Exibindo JSON enviado e resposta no debug
    Debug.Print "JSON Enviado: " & JSON
    Debug.Print "Resposta da API: " & http.responseText
    
    ' Processando resposta JSON
    responseText = http.responseText
    Set result = JsonConverter.ParseJson(responseText)
    
    Debug.Print VarType(result)
    ' Extraindo texto da resposta
    If Not result Is Nothing Then
        If result.Exists("candidates") Then
            Dim candidate As Object
            Dim contentPart As Object
            Dim fullText As String
            
            ' Iterando sobre os candidatos
            For Each candidate In result("candidates")
                If candidate.Exists("content") Then
                    ' Iterando sobre as partes do conteúdo
                    For Each contentPart In candidate("content")("parts")
                        If contentPart.Exists("text") Then
                            fullText = fullText & contentPart("text") & vbNewLine
                        End If
                    Next
                End If
            Next
            ResumirArquivoGemini = fullText
        Else
            ResumirArquivoGemini = "Erro na resposta da API."
        End If
    End If

End Function

Sub TestarGemini()
    ' Obtendo o prompt da célula B4
    resumo = ThisWorkbook.Sheets(1).Range("B4").Value
    Debug.Print "Resumo original: " & resumo
    
    ' Enviando prompt e obtendo resposta
    ThisWorkbook.Sheets(1).Range("B10").Value = ResumirArquivoGemini
End Sub

