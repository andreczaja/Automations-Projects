Attribute VB_Name = "Requisição"
Function ChamarOpenAI(prompt As String) As String
    Dim http As Object
    Dim JSON As Object
    Dim url As String
    Dim apiKey As String
    Dim data As String
    
    ' Definir URL e Chave da API
    url = "https://api.openai.com/v1/chat/completions"
    apiKey = "API_KEY"
    
    ' Criar JSON para envio
    data = "{""model"": ""gpt-3.5-turbo"", ""messages"": [{""role"": ""system"", ""content"": ""Você é um assistente.""}, {""role"": ""user"", ""content"": """ & prompt & """}]}"
    
    ' Criar objeto HTTP
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    
    ' Enviar requisição
    http.Send data
    
    ' Converter resposta JSON
    Set JSON = JsonConverter.ParseJson(http.responseText)
    
    ' Extrair resposta do modelo
    ChamarOpenAI = JSON("choices")(1)("message")("content")
End Function

Sub TesteOpenAI()
    Dim resposta As String
    resposta = ChamarOpenAI("Qual é a capital do Uruguai?")
    Debug.Print resposta ' Exibe a resposta na Janela de Verificação Imediata
End Sub

