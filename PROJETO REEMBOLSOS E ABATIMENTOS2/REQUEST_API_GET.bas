Attribute VB_Name = "REQUEST_API_GET"
Public API_token As String, nome_anexo As String, generatorReferenceCode As String, response  As String, texto_tramite As String
Public arquivo_anexo_chamado_atual As Workbook
Public rngEncontrado As Range
Public aba_1_arquivo_anexo_chamado_atual As Worksheet
Public extensoesValidas As Variant
Public Const pastaAnexos As String = "\Anexos Chamados\"
Public anexos, anexo
Public json As Object
Public Function APIBuscaChamadoUnico() As Boolean
    ' Declara��o de vari�veis
    Dim http As Object ' Objeto para realizar requisi��es HTTP
    Dim i As Long, j As Long ' Contadores para loops
    Dim proceedings, tramite ' Vari�veis para armazenar objetos JSON de procedimentos e tr�mites
    Dim link, extensao As String ' Vari�veis para armazenar links e extens�es de arquivos
    Dim numero_tramites_chamado, ultimo_tramite_verificado, qtde_anexos As Long ' Vari�veis num�ricas para informa��es sobre tr�mites e anexos
    Dim anexoValidoEncontrado As Boolean ' Flag para indicar se um anexo v�lido foi encontrado

    ' Inicializa a quantidade de anexos encontrados
    qtde_anexos = 0
    ' Define o valor padr�o de retorno da fun��o como Falso
    APIBuscaChamadoUnico = False
    ' Define um array com as extens�es de arquivo v�lidas
    extensoesValidas = Array(".xlsx", ".xls", ".xlsm", ".xlsb", ".csv")
    ' Define um c�digo de refer�ncia do gerador (parece ser uma identifica��o)
    generatorReferenceCode = "CONTASARECEBERELECT-638618399315"
    ' Obt�m o token de autentica��o da API da planilha "API KEY"
    API_token = ThisWorkbook.Sheets("API KEY").Range("A1").Value

    ' Cria um objeto XMLHTTP para realizar a requisi��o GET
    Set http = CreateObject("MSXML2.XMLHTTP")
    ' Monta o link da API para buscar informa��es do chamado espec�fico
    link = "https://electrolux.ellevo.com/api/v1/ticket/ticket-list/" & chamado
    ' Envia a requisi��o GET para a API e armazena a resposta
    response = EnviarRequisicao("GET", http, link, "")
    ' Se a resposta da API estiver vazia, encerra a fun��o
    If response = "" Then Exit Function

    ' Utiliza a biblioteca JsonConverter para analisar a resposta JSON
    Set json = JsonConverter.ParseJson(response)
    ' Define o objeto 'proceedings' com a se��o de procedimentos do JSON
    Set proceedings = json("proceedings")
    ' Obt�m o n�mero total de tr�mites do chamado
    numero_tramites_chamado = proceedings.Count

    ' Verifica se o chamado ainda n�o existe no dicion�rio de chamados pendentes (primeiro tr�mite sendo analisado)
    If Not dictChamadosPendentes.exists(CDbl(chamado)) Then
        ' Verifica se existe a se��o de anexos no n�vel do ticket principal
        If Not IsNull(json("integrationApiTicketAttachments")) Then
             'Define o objeto 'anexos' com os anexos do ticket principal
            Set anexos = json("integrationApiTicketAttachments")
             'Chama a fun��o para verificar se h� um arquivo v�lido entre os anexos
            If BuscaArquivoValido Then
                 'Se um arquivo v�lido for encontrado, define o retorno da fun��o como Verdadeiro
                APIBuscaChamadoUnico = True
            Else
                 'Se nenhum arquivo v�lido for encontrado, gera um token de tr�mite com o status de anexo incorreto
                Call GerarTokenTramite("ANEXO_INCORRETO")
                 'Define o retorno da fun��o como Falso
                APIBuscaChamadoUnico = False
            End If
        Else
            ' Se n�o houver anexos no n�vel do ticket principal, itera pelos tr�mites
            For i = 1 To numero_tramites_chamado
                ' Define o objeto 'tramite' com o tr�mite atual
                Set tramite = proceedings(i)
                ' Verifica se o tr�mite possui a se��o de anexos
                If Not IsNull(tramite("integrationApiProceedingAttachments")) Then
                    ' Define o objeto 'anexos' com os anexos do tr�mite
                    Set anexos = tramite("integrationApiProceedingAttachments")
                    ' Incrementa a contagem total de anexos
                    qtde_anexos = qtde_anexos + anexos.Count
                    ' Chama a fun��o para verificar se h� um arquivo v�lido entre os anexos do tr�mite
                    If BuscaArquivoValido Then
                        ' Se um arquivo v�lido for encontrado, define o retorno da fun��o como Verdadeiro
                        APIBuscaChamadoUnico = True
                        ' Chama a sub-rotina para incluir ou atualizar o chamado no dicion�rio de pendentes
                        Call IncluirAtualizarChamadoPendente
                        ' Encerra a fun��o
                        Exit Function
                    End If
                End If
            Next i
            ' Se nenhum arquivo v�lido for encontrado ap�s verificar todos os tr�mites, gera um token de aviso de falta de anexo
            Call GerarTokenTramite("AVISO_FALTA_DE_ANEXO")
            ' Define o retorno da fun��o como Falso
            APIBuscaChamadoUnico = False
        End If
    Else
        ' Se o chamado j� existe no dicion�rio de pendentes, verifica se h� novos tr�mites
        If numero_tramites_chamado > dictChamadosPendentes(CDbl(chamado)) Then
            ' Obt�m o n�mero do �ltimo tr�mite verificado
            ultimo_tramite_verificado = dictChamadosPendentes(CDbl(chamado))
            ' Garante que o �ndice inicial seja pelo menos 1
            If ultimo_tramite_verificado < 1 Then ultimo_tramite_verificado = 1
            ' Itera pelos novos tr�mites
            For i = ultimo_tramite_verificado + 1 To numero_tramites_chamado
                ' Define o objeto 'tramite' com o tr�mite atual
                Set tramite = proceedings(i)
                ' Verifica se o tr�mite possui a se��o de anexos
                If Not IsNull(tramite("integrationApiProceedingAttachments")) Then
                    ' Define o objeto 'anexos' com os anexos do tr�mite
                    Set anexos = tramite("integrationApiProceedingAttachments")
                    ' Incrementa a contagem total de anexos
                    qtde_anexos = qtde_anexos + anexos.Count
                    ' Chama a fun��o para verificar se h� um arquivo v�lido entre os anexos do tr�mite
                    If BuscaArquivoValido Then
                        ' Se um arquivo v�lido for encontrado, define o retorno da fun��o como Verdadeiro
                        APIBuscaChamadoUnico = True
                        ' Chama a sub-rotina para incluir ou atualizar o chamado no dicion�rio de pendentes
                        Call IncluirAtualizarChamadoPendente
                        ' Encerra a fun��o
                        Exit Function
                    End If
                End If
            Next i
            Call APITrocaResponsavelChamado(3)
            ' Se nenhum novo arquivo v�lido for encontrado, define o retorno da fun��o como Falso
            APIBuscaChamadoUnico = False
        Else
            ' Se n�o houver novos tr�mites, define o retorno da fun��o como Falso
            APIBuscaChamadoUnico = False
        End If

    End If
End Function
Private Function BuscaArquivoValido() As Boolean
    ' Define o valor padr�o de retorno da fun��o como Falso
    BuscaArquivoValido = False
    ' Itera por cada anexo na cole��o 'anexos'
    For Each anexo In anexos
        ' Obt�m a extens�o do arquivo em letras min�sculas
        extensao = LCase(anexo("extension"))
        ' Chama a fun��o para verificar se a extens�o � v�lida
        If IsExtensaoValida(extensao) Then
            ' Obt�m o nome do anexo
            nome_anexo = anexo("name")
            ' Obt�m o link para download do anexo
            link = anexo("link")
            ' Chama a fun��o para baixar o anexo
            If DownloadAnexo(link) Then
                ' Se o download for bem-sucedido, define o retorno da fun��o como Verdadeiro
                BuscaArquivoValido = True
                ' Encerra a fun��o, pois um arquivo v�lido foi encontrado
                Exit Function
            End If
        End If
    Next anexo

End Function
Private Function IsExtensaoValida(ByVal extensao As String) As Boolean
    ' Declara��o de um contador para o loop
    Dim l As Long
    ' Itera por cada elemento no array 'extensoesValidas'
    For l = LBound(extensoesValidas) To UBound(extensoesValidas)
        ' Verifica se a extens�o passada como par�metro � igual a um dos elementos do array
        If extensao = extensoesValidas(l) Then
            ' Se a extens�o for v�lida, define o retorno da fun��o como Verdadeiro
            IsExtensaoValida = True
            ' Encerra a fun��o, pois a extens�o foi validada
            Exit Function
        End If
    Next l
    ' Se o loop terminar sem encontrar uma correspond�ncia, a extens�o n�o � v�lida
    IsExtensaoValida = False
End Function


Private Function DownloadAnexo(ByVal link As String) As Boolean
    ' Declara��o de objetos
    Dim http As Object ' Objeto para realizar requisi��es HTTP
    Dim stream As Object ' Objeto para manipular streams de dados
    Dim filePath As String ' Vari�vel para armazenar o caminho do arquivo

    ' Cria um objeto XMLHTTP para realizar a requisi��o GET
    Set http = CreateObject("MSXML2.XMLHTTP")
    ' Abre uma conex�o GET para o link fornecido, de forma s�ncrona (aguarda a conclus�o)
    http.Open "GET", link, False
    ' Define o header de autoriza��o com o token da API
    http.setRequestHeader "Authorization", "Bearer " & API_token
    ' Envia a requisi��o HTTP
    http.Send

    ' Verifica se o status da resposta HTTP n�o � 200 (OK), indicando um erro
    If http.Status <> 200 Then Exit Function

    ' Define o caminho completo para salvar o arquivo baixado
    filePath = ThisWorkbook.Path & pastaAnexos & chamado & ".xls"
    ' Verifica se o arquivo j� existe e, se existir, o exclui
    If Dir(filePath) <> "" Then Kill filePath

    ' Cria um objeto ADODB.Stream para manipular o corpo da resposta
    Set stream = CreateObject("ADODB.Stream")
    ' Define o tipo do stream como bin�rio (1)
    stream.Type = 1
    ' Abre o stream
    stream.Open
    ' Escreve o corpo da resposta HTTP no stream
    stream.Write http.responseBody
    ' Salva o conte�do do stream em um arquivo no caminho especificado, sobrescrevendo se j� existir (2)
    stream.SaveToFile filePath, 2
    ' Fecha o stream
    stream.Close
    ' Abre o arquivo baixado como um workbook
    Set arquivo_anexo_chamado_atual = Workbooks.Open(filePath)
    ' Procura pela c�lula que cont�m o texto "N�MERO DA OCORRENCIA ( OC )" na primeira planilha
    Set rngEncontrado = arquivo_anexo_chamado_atual.Sheets(1).Range("A:AA").Find("N�MERO DA OCORRENCIA ( OC )", , , xlWhole)
    ' Verifica se a c�lula n�o foi encontrada OU se o arquivo possui mais de uma planilha
    If rngEncontrado Is Nothing Or arquivo_anexo_chamado_atual.Sheets.Count > 1 Then
        ' Fecha o arquivo sem salvar as altera��es
        arquivo_anexo_chamado_atual.Close False
        ' Define o retorno da fun��o como Falso (download falhou ou arquivo inv�lido)
        DownloadAnexo = False
        ' Chama a sub-rotina para registrar no relat�rio que o chamado possui um anexo incorreto/fora do padr�o
        Call AlimentarDicionario_Relatorio_Processamento("Chamados com anexo incorreto/fora do padr�o: ", chamado)
        ' Encerra a fun��o
        Exit Function
    End If

    ' Se todas as verifica��es passarem, define o retorno da fun��o como Verdadeiro (download bem-sucedido e arquivo aparentemente v�lido)
    DownloadAnexo = True
End Function
