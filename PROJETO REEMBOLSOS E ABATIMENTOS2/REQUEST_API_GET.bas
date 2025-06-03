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
    ' Declaração de variáveis
    Dim http As Object ' Objeto para realizar requisições HTTP
    Dim i As Long, j As Long ' Contadores para loops
    Dim proceedings, tramite ' Variáveis para armazenar objetos JSON de procedimentos e trâmites
    Dim link, extensao As String ' Variáveis para armazenar links e extensões de arquivos
    Dim numero_tramites_chamado, ultimo_tramite_verificado, qtde_anexos As Long ' Variáveis numéricas para informações sobre trâmites e anexos
    Dim anexoValidoEncontrado As Boolean ' Flag para indicar se um anexo válido foi encontrado

    ' Inicializa a quantidade de anexos encontrados
    qtde_anexos = 0
    ' Define o valor padrão de retorno da função como Falso
    APIBuscaChamadoUnico = False
    ' Define um array com as extensões de arquivo válidas
    extensoesValidas = Array(".xlsx", ".xls", ".xlsm", ".xlsb", ".csv")
    ' Define um código de referência do gerador (parece ser uma identificação)
    generatorReferenceCode = "CONTASARECEBERELECT-638618399315"
    ' Obtém o token de autenticação da API da planilha "API KEY"
    API_token = ThisWorkbook.Sheets("API KEY").Range("A1").Value

    ' Cria um objeto XMLHTTP para realizar a requisição GET
    Set http = CreateObject("MSXML2.XMLHTTP")
    ' Monta o link da API para buscar informações do chamado específico
    link = "https://electrolux.ellevo.com/api/v1/ticket/ticket-list/" & chamado
    ' Envia a requisição GET para a API e armazena a resposta
    response = EnviarRequisicao("GET", http, link, "")
    ' Se a resposta da API estiver vazia, encerra a função
    If response = "" Then Exit Function

    ' Utiliza a biblioteca JsonConverter para analisar a resposta JSON
    Set json = JsonConverter.ParseJson(response)
    ' Define o objeto 'proceedings' com a seção de procedimentos do JSON
    Set proceedings = json("proceedings")
    ' Obtém o número total de trâmites do chamado
    numero_tramites_chamado = proceedings.Count

    ' Verifica se o chamado ainda não existe no dicionário de chamados pendentes (primeiro trâmite sendo analisado)
    If Not dictChamadosPendentes.exists(CDbl(chamado)) Then
        ' Verifica se existe a seção de anexos no nível do ticket principal
        If Not IsNull(json("integrationApiTicketAttachments")) Then
             'Define o objeto 'anexos' com os anexos do ticket principal
            Set anexos = json("integrationApiTicketAttachments")
             'Chama a função para verificar se há um arquivo válido entre os anexos
            If BuscaArquivoValido Then
                 'Se um arquivo válido for encontrado, define o retorno da função como Verdadeiro
                APIBuscaChamadoUnico = True
            Else
                 'Se nenhum arquivo válido for encontrado, gera um token de trâmite com o status de anexo incorreto
                Call GerarTokenTramite("ANEXO_INCORRETO")
                 'Define o retorno da função como Falso
                APIBuscaChamadoUnico = False
            End If
        Else
            ' Se não houver anexos no nível do ticket principal, itera pelos trâmites
            For i = 1 To numero_tramites_chamado
                ' Define o objeto 'tramite' com o trâmite atual
                Set tramite = proceedings(i)
                ' Verifica se o trâmite possui a seção de anexos
                If Not IsNull(tramite("integrationApiProceedingAttachments")) Then
                    ' Define o objeto 'anexos' com os anexos do trâmite
                    Set anexos = tramite("integrationApiProceedingAttachments")
                    ' Incrementa a contagem total de anexos
                    qtde_anexos = qtde_anexos + anexos.Count
                    ' Chama a função para verificar se há um arquivo válido entre os anexos do trâmite
                    If BuscaArquivoValido Then
                        ' Se um arquivo válido for encontrado, define o retorno da função como Verdadeiro
                        APIBuscaChamadoUnico = True
                        ' Chama a sub-rotina para incluir ou atualizar o chamado no dicionário de pendentes
                        Call IncluirAtualizarChamadoPendente
                        ' Encerra a função
                        Exit Function
                    End If
                End If
            Next i
            ' Se nenhum arquivo válido for encontrado após verificar todos os trâmites, gera um token de aviso de falta de anexo
            Call GerarTokenTramite("AVISO_FALTA_DE_ANEXO")
            ' Define o retorno da função como Falso
            APIBuscaChamadoUnico = False
        End If
    Else
        ' Se o chamado já existe no dicionário de pendentes, verifica se há novos trâmites
        If numero_tramites_chamado > dictChamadosPendentes(CDbl(chamado)) Then
            ' Obtém o número do último trâmite verificado
            ultimo_tramite_verificado = dictChamadosPendentes(CDbl(chamado))
            ' Garante que o índice inicial seja pelo menos 1
            If ultimo_tramite_verificado < 1 Then ultimo_tramite_verificado = 1
            ' Itera pelos novos trâmites
            For i = ultimo_tramite_verificado + 1 To numero_tramites_chamado
                ' Define o objeto 'tramite' com o trâmite atual
                Set tramite = proceedings(i)
                ' Verifica se o trâmite possui a seção de anexos
                If Not IsNull(tramite("integrationApiProceedingAttachments")) Then
                    ' Define o objeto 'anexos' com os anexos do trâmite
                    Set anexos = tramite("integrationApiProceedingAttachments")
                    ' Incrementa a contagem total de anexos
                    qtde_anexos = qtde_anexos + anexos.Count
                    ' Chama a função para verificar se há um arquivo válido entre os anexos do trâmite
                    If BuscaArquivoValido Then
                        ' Se um arquivo válido for encontrado, define o retorno da função como Verdadeiro
                        APIBuscaChamadoUnico = True
                        ' Chama a sub-rotina para incluir ou atualizar o chamado no dicionário de pendentes
                        Call IncluirAtualizarChamadoPendente
                        ' Encerra a função
                        Exit Function
                    End If
                End If
            Next i
            Call APITrocaResponsavelChamado(3)
            ' Se nenhum novo arquivo válido for encontrado, define o retorno da função como Falso
            APIBuscaChamadoUnico = False
        Else
            ' Se não houver novos trâmites, define o retorno da função como Falso
            APIBuscaChamadoUnico = False
        End If

    End If
End Function
Private Function BuscaArquivoValido() As Boolean
    ' Define o valor padrão de retorno da função como Falso
    BuscaArquivoValido = False
    ' Itera por cada anexo na coleção 'anexos'
    For Each anexo In anexos
        ' Obtém a extensão do arquivo em letras minúsculas
        extensao = LCase(anexo("extension"))
        ' Chama a função para verificar se a extensão é válida
        If IsExtensaoValida(extensao) Then
            ' Obtém o nome do anexo
            nome_anexo = anexo("name")
            ' Obtém o link para download do anexo
            link = anexo("link")
            ' Chama a função para baixar o anexo
            If DownloadAnexo(link) Then
                ' Se o download for bem-sucedido, define o retorno da função como Verdadeiro
                BuscaArquivoValido = True
                ' Encerra a função, pois um arquivo válido foi encontrado
                Exit Function
            End If
        End If
    Next anexo

End Function
Private Function IsExtensaoValida(ByVal extensao As String) As Boolean
    ' Declaração de um contador para o loop
    Dim l As Long
    ' Itera por cada elemento no array 'extensoesValidas'
    For l = LBound(extensoesValidas) To UBound(extensoesValidas)
        ' Verifica se a extensão passada como parâmetro é igual a um dos elementos do array
        If extensao = extensoesValidas(l) Then
            ' Se a extensão for válida, define o retorno da função como Verdadeiro
            IsExtensaoValida = True
            ' Encerra a função, pois a extensão foi validada
            Exit Function
        End If
    Next l
    ' Se o loop terminar sem encontrar uma correspondência, a extensão não é válida
    IsExtensaoValida = False
End Function


Private Function DownloadAnexo(ByVal link As String) As Boolean
    ' Declaração de objetos
    Dim http As Object ' Objeto para realizar requisições HTTP
    Dim stream As Object ' Objeto para manipular streams de dados
    Dim filePath As String ' Variável para armazenar o caminho do arquivo

    ' Cria um objeto XMLHTTP para realizar a requisição GET
    Set http = CreateObject("MSXML2.XMLHTTP")
    ' Abre uma conexão GET para o link fornecido, de forma síncrona (aguarda a conclusão)
    http.Open "GET", link, False
    ' Define o header de autorização com o token da API
    http.setRequestHeader "Authorization", "Bearer " & API_token
    ' Envia a requisição HTTP
    http.Send

    ' Verifica se o status da resposta HTTP não é 200 (OK), indicando um erro
    If http.Status <> 200 Then Exit Function

    ' Define o caminho completo para salvar o arquivo baixado
    filePath = ThisWorkbook.Path & pastaAnexos & chamado & ".xls"
    ' Verifica se o arquivo já existe e, se existir, o exclui
    If Dir(filePath) <> "" Then Kill filePath

    ' Cria um objeto ADODB.Stream para manipular o corpo da resposta
    Set stream = CreateObject("ADODB.Stream")
    ' Define o tipo do stream como binário (1)
    stream.Type = 1
    ' Abre o stream
    stream.Open
    ' Escreve o corpo da resposta HTTP no stream
    stream.Write http.responseBody
    ' Salva o conteúdo do stream em um arquivo no caminho especificado, sobrescrevendo se já existir (2)
    stream.SaveToFile filePath, 2
    ' Fecha o stream
    stream.Close
    ' Abre o arquivo baixado como um workbook
    Set arquivo_anexo_chamado_atual = Workbooks.Open(filePath)
    ' Procura pela célula que contém o texto "NÚMERO DA OCORRENCIA ( OC )" na primeira planilha
    Set rngEncontrado = arquivo_anexo_chamado_atual.Sheets(1).Range("A:AA").Find("NÚMERO DA OCORRENCIA ( OC )", , , xlWhole)
    ' Verifica se a célula não foi encontrada OU se o arquivo possui mais de uma planilha
    If rngEncontrado Is Nothing Or arquivo_anexo_chamado_atual.Sheets.Count > 1 Then
        ' Fecha o arquivo sem salvar as alterações
        arquivo_anexo_chamado_atual.Close False
        ' Define o retorno da função como Falso (download falhou ou arquivo inválido)
        DownloadAnexo = False
        ' Chama a sub-rotina para registrar no relatório que o chamado possui um anexo incorreto/fora do padrão
        Call AlimentarDicionario_Relatorio_Processamento("Chamados com anexo incorreto/fora do padrão: ", chamado)
        ' Encerra a função
        Exit Function
    End If

    ' Se todas as verificações passarem, define o retorno da função como Verdadeiro (download bem-sucedido e arquivo aparentemente válido)
    DownloadAnexo = True
End Function
