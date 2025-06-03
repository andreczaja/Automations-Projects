Attribute VB_Name = "REQUEST_API_POST"
Public apiUrl As String
Public http As Object
Public fileData() As Byte
Public stream As Object
Public bodyStream As Object
Public byteArray() As Byte
Public filePath As String
Public boundary As String
Public requestBody As String
Public closingBoundary, data_api As String
Public status_chamado As String
Public status_OC, chamado_anterior, data_solicitacao_reembolso_abatimento As String


Public Sub GerarTokenTramite(ByRef acao As String)

    ' Declaração de variáveis
    Dim token_tramite As String ' Variável para armazenar o token do trâmite gerado
    Dim chamado_anterior As Variant ' Variável para armazenar o chamado anterior

    ' Inicialização de variáveis
    generatorReferenceCode = "CONTASARECEBERELECT-638618399315"
    API_token = ThisWorkbook.Sheets("API KEY").Range("A1").Value ' Obtém o token da planilha "API KEY"
    apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/" & chamado & "/proceeding" ' Monta a URL da API para o chamado específico

    ' Cria o objeto HTTP para comunicação com a API
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Estrutura de seleção para definir o comportamento com base na ação
    Select Case acao
        Case "OC_JA_CONSULTADA" ' Caso a OC já tenha sido consultada anteriormente

            ' Se a quantidade de NFDs (Notas Fiscais de Devolução) for maior que 1, encerra a sub-rotina
            If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If

            ' Busca informações da OC na aba histórica
            chamado_anterior = Application.WorksheetFunction.VLookup(CLng(numero_OC), aba_historica.Columns("A:B"), 2, False)
            status_OC = Application.WorksheetFunction.VLookup(CLng(numero_OC), aba_historica.Columns("A:C"), 3, False)
            data_solicitacao_reembolso_abatimento = VBA.Format(VBA.CDate(Application.WorksheetFunction.VLookup(CLng(numero_OC), aba_historica.Columns("A:D"), 4, False)), "dd/mm/yyyy")

            ' Define o texto do trâmite com base no status da OC
            If status_OC = "REEMBOLSO" Then
                texto_tramite = "Prezado cliente," & vbNewLine & _
                    "A OC informada já havia sido consultada anteriormente no chamado " & chamado_anterior & ". Consta, portanto que o reembolso foi enviado para ser pago no dia " & data_solicitacao_reembolso_abatimento & _
                    ". Favor conferir chamado mencionado."
            ElseIf status_OC = "ABATIMENTO" Then
                texto_tramite = "Prezado cliente," & vbNewLine & _
                    "A OC informada já havia sido consultada anteriormente no chamado " & chamado_anterior & ". Consta, portanto que o abatimento foi realizado no dia " & data_solicitacao_reembolso_abatimento & _
                    ". Favor conferir chamado mencionado."
            ElseIf status_OC = "SEM CREDITOS EM ABERTO ENCONTRADOS" Then
                texto_tramite = "Prezado cliente," & vbNewLine & _
                    "A OC informada não possui créditos de devolução pendentes associados a ela."
            End If

            ' Monta o corpo da requisição POST para criar o trâmite
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                """private"":false," & _
                """status"":9," & _
                """description"":""" & texto_tramite & """}"

            ' Envia a requisição POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "SEM_DADOS_BANCARIOS" ' Caso não haja dados bancários cadastrados
            texto_tramite = "Prezado cliente," & vbNewLine & _
                "Efetuamos a consulta da(s) OC(s) informadas e foi analisado que existe um saldo pendente de reembolso no valor de R$" & soma_cred_dev & vbNewLine & _
                ".Porém não foram encontrados seus dados bancários cadastrados em nosso sistema." & vbNewLine & "Gentileza acionar a equipe do comercial para inserção dos dados"
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                """private"":false," & _
                """status"":9," & _
                """description"":""" & texto_tramite & """}"
            

            ' Envia a requisição POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "ERRO_ZSD164_1" ' Caso ocorra erro na transação ZSD164 (OC não encontrada)
            If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If
            texto_tramite = "Prezado cliente," & vbNewLine & _
                            "Efetuamos a consulta da(s) OC(s) " & OCs_erro_zsd164_1 & " as quais não se encontram disponíveis ou não foram localizadas." & vbNewLine & _
                            "Gentileza acionar a equipe de logística conforme sua regional, segue contatos:" & vbNewLine & _
                            "bruna.c.santos@electrolux.com , jose.lima@electrolux.com e maria.regis@electrolux.com"
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                   """private"":false," & _
                   """status"":9," & _
                   """description"":""" & texto_tramite & """}"
            

            ' Envia a requisição POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "ERRO_ZSD164_2" ' Caso ocorra erro na transação ZSD164 (devolução não finalizada)
            If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If
            texto_tramite = "Prezado cliente," & vbNewLine & _
                            "Efetuamos a consulta da(s) OC(s) <strong>" & OCs_erro_zsd164_2 & "</strong> as quais não se encontram disponíveis." & vbNewLine & "<strong>Status: Devolução não finalizada e não registrada.</strong> " & vbNewLine & _
                            "Gentileza acionar a equipe de logística conforme sua regional, segue contatos:" & vbNewLine & _
                            "bruna.c.santos@electrolux.com, jose.lima@electrolux.com e maria.regis@electrolux.com"
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                  """private"":false," & _
                  """status"":9," & _
                  """description"":""" & texto_tramite & """}"

            ' Envia a requisição POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "AVISO_OC_SEM_CREDITOS_ASSOCIADOS" ' Caso a OC não tenha créditos associados
             If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If
            texto_tramite = "Prezado cliente," & vbNewLine & _
                            "Efetuamos a consulta da OC <strong>" & numero_OC & "</strong> e não foram encontrados créditos associados a essa OC"
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                        """private"":false," & _
                        """status"":9," & _
                        """description"":""" & texto_tramite & """}"
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)
            
        Case "AVISO_OC_INCORRETA"  ' Caso a OC esteja incorreta
            If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If

            texto_tramite = "Foi identificado que a <strong> OC </strong>informada está <strong>incorreta ou inexistente</strong>. Pede-se que reabra um novo chamado constando a OC correta e necessária para a nossa análise."
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                """private"":false," & _
                """status"":9," & _
                """description"":""" & texto_tramite & """}"

            ' Envia a requisição POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "AVISO_FALTA_DE_ANEXO", "ANEXO_INCORRETO" ' Caso falte anexo ou o anexo esteja incorreto
        
            apiUrl = apiUrl & "/attachment/upload" ' Adiciona o endpoint de upload de anexos à URL
            
            texto_tramite = "Impossível prosseguir com o chamado!" & vbNewLine & _
                            "É necessário anexar o documento contendo as <strong>OCs </strong> a serem consultadas <strong>conforme arquivo padrão disponibilizado em > Clientes B2B > Modelo solicitação de devolução</strong>" & vbCrLf & _
                            "Segue também modelo em anexo:"
                        
            ' Ler arquivo binário
            fileData = LerArquivoBinario(caminho_arquivo_modelo)

            ' Criar requisição multipart/form-data
            boundary = "---------------------------" & Format(Now, "yyyymmddhhmmss") ' Gera um delimitador único
            requestBody = "--" & boundary & vbCrLf & _
                            "Content-Disposition: form-data; name=""file""; filename=""" & Mid(caminho_arquivo_modelo, InStrRev(caminho_arquivo_modelo, "\") + 1) & """" & vbCrLf & _
                            "Content-Type: application/octet-stream" & vbCrLf & vbCrLf ' Cabeçalho da requisição para o arquivo

            closingBoundary = vbCrLf & "--" & boundary & "--" & vbCrLf ' Delimitador de fechamento

            ' Criar Stream para o corpo da requisição
            Set bodyStream = CriarBodyStream(requestBody, fileData, closingBoundary) ' Concatena os dados

            ' Enviar requisição
            http.Open "POST", apiUrl, False
            http.setRequestHeader "Authorization", "Bearer " & API_token
            http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
            http.Send bodyStream.Read(bodyStream.Size) ' Envia o stream com os dados

            'Obtem o token de resposta
            token_tramite = Replace(http.responseText, """", "")
            
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                """private"":false," & _
                """status"":5," & _
                """reasonForWaitingReferenceCode"":""AGUARDANDORETORNODO-638217233492""," & _
                """description"":""" & texto_tramite & """," & _
                """attachmentsIds"":[""" & token_tramite & """]}"
                
            ' Monta a URL da API para registrar o trâmite no chamado específico.
            apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/" & chamado & "/proceeding"
            ' Envia a requisição POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)
            ' Chama a sub-rotina para incluir/atualizar o chamado na lista de pendentes
            Call IncluirAtualizarChamadoPendente
        Case "NENHUMA_OC_INFORMADA"
            texto_tramite = "Prezado Cliente," & vbCrLf & "É necessário informar as OCs na respectiva coluna do arquivo, sem ela não é possível seguir com a análise de créditos associados " & _
                "e processar possíveis abatimentos/reembolsos." & vbCrLf & " Favor em novo trâmite dentro desse chamado anexar o arquivo novamente com a coluna de OCs preenchida para poder dar seguimento."
            If dictChamadosPendentes.exists(chamado) Then
                requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                            """private"":false," & _
                            """status"":1," & _
                            """description"":""" & texto_tramite & """}"
                Call EnviarRequisicao("POST", http, apiUrl, requestBody)
                Call APITrocaResponsavelChamado(3)
            Else
                requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                            """private"":false," & _
                            """status"":5," & _
                            """reasonForWaitingReferenceCode"":""AGUARDANDORETORNODO-638217233492""," & _
                            """description"":""" & texto_tramite & """}"
                ' Chama a sub-rotina para incluir/atualizar o chamado na lista de pendentes
                Call EnviarRequisicao("POST", http, apiUrl, requestBody)
                Call IncluirAtualizarChamadoPendente
            End If
            ' Envia a requisição POST para a API
            
            
            
        Case "ENVIO_ANEXO_OCS_VERIFICADAS" ' Caso os anexos das OCs verificadas devam ser enviados
            apiUrl = apiUrl & "/attachment/upload" ' Adiciona o endpoint de upload de anexos à URL

            If qtde_NFD_OC_chamado = "01" Then
                If condicao_OCs_reembolso Then
                    texto_tramite = "Prezado Cliente," & vbCrLf & "A OC informada foi enviada para aprovação de reembolso. Notificaremos por meio desse chamado a respectiva data prevista de pagamento." & vbCrLf
                Else
                    If linhas_abertas And linhas_compensadas Then
                        texto_tramite = "Prezado Cliente," & vbCrLf & "Segue em anexo arquivo com as devidas verificações conforme OCs informadas." & vbCrLf & _
                                        "Foi analisado que existem créditos já resolvidos e outros ainda pendentes de serem abatidos/reembolsados os quais já foram enviados para pagamento/abatimento. Favor conferir em arquivo anexo os respectivos detalhes."
                        Call AlimentarDicionario_Relatorio_Processamento("Chamados com créditos em aberto e também já utilizados: ", chamado)
                    ElseIf linhas_abertas And Not linhas_compensadas Then
                        If condicao_payer = "abatidos" Then
                            texto_tramite = "Prezado Cliente," & vbCrLf & "Conforme analisado, haviam créditos pendentes a serem abatidos no valor de <strong>R$" & Round(Abs(soma_cred_dev), 2) & "</strong> os quais foram realizados segundo detalhe em arquivo anexo. Favor conferir." & vbCrLf & _
                            "Obs: Os boletos abatidos parcialmente, estarão disponíveis no Portal do Cliente em até 03 dias úteis."
                        End If
                        Call AlimentarDicionario_Relatorio_Processamento("Chamados apenas com créditos em aberto: ", chamado)
                    ElseIf Not linhas_abertas And linhas_compensadas Then
                         texto_tramite = "Prezado Cliente," & vbCrLf & "Segue em anexo arquivo com as devidas verificações conforme OCs informadas." & vbCrLf & _
                                        "Foi analisado que NÃO existem créditos pendentes a serem abatidos/reembolsados (favor conferir linhas mencionadas diretamente no arquivo anexo)." & vbCrLf & _
                                        "Gentileza verificar se seu caso for supplier abrir no serviço - OTC_Order to Cash - Contas a Receber  / Brasil / Supplier Card - Elux Card / Abatimento de devolução Supplier"

                        Call AlimentarDicionario_Relatorio_Processamento("Chamados apenas com créditos já utilizados: ", chamado)
                    End If
                End If
            ElseIf qtde_NFD_OC_chamado = "Acima de 01" Then
                Set regex = CreateObject("VBScript.RegExp")
                regex.Pattern = "^/*$"
                regex.IgnoreCase = True
                regex.Global = False

                texto_tramite = "Prezado Cliente," & vbCrLf & "Segue em anexo arquivo com as devidas verificações conforme OCs informadas."
                If condicao_payer = "abatidos" Then
        
                    texto_tramite = texto_tramite & vbCrLf & "Conforme analisado, haviam créditos pendentes a serem abatidos no valor de <strong>R$" & Round(Abs(soma_cred_dev), 2) & "</strong> os quais foram realizados segundo detalhe em arquivo anexo. Favor conferir." & vbCrLf & "OBS: em caso de boletos abatidos parcialmente, estes estarão disponíveis no Portal do Cliente em até 03 dias úteis."
                End If
                If condicao_OC_incorreta Then
                    texto_tramite = texto_tramite & "Informamos que existem OCs incorretas."
                    If Not regex.Test(OCs_incorretas) Then
                        OCs_incorretas = Mid(OCs_incorretas, 2, 9999)
                        texto_tramite = texto_tramite & " São elas: " & OCs_incorretas & ". Favor reenviar arquivo modelo com as OCs corretas em trâmite dentro desse chamado."
                    End If
                End If
                If condicao_erro_zsd164_1 Then
                    texto_tramite = texto_tramite & vbCrLf & "Existem OCs que não estão disponíveis e/ou não encontradas."
                    If Not regex.Test(OCs_erro_zsd164_1) Then
                        OCs_erro_zsd164_1 = Mid(OCs_erro_zsd164_1, 2, 9999)
                        texto_tramite = texto_tramite & " São elas: " & OCs_erro_zsd164_1
                    End If
                End If
                If condicao_erro_zsd164_2 Then
                      texto_tramite = texto_tramite & vbCrLf & "Existem OCs não finalizadas e não registradas."
                    If Not regex.Test(OCs_erro_zsd164_2) Then
                        OCs_erro_zsd164_2 = Mid(OCs_erro_zsd164_2, 2, 9999)
                        texto_tramite = texto_tramite & " São elas: " & OCs_erro_zsd164_2
                    End If
                End If
                If condicao_erro_zsd164_1 Or condicao_erro_zsd164_2 Then
                    texto_tramite = texto_tramite & vbCrLf & "Para estes casos, gentileza acionar a equipe de logística conforme sua regional, segue contatos:" & vbNewLine & _
                            "bruna.c.santos@electrolux.com , jose.lima@electrolux.com e maria.regis@electrolux.com"
                End If
                If condicao_OCs_reembolso Then
                    texto_tramite = texto_tramite & vbCrLf & "Existem OCs enviadas para aprovação de reembolso. Assim que a solicitação for aprovada," & _
                        " notificaremos por meio desse chamado tal como a respectiva data prevista de pagamento."
                End If
            End If

            ' Ler arquivo binário
            fileData = LerArquivoBinario(caminho_arquivo)

            ' Criar requisição multipart/form-data
            boundary = "---------------------------" & Format(Now, "yyyymmddhhmmss") ' Gera um delimitador único
            requestBody = "--" & boundary & vbCrLf & _
                            "Content-Disposition: form-data; name=""file""; filename=""" & Mid(caminho_arquivo, InStrRev(caminho_arquivo, "\") + 1) & """" & vbCrLf & _
                            "Content-Type: application/octet-stream" & vbCrLf & vbCrLf ' Cabeçalho da requisição para o arquivo

            closingBoundary = vbCrLf & "--" & boundary & "--" & vbCrLf ' Delimitador de fechamento

            ' Criar Stream para o corpo da requisição
            Set bodyStream = CriarBodyStream(requestBody, fileData, closingBoundary) ' Concatena os dados

            ' Enviar requisição
            http.Open "POST", apiUrl, False
            http.setRequestHeader "Authorization", "Bearer " & API_token
            http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
            http.Send bodyStream.Read(bodyStream.Size) ' Envia o stream com os dados

            'Obtem o token de resposta
            token_tramite = Replace(http.responseText, """", "")
            'Chama a função para registrar o tramite
            Call RegistrarTramite(token_tramite, texto_tramite)

    End Select

    ' Liberar memória
    Set http = Nothing
End Sub


Private Sub RegistrarTramite(ByVal token_tramite As String, ByVal texto_tramite As String)
    ' Sub-rotina para registrar um trâmite em um chamado existente.

    Dim apiUrl As String ' Declara a variável apiUrl para armazenar a URL da API.
    Dim requestBody As String ' Declara a variável requestBody para armazenar o corpo da requisição POST.
    Dim resposta As String ' Declara a variável resposta para armazenar a resposta da API.
    Dim http As Object

    ' Monta a URL da API para registrar o trâmite no chamado específico.
    apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/" & chamado & "/proceeding"

    
    
    If qtde_NFD_OC_chamado = "Acima de 01" And condicao_OC_incorreta Then
        If dictChamadosPendentes.exists(chamado) Then
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                """private"":false," & _
                """status"":1," & _
                """description"":""" & texto_tramite & """," & _
                """attachmentsIds"":[""" & token_tramite & """]}"
            Call APITrocaResponsavelChamado(3)
        Else
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                """private"":false," & _
                """status"":5," & _
                """reasonForWaitingReferenceCode"":""AGUARDANDORETORNODO-638217233492""," & _
                """description"":""" & texto_tramite & """," & _
                """attachmentsIds"":[""" & token_tramite & """]}"
            Call IncluirAtualizarChamadoPendente
        End If
        GoTo enviar
    End If
    ' Verifica se o status do chamado já foi definido. Se não, define como "9".
    
    linha_fim_aba_historico_chamados_pendentes = aba_historico_chamados_pendentes.Range("A1048576").End(xlUp).Row
    For linha = linha_fim_aba_historico_chamados_pendentes To 2 Step -1
        If aba_historico_chamados_pendentes.Range("A" & linha).Value = CLng(chamado) Then
            aba_historico_chamados_pendentes.Rows(linha).Delete
        End If
    Next linha
    
    If condicao_OCs_reembolso Then
        requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
            """private"":false," & _
            """status"":1 ," & _
            """description"":""" & texto_tramite & """," & _
            """attachmentsIds"":[""" & token_tramite & """]}"
    Else
        requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
        """private"":false," & _
        """status"":9 ," & _
        """description"":""" & texto_tramite & """," & _
        """attachmentsIds"":[""" & token_tramite & """]}"
        
    End If
    
    

enviar:

    ' Chama a função EnviarRequisicao para enviar a requisição POST para a API e obtém a resposta.
    resposta = EnviarRequisicao("POST", http, apiUrl, requestBody)

    ' Libera a memória do objeto HTTP.
    Set http = Nothing
End Sub
Public Sub AbrirChamadoContasAPagar()

    Dim texto_chamado, token_anexo As String
    Dim ticket_aberto
    ' Criar objeto HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    generatorReferenceCode = "CONTASARECEBERELECT-638618399315"
    API_token = ThisWorkbook.Sheets("API KEY").Range("A1").Value
    apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/attachment/upload"
    ' Ler arquivo binário
    fileData = LerArquivoBinario(caminho_arquivo)
    
    ' Criar requisição multipart/form-data
    boundary = "---------------------------" & Format(Now, "yyyymmddhhmmss")
    requestBody = "--" & boundary & vbCrLf & _
                  "Content-Disposition: form-data; name=""file""; filename=""" & Mid(caminho_arquivo, InStrRev(caminho_arquivo, "\") + 1) & """" & vbCrLf & _
                  "Content-Type: application/octet-stream" & vbCrLf & vbCrLf


    closingBoundary = vbCrLf & "--" & boundary & "--" & vbCrLf

    ' Criar Stream para o corpo da requisição
    Set bodyStream = CriarBodyStream(requestBody, fileData, closingBoundary)

    ' Enviar requisição
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Authorization", "Bearer " & API_token
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.Send bodyStream.Read(bodyStream.Size)


    ' Liberar memória
    bodyStream.Close
    Set bodyStream = Nothing

    token_anexo = Replace(http.responseText, """", "")
    data_api = Format(Form_SAP.txt_box_data_agrupado_pgto_SAP.Value, "yyyy-MM-dd")
    texto_chamado = "Segue em anexo linhas de reembolso que somam R$" & Abs(soma_cred_dev) & " e que devem ser pagas ao(s) respectivo(s) cliente(s)"

    requestBody = "{""title"":""PTP_Payment /  Brasil / Solicitação de pagamento nacional""," & _
              """private"":false," & """generatorReferenceCode"":""" & generatorReferenceCode & """," & _
              """CustomerReferenceCode"":""Electrolux""," & """RequesterReferenceCode"":""" & generatorReferenceCode & """," & _
              """serviceReferenceCode"":""SOLICITACAODEPAGAME-638254176526""," & """requestTypeReferenceCode"":""PAGAMENTONORMAL-638152833228""," & _
              """severityReferenceCode"":""PAGAMENTOAGRUPADO05-638749843822""," & """status"":0," & _
               """attachmentsIds"":[""" & token_anexo & """]," & _
              """forms"":[{" & """referenceCode"":""SOLICITACAODEPAGAME-638230145176""," & _
              """fieldsValues"":{" & _
                  """SELECAO-638230148228"":""Fornecedores/Clientes""," & _
                  """CLASSIFICACAODOPAGA-638254105887"":""Reembolso clientes (OTC)""," & _
                  """EMPRESA-638253148448"":""BR10""," & _
                  """CODIGOSAP-638253214686"":""" & payer_associado_OC & """," & _
                  """NOMEDOFORNECEDOR-638230197871"":""------""," & _
                  """NUMERODANF-638230901764"":""------""," & _
                  """NDODOCUMENTOSAP-638458111407"":""------""," & _
                  """VALOR-638230113178"":""" & Replace(Abs(soma_cred_dev), ".", ",") & """," & _
                  """DATADOPAGAMENTO-638230160781"":""" & data_api & """," & _
                  """FORMADEPAGAMENTO-638230172310"":""TED""," & _
                  """NUMERODOPROCESSO-638230190940"":""------""," & _
                  """BANCO-638230161183"":""------""," & _
                  """OBSERVACOES-638283790711"":""" & texto_chamado & """" & _
              "}}]}"
    apiUrl = "https://electrolux.ellevo.com/api/v1/ticket"
    response = EnviarRequisicao("POST", http, apiUrl, requestBody)
    Set json = JsonConverter.ParseJson(response)
    ticket_aberto = json("number")
    chamado_ellevo_aberto_contas_pagar = ticket_aberto
    

End Sub


Public Sub CriarTramiteNotificacaoReembolsoAprovado()

    If qtde_NFD_OC_chamado = "1" Then
        qtde_NFD_OC_chamado = "01"
    End If
    status_chamado = "9"

    caminho_arquivo = Replace(pasta_arquivos_clientes, "Arquivos Clientes", "Anexos Detalhe Reembolso") & "\" & VBA.Format(data_solicitacao_reembolso, "dd.mm.yyyy") & "\" & doc_f65 & ".xlsx"

    Set http = CreateObject("MSXML2.XMLHTTP")
    
    generatorReferenceCode = "CONTASARECEBERELECT-638618399315"
    API_token = ThisWorkbook.Sheets("API KEY").Range("A1").Value
    apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/" & chamado & "/proceeding/attachment/upload"
    ' Ler arquivo binário
    fileData = LerArquivoBinario(caminho_arquivo)
    
    ' Criar requisição multipart/form-data
    boundary = "---------------------------" & Format(Now, "yyyymmddhhmmss")
    requestBody = "--" & boundary & vbCrLf & _
                  "Content-Disposition: form-data; name=""file""; filename=""" & Mid(caminho_arquivo, InStrRev(caminho_arquivo, "\") + 1) & """" & vbCrLf & _
                  "Content-Type: application/octet-stream" & vbCrLf & vbCrLf


    closingBoundary = vbCrLf & "--" & boundary & "--" & vbCrLf

    ' Criar Stream para o corpo da requisição
    Set bodyStream = CriarBodyStream(requestBody, fileData, closingBoundary)

    http.Open "POST", apiUrl, False
    http.setRequestHeader "Authorization", "Bearer " & API_token
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.Send bodyStream.Read(bodyStream.Size)
    ' Liberar memória
    bodyStream.Close
    Set bodyStream = Nothing

    token_tramite = Replace(http.responseText, """", "")
    
    texto_tramite = "Prezado Cliente, Segue demonstrativo de reembolso aprovado para pagamento no valor de R$" & soma_cred_dev & ", o mesmo será pago no dia " & Form_SAP.txt_box_data_agrupado_pgto_SAP & "."

    requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                      """private"":false," & _
                      """status"":""" & status_chamado & """," & _
                      """description"":""" & texto_tramite & """," & _
                      """attachmentsIds"":[""" & token_tramite & """]}"
                      
    apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/" & chamado & "/proceeding"
    resposta = EnviarRequisicao("POST", http, apiUrl, requestBody)

End Sub

Public Function EnviarRequisicao(metodo As String, ByVal http As Object, ByVal apiUrl As String, ByVal requestBody As String) As String

    Set http = CreateObject("MSXML2.XMLHTTP")
    ' Configura a requisição
    http.Open metodo, apiUrl, False
    http.setRequestHeader "Authorization", "Bearer " & API_token
    http.setRequestHeader "Content-Type", "application/json"

    If metodo = "POST" Then
        ' Envia os dados
        http.Send requestBody
        EnviarRequisicao = http.responseText
    ElseIf metodo = "GET" Then
        http.Send
        EnviarRequisicao = http.responseText
    End If
    On Error GoTo 0
    Set http = Nothing
End Function


' Função para ler arquivo como binário
Private Function LerArquivoBinario(ByVal filePath As String) As Byte()

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binário
    stream.Open
    stream.LoadFromFile filePath
    stream.Position = 0
    fileData = stream.Read(stream.Size)
    stream.Close
    Set stream = Nothing

    LerArquivoBinario = fileData
End Function

' Função para criar o body da requisição multipart/form-data
Private Function CriarBodyStream(ByVal requestBody As String, ByRef fileData() As Byte, ByVal closingBoundary As String) As Object

    Set bodyStream = CreateObject("ADODB.Stream")
    bodyStream.Type = 1 ' Binário
    bodyStream.Open

    ' Adicionar cabeçalho
    byteArray = StrConv(requestBody, vbFromUnicode)
    bodyStream.Write byteArray

    ' Adicionar os bytes do arquivo
    bodyStream.Write fileData

    ' Adicionar o boundary de encerramento
    byteArray = StrConv(closingBoundary, vbFromUnicode)
    bodyStream.Write byteArray

    ' Retornar o stream
    bodyStream.Position = 0
    Set CriarBodyStream = bodyStream
End Function


