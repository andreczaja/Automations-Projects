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

    ' Declara��o de vari�veis
    Dim token_tramite As String ' Vari�vel para armazenar o token do tr�mite gerado
    Dim chamado_anterior As Variant ' Vari�vel para armazenar o chamado anterior

    ' Inicializa��o de vari�veis
    generatorReferenceCode = "CONTASARECEBERELECT-638618399315"
    API_token = ThisWorkbook.Sheets("API KEY").Range("A1").Value ' Obt�m o token da planilha "API KEY"
    apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/" & chamado & "/proceeding" ' Monta a URL da API para o chamado espec�fico

    ' Cria o objeto HTTP para comunica��o com a API
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Estrutura de sele��o para definir o comportamento com base na a��o
    Select Case acao
        Case "OC_JA_CONSULTADA" ' Caso a OC j� tenha sido consultada anteriormente

            ' Se a quantidade de NFDs (Notas Fiscais de Devolu��o) for maior que 1, encerra a sub-rotina
            If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If

            ' Busca informa��es da OC na aba hist�rica
            chamado_anterior = Application.WorksheetFunction.VLookup(CLng(numero_OC), aba_historica.Columns("A:B"), 2, False)
            status_OC = Application.WorksheetFunction.VLookup(CLng(numero_OC), aba_historica.Columns("A:C"), 3, False)
            data_solicitacao_reembolso_abatimento = VBA.Format(VBA.CDate(Application.WorksheetFunction.VLookup(CLng(numero_OC), aba_historica.Columns("A:D"), 4, False)), "dd/mm/yyyy")

            ' Define o texto do tr�mite com base no status da OC
            If status_OC = "REEMBOLSO" Then
                texto_tramite = "Prezado cliente," & vbNewLine & _
                    "A OC informada j� havia sido consultada anteriormente no chamado " & chamado_anterior & ". Consta, portanto que o reembolso foi enviado para ser pago no dia " & data_solicitacao_reembolso_abatimento & _
                    ". Favor conferir chamado mencionado."
            ElseIf status_OC = "ABATIMENTO" Then
                texto_tramite = "Prezado cliente," & vbNewLine & _
                    "A OC informada j� havia sido consultada anteriormente no chamado " & chamado_anterior & ". Consta, portanto que o abatimento foi realizado no dia " & data_solicitacao_reembolso_abatimento & _
                    ". Favor conferir chamado mencionado."
            ElseIf status_OC = "SEM CREDITOS EM ABERTO ENCONTRADOS" Then
                texto_tramite = "Prezado cliente," & vbNewLine & _
                    "A OC informada n�o possui cr�ditos de devolu��o pendentes associados a ela."
            End If

            ' Monta o corpo da requisi��o POST para criar o tr�mite
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                """private"":false," & _
                """status"":9," & _
                """description"":""" & texto_tramite & """}"

            ' Envia a requisi��o POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "SEM_DADOS_BANCARIOS" ' Caso n�o haja dados banc�rios cadastrados
            texto_tramite = "Prezado cliente," & vbNewLine & _
                "Efetuamos a consulta da(s) OC(s) informadas e foi analisado que existe um saldo pendente de reembolso no valor de R$" & soma_cred_dev & vbNewLine & _
                ".Por�m n�o foram encontrados seus dados banc�rios cadastrados em nosso sistema." & vbNewLine & "Gentileza acionar a equipe do comercial para inser��o dos dados"
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                """private"":false," & _
                """status"":9," & _
                """description"":""" & texto_tramite & """}"
            

            ' Envia a requisi��o POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "ERRO_ZSD164_1" ' Caso ocorra erro na transa��o ZSD164 (OC n�o encontrada)
            If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If
            texto_tramite = "Prezado cliente," & vbNewLine & _
                            "Efetuamos a consulta da(s) OC(s) " & OCs_erro_zsd164_1 & " as quais n�o se encontram dispon�veis ou n�o foram localizadas." & vbNewLine & _
                            "Gentileza acionar a equipe de log�stica conforme sua regional, segue contatos:" & vbNewLine & _
                            "bruna.c.santos@electrolux.com , jose.lima@electrolux.com e maria.regis@electrolux.com"
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                   """private"":false," & _
                   """status"":9," & _
                   """description"":""" & texto_tramite & """}"
            

            ' Envia a requisi��o POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "ERRO_ZSD164_2" ' Caso ocorra erro na transa��o ZSD164 (devolu��o n�o finalizada)
            If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If
            texto_tramite = "Prezado cliente," & vbNewLine & _
                            "Efetuamos a consulta da(s) OC(s) <strong>" & OCs_erro_zsd164_2 & "</strong> as quais n�o se encontram dispon�veis." & vbNewLine & "<strong>Status: Devolu��o n�o finalizada e n�o registrada.</strong> " & vbNewLine & _
                            "Gentileza acionar a equipe de log�stica conforme sua regional, segue contatos:" & vbNewLine & _
                            "bruna.c.santos@electrolux.com, jose.lima@electrolux.com e maria.regis@electrolux.com"
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                  """private"":false," & _
                  """status"":9," & _
                  """description"":""" & texto_tramite & """}"

            ' Envia a requisi��o POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "AVISO_OC_SEM_CREDITOS_ASSOCIADOS" ' Caso a OC n�o tenha cr�ditos associados
             If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If
            texto_tramite = "Prezado cliente," & vbNewLine & _
                            "Efetuamos a consulta da OC <strong>" & numero_OC & "</strong> e n�o foram encontrados cr�ditos associados a essa OC"
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                        """private"":false," & _
                        """status"":9," & _
                        """description"":""" & texto_tramite & """}"
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)
            
        Case "AVISO_OC_INCORRETA"  ' Caso a OC esteja incorreta
            If qtde_NFD_OC_chamado = "Acima de 01" Then
                Exit Sub
            End If

            texto_tramite = "Foi identificado que a <strong> OC </strong>informada est� <strong>incorreta ou inexistente</strong>. Pede-se que reabra um novo chamado constando a OC correta e necess�ria para a nossa an�lise."
            requestBody = "{""generatorReferenceCode"":""" & generatorReferenceCode & """," & _
                """private"":false," & _
                """status"":9," & _
                """description"":""" & texto_tramite & """}"

            ' Envia a requisi��o POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)

        Case "AVISO_FALTA_DE_ANEXO", "ANEXO_INCORRETO" ' Caso falte anexo ou o anexo esteja incorreto
        
            apiUrl = apiUrl & "/attachment/upload" ' Adiciona o endpoint de upload de anexos � URL
            
            texto_tramite = "Imposs�vel prosseguir com o chamado!" & vbNewLine & _
                            "� necess�rio anexar o documento contendo as <strong>OCs </strong> a serem consultadas <strong>conforme arquivo padr�o disponibilizado em > Clientes B2B > Modelo solicita��o de devolu��o</strong>" & vbCrLf & _
                            "Segue tamb�m modelo em anexo:"
                        
            ' Ler arquivo bin�rio
            fileData = LerArquivoBinario(caminho_arquivo_modelo)

            ' Criar requisi��o multipart/form-data
            boundary = "---------------------------" & Format(Now, "yyyymmddhhmmss") ' Gera um delimitador �nico
            requestBody = "--" & boundary & vbCrLf & _
                            "Content-Disposition: form-data; name=""file""; filename=""" & Mid(caminho_arquivo_modelo, InStrRev(caminho_arquivo_modelo, "\") + 1) & """" & vbCrLf & _
                            "Content-Type: application/octet-stream" & vbCrLf & vbCrLf ' Cabe�alho da requisi��o para o arquivo

            closingBoundary = vbCrLf & "--" & boundary & "--" & vbCrLf ' Delimitador de fechamento

            ' Criar Stream para o corpo da requisi��o
            Set bodyStream = CriarBodyStream(requestBody, fileData, closingBoundary) ' Concatena os dados

            ' Enviar requisi��o
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
                
            ' Monta a URL da API para registrar o tr�mite no chamado espec�fico.
            apiUrl = "https://electrolux.ellevo.com/api/v1/ticket/" & chamado & "/proceeding"
            ' Envia a requisi��o POST para a API
            Call EnviarRequisicao("POST", http, apiUrl, requestBody)
            ' Chama a sub-rotina para incluir/atualizar o chamado na lista de pendentes
            Call IncluirAtualizarChamadoPendente
        Case "NENHUMA_OC_INFORMADA"
            texto_tramite = "Prezado Cliente," & vbCrLf & "� necess�rio informar as OCs na respectiva coluna do arquivo, sem ela n�o � poss�vel seguir com a an�lise de cr�ditos associados " & _
                "e processar poss�veis abatimentos/reembolsos." & vbCrLf & " Favor em novo tr�mite dentro desse chamado anexar o arquivo novamente com a coluna de OCs preenchida para poder dar seguimento."
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
            ' Envia a requisi��o POST para a API
            
            
            
        Case "ENVIO_ANEXO_OCS_VERIFICADAS" ' Caso os anexos das OCs verificadas devam ser enviados
            apiUrl = apiUrl & "/attachment/upload" ' Adiciona o endpoint de upload de anexos � URL

            If qtde_NFD_OC_chamado = "01" Then
                If condicao_OCs_reembolso Then
                    texto_tramite = "Prezado Cliente," & vbCrLf & "A OC informada foi enviada para aprova��o de reembolso. Notificaremos por meio desse chamado a respectiva data prevista de pagamento." & vbCrLf
                Else
                    If linhas_abertas And linhas_compensadas Then
                        texto_tramite = "Prezado Cliente," & vbCrLf & "Segue em anexo arquivo com as devidas verifica��es conforme OCs informadas." & vbCrLf & _
                                        "Foi analisado que existem cr�ditos j� resolvidos e outros ainda pendentes de serem abatidos/reembolsados os quais j� foram enviados para pagamento/abatimento. Favor conferir em arquivo anexo os respectivos detalhes."
                        Call AlimentarDicionario_Relatorio_Processamento("Chamados com cr�ditos em aberto e tamb�m j� utilizados: ", chamado)
                    ElseIf linhas_abertas And Not linhas_compensadas Then
                        If condicao_payer = "abatidos" Then
                            texto_tramite = "Prezado Cliente," & vbCrLf & "Conforme analisado, haviam cr�ditos pendentes a serem abatidos no valor de <strong>R$" & Round(Abs(soma_cred_dev), 2) & "</strong> os quais foram realizados segundo detalhe em arquivo anexo. Favor conferir." & vbCrLf & _
                            "Obs: Os boletos abatidos parcialmente, estar�o dispon�veis no Portal do Cliente em at� 03 dias �teis."
                        End If
                        Call AlimentarDicionario_Relatorio_Processamento("Chamados apenas com cr�ditos em aberto: ", chamado)
                    ElseIf Not linhas_abertas And linhas_compensadas Then
                         texto_tramite = "Prezado Cliente," & vbCrLf & "Segue em anexo arquivo com as devidas verifica��es conforme OCs informadas." & vbCrLf & _
                                        "Foi analisado que N�O existem cr�ditos pendentes a serem abatidos/reembolsados (favor conferir linhas mencionadas diretamente no arquivo anexo)." & vbCrLf & _
                                        "Gentileza verificar se seu caso for supplier abrir no servi�o - OTC_Order to Cash - Contas a Receber  / Brasil / Supplier Card - Elux Card / Abatimento de devolu��o Supplier"

                        Call AlimentarDicionario_Relatorio_Processamento("Chamados apenas com cr�ditos j� utilizados: ", chamado)
                    End If
                End If
            ElseIf qtde_NFD_OC_chamado = "Acima de 01" Then
                Set regex = CreateObject("VBScript.RegExp")
                regex.Pattern = "^/*$"
                regex.IgnoreCase = True
                regex.Global = False

                texto_tramite = "Prezado Cliente," & vbCrLf & "Segue em anexo arquivo com as devidas verifica��es conforme OCs informadas."
                If condicao_payer = "abatidos" Then
        
                    texto_tramite = texto_tramite & vbCrLf & "Conforme analisado, haviam cr�ditos pendentes a serem abatidos no valor de <strong>R$" & Round(Abs(soma_cred_dev), 2) & "</strong> os quais foram realizados segundo detalhe em arquivo anexo. Favor conferir." & vbCrLf & "OBS: em caso de boletos abatidos parcialmente, estes estar�o dispon�veis no Portal do Cliente em at� 03 dias �teis."
                End If
                If condicao_OC_incorreta Then
                    texto_tramite = texto_tramite & "Informamos que existem OCs incorretas."
                    If Not regex.Test(OCs_incorretas) Then
                        OCs_incorretas = Mid(OCs_incorretas, 2, 9999)
                        texto_tramite = texto_tramite & " S�o elas: " & OCs_incorretas & ". Favor reenviar arquivo modelo com as OCs corretas em tr�mite dentro desse chamado."
                    End If
                End If
                If condicao_erro_zsd164_1 Then
                    texto_tramite = texto_tramite & vbCrLf & "Existem OCs que n�o est�o dispon�veis e/ou n�o encontradas."
                    If Not regex.Test(OCs_erro_zsd164_1) Then
                        OCs_erro_zsd164_1 = Mid(OCs_erro_zsd164_1, 2, 9999)
                        texto_tramite = texto_tramite & " S�o elas: " & OCs_erro_zsd164_1
                    End If
                End If
                If condicao_erro_zsd164_2 Then
                      texto_tramite = texto_tramite & vbCrLf & "Existem OCs n�o finalizadas e n�o registradas."
                    If Not regex.Test(OCs_erro_zsd164_2) Then
                        OCs_erro_zsd164_2 = Mid(OCs_erro_zsd164_2, 2, 9999)
                        texto_tramite = texto_tramite & " S�o elas: " & OCs_erro_zsd164_2
                    End If
                End If
                If condicao_erro_zsd164_1 Or condicao_erro_zsd164_2 Then
                    texto_tramite = texto_tramite & vbCrLf & "Para estes casos, gentileza acionar a equipe de log�stica conforme sua regional, segue contatos:" & vbNewLine & _
                            "bruna.c.santos@electrolux.com , jose.lima@electrolux.com e maria.regis@electrolux.com"
                End If
                If condicao_OCs_reembolso Then
                    texto_tramite = texto_tramite & vbCrLf & "Existem OCs enviadas para aprova��o de reembolso. Assim que a solicita��o for aprovada," & _
                        " notificaremos por meio desse chamado tal como a respectiva data prevista de pagamento."
                End If
            End If

            ' Ler arquivo bin�rio
            fileData = LerArquivoBinario(caminho_arquivo)

            ' Criar requisi��o multipart/form-data
            boundary = "---------------------------" & Format(Now, "yyyymmddhhmmss") ' Gera um delimitador �nico
            requestBody = "--" & boundary & vbCrLf & _
                            "Content-Disposition: form-data; name=""file""; filename=""" & Mid(caminho_arquivo, InStrRev(caminho_arquivo, "\") + 1) & """" & vbCrLf & _
                            "Content-Type: application/octet-stream" & vbCrLf & vbCrLf ' Cabe�alho da requisi��o para o arquivo

            closingBoundary = vbCrLf & "--" & boundary & "--" & vbCrLf ' Delimitador de fechamento

            ' Criar Stream para o corpo da requisi��o
            Set bodyStream = CriarBodyStream(requestBody, fileData, closingBoundary) ' Concatena os dados

            ' Enviar requisi��o
            http.Open "POST", apiUrl, False
            http.setRequestHeader "Authorization", "Bearer " & API_token
            http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
            http.Send bodyStream.Read(bodyStream.Size) ' Envia o stream com os dados

            'Obtem o token de resposta
            token_tramite = Replace(http.responseText, """", "")
            'Chama a fun��o para registrar o tramite
            Call RegistrarTramite(token_tramite, texto_tramite)

    End Select

    ' Liberar mem�ria
    Set http = Nothing
End Sub


Private Sub RegistrarTramite(ByVal token_tramite As String, ByVal texto_tramite As String)
    ' Sub-rotina para registrar um tr�mite em um chamado existente.

    Dim apiUrl As String ' Declara a vari�vel apiUrl para armazenar a URL da API.
    Dim requestBody As String ' Declara a vari�vel requestBody para armazenar o corpo da requisi��o POST.
    Dim resposta As String ' Declara a vari�vel resposta para armazenar a resposta da API.
    Dim http As Object

    ' Monta a URL da API para registrar o tr�mite no chamado espec�fico.
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
    ' Verifica se o status do chamado j� foi definido. Se n�o, define como "9".
    
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

    ' Chama a fun��o EnviarRequisicao para enviar a requisi��o POST para a API e obt�m a resposta.
    resposta = EnviarRequisicao("POST", http, apiUrl, requestBody)

    ' Libera a mem�ria do objeto HTTP.
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
    ' Ler arquivo bin�rio
    fileData = LerArquivoBinario(caminho_arquivo)
    
    ' Criar requisi��o multipart/form-data
    boundary = "---------------------------" & Format(Now, "yyyymmddhhmmss")
    requestBody = "--" & boundary & vbCrLf & _
                  "Content-Disposition: form-data; name=""file""; filename=""" & Mid(caminho_arquivo, InStrRev(caminho_arquivo, "\") + 1) & """" & vbCrLf & _
                  "Content-Type: application/octet-stream" & vbCrLf & vbCrLf


    closingBoundary = vbCrLf & "--" & boundary & "--" & vbCrLf

    ' Criar Stream para o corpo da requisi��o
    Set bodyStream = CriarBodyStream(requestBody, fileData, closingBoundary)

    ' Enviar requisi��o
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Authorization", "Bearer " & API_token
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.Send bodyStream.Read(bodyStream.Size)


    ' Liberar mem�ria
    bodyStream.Close
    Set bodyStream = Nothing

    token_anexo = Replace(http.responseText, """", "")
    data_api = Format(Form_SAP.txt_box_data_agrupado_pgto_SAP.Value, "yyyy-MM-dd")
    texto_chamado = "Segue em anexo linhas de reembolso que somam R$" & Abs(soma_cred_dev) & " e que devem ser pagas ao(s) respectivo(s) cliente(s)"

    requestBody = "{""title"":""PTP_Payment /  Brasil / Solicita��o de pagamento nacional""," & _
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
    ' Ler arquivo bin�rio
    fileData = LerArquivoBinario(caminho_arquivo)
    
    ' Criar requisi��o multipart/form-data
    boundary = "---------------------------" & Format(Now, "yyyymmddhhmmss")
    requestBody = "--" & boundary & vbCrLf & _
                  "Content-Disposition: form-data; name=""file""; filename=""" & Mid(caminho_arquivo, InStrRev(caminho_arquivo, "\") + 1) & """" & vbCrLf & _
                  "Content-Type: application/octet-stream" & vbCrLf & vbCrLf


    closingBoundary = vbCrLf & "--" & boundary & "--" & vbCrLf

    ' Criar Stream para o corpo da requisi��o
    Set bodyStream = CriarBodyStream(requestBody, fileData, closingBoundary)

    http.Open "POST", apiUrl, False
    http.setRequestHeader "Authorization", "Bearer " & API_token
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.Send bodyStream.Read(bodyStream.Size)
    ' Liberar mem�ria
    bodyStream.Close
    Set bodyStream = Nothing

    token_tramite = Replace(http.responseText, """", "")
    
    texto_tramite = "Prezado Cliente, Segue demonstrativo de reembolso aprovado para pagamento no valor de R$" & soma_cred_dev & ", o mesmo ser� pago no dia " & Form_SAP.txt_box_data_agrupado_pgto_SAP & "."

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
    ' Configura a requisi��o
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


' Fun��o para ler arquivo como bin�rio
Private Function LerArquivoBinario(ByVal filePath As String) As Byte()

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Bin�rio
    stream.Open
    stream.LoadFromFile filePath
    stream.Position = 0
    fileData = stream.Read(stream.Size)
    stream.Close
    Set stream = Nothing

    LerArquivoBinario = fileData
End Function

' Fun��o para criar o body da requisi��o multipart/form-data
Private Function CriarBodyStream(ByVal requestBody As String, ByRef fileData() As Byte, ByVal closingBoundary As String) As Object

    Set bodyStream = CreateObject("ADODB.Stream")
    bodyStream.Type = 1 ' Bin�rio
    bodyStream.Open

    ' Adicionar cabe�alho
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


