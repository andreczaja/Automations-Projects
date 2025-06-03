Attribute VB_Name = "PASSO_0_VERIFICACOES_INICIAIS"
Option Explicit
Public aba_consolidado, aba_plan_distribuicao, aba_historico_chamados_pendentes, aba_dados_bancarios, aba_reembolsos_pendentes, aba_relatorio_processamento, aba_historica As Worksheet
Public arquivo_cliente As Workbook
Public tabela_aba_consolidado, tabela_aba_plan_distribuicao As ListObject
Public i, i2, i3, i4, i5, linha, linha_fim, linhas_visiveis, linha_fim_aba_historico_chamados_pendentes, linha_fim_aba_chamados_supplier, ultima_linha_preenchida_atribuicoes_exclusas As Long
Public x_doc_compensacao, x_data_compensacao, x_texto, x_tipo_doc, x_cliente, x_nome, x_referencia, x_montante, x_venc_liquido, x_bloq_adv, x_chave_ref_1, x_chave_ref_2, x_chave_ref_3, x_num_doc, x_item, x_atribuicao As Integer
Public linha_fim_aba_creditos_utilizados_arquivo_cliente, linha_fim_aba_creditos_em_aberto_arquivo_cliente, linha_fim_aba_reembolso_arquivo_cliente, linha_fim_aba_abatimento_arquivo_cliente As Long
Public chamado, cnpj, numero_OC, tipo_data_sap, payer_associado_OC, fatura_devolucao, ordem_devolucao, qtde_NFD_OC_chamado, numero_NFD, pasta_diaria, data_solicitacao_reembolso As String
Public dictGeralChamados, dictChamadoAtual, dictChamadosPendentes, dictOCNumeroDOC_ZFI105 As Object
Public SapGui, Applic, connection, session, session_2, session_3, regex, ocorrencia As Object
Public array_payers_encontrados(), array_chamados(), array_cabecalho_arquivo_cliente(), array_doc_compensacoes(), array_num_doc_zsd336(), array_num_doc_zfi105(), _
    array_geral_linhas_abertas_FBL5N(), array_linhas_compensadas_FB03(), array_linha_atual(), array_linhas_chamado_contas_pagar(), array_nome_abas(), array_contas_supplier(), array_OC_linhas_abertas(), array_OC_linhas_compensadas() As Variant
Public array_atribuicoes_proibidas() As Variant
Public Dicionario_Relatorio_Processamento As Object
Public qtde_page_downs, ultimo_tramite_verificado, contador_erro_zfi156 As Integer
Public linhas_compensadas, linhas_abertas, condicao_OC_incorreta, condicao_cliente_sem_dados_bancarios, condicao_erro_zsd164_1, condicao_erro_zsd164_2, condicao_processamento_nao_escolhido, condicao_OCs_reembolso, condicao_doc_compensacao_repetido, condicao_chamado_supplier, erro_zfi156 As Boolean
Public pasta_arquivos_clientes, OCs_incorretas, caminho_arquivo, condicao_abat_reemb, pasta_anexos_detalhe_reembolso, chamado_ellevo_aberto_contas_pagar, doc_compensacao, OCs_erro_zsd164_1, OCs_erro_zsd164_2, num_doc_supplier, caminho_arquivo_modelo  As String
Public linha_fim_array_payers_encontrados, linha_fim_docs_sap As Long
Public rng As Range

Sub MAIN()

    ' Desativa a atualiza��o da tela para melhorar o desempenho.
    Application.ScreenUpdating = False


    ' Define as vari�veis de objeto para as planilhas e tabelas do Excel.
    Set aba_consolidado = ThisWorkbook.Sheets("Consolidado Chamados Ellevo")
    Set tabela_aba_consolidado = aba_consolidado.ListObjects("Consolidado")
    Set aba_plan_distribuicao = ThisWorkbook.Sheets("Plan Distribui��o")
    Set tabela_aba_plan_distribuicao = aba_plan_distribuicao.ListObjects("Plan_Distribui��o")
    Set aba_historico_chamados_pendentes = ThisWorkbook.Sheets("Hist�rico Chamados Pendentes")
    Set aba_dados_bancarios = ThisWorkbook.Sheets("Check Dados Banc�rios")
    Set aba_reembolsos_pendentes = ThisWorkbook.Sheets("Reembolsos Pendentes")
    Set aba_historica = ThisWorkbook.Sheets("Aba Historica")
    ' Cria um dicion�rio para armazenar informa��es do relat�rio de processamento.
    Set Dicionario_Relatorio_Processamento = CreateObject("Scripting.Dictionary")


    ' Exibe o formul�rio "Form_SAP".
    Form_SAP.Show

    ' Verifica se o checkbox "processamento_novos_chamados" no formul�rio est� marcado.
    If Form_SAP.checkbox_processamento_novos_chamados Then
        ' Se estiver marcado, atualiza a tabela de consulta da aba "Consolidado" em segundo plano.
        tabela_aba_plan_distribuicao.QueryTable.BackgroundQuery = False
        tabela_aba_plan_distribuicao.QueryTable.Refresh False
        tabela_aba_consolidado.QueryTable.BackgroundQuery = False
        tabela_aba_consolidado.QueryTable.Refresh False
    End If

    ' Desativa os alertas do Excel para evitar interrup��es.
    Application.DisplayAlerts = False

    ' Obt�m a �ltima linha preenchida nas abas "Consolidado" e "Reembolsos Pendentes".
    linha_fim = aba_consolidado.Range("A1048576").End(xlUp).Row
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row

    ' Verifica se o Excel est� rodando na Citrix (32 bits) ou n�o.
    If InStr(1, Application.OperatingSystem, "32") Then
        ' Se estiver na Citrix, busca o caminho da pasta de arquivos dos clientes.
        pasta_arquivos_clientes = BuscarPasta("", True)
        ' Define o caminho da pasta de anexos de detalhes de reembolso.
        pasta_anexos_detalhe_reembolso = Replace(pasta_arquivos_clientes, "Arquivos Clientes", "Anexos Detalhe Reembolso") & "\" & VBA.Format(VBA.Date, "dd.mm.yyyy")
        
        caminho_arquivo_modelo = Replace(pasta_arquivos_clientes, "Arquivos Clientes", "Modelo Solicitacao Ellevo.xlsx")
    ElseIf InStr(1, Application.OperatingSystem, "64") Then
        ' Se n�o estiver na Citrix (64 bits), exibe uma mensagem e encerra a execu��o.
        MsgBox "A automa��o deve ser executada com o Excel da Citrix."
        End
    End If
    
    ' Tratamento de erro: ignora o erro se o SAP GUI n�o estiver aberto.
    On Error Resume Next
    Set SapGui = GetObject("SAPGUI")
    If Err.number = -2147221020 Then
        ' Se o SAP GUI n�o estiver aberto, exibe uma mensagem e encerra a execu��o.
        MsgBox "Seu SAP est� fechado. Favor abri-lo e iniciar novamente a automa��o"
        End
    End If
    On Error GoTo 0
    ' Obt�m o objeto de scripting do SAP.
    Set Applic = SapGui.GetScriptingEngine
    ' Tratamento de erro: ignora o erro se a conex�o com o SAP n�o existir.
    On Error Resume Next
    Set connection = Applic.Connections(0)
    If Err.number = 614 Then
        ' Se a conex�o n�o existir, abre a conex�o com o SAP.
        Set connection = Applic.OpenConnection("002. P1L - SAP ECC Latin America (Single Sign On)", True)
        Set session = connection.Children(0)
    End If
    On Error GoTo 0
    ' Define o objeto de sess�o do SAP.
    Set session = connection.Children(0)
    ' Verifica o formato de data padr�o do SAP.
    tipo_data_sap = VerificarFormatoPadraoSAP
    ' Define as sess�es do SAP para as transa��es ZSD164, ZSD336, ZFI105 e FB03.
    Set session_2 = InteracaoTelasSAP(Nothing, 2, "ZSD164")
    Set session_3 = InteracaoTelasSAP(Nothing, 3, "ZSD336")
    
    session.findById("wnd[0]").maximize
    session_2.findById("wnd[0]").maximize
    session_3.findById("wnd[0]").maximize


    ' Chama a subrotina para armazenar informa��es dos chamados.
    Call Armazenar_Infos_Chamados
    ' Verifica se o checkbox "enviar_aprov_reembolsos_antigos" no formul�rio est� marcado.
    If Form_SAP.checkbox_enviar_aprov_reembolsos_antigos Then
        ' Se estiver marcado, chama a subrotina para verificar as linhas do SBWP.
        Call VerificarLinhasSBWP(session, "GERAL")
    End If
    ' Define um array com as atribui��es proibidas para chamados de contas a pagar.
    array_atribuicoes_proibidas = Array("*PROCESSADO AUTO*", "ELLEVO*", "*REEMBOLSO*", "*AUTOMACAO DEV*", "*UTILIZAR*", "*AG PROCESS SBWP*", "*ABATIDO PARCIAL*", "*ABATIDO TOTAL*")
    ' Define arrays com os cabe�alhos dos arquivos de cliente e abatimento.
    array_cabecalho_arquivo_cliente = Array("DocCompens", "Compensa�.", "Texto", "Tip", "Cliente", "Nome 1", "N� nota fiscal", "Mont.moeda doc.", "VencL�quid", "Bloq.", "Chv.ref.1", "Chv.ref.2", "Chave refer�ncia 3", "N� doc.", "Itm", "Atribui��o", "OC Associada")
    array_cabecalho_abatimento = Array("DocCompens", "Compensa�.", "Texto", "Tip", "Cliente", "Nome 1", "N� nota fiscal", "Mont.moeda doc.", "VencL�quid", "Bloq.", "Chv.ref.1", "Chv.ref.2", "Chave refer�ncia 3", "N� doc.", "Itm", "Atribui��o", "OC Associada", "Status", "Valor Residual")
     ' Define um array com as contas de fornecedores.
    array_contas_supplier = Array("50181303", "50181700")
    'Verifica se o checkbox de cria��o de chamados de reembolso est� marcado
    If Form_SAP.checkbox_verificar_abrir_chamado_reembolsos_aprovados Then
        ' Se estiver marcado, chama a subrotina para criar os chamados
        Call CriarChamadoReembolsosAprovados
    End If

    ' 'AQUI TEM QUE ENTRAR SUB QUE FAZ BUSCA DAS LINHAS NA FBL5N COM O DOC DE COMPENSACAO ARMAZENADO NO ARRAY_DOCS_F65 E COPIAR E JOGAR NUMA PLANILHA
    ' 'O DETALHE DAS LINHAS REEMBOLSADAS PARA ABRIR CHAMADO COM O CONTAS A PAGAR E PARA ENVIAR DETALHE PARA O CLIENTE

    ' Define um array para armazenar os pagadores com dados banc�rios.
    array_payers_com_dados_bancarios = Array()
    ' Chama a subrotina para preencher o array com os pagadores.
    Call PreencherArrayPayersCOMDadosBancarios
    ' Verifica se o checkbox de processamento de novos chamados est� marcado
    If Form_SAP.checkbox_processamento_novos_chamados Then
       'Se estiver marcado, chama a subrotina para processar os chamados
       Call ProcessarChamados
    End If
    ' Chama a subrotina para preencher a aba de relat�rio de processamento.
    Call PreencherAbaRelatorioProcessamento
    ' Ativa a aba de relat�rio de processamento.
    aba_relatorio_processamento.Activate
    ' Exibe uma mensagem informando que o processo foi conclu�do.
    MsgBox ("Processo de verifica��o e resposta de chamados Ellevo conclu�do!" & vbNewLine & "Por favor, verifique a aba Relat�rio de Processamento para visualizar todos os dados."), vbInformation

    ' Reativa a atualiza��o da tela e os alertas do Excel.
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True


End Sub
Sub Armazenar_Infos_Chamados()

    Dim status_chamado As String
    
    Set dictChamadosPendentes = CreateObject("Scripting.Dictionary")
    
    linha_fim_aba_historico_chamados_pendentes = aba_historico_chamados_pendentes.Range("A1048576").End(xlUp).Row
    If linha_fim_aba_historico_chamados_pendentes > 1 Then
        ' preenchendo o dict de chamado pendentes de outras rodadas com a estrutura de chave - valor
        ' ou seja {'123123':2,'456456',3}
        For linha = 2 To linha_fim_aba_historico_chamados_pendentes
            chamado = aba_historico_chamados_pendentes.Range("A" & linha).Value
            ultimo_tramite_verificado = aba_historico_chamados_pendentes.Range("B" & linha).Value
            If Not dictChamadosPendentes.exists(chamado) Then
                dictChamadosPendentes.Add chamado, ultimo_tramite_verificado
            End If
        Next linha
    End If
    ' preenchendo o array_chamados com todos os chamados encontrados na aba consolidado
    array_chamados = Array()
    
    For linha = 2 To linha_fim
        chamado = aba_consolidado.Range("A" & linha).Value
        status_chamado = aba_consolidado.Range("E" & linha).Value
        If Not UBound(VBA.Filter(array_chamados, chamado)) >= 0 Then
            array_chamados = Add_ao_Array(array_chamados, CStr(chamado))
        End If
    Next linha
    
    
    ' armazenando no dictGeralChamados todas as informa��es do chamado com a estrutura de chave - valor
    ' ou seja {'123123':
    '           {'info1':222222,'info2':333333,'info3':444444},
    '           '456456':
    '           {'info1':222222,'info2':333333,'info3':444444},
    '         }
    Set dictGeralChamados = CreateObject("Scripting.Dictionary")
    For i = LBound(array_chamados) To UBound(array_chamados)
        chamado = array_chamados(i)
        dictGeralChamados.Add chamado, PreencherDicionario()
    Next i
    
End Sub


Sub ProcessarChamados()

    Dim chave As Variant, chave_2 As Variant
    Dim status_chamado_string As String

' VERIFICACAO POR CHAMADO
    For Each chave In dictGeralChamados.keys
    
        chamado = chave
        Call AlimentarDicionario_Relatorio_Processamento("Chamados processados pela automa��o: ", chamado)
        OCs_incorretas = ""
        OCs_erro_zsd164_1 = ""
        OCs_erro_zsd164_2 = ""
        cnpj = ""
        numero_OC = ""
        numero_NFD = ""
        qtde_NFD_OC_chamado = ""
        status_chamado = ""
        condicao_OC_incorreta = False
        condicao_erro_zsd164_1 = False
        condicao_erro_zsd164_2 = False
        condicao_cliente_sem_dados_bancarios = False
        conta_bloqueada = False
        condicao_OCs_reembolso = False
        erro_zfi156 = False
        Set dictChamadoAtual = dictGeralChamados(chave)
        For Each chave_2 In dictChamadoAtual.keys
            Select Case chave_2
                Case "CNPJ": cnpj = "*" & Left(dictChamadoAtual(chave_2), 10) & "*"
                Case "N�mero OC": numero_OC = dictChamadoAtual(chave_2)
                Case "N� ND/NFD": numero_NFD = dictChamadoAtual(chave_2)
                Case "Quantidade de ND/notas para consulta": qtde_NFD_OC_chamado = dictChamadoAtual(chave_2)
            End Select
        Next chave_2
        status_chamado_string = Application.WorksheetFunction.VLookup(CLng(chamado), aba_consolidado.Columns("A:E"), 5, False)
        If status_chamado_string = "WAITING" Or (status_chamado_string = "INPROGRESS" And Not dictChamadosPendentes.exists(CLng(chamado))) Then
            GoTo proximo_chamado
        End If
        If VerificarCNPJ Then
            Select Case qtde_NFD_OC_chamado
                Case "01"
                    If Application.WorksheetFunction.CountIf(aba_historica.Columns("A:A"), numero_OC) > 0 And Not numero_OC = "" Then
                        Call GerarTokenTramite("OC_JA_CONSULTADA")
                        GoTo proximo_chamado
                    End If
                    array_geral_linhas_abertas_FBL5N = Array(array_cabecalho_arquivo_cliente)
                    array_linhas_compensadas_FB03 = Array(array_cabecalho_arquivo_cliente)
                    array_linhas_detalhe_abatimento = Array(array_cabecalho_abatimento)
                    condicao_payer = ""
                    linhas_compensadas = False
                    linhas_abertas = False
                    conta_bloqueada = False
                    condicao_cliente_sem_dados_bancarios = False
                    array_doc_compensacoes = Array()
                    array_num_doc_zfi105 = Array()
                    numero_OC = TratativasOC(numero_OC)
                    If EtapaZSD164 Then
                        Call EtapaZSD336
                        Call EtapaZFI105
                        If VerificarCondicaoClienteFBL5N Then
                            If linhas_abertas Then
                                Call ProcessarAbatimentoOuReembolso
                            End If
                            If erro_zfi156 Then
                                GoTo proximo_chamado
                            End If
                            If Not condicao_cliente_sem_dados_bancarios Then
                                If Not conta_bloqueada Then
                                    Call AlimentarArquivoCliente
                                    Call SalvarFecharArquivoCliente
                                    If condicao_OCs_reembolso Then
                                        status_chamado = "1"
                                    End If
                                    Call GerarTokenTramite("ENVIO_ANEXO_OCS_VERIFICADAS")
                                Else
                                    session.findById("wnd[0]").sendVKey 5
                                    Call AlterarAtribuicao(session, "CTA BLOQUEADA")
                                    MsgBox "Execu��o do RPA de compensa��o em andamento, favor rodar a automa��o quando o RPA for finalizado."
                                    GoTo fim
                                End If
                            Else
                                Call GerarTokenTramite("SEM_DADOS_BANCARIOS")
                            
                            End If ' Fechando erro zfi156
                        End If ' fechando if VerificarCondicaoClienteFBL5N
                    End If ' fechando if EtapaZSD164
                Case "Acima de 01"
                    array_nome_abas = Array()
                    array_geral_linhas_abertas_FBL5N = Array(array_cabecalho_arquivo_cliente)
                    array_linhas_compensadas_FB03 = Array(array_cabecalho_arquivo_cliente)
                    array_linhas_detalhe_abatimento = Array(array_cabecalho_abatimento)
                    If APIBuscaChamadoUnico Then
                        Set aba_1_arquivo_anexo_chamado_atual = arquivo_anexo_chamado_atual.Sheets(1)
                        If aba_1_arquivo_anexo_chamado_atual.Cells(rngEncontrado.Row, rngEncontrado.Column + 3).Value = "" Then
                            aba_1_arquivo_anexo_chamado_atual.Cells(rngEncontrado.Row, rngEncontrado.Column + 3).Value = "AN�LISE ELECTROLUX"
                        End If
                        linha_fim = aba_1_arquivo_anexo_chamado_atual.Cells(1048576, rngEncontrado.Column).End(xlUp).Row
                        If rngEncontrado.Row = linha_fim Then
                            arquivo_anexo_chamado_atual.Close
                            Call GerarTokenTramite("NENHUMA_OC_INFORMADA")
                            GoTo proximo_chamado
                        End If
                        Set dictOCNumeroDOC_ZFI105 = CreateObject("Scripting.Dictionary")
                        ' primeira itera��o nas linhas que verificar� se todas as linhas tem OCs em conformidade, se sim
                        ' segue para a proxima etapa, se n�o, cria o tramite detalhando as OCs erradas e tamb�m acrescenta na aba
                        ' chamados pendentes o chamado em quest�o
                        For linha = rngEncontrado.Row + 1 To linha_fim
                            numero_OC = aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column).Value
                            numero_OC = TratativasOC(numero_OC)
                            If Application.WorksheetFunction.CountIf(aba_historica.Columns("A:A"), numero_OC) > 0 And Not numero_OC = "" Then
                                chamado_anterior = Application.WorksheetFunction.VLookup(CLng(numero_OC), aba_historica.Columns("A:B"), 2, False)
                                status_OC = Application.WorksheetFunction.VLookup(CLng(numero_OC), aba_historica.Columns("A:C"), 3, False)
                                data_solicitacao_reembolso_abatimento = Application.WorksheetFunction.VLookup(CLng(numero_OC), aba_historica.Columns("A:D"), 4, False)
                                If status_OC = "REEMBOLSO" Then
                                    aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3).Value = "OC j� consultada anteriormente no chamado " & chamado_anterior & _
                                    ", programando reembolso para pagamento no dia " & data_solicitacao_reembolso_abatimento & " conforme detalhado em anexo do chamado mencionado."
                                ElseIf status_OC = "ABATIMENTO" Then
                                    aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3).Value = "OC j� consultada anteriormente no chamado " & chamado_anterior & _
                                    ", realizando abatimento no dia " & data_solicitacao_reembolso_abatimento & " conforme detalhado em anexo do chamado mencionado."
                                ElseIf status_OC = "SEM CREDITOS EM ABERTO ENCONTRADOS" Then
                                    aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3).Value = "OC j� consultada anteriormente. N�o foram encontrados cr�ditos de devolu��o associados � ela."
                                End If
                            Else
                                If numero_OC = "" Then
                                    aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3).Value = "Campo de OC vazio"
                                Else
                                    array_num_doc_zfi105 = Array()
                                    If EtapaZSD164 Then
                                        Call EtapaZSD336
                                        Call EtapaZFI105
                                        For i = LBound(array_num_doc_zfi105) To UBound(array_num_doc_zfi105)
                                            If Not dictOCNumeroDOC_ZFI105.exists(array_num_doc_zfi105(i)) Then
                                                dictOCNumeroDOC_ZFI105.Add array_num_doc_zfi105(i), numero_OC
                                            End If
                                        Next i
                                    End If
                                End If
                            End If
                        Next linha
                        linhas_compensadas = False
                        linhas_abertas = False
                        condicao_doc_compensacao_repetido = False
                        condicao_chamado_supplier = False
                        condicao_payer = ""
                        array_doc_compensacoes = Array()
                        Dim condicao_linhas_sem_status As Boolean
                        condicao_linhas_sem_status = False
                        For linha = rngEncontrado.Row + 1 To linha_fim
                            If aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3).Value = "" Then
                                condicao_linhas_sem_status = True
                            End If
                        Next linha
                        Set arquivo_cliente = arquivo_anexo_chamado_atual
                        If Not condicao_linhas_sem_status Then
                            If dictChamadosPendentes.exists(chamado) Then
                                If arquivo_cliente Is Nothing Then
                                    Set arquivo_cliente = arquivo_anexo_chamado_atual
                                End If
                                arquivo_cliente.Close
                                ' Libera a vari�vel do objeto
                                Set arquivo_cliente = Nothing
                                Call APITrocaResponsavelChamado(3)
                                GoTo proximo_chamado
                            End If
                            status_chamado = "9"
                            Call SalvarFecharArquivoCliente
                            Call GerarTokenTramite("ENVIO_ANEXO_OCS_VERIFICADAS")
                            GoTo proximo_chamado
                        End If
                        Call VerificarCondicaoClienteFBL5N
                        ' CRIAR ARQUIVO COM AS LINHAS ABERTAS E COMPENSADAS DE TODAS AS OCS CONSULTADAS
                        If condicao_chamado_supplier Then
                            Set arquivo_cliente = arquivo_anexo_chamado_atual
                            arquivo_cliente.Close
                            GoTo proximo_chamado
                        End If
                        If linhas_abertas Then
                            Call ProcessarAbatimentoOuReembolso
                        End If
                        If erro_zfi156 Then
                            GoTo proximo_chamado
                        End If
                        If condicao_cliente_sem_dados_bancarios Then
                            Call GerarTokenTramite("SEM_DADOS_BANCARIOS")
                            Set arquivo_cliente = arquivo_anexo_chamado_atual
                            arquivo_cliente.Close
                            GoTo proximo_chamado
                        ElseIf conta_bloqueada Then
                            Set arquivo_cliente = arquivo_anexo_chamado_atual
                            arquivo_cliente.Close
                            MsgBox "Execu��o do RPA de compensa��o em andamento, favor rodar a automa��o quando o RPA for finalizado."
                            GoTo fim
                        End If
                        If condicao_OCs_reembolso Then
                            status_chamado = "1"
                        End If
                        Call AlimentarArquivoCliente
                        For linha = rngEncontrado.Row + 1 To linha_fim
                            If aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3) = "" Then
                                numero_OC = aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column).Value
                                numero_OC = TratativasOC(numero_OC)
                                For i = LBound(array_linhas_compensadas_FB03) To UBound(array_linhas_compensadas_FB03)
                                    Dim OC_consultada As String
                                    OC_consultada = array_linhas_compensadas_FB03(i)(13)
                                    If dictOCNumeroDOC_ZFI105(OC_consultada) = numero_OC Then
                                        aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3) = "OC com cr�dito de devolu��o j� utilizados em abatimentos/reembolsos anteriores. Conferir aba 'Cr�ditos Ja Utilizados'"
                                        Exit For
                                    End If
                                Next i
                            End If
                            If aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3) = "" Then
                                numero_OC = aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column).Value
                                numero_OC = TratativasOC(numero_OC)
                                For i = LBound(array_geral_linhas_abertas_FBL5N) To UBound(array_geral_linhas_abertas_FBL5N)
                                    If dictOCNumeroDOC_ZFI105(array_geral_linhas_abertas_FBL5N(i)(13)) = numero_OC Then
                                        If condicao_payer = "abatidos" Then
                                            aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3) = "OC com cr�dito de devolu��o abatidos de t�tulos em aberto. Favor Consultar aba 'Detalhe Abatimento'"
                                            Call AlimentarAbaHistorica("ABATIMENTO")
                                            Exit For
                                        ElseIf condicao_payer = "reembolsados" Then
                                            aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3) = "OC com cr�dito de devolu��o enviados para aprova��o do reembolso. Favor Consultar aba 'Detalhe Reembolso'"
                                            Call AlimentarAbaHistorica("REEMBOLSO")
                                            Exit For
                                        End If
                                    End If
                                Next i
                            End If
                        Next linha
                        aba_1_arquivo_anexo_chamado_atual.Columns(rngEncontrado.Column + 3).AutoFit
                        Call SalvarFecharArquivoCliente
                        Call GerarTokenTramite("ENVIO_ANEXO_OCS_VERIFICADAS")
                    End If
                End Select ' finalizacao verificacao qtde_NFD_OC_chamado
            End If 'finalizacao VerificarCNPJ
proximo_chamado:
        Set dictOCNumeroDOC_ZFI105 = Nothing
        Debug.Print chamado
    Next chave
fim:
                
End Sub

Private Function VerificarCNPJ() As Boolean

    Dim payer As String, resultado As String
    Dim qtde_payers As Long

    ' Entra na transa��o FBL5N do SAP
    session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
    ' Envia um Enter para executar a transa��o
    session.findById("wnd[0]").sendVKey 0
    ' Clica no bot�o "Sele��o Din�mica"
    session.findById("wnd[0]").sendVKey 4
    ' Trata poss�vel erro caso a aba n�o exista
    On Error Resume Next
    ' Seleciona a aba "Outros campos"
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006").Select
    ' Restaura o tratamento de erros padr�o
    On Error GoTo 0
    ' Preenche o campo de CNPJ com o valor da vari�vel "cnpj"
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = cnpj
    ' Clica no bot�o "Executar"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    ' Verifica se a mensagem na barra de status indica que nenhum valor foi encontrado para a sele��o
    If session.findById("wnd[0]/sbar").text = "Nenhum valor para esta sele��o" Then
        ' Define o valor da fun��o como Falso, indicando que o CNPJ n�o foi encontrado
        VerificarCNPJ = False
        ' Sai da fun��o
        Exit Function
    End If
    
    ' Verifica a posi��o da coluna "Cliente" na tabela
    x_cliente = VerificarColuna(1, 1, 300, "wnd[1]/usr/lbl[", "Cliente")
    ' Preenche um array com os payers (clientes) encontrados na FBL5N
    Call PreencherPayersCNPJ(session, 3, 100, "wnd[1]/usr/lbl[", x_cliente)

    ' Inicialmente define o valor da fun��o como Falso
    VerificarCNPJ = False
    ' Loop atrav�s de cada payer encontrado no array
    For i2 = LBound(array_payers_encontrados) To UBound(array_payers_encontrados)
        payer = array_payers_encontrados(i2)
        ' Trata poss�vel erro caso o valor n�o seja encontrado na planilha
        On Error Resume Next
        ' Busca o payer na coluna A da aba "aba_plan_distribuicao" e retorna o valor da coluna D
        resultado = Application.VLookup(payer, aba_plan_distribuicao.Columns("A:D"), 4, False)
        ' Restaura o tratamento de erros padr�o
        On Error GoTo 0
        ' Verifica se o resultado da busca n�o est� vazio
        If resultado <> "" Then
            ' Define o valor da fun��o como Verdadeiro, indicando que o CNPJ foi encontrado na planilha de distribui��o
            VerificarCNPJ = True
            ' Sai da fun��o
            Exit Function
        End If
    Next i2
    ' Se o loop terminar sem encontrar o CNPJ na planilha, registra no relat�rio que o chamado n�o tem CNPJ cadastrado
    Call AlimentarDicionario_Relatorio_Processamento("Chamados sem CNPJs cadastrado como Oficinas Autorizadas ou Canal Direto I: ", chamado)
    
End Function

Private Function EtapaZSD164() As Boolean

    Dim elemento_tabela_zsd164_superior As Object, elemento_tabela_zsd164_inferior As Object
    
    ' Inicializa a fun��o como Verdadeira
    EtapaZSD164 = True
    
    ' Verifica se o comprimento do n�mero da OC � diferente de 6
    If Len(numero_OC) <> 6 Then
        ' Se o n�mero da OC estiver vazio, atribui uma mensagem informativa
        If numero_OC = "" Then
            numero_OC = "Informado campo de OC vazio"
        End If
        ' Se a quantidade de NFD/OC no chamado for igual a "01"
        If qtde_NFD_OC_chamado = "01" Then
            ' Armazena o n�mero da OC incorreta
            OCs_incorretas = numero_OC
            ' Define a fun��o como Falsa
            EtapaZSD164 = False
            ' Gera um token de tr�mite indicando OC incorreta
            Call GerarTokenTramite("AVISO_OC_INCORRETA")
            ' Registra no relat�rio que o chamado possui OC informada incorretamente
            Call AlimentarDicionario_Relatorio_Processamento("Chamados com OCs informadas incorretamente: ", chamado)
            ' Sai da fun��o
            Exit Function
        ' Sen�o, se a quantidade de NFD/OC no chamado for "Acima de 01"
        ElseIf qtde_NFD_OC_chamado = "Acima de 01" Then
            ' Define a condi��o de OC incorreta como Verdadeira
            condicao_OC_incorreta = True
            ' Se o n�mero da OC n�o estiver vazio, concatena na string de OCs incorretas
            If numero_OC <> "" Then
                OCs_incorretas = OCs_incorretas & "/" & numero_OC
            End If
            ' Define a fun��o como Falsa
            EtapaZSD164 = False
            ' Informa na planilha de anexo do chamado que a OC est� incorreta
            aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3).Value = "OC Incorreta"
            Call AlimentarAbaHistorica("OC INCORRETA")
            ' Registra no relat�rio a OC informada incorretamente para o chamado
            Call AlimentarDicionario_Relatorio_Processamento("OCs informadas incorretamente referente ao chamado " & chamado & ": ", numero_OC)
            ' Sai da fun��o
            Exit Function
        End If
    End If

    ' Entra na transa��o ZSD164 do SAP (segunda sess�o)
    session_2.findById("wnd[0]/tbar[0]/okcd").text = "/N ZSD164"
    ' Envia um Enter para executar a transa��o
    session_2.findById("wnd[0]").sendVKey 0
    ' Seleciona a op��o de menu para buscar ocorr�ncias
    session_2.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    ' Preenche o campo de tipo de ocorr�ncia com "zsd164"
    session_2.findById("wnd[1]/usr/txtV-LOW").text = "zsd164"
    ' Limpa o campo de nome
    session_2.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    ' Clica no bot�o de executar
    session_2.findById("wnd[1]/tbar[0]/btn[8]").press
    ' Preenche o campo de n�mero da ocorr�ncia com o n�mero da OC
    session_2.findById("wnd[0]/usr/txtS_OCCUR-LOW").text = numero_OC
    ' Preenche a data inicial da pesquisa (01/01/2017)
    session_2.findById("wnd[0]/usr/ctxtS_PERIOD-LOW").text = VBA.Format(VBA.DateSerial(2017, 1, 1), tipo_data_sap)
    ' Preenche a data final da pesquisa (data atual + 500 dias)
    session_2.findById("wnd[0]/usr/ctxtS_PERIOD-HIGH").text = VBA.Format(Date + 500, tipo_data_sap)
    ' Preenche a sociedade com "BR10"
    session_2.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "BR10"
    ' Clica no bot�o de executar
    session_2.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Trata poss�vel erro se a mensagem n�o for encontrada
    On Error Resume Next
    ' Tenta focar no campo de texto da mensagem
    session_2.findById("wnd[1]/usr/txtMESSTXT1").SetFocus
    ' Se n�o houver erro e a quantidade de NFD/OC for "01" (indica erro na ZSD164)
    If Err.number = 0 And qtde_NFD_OC_chamado = "01" Then
        ' Define a fun��o como Falsa
        EtapaZSD164 = False
        ' Armazena o n�mero da OC com erro na etapa ZSD164 (cen�rio 1)
        OCs_erro_zsd164_1 = numero_OC
        ' Gera um token de tr�mite indicando erro na ZSD164 (cen�rio 1)
        Call GerarTokenTramite("ERRO_ZSD164_1")
        ' Sai da fun��o
        Exit Function
    ' Sen�o, se n�o houver erro e a quantidade de NFD/OC for "Acima de 01"
    ElseIf Err.number = 0 And qtde_NFD_OC_chamado = "Acima de 01" Then
        ' Define a condi��o de erro na ZSD164 (cen�rio 1) como Verdadeira
        condicao_erro_zsd164_1 = True
        ' Se o n�mero da OC n�o estiver vazio, concatena na string de OCs com erro
        If numero_OC <> "" Then
            OCs_erro_zsd164_1 = OCs_erro_zsd164_1 & "/" & numero_OC
        End If
        ' Define a fun��o como Falsa
        EtapaZSD164 = False
        ' Informa na planilha de anexo do chamado que a OC n�o est� dispon�vel/encontrada
        aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3).Value = "OC n�o dispon�vel e/ou n�o encontrada"
        ' Registra no relat�rio a OC com erro na etapa ZSD164 para o chamado
        Call AlimentarDicionario_Relatorio_Processamento("OCs com erro etapa ZSD164 referente ao Chamado Ellevo " & chamado & ": ", numero_OC)
        ' Sai da fun��o
        Exit Function
    End If
    ' Restaura o tratamento de erros padr�o
    On Error GoTo 0
    
    ' Define o objeto da tabela superior da ZSD164
    Set elemento_tabela_zsd164_superior = session_2.findById("wnd[0]/usr/subOCCURRENCE_HEADER:ZSDRMR_OCCURRENCE_COCKPIT:9001/cntlC_HEADER/shellcont/shell")
    ' Obt�m o n�mero da fatura de devolu��o da linha atual da tabela
    fatura_devolucao = elemento_tabela_zsd164_superior.GetCellValue(elemento_tabela_zsd164_superior.currentCellRow, "BILLING")
    ' Obt�m o n�mero da ordem de devolu��o da linha atual da tabela
    ordem_devolucao = elemento_tabela_zsd164_superior.GetCellValue(elemento_tabela_zsd164_superior.currentCellRow, "VBELN")
    
    ' Se a fatura de devolu��o estiver vazia e a quantidade de NFD/OC for "01" (indica erro na ZSD164)
    If fatura_devolucao = "" And qtde_NFD_OC_chamado = "01" Then
        ' Define a fun��o como Falsa
        EtapaZSD164 = False
        ' Armazena o n�mero da OC com erro na etapa ZSD164 (cen�rio 2)
        OCs_erro_zsd164_2 = numero_OC
        ' Gera um token de tr�mite indicando erro na ZSD164 (cen�rio 2)
        Call GerarTokenTramite("ERRO_ZSD164_2")
        ' Registra no relat�rio a OC com erro na etapa ZSD164 para o chamado
        Call AlimentarDicionario_Relatorio_Processamento("OCs com erro etapa ZSD164 referente ao Chamado Ellevo " & chamado & ": ", numero_OC)
    ' Sen�o, se a fatura de devolu��o estiver vazia e a quantidade de NFD/OC for "Acima de 01"
    ElseIf fatura_devolucao = "" And qtde_NFD_OC_chamado = "Acima de 01" Then
        ' Define a condi��o de erro na ZSD164 (cen�rio 2) como Verdadeira
        condicao_erro_zsd164_2 = True
        ' Se o n�mero da OC n�o estiver vazio, concatena na string de OCs com erro
        If numero_OC <> "" Then
            OCs_erro_zsd164_2 = OCs_erro_zsd164_2 & "/" & numero_OC
        End If
        ' Define a fun��o como Falsa
        EtapaZSD164 = False
        ' Informa na planilha de anexo do chamado que a OC n�o foi finalizada/registrada
        aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3).Value = "OC n�o finalizada e n�o registrada"
        ' Registra no relat�rio a OC com erro na etapa ZSD164 para o chamado
        Call AlimentarDicionario_Relatorio_Processamento("OCs com erro etapa ZSD164 referente ao Chamado Ellevo " & chamado & ": ", numero_OC)
    End If
    
End Function

Private Sub EtapaZSD336()

    Dim elemento_tabela_zsd336 As Object

    ' Entra na transa��o ZSD336 do SAP (terceira sess�o)
    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N ZSD336"
    ' Envia um Enter para executar a transa��o
    session_3.findById("wnd[0]").sendVKey 0
    ' Clica no bot�o para exibir as variantes de layout
    session_3.findById("wnd[0]/tbar[1]/btn[17]").press
    ' Define o objeto da tabela de variantes de layout
    Set elemento_tabela_zsd336 = session_3.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
    ' Loop atrav�s das linhas da tabela de variantes (at� 10 linhas)
    For i2 = 0 To 10
        ' Verifica se a variante "ZSD336" � encontrada na coluna "VARIANT"
        If elemento_tabela_zsd336.GetCellValue(i2, "VARIANT") = "ZSD336" Then
            ' Define a coluna atual como "VARIANT"
            elemento_tabela_zsd336.currentCellColumn = "VARIANT"
            ' Seleciona a linha atual
            elemento_tabela_zsd336.selectedRows = i2
            ' Define a linha atual
            elemento_tabela_zsd336.currentCellRow = i2
            ' Simula um duplo clique na c�lula atual para selecionar a variante
            elemento_tabela_zsd336.doubleClickCurrentCell
            ' Sai do loop ap�s encontrar a variante
            Exit For
        End If
    Next i2
    ' Preenche o campo de ordem de devolu��o com o valor da vari�vel "ordem_devolucao"
    session_3.findById("wnd[0]/usr/ctxtS_VBELN-LOW").text = ordem_devolucao
    ' Clica no bot�o de executar
    session_3.findById("wnd[0]/tbar[1]/btn[8]").press
    ' Libera a vari�vel de objeto
    Set elemento_tabela_zsd336 = Nothing
    ' Define o objeto da tabela de documentos da ZSD336
    Set elemento_tabela_zsd336 = session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
    ' Inicializa um array para armazenar os n�meros dos documentos da ZSD336
    array_num_doc_zsd336 = Array()
    ' Chama a fun��o para preencher o array com os n�meros dos documentos da coluna "DOCNUM"
    Call PreencherArrayDocsSap(array_num_doc_zsd336, elemento_tabela_zsd336, "DOCNUM")
    
    
End Sub
Private Sub EtapaZFI105()

    Dim elemento_tabela_zfi105 As Object
    
    ' Chama a sub-rotina para listar informa��es da coluna BB com base no array de documentos da ZSD336
    Call ListarInfosColunaBB(array_num_doc_zsd336)
    ' Obt�m a �ltima linha preenchida na coluna BB da aba consolidado
    linha_fim_docs_sap = aba_consolidado.Range("BB1048576").End(xlUp).Row
    ' Copia o intervalo de documentos da coluna BB (a partir da segunda linha)
    aba_consolidado.Range("BB2:BB" & linha_fim_docs_sap).Copy
    ' Entra na transa��o ZFI105 do SAP (quarta sess�o)
    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N ZFI105"
    ' Envia um Enter para executar a transa��o
    session_3.findById("wnd[0]").sendVKey 0
    ' Seleciona a op��o de menu para buscar documentos financeiros
    session_3.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    ' Preenche o campo de variante com "zfi105robo"
    session_3.findById("wnd[1]/usr/txtV-LOW").text = "zfi105robo"
    ' Limpa o campo de nome
    session_3.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    ' Clica no bot�o de executar
    session_3.findById("wnd[1]/tbar[0]/btn[8]").press
    ' Clica no bot�o de m�ltipla sele��o para o campo de n�mero do documento
    session_3.findById("wnd[0]/usr/btn%_S_DOCNUM_%_APP_%-VALU_PUSH").press
    ' Clica no bot�o de colar da �rea de transfer�ncia
    session_3.findById("wnd[1]/tbar[0]/btn[16]").press
    ' Clica no bot�o de selecionar tudo
    session_3.findById("wnd[1]/tbar[0]/btn[24]").press
    ' Clica no bot�o de copiar
    session_3.findById("wnd[1]/tbar[0]/btn[8]").press
    ' Clica no bot�o de executar
    session_3.findById("wnd[0]/tbar[1]/btn[8]").press
    ' Define o objeto da tabela de documentos financeiros
    Set elemento_tabela_zfi105 = session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    ' Chama a sub-rotina para preencher o array de documentos da ZFI105 com os valores da coluna "BELNR"
    Call PreencherArrayDocsSap(array_num_doc_zfi105, elemento_tabela_zfi105, "BELNR")

    
End Sub

Private Function VerificarCondicaoClienteFBL5N() As Boolean

    Dim data_compensacao As String, texto As String, tipo_doc As String, cliente As String, nome As String, referencia As String, montante As String, venc_liquido As String, _
        bloq_adv As String, chave_ref_1 As String, chave_ref_2 As String, chave_ref_3 As String, num_doc As String, item As String, atribuicao As String, texto_sbar As String
    Dim quantidade_linhas As Long

    

    
    session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N" ' Entra na transa��o FBL5N do SAP (primeira sess�o)
    session.findById("wnd[0]").sendVKey 0 ' Envia um Enter para executar a transa��o
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select ' Seleciona a op��o de menu para buscar linhas em aberto
    session.findById("wnd[1]/usr/txtV-LOW").text = "id328" ' Preenche o campo de variante com "id328"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "" ' Limpa o campo de nome
    session.findById("wnd[1]/tbar[0]/btn[8]").press ' Clica no bot�o de executar
    session.findById("wnd[0]/usr/chkX_SHBV").Selected = True ' Marca a caixa de sele��o para exibir itens especiais
    session.findById("wnd[0]/tbar[1]/btn[16]").press ' Clica no bot�o de executar
    
    
    ' Se a quantidade de NFD/OC no chamado for "Acima de 01"
    If qtde_NFD_OC_chamado = "Acima de 01" Then
        Dim chave_3, array_chave_atual() As Variant
        array_chave_atual = Array()
        For Each chave_3 In dictOCNumeroDOC_ZFI105.keys
            array_chave_atual = Add_ao_Array(array_chave_atual, chave_3)
        Next chave_3
        Call ListarInfosColunaBB(array_chave_atual)
    ElseIf qtde_NFD_OC_chamado = "01" Then
        Call ListarInfosColunaBB(array_num_doc_zfi105) ' Chama a sub-rotina para listar informa��es da coluna BB com base no array de documentos da ZFI105
    End If
    
    linha_fim_docs_sap = aba_consolidado.Range("BB1048576").End(xlUp).Row ' Obt�m a �ltima linha preenchida na coluna BB da aba consolidado
    
    aba_consolidado.Range("BB2:BB" & linha_fim_docs_sap).Copy ' Copia o intervalo de documentos da coluna BB (a partir da segunda linha)

    ' Clica no bot�o de m�ltipla sele��o para o campo de n�mero do documento
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN003_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press ' Clica no bot�o de colar da �rea de transfer�ncia
    session.findById("wnd[1]/tbar[0]/btn[24]").press ' Clica no bot�o de selecionar tudo
    session.findById("wnd[1]/tbar[0]/btn[8]").press  ' Clica no bot�o de copiar
    ' Preenche o campo de tipo de documento com "R1" (faturas)
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN006-LOW").text = "R1"
    session.findById("wnd[0]/usr/radX_AISEL").Select ' Seleciona o radio button para considerar os payers informados
    session.findById("wnd[0]/usr/chkX_SHBV").Selected = False ' Desmarca a caixa de sele��o para exibir itens especiais (j� foi feita anteriormente)
    
    
    Call ListarInfosColunaBB(array_payers_encontrados) ' Chama a sub-rotina para listar os payers encontrados na coluna BB
    linha_fim_array_payers_encontrados = aba_consolidado.Range("BB1048576").End(xlUp).Row ' Obt�m a �ltima linha preenchida na coluna BB da aba consolidado para os payers
    aba_consolidado.Range("BB2:BB" & linha_fim_array_payers_encontrados).Copy ' Copia o intervalo de payers da coluna BB (a partir da segunda linha)
    
    
    session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press ' Clica no bot�o de m�ltipla sele��o para o campo de cliente
    session.findById("wnd[1]/tbar[0]/btn[16]").press ' Clica no bot�o de colar da �rea de transfer�ncia
    ' Preenche valores espec�ficos para exclus�o de clientes (Supplier)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "50181303"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "50181700"
    
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").SetFocus ' Define o foco no pr�ximo campo
    session.findById("wnd[1]/tbar[0]/btn[24]").press ' Clica no bot�o de selecionar tudo
    session.findById("wnd[1]/tbar[0]/btn[8]").press ' Clica no bot�o de copiar
    session.findById("wnd[0]/tbar[1]/btn[8]").press ' Clica no bot�o de executar

    Debug.Print chamado
    ' Se a barra de status indicar que linhas foram exibidas
    If Left(session.findById("wnd[0]/sbar").text, 12) = "S�o exibidas" Then
        ' Chama a sub-rotina para definir o eixo X das colunas da tabela
        Call SetEixoXColunas ' Obt�m o payer associado � OC da linha atual da tabela
        payer_associado_OC = session.findById("wnd[0]/usr/lbl[" & x_cliente & ",4]").text ' Verifica se o payer encontrado pertence � lista de contas Supplier
        If payer_associado_OC = "50203843" Then
            VerificarCondicaoClienteFBL5N = False ' Define a fun��o como Falsa
            Exit Function
        End If
        If qtde_NFD_OC_chamado = "Acima de 01" Then
            session.findById("wnd[0]/usr/lbl[" & x_cliente & ",2]").SetFocus
            session.findById("wnd[0]").sendVKey 2
            session.findById("wnd[0]/tbar[1]/btn[38]").press
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "50181303"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "50181700"
            session.findById("wnd[2]/tbar[0]/btn[8]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            On Error Resume Next
            session.findById("wnd[0]/usr/lbl[" & x_cliente & ",4]").SetFocus
            If Err.number = 0 Then
                payer_associado_OC = session.findById("wnd[0]/usr/lbl[" & x_cliente & ",4]").text
            End If
            On Error GoTo 0
        End If
        If UBound(VBA.Filter(array_contas_supplier, payer_associado_OC)) = 0 Then ' Chama a sub-rotina para trocar o respons�vel do chamado via API (Supplier)
            If x_num_doc = 0 Then
                Call SetEixoXColunas
            End If
            num_doc_supplier = session.findById("wnd[0]/usr/lbl[" & x_num_doc & ",4]").text
            Call APITrocaResponsavelChamado(2) ' Registra no relat�rio que o chamado possui OCs relacionadas a cr�ditos com c�digo da Supplier
            Call AlimentarDicionario_Relatorio_Processamento("Chamados Ellevo com OCs relacionadas a cr�ditos com c�digo da Supplier: ", chamado)
            VerificarCondicaoClienteFBL5N = False ' Define a fun��o como Falsa
            Exit Function
        End If
        
        session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select
        session.findById("wnd[0]/usr/lbl[" & x_doc_compensacao & ",2]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        session.findById("wnd[0]/tbar[1]/btn[38]").press
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
        i2 = 0
        For i = 0 To 9
            If i2 > 9 Then
                Exit For
            End If
            On Error Resume Next
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").SetFocus
            If Err.number <> 0 Then
                session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = i
                i = 1
            End If
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").text = "*" & i2 & "*"
            i2 = i2 + 1
            On Error GoTo 0
        Next i
        session.findById("wnd[2]/tbar[0]/btn[8]").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        On Error Resume Next
        session.findById("wnd[0]/usr/lbl[" & x_tipo_doc & ",4]").SetFocus
        If Err.number = 0 Then
            Call PreencherArrayLinhasCondicaoAtual(session, i4, i5, "VERIFICACAO") ' Chama a sub-rotina para preencher um array com informa��es da linha atual da tabela
        End If
        On Error GoTo 0
               
        session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select
        
        session.findById("wnd[0]/usr/lbl[" & x_doc_compensacao & ",2]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        session.findById("wnd[0]/tbar[1]/btn[38]").press
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select

        i2 = 0
        For i = 0 To 9
            If i2 > 9 Then
                Exit For
            End If
            On Error Resume Next
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1," & i & "]").SetFocus
            If Err.number <> 0 Then
                session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.Position = i
                i = 1
            End If
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1," & i & "]").text = "*" & i2 & "*"
            i2 = i2 + 1
            On Error GoTo 0
        Next i
        session.findById("wnd[2]/tbar[0]/btn[8]").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        On Error Resume Next
        session.findById("wnd[0]/usr/lbl[" & x_tipo_doc & ",4]").SetFocus
        If Err.number = 0 Then
            Call PreencherArrayLinhasCondicaoAtual(session, i4, i5, "VERIFICACAO") ' Chama a sub-rotina para preencher um array com informa��es da linha atual da tabela
        End If
        On Error GoTo 0
        
        If linhas_compensadas Then ' Se houver linhas compensadas
            
            For i = LBound(array_doc_compensacoes) To UBound(array_doc_compensacoes)
                doc_compensacao = array_doc_compensacoes(i)
                Call EtapaFB03(session_3, doc_compensacao, "VERIFICACAO")
            Next i
        End If
        VerificarCondicaoClienteFBL5N = True
    ' Sen�o, se a quantidade de NFD/OC for "01" (e n�o houver cr�ditos associados)
    ElseIf qtde_NFD_OC_chamado = "01" Then
        
        VerificarCondicaoClienteFBL5N = False ' Define a fun��o como Falsa
        OCs_incorretas = numero_OC ' Armazena o n�mero da OC incorreta
        Call GerarTokenTramite("AVISO_OC_SEM_CREDITOS_ASSOCIADOS") ' Gera um token de tr�mite indicando OC sem cr�ditos associados
        Call AlimentarDicionario_Relatorio_Processamento("Chamados com OCs sem cr�ditos associados: ", chamado) ' Registra no relat�rio que o chamado possui OC sem cr�ditos associados
        Call AlimentarAbaHistorica("SEM CREDITOS EM ABERTO ENCONTRADOS") ' Registra na aba hist�rica que os cr�ditos foram liquidados anteriormente
    ' Sen�o, se a quantidade de NFD/OC for "Acima de 01" (e n�o houver cr�ditos associados)
    ElseIf qtde_NFD_OC_chamado = "Acima de 01" Then
        
        condicao_OC_incorreta = True ' Define a condi��o de OC incorreta como Verdadeira
        VerificarCondicaoClienteFBL5N = False ' Define a fun��o como Falsa
        Call AlimentarDicionario_Relatorio_Processamento("Chamados com OCs sem cr�ditos associados: ", chamado) ' Registra no relat�rio que o chamado possui OCs informadas incorretamente
        linha_fim = aba_1_arquivo_anexo_chamado_atual.Cells(1048576, rngEncontrado.Column).End(xlUp).Row
        For linha = rngEncontrado.Row + 1 To linha_fim
            If aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3) = "" Then
                numero_OC = aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column).Value
                numero_OC = TratativasOC(numero_OC)
                OCs_incorretas = numero_OC & "/" & numero_OC
                Call AlimentarAbaHistorica("SEM CREDITOS EM ABERTO ENCONTRADOS") ' Registra na aba hist�rica que os cr�ditos foram liquidados anteriormente
                aba_1_arquivo_anexo_chamado_atual.Cells(linha, rngEncontrado.Column + 3) = "OC sem cr�ditos de devolu��o associados para abatimento/reembolso"
            End If
        Next linha
    End If

    
    
    
End Function

Sub PreencherAbaRelatorioProcessamento()
    Dim chave_relatorio As Variant
    
    Set aba_relatorio_processamento = ThisWorkbook.Sheets("Relat�rio de Processamento")
    
    aba_relatorio_processamento.Columns("A:B").ClearContents

    linha = 1
    For Each chave_relatorio In Dicionario_Relatorio_Processamento.keys
        aba_relatorio_processamento.Range("A" & linha).Value = chave_relatorio
        aba_relatorio_processamento.Range("B" & linha).Value = Dicionario_Relatorio_Processamento(chave_relatorio)
        linha = linha + 1
    Next chave_relatorio
    
    aba_relatorio_processamento.Columns("A:B").AutoFit

End Sub

Sub SalvarFecharArquivoCliente()
    pasta_diaria = pasta_arquivos_clientes & "\" & VBA.Format(VBA.Date, "dd.mm.yyyy")
    ' Define o caminho completo do arquivo a ser salvo
    caminho_arquivo = pasta_diaria & "\Resposta Chamado " & chamado & ".xlsx"
    ' Cria a pasta di�ria se ela n�o existir
    If Dir(pasta_diaria, vbDirectory) = "" Then
        MkDir (pasta_diaria)
    End If
    ' Exclui o arquivo se ele j� existir
    If Dir(caminho_arquivo, vbDirectory) <> "" Then
        Kill caminho_arquivo
    End If
    If arquivo_cliente Is Nothing Then
        Set arquivo_cliente = arquivo_anexo_chamado_atual
    End If
    ' Salva o arquivo
    arquivo_cliente.SaveAs caminho_arquivo
    ' Fecha o arquivo
    arquivo_cliente.Close
    ' Libera a vari�vel do objeto
    Set arquivo_cliente = Nothing
End Sub
    
