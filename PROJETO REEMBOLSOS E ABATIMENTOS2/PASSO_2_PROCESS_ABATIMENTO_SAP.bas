Attribute VB_Name = "PASSO_2_PROCESS_ABATIMENTO_SAP"
Public valor_residual_F32, valor_residual_AB, valor, ultimo_valor As Single
Public abater_integrais, conta_bloqueada, ultima_linha As Boolean
Public status_titulo As String
Public array_cabecalho_abatimento() As Variant
Sub Processamento_Abatimentos_SAP()

    ' Verifica se a conta não está bloqueada para a transação "F-32"
    If Not VerificarContaBloqueada("F-32") Then

        ' Simula a tecla F8 (executar) na segunda sessão do SAP
        session_2.findById("wnd[0]").sendVKey 80
        ' Inicializa a flag para controlar se todos os itens devem ser abatidos integralmente
        abater_integrais = False
        ' Inicializa a variável para somar os valores de débito em contas a receber (AR)
        soma_debito_AR = 0
        ' Inicializa um contador
        i = 0
        ' Inicializa um índice para o array geral (não parece ser usado diretamente aqui)
        linha_index_array_geral = 1
        ' Loop que continua até que o contador 'i' seja maior que a quantidade de page downs
        Do Until i > qtde_page_downs
            ' Loop através das linhas visíveis na tela (da linha 4 até a última linha visível)
            For i2 = 4 To linhas_visiveis
                ' Obtém o valor da coluna de montante na linha atual, remove separadores de milhar e substitui vírgula por ponto para conversão para Single
                Dim valor_string As String
                
                valor_string = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_montante & "," & i2 & "]").text)
                valor_string = Replace(valor_string, "-", "")
                valor_string = Replace(valor_string, ".", "")
                valor_string = Replace(valor_string, ",", ".")
                valor = CSng(valor_string)
                
                ' Se a soma do crédito de devolução, débito em AR e o valor atual for menor que zero
                If (soma_cred_dev + soma_debito_AR + valor) < 0 Then
                    ' Seleciona a caixa de seleção na coluna 1 da linha atual
                    session_2.findById("wnd[0]/usr/chk[1," & i2 & "]").Selected = True
                    ' Define a flag 'abater_integrais' como Verdadeira, indicando que pelo menos um item será abatido integralmente
                    abater_integrais = True
                ' Senão, se a soma for maior que zero
                ElseIf (soma_cred_dev + soma_debito_AR + valor) > 0 Then
                    ' Se a flag 'abater_integrais' for Verdadeira
                    If abater_integrais Then
                        ' Simula a tecla F5 (processar) na primeira sessão do SAP
                        session.findById("wnd[0]").sendVKey 5
                        ' Chama a sub-rotina para alterar a atribuição para "ABATIDO TOTAL" na primeira sessão
                        Call AlterarAtribuicao(session, "ABATIDO TOTAL")
                        ' Chama a sub-rotina para alterar a atribuição para "ABATIDO TOTAL" na segunda sessão
                        Call AlterarAtribuicao(session_2, "ABATIDO TOTAL")
                    ' Senão (se 'abater_integrais' for Falsa)
                    ElseIf Not abater_integrais Then
                        ' Simula a tecla F5 (processar) na primeira sessão do SAP
                        session.findById("wnd[0]").sendVKey 5
                        ' Chama a sub-rotina para alterar a atribuição para "ABATIDO PARCIAL" na primeira sessão
                        Call AlterarAtribuicao(session, "ABATIDO PARCIAL")
                    End If
                    ' Calcula o valor residual após o abatimento
                    valor_residual_AB = soma_cred_dev + soma_debito_AR + valor
                    ' Chama a sub-rotina para preencher um array com informações da linha atual para abatimento na primeira sessão
                    Call PreencherArrayLinhasCondicaoAtual(session, i4, i5, "ABATIMENTO")
                    ' Seleciona a caixa de seleção na coluna 1 da linha atual na segunda sessão
                    session_2.findById("wnd[0]/usr/chk[1," & i2 & "]").Selected = True
                    ' Chama a sub-rotina para alterar a atribuição para "ABATIDO PARCIAL" na segunda sessão
                    Call AlterarAtribuicao(session_2, "ABATIDO PARCIAL")
                    ' Simula a tecla Backspace (apagar seleção) na segunda sessão
                    session_2.findById("wnd[0]").sendVKey 8
                    ' Define o foco no label da coluna de atribuição na segunda linha da segunda sessão
                    session_2.findById("wnd[0]/usr/lbl[" & x_atribuicao & ",2]").SetFocus
                    ' Simula a tecla Shift+F2 (ir para o campo de atribuição) na segunda sessão
                    session_2.findById("wnd[0]").sendVKey 2
                    ' Clica no botão de ajuda de pesquisa para o campo de atribuição na segunda sessão
                    session_2.findById("wnd[0]/tbar[1]/btn[38]").press
                    ' Clica no botão de múltipla seleção para o campo de atribuição na janela de pesquisa
                    session_2.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
                    ' Preenche os campos de seleção com "ABATIDO TOTAL" e "ABATIDO PARCIAL" na janela de múltipla seleção
                    session_2.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "ABATIDO TOTAL"
                    session_2.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "ABATIDO PARCIAL"
                    ' Clica no botão de executar na janela de múltipla seleção
                    session_2.findById("wnd[2]/tbar[0]/btn[8]").press
                    ' Clica no botão de copiar na janela de pesquisa
                    session_2.findById("wnd[1]/tbar[0]/btn[0]").press
                    ' Define o foco novamente no label da coluna de atribuição na segunda linha da segunda sessão
                    session_2.findById("wnd[0]/usr/lbl[" & x_atribuicao & ",2]").SetFocus
                    ' Simula a tecla Shift+F2 novamente na segunda sessão
                    session_2.findById("wnd[0]").sendVKey 2
                    ' Clica no botão de gravar (disquete) na segunda sessão
                    session_2.findById("wnd[0]/tbar[1]/btn[40]").press
                    ' Chama a sub-rotina para preencher um array com informações da linha atual para abatimento na segunda sessão
                    Call PreencherArrayLinhasCondicaoAtual(session_2, i4, i5, "ABATIMENTO")
                    ' Sai do loop Do...Until
                    Exit Do
                End If
                ' Adiciona o valor atual à soma dos débitos em AR
                soma_debito_AR = soma_debito_AR + valor
            Next i2
            ' Incrementa o contador de page downs
            i = i + 1
            ' Simula a tecla Page Down na segunda sessão
            session_2.findById("wnd[0]").sendVKey 82
        Loop
        ' Simula a tecla F8 (executar) na primeira sessão
        session.findById("wnd[0]").sendVKey 80
        ' Simula a tecla F8 (executar) na segunda sessão
        session_2.findById("wnd[0]").sendVKey 80

        ' Se a flag 'abater_integrais' for Verdadeira
        If abater_integrais Then
            ' Chama a sub-rotina F32 (provavelmente para realizar o lançamento do abatimento)
            Call F32
            ' Se a conta não estiver bloqueada após a execução de F32
            If Not conta_bloqueada Then
                ' Chama a sub-rotina ZFI156 (provavelmente para gerar algum documento ou registro), passando Verdadeiro como parâmetro
                Call ZFI156(True)
            ' Senão (se a conta estiver bloqueada)
            Else
                ' Sai da sub-rotina
                Exit Sub
            End If
        ' Senão (se 'abater_integrais' for Falsa)
        Else
            ' Chama a sub-rotina ZFI156, passando Falso como parâmetro
            Call ZFI156(False)
        End If
        
        If erro_zfi156 Then
            Exit Sub
        End If
        ' Chama a sub-rotina para alimentar um dicionário com o número do documento de compensação de abatimento gerado
        Call AlimentarDicionario_Relatorio_Processamento("Documentos de compensação de abatimento gerados: ", doc_compensacao_abatimento)

        ' Entra na transação FBL5N na terceira sessão do SAP
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
        ' Seleciona a opção de menu para listar linhas em aberto
        session_3.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
        ' Preenche o campo de variante com "id328"
        session_3.findById("wnd[1]/usr/txtV-LOW").text = "id328"
        ' Limpa o campo de nome
        session_3.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        ' Clica no botão de executar
        session_3.findById("wnd[1]/tbar[0]/btn[8]").press
        ' Preenche o campo de cliente com o payer associado à OC
        session_3.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = payer_associado_OC
        ' Clica no botão para exibir mais opções de seleção
        session_3.findById("wnd[0]/tbar[1]/btn[16]").press
        ' Preenche o campo de número do documento com o documento de compensação de abatimento
        session_3.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN003-LOW").text = doc_compensacao_abatimento
        ' Preenche o campo de tipo de documento com "AB" (Abatimento)
        session_3.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN006-LOW").text = "AB"
        ' Preenche o campo de data de vencimento líquido com a data atual + 5 dias
        session_3.findById("wnd[0]/usr/ctxtPA_STIDA").text = Format(Date + 5, tipo_data_sap)
        ' Clica no botão de executar
        session_3.findById("wnd[0]/tbar[1]/btn[8]").press
        ' Chama a sub-rotina para preencher um array com informações da linha atual para abatimento na terceira sessão
        Call PreencherArrayLinhasCondicaoAtual(session_3, i4, i5, "ABATIMENTO")
    End If


End Sub

Sub F32()

    ' Entra na transação F-32 (Compensar Contas de Cliente) na terceira sessão
    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N F-32"
    ' Simula a tecla Enter
    session_3.findById("wnd[0]").sendVKey 0
    ' Seleciona o radio button para contas de cliente
    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[3,0]").Select
    ' Preenche o campo de cliente com o payer associado à OC
    session_3.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = payer_associado_OC
    ' Preenche o campo de data de lançamento com a data atual no formato SAP
    session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = Format(Date, tipo_data_sap)
    ' Preenche o campo de mês do documento com o mês atual
    session_3.findById("wnd[0]/usr/txtBKPF-MONAT").text = Month(Date)
    ' Preenche o campo de código da empresa com "BR10"
    session_3.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "BR10"
    ' Preenche o campo de moeda com "BRL"
    session_3.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BRL"
    ' Clica no botão para exibir as partidas em aberto
    session_3.findById("wnd[0]/tbar[1]/btn[16]").press

    ' Se a barra de status não estiver vazia (indicando alguma mensagem, possivelmente erro de conta bloqueada)
    If session_3.findById("wnd[0]/sbar").text <> "" Then
        ' Define a flag de conta bloqueada como Verdadeira
        conta_bloqueada = True
        ' Chama a sub-rotina para registrar no relatório que o payer tem conta bloqueada para processamento na F-32
        Call AlimentarDicionario_Relatorio_Processamento("Payers com contas bloqueada para processamento na F-32: ", payer_associado_OC)
        ' Sai da sub-rotina
        Exit Sub
    End If

    ' Preenche o campo de seleção com "ABATIDO TOTAL"
    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]").text = "ABATIDO TOTAL"
    ' Simula a tecla Enter
    session_3.findById("wnd[0]").sendVKey 0

    ' Clica no botão para exibir as partidas selecionadas
    session_3.findById("wnd[0]/tbar[1]/btn[16]").press


    ' Obtém o valor residual para compensação
    valor_residual_F32 = CSng(VBA.Trim(Replace(Replace(session_3.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-DIFFB").text, ".", ""), ",", ".")))

    ' Clica no botão para selecionar todas as partidas
    session_3.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
    ' Clica no botão para atribuir o valor residual
    session_3.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press
    ' Seleciona a aba "Valor a ser imputado"
    session_3.findById("wnd[0]/usr/tabsTS/tabpREST").Select


    ' Obtém a quantidade de partidas
    qtde_partidas = session_3.findById("wnd[0]/usr/tabsTS/tabpREST/ssubPAGE:SAPDF05X:6106/txtRF05A-ANZPO").text

    ' Loop através das partidas
    For i2 = 0 To CInt(qtde_partidas)
        ' Imprime o tipo de documento para depuração
        Debug.Print session_3.findById("wnd[0]/usr/tabsTS/tabpREST/ssubPAGE:SAPDF05X:6106/tblSAPDF05XTC_6106/txtRFOPS_DK-BLART[3," & i2 & "]").text
        ' Se o tipo de documento for "R1" (fatura)
        If session_3.findById("wnd[0]/usr/tabsTS/tabpREST/ssubPAGE:SAPDF05X:6106/tblSAPDF05XTC_6106/txtRFOPS_DK-BLART[3," & i2 & "]").text = "R1" Then
            ' Define o foco no campo de valor a ser imputado
            session_3.findById("wnd[0]/usr/tabsTS/tabpREST/ssubPAGE:SAPDF05X:6106/tblSAPDF05XTC_6106/txtDF05B-PSDIF[8," & i2 & "]").SetFocus
            ' Simula a tecla Shift+F2 (provavelmente para inserir o valor residual)
            session_3.findById("wnd[0]").sendVKey 2
            ' Sai do loop
            Exit For
        End If
    Next i2

    ' Seleciona a opção de menu "Documento -> Simular"
    session_3.findById("wnd[0]/mbar/menu[0]/menu[1]").Select
    ' Simula a tecla Shift+F11 (provavelmente para gravar o documento)
    session_3.findById("wnd[0]").sendVKey 21
    ' Define o foco no primeiro campo de texto de atribuição
    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0700/txtRF05A-AZEI1[0,0]").SetFocus
    ' Simula a tecla Shift+F2
    session_3.findById("wnd[0]").sendVKey 2
    ' Preenche o campo de texto de atribuição com "ABATIDO PARCIAL"
    session_3.findById("wnd[0]/usr/txtBSEG-ZUONR").text = "ABATIDO PARCIAL"
    ' Simula a tecla Shift+F2
    session_3.findById("wnd[0]").sendVKey 2
    ' Se a barra de status começar com "Base de desconto"
    If Left(session_3.findById("wnd[0]/sbar").text, 16) = "Base de desconto" Then
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
    End If
    ' Trata possível erro se a janela existir
    On Error Resume Next
    ' Fecha a janela (se aberta)
    session_3.findById("wnd[1]").Close
    ' Desativa o tratamento de erros
    On Error GoTo 0
    ' Clica no botão de gravar
    session_3.findById("wnd[0]/tbar[0]/btn[11]").press
    ' Se a barra de status começar com "Base de desconto" novamente
    If Left(session_3.findById("wnd[0]/sbar").text, 16) = "Base de desconto" Then
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
        ' Clica no botão de gravar novamente
        session_3.findById("wnd[0]/tbar[0]/btn[11]").press
    End If

End Sub

Sub ZFI156(ByVal abater_integrais As Boolean)

    contador_erro_zfi156 = 1

    ' Se o parâmetro 'abater_integrais' for Verdadeiro (indica que o título foi abatido integralmente)
    If abater_integrais Then

        ' Entra na transação ZFI156 (transação customizada) na terceira sessão do SAP
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N ZFI156"
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0

        ' *** ETAPA BAIXA DE TITULO QUE FOI ABATIDO INTEGRALMENTE ***

        ' Clica no botão "Baixa de Título Após Compensação"
        session_3.findById("wnd[0]/usr/btnBT_BX_TIT_APOS_COMPENSACAO").press
        ' Preenche o campo de código da empresa com "BR10"
        session_3.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "BR10"
        ' Preenche o campo de cliente com o payer associado à OC
        session_3.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").text = payer_associado_OC
        ' Preenche o campo de atribuição com "ABATIDO TOTAL"
        session_3.findById("wnd[0]/usr/txtS_ZUONR-LOW").text = "ABATIDO TOTAL"
        ' Clica no botão de múltipla seleção para o tipo de documento
        session_3.findById("wnd[0]/usr/btn%_S_BLART_%_APP_%-VALU_PUSH").press
        ' Preenche os campos de seleção com "R1" (Fatura) e "RV" (Nota de Crédito)
        session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "R1"
        session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "RV"
        ' Clica no botão de executar na janela de múltipla seleção
        session_3.findById("wnd[1]/tbar[0]/btn[8]").press
        ' Clica no botão de executar
        session_3.findById("wnd[0]/tbar[1]/btn[8]").press
        ' Seleciona todas as linhas na grid
        session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
        ' Clica no botão de "Baixar" (ou função similar para baixa de títulos)
        session_3.findById("wnd[0]/tbar[1]/btn[13]").press

    End If

    ' *** ETAPA ABATIMENTO DE TITULO QUE FOI ABATIDO PARCIALMENTE ***
processar_novamente:
    ' Entra na transação ZFI156 novamente
    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N ZFI156"
    ' Simula a tecla Enter
    session_3.findById("wnd[0]").sendVKey 0
    ' Clica no botão "Abatimento"
    session_3.findById("wnd[0]/usr/btnBT_ABATIMENTO").press
    ' Preenche o campo de código da empresa com "BR10"
    session_3.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "BR10"
    ' Preenche o campo de cliente com o payer associado à OC
    session_3.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").text = payer_associado_OC
    ' Preenche o campo de atribuição com "ABATIDO PARCIAL"
    session_3.findById("wnd[0]/usr/txtS_ZUONR-LOW").text = "ABATIDO PARCIAL"
    ' Clica no botão de múltipla seleção para o tipo de documento
    session_3.findById("wnd[0]/usr/btn%_S_BLART_%_APP_%-VALU_PUSH").press
    ' Preenche o primeiro tipo de documento com "RV" (Nota de Crédito)
    session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "RV"
    ' Se o abatimento foi integral, busca também por "AB" (Documento de Abatimento)
    If abater_integrais Then
        session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "AB"
    ' Senão (abatimento parcial), busca por "R1" (Fatura)
    Else
        session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "R1"
    End If

    ' Clica no botão de executar na janela de múltipla seleção
    session_3.findById("wnd[1]/tbar[0]/btn[8]").press
    ' Clica no botão de executar
    session_3.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Se o abatimento foi integral
    If abater_integrais Then
        ' Define o objeto da tabela grid
        Set elemento_tabela = session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        ' Define a primeira linha como a linha atual
        elemento_tabela.currentCellRow = 0
        ' Loop para percorrer até 1000 linhas da tabela
        For i5 = 1 To 1000
            On Error Resume Next
            elemento_tabela.setCurrentCell i5, "WRBTR"
            If Err.number <> 0 Then
                On Error GoTo 0
                If contador_erro_zfi156 < 20 Then
                    GoTo processar_novamente
                Else
                    erro_zfi156 = True
                    Call AlimentarDicionario_Relatorio_Processamento("Chamado com OCs em condição de abatimento que apresentaram erro no Cockpit/ZFI156: ", chamado)
                    Exit Sub
                End If
            End If
            
            ' Se o último caractere do valor for "-", indica um valor negativo (crédito)
            If VBA.Right(elemento_tabela.GetCellValue(i5, "WRBTR"), 1) = "-" Then
                ' Seleciona a linha atual e a linha 0
                elemento_tabela.selectedRows = "0," & i5
                ' Sai do loop
                Exit For
            End If
        Next i5
    ' Senão (abatimento parcial)
    Else
        ' Seleciona todas as linhas na grid
        session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    End If

    ' Clica no botão de "Processar" (ou função similar para realizar o abatimento)
    session_3.findById("wnd[0]/tbar[1]/btn[13]").press

    ' Se os últimos 4 caracteres da barra de status não forem "BR10" (indicando possível erro ou necessidade de nova tentativa)
    If Right(session_3.findById("wnd[0]/sbar").text, 4) <> "BR10" Then
        ' Volta para a linha 'processar_novamente' para tentar o processamento novamente
        GoTo processar_novamente
    ' Senão (processamento bem-sucedido)
    Else
        ' Extrai o número do documento de compensação de abatimento da barra de status
        doc_compensacao_abatimento = Mid(session_3.findById("wnd[0]/sbar").text, 11, 9)
    End If

End Sub


