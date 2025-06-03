Attribute VB_Name = "PASSO_2_PROCESS_ABATIMENTO_SAP"
Public valor_residual_F32, valor_residual_AB, valor, ultimo_valor As Single
Public abater_integrais, conta_bloqueada, ultima_linha As Boolean
Public status_titulo As String
Public array_cabecalho_abatimento() As Variant
Sub Processamento_Abatimentos_SAP()

    ' Verifica se a conta n�o est� bloqueada para a transa��o "F-32"
    If Not VerificarContaBloqueada("F-32") Then

        ' Simula a tecla F8 (executar) na segunda sess�o do SAP
        session_2.findById("wnd[0]").sendVKey 80
        ' Inicializa a flag para controlar se todos os itens devem ser abatidos integralmente
        abater_integrais = False
        ' Inicializa a vari�vel para somar os valores de d�bito em contas a receber (AR)
        soma_debito_AR = 0
        ' Inicializa um contador
        i = 0
        ' Inicializa um �ndice para o array geral (n�o parece ser usado diretamente aqui)
        linha_index_array_geral = 1
        ' Loop que continua at� que o contador 'i' seja maior que a quantidade de page downs
        Do Until i > qtde_page_downs
            ' Loop atrav�s das linhas vis�veis na tela (da linha 4 at� a �ltima linha vis�vel)
            For i2 = 4 To linhas_visiveis
                ' Obt�m o valor da coluna de montante na linha atual, remove separadores de milhar e substitui v�rgula por ponto para convers�o para Single
                Dim valor_string As String
                
                valor_string = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_montante & "," & i2 & "]").text)
                valor_string = Replace(valor_string, "-", "")
                valor_string = Replace(valor_string, ".", "")
                valor_string = Replace(valor_string, ",", ".")
                valor = CSng(valor_string)
                
                ' Se a soma do cr�dito de devolu��o, d�bito em AR e o valor atual for menor que zero
                If (soma_cred_dev + soma_debito_AR + valor) < 0 Then
                    ' Seleciona a caixa de sele��o na coluna 1 da linha atual
                    session_2.findById("wnd[0]/usr/chk[1," & i2 & "]").Selected = True
                    ' Define a flag 'abater_integrais' como Verdadeira, indicando que pelo menos um item ser� abatido integralmente
                    abater_integrais = True
                ' Sen�o, se a soma for maior que zero
                ElseIf (soma_cred_dev + soma_debito_AR + valor) > 0 Then
                    ' Se a flag 'abater_integrais' for Verdadeira
                    If abater_integrais Then
                        ' Simula a tecla F5 (processar) na primeira sess�o do SAP
                        session.findById("wnd[0]").sendVKey 5
                        ' Chama a sub-rotina para alterar a atribui��o para "ABATIDO TOTAL" na primeira sess�o
                        Call AlterarAtribuicao(session, "ABATIDO TOTAL")
                        ' Chama a sub-rotina para alterar a atribui��o para "ABATIDO TOTAL" na segunda sess�o
                        Call AlterarAtribuicao(session_2, "ABATIDO TOTAL")
                    ' Sen�o (se 'abater_integrais' for Falsa)
                    ElseIf Not abater_integrais Then
                        ' Simula a tecla F5 (processar) na primeira sess�o do SAP
                        session.findById("wnd[0]").sendVKey 5
                        ' Chama a sub-rotina para alterar a atribui��o para "ABATIDO PARCIAL" na primeira sess�o
                        Call AlterarAtribuicao(session, "ABATIDO PARCIAL")
                    End If
                    ' Calcula o valor residual ap�s o abatimento
                    valor_residual_AB = soma_cred_dev + soma_debito_AR + valor
                    ' Chama a sub-rotina para preencher um array com informa��es da linha atual para abatimento na primeira sess�o
                    Call PreencherArrayLinhasCondicaoAtual(session, i4, i5, "ABATIMENTO")
                    ' Seleciona a caixa de sele��o na coluna 1 da linha atual na segunda sess�o
                    session_2.findById("wnd[0]/usr/chk[1," & i2 & "]").Selected = True
                    ' Chama a sub-rotina para alterar a atribui��o para "ABATIDO PARCIAL" na segunda sess�o
                    Call AlterarAtribuicao(session_2, "ABATIDO PARCIAL")
                    ' Simula a tecla Backspace (apagar sele��o) na segunda sess�o
                    session_2.findById("wnd[0]").sendVKey 8
                    ' Define o foco no label da coluna de atribui��o na segunda linha da segunda sess�o
                    session_2.findById("wnd[0]/usr/lbl[" & x_atribuicao & ",2]").SetFocus
                    ' Simula a tecla Shift+F2 (ir para o campo de atribui��o) na segunda sess�o
                    session_2.findById("wnd[0]").sendVKey 2
                    ' Clica no bot�o de ajuda de pesquisa para o campo de atribui��o na segunda sess�o
                    session_2.findById("wnd[0]/tbar[1]/btn[38]").press
                    ' Clica no bot�o de m�ltipla sele��o para o campo de atribui��o na janela de pesquisa
                    session_2.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
                    ' Preenche os campos de sele��o com "ABATIDO TOTAL" e "ABATIDO PARCIAL" na janela de m�ltipla sele��o
                    session_2.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "ABATIDO TOTAL"
                    session_2.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "ABATIDO PARCIAL"
                    ' Clica no bot�o de executar na janela de m�ltipla sele��o
                    session_2.findById("wnd[2]/tbar[0]/btn[8]").press
                    ' Clica no bot�o de copiar na janela de pesquisa
                    session_2.findById("wnd[1]/tbar[0]/btn[0]").press
                    ' Define o foco novamente no label da coluna de atribui��o na segunda linha da segunda sess�o
                    session_2.findById("wnd[0]/usr/lbl[" & x_atribuicao & ",2]").SetFocus
                    ' Simula a tecla Shift+F2 novamente na segunda sess�o
                    session_2.findById("wnd[0]").sendVKey 2
                    ' Clica no bot�o de gravar (disquete) na segunda sess�o
                    session_2.findById("wnd[0]/tbar[1]/btn[40]").press
                    ' Chama a sub-rotina para preencher um array com informa��es da linha atual para abatimento na segunda sess�o
                    Call PreencherArrayLinhasCondicaoAtual(session_2, i4, i5, "ABATIMENTO")
                    ' Sai do loop Do...Until
                    Exit Do
                End If
                ' Adiciona o valor atual � soma dos d�bitos em AR
                soma_debito_AR = soma_debito_AR + valor
            Next i2
            ' Incrementa o contador de page downs
            i = i + 1
            ' Simula a tecla Page Down na segunda sess�o
            session_2.findById("wnd[0]").sendVKey 82
        Loop
        ' Simula a tecla F8 (executar) na primeira sess�o
        session.findById("wnd[0]").sendVKey 80
        ' Simula a tecla F8 (executar) na segunda sess�o
        session_2.findById("wnd[0]").sendVKey 80

        ' Se a flag 'abater_integrais' for Verdadeira
        If abater_integrais Then
            ' Chama a sub-rotina F32 (provavelmente para realizar o lan�amento do abatimento)
            Call F32
            ' Se a conta n�o estiver bloqueada ap�s a execu��o de F32
            If Not conta_bloqueada Then
                ' Chama a sub-rotina ZFI156 (provavelmente para gerar algum documento ou registro), passando Verdadeiro como par�metro
                Call ZFI156(True)
            ' Sen�o (se a conta estiver bloqueada)
            Else
                ' Sai da sub-rotina
                Exit Sub
            End If
        ' Sen�o (se 'abater_integrais' for Falsa)
        Else
            ' Chama a sub-rotina ZFI156, passando Falso como par�metro
            Call ZFI156(False)
        End If
        
        If erro_zfi156 Then
            Exit Sub
        End If
        ' Chama a sub-rotina para alimentar um dicion�rio com o n�mero do documento de compensa��o de abatimento gerado
        Call AlimentarDicionario_Relatorio_Processamento("Documentos de compensa��o de abatimento gerados: ", doc_compensacao_abatimento)

        ' Entra na transa��o FBL5N na terceira sess�o do SAP
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
        ' Seleciona a op��o de menu para listar linhas em aberto
        session_3.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
        ' Preenche o campo de variante com "id328"
        session_3.findById("wnd[1]/usr/txtV-LOW").text = "id328"
        ' Limpa o campo de nome
        session_3.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        ' Clica no bot�o de executar
        session_3.findById("wnd[1]/tbar[0]/btn[8]").press
        ' Preenche o campo de cliente com o payer associado � OC
        session_3.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = payer_associado_OC
        ' Clica no bot�o para exibir mais op��es de sele��o
        session_3.findById("wnd[0]/tbar[1]/btn[16]").press
        ' Preenche o campo de n�mero do documento com o documento de compensa��o de abatimento
        session_3.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN003-LOW").text = doc_compensacao_abatimento
        ' Preenche o campo de tipo de documento com "AB" (Abatimento)
        session_3.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN006-LOW").text = "AB"
        ' Preenche o campo de data de vencimento l�quido com a data atual + 5 dias
        session_3.findById("wnd[0]/usr/ctxtPA_STIDA").text = Format(Date + 5, tipo_data_sap)
        ' Clica no bot�o de executar
        session_3.findById("wnd[0]/tbar[1]/btn[8]").press
        ' Chama a sub-rotina para preencher um array com informa��es da linha atual para abatimento na terceira sess�o
        Call PreencherArrayLinhasCondicaoAtual(session_3, i4, i5, "ABATIMENTO")
    End If


End Sub

Sub F32()

    ' Entra na transa��o F-32 (Compensar Contas de Cliente) na terceira sess�o
    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N F-32"
    ' Simula a tecla Enter
    session_3.findById("wnd[0]").sendVKey 0
    ' Seleciona o radio button para contas de cliente
    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[3,0]").Select
    ' Preenche o campo de cliente com o payer associado � OC
    session_3.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = payer_associado_OC
    ' Preenche o campo de data de lan�amento com a data atual no formato SAP
    session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = Format(Date, tipo_data_sap)
    ' Preenche o campo de m�s do documento com o m�s atual
    session_3.findById("wnd[0]/usr/txtBKPF-MONAT").text = Month(Date)
    ' Preenche o campo de c�digo da empresa com "BR10"
    session_3.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "BR10"
    ' Preenche o campo de moeda com "BRL"
    session_3.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BRL"
    ' Clica no bot�o para exibir as partidas em aberto
    session_3.findById("wnd[0]/tbar[1]/btn[16]").press

    ' Se a barra de status n�o estiver vazia (indicando alguma mensagem, possivelmente erro de conta bloqueada)
    If session_3.findById("wnd[0]/sbar").text <> "" Then
        ' Define a flag de conta bloqueada como Verdadeira
        conta_bloqueada = True
        ' Chama a sub-rotina para registrar no relat�rio que o payer tem conta bloqueada para processamento na F-32
        Call AlimentarDicionario_Relatorio_Processamento("Payers com contas bloqueada para processamento na F-32: ", payer_associado_OC)
        ' Sai da sub-rotina
        Exit Sub
    End If

    ' Preenche o campo de sele��o com "ABATIDO TOTAL"
    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]").text = "ABATIDO TOTAL"
    ' Simula a tecla Enter
    session_3.findById("wnd[0]").sendVKey 0

    ' Clica no bot�o para exibir as partidas selecionadas
    session_3.findById("wnd[0]/tbar[1]/btn[16]").press


    ' Obt�m o valor residual para compensa��o
    valor_residual_F32 = CSng(VBA.Trim(Replace(Replace(session_3.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-DIFFB").text, ".", ""), ",", ".")))

    ' Clica no bot�o para selecionar todas as partidas
    session_3.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
    ' Clica no bot�o para atribuir o valor residual
    session_3.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press
    ' Seleciona a aba "Valor a ser imputado"
    session_3.findById("wnd[0]/usr/tabsTS/tabpREST").Select


    ' Obt�m a quantidade de partidas
    qtde_partidas = session_3.findById("wnd[0]/usr/tabsTS/tabpREST/ssubPAGE:SAPDF05X:6106/txtRF05A-ANZPO").text

    ' Loop atrav�s das partidas
    For i2 = 0 To CInt(qtde_partidas)
        ' Imprime o tipo de documento para depura��o
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

    ' Seleciona a op��o de menu "Documento -> Simular"
    session_3.findById("wnd[0]/mbar/menu[0]/menu[1]").Select
    ' Simula a tecla Shift+F11 (provavelmente para gravar o documento)
    session_3.findById("wnd[0]").sendVKey 21
    ' Define o foco no primeiro campo de texto de atribui��o
    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0700/txtRF05A-AZEI1[0,0]").SetFocus
    ' Simula a tecla Shift+F2
    session_3.findById("wnd[0]").sendVKey 2
    ' Preenche o campo de texto de atribui��o com "ABATIDO PARCIAL"
    session_3.findById("wnd[0]/usr/txtBSEG-ZUONR").text = "ABATIDO PARCIAL"
    ' Simula a tecla Shift+F2
    session_3.findById("wnd[0]").sendVKey 2
    ' Se a barra de status come�ar com "Base de desconto"
    If Left(session_3.findById("wnd[0]/sbar").text, 16) = "Base de desconto" Then
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
    End If
    ' Trata poss�vel erro se a janela existir
    On Error Resume Next
    ' Fecha a janela (se aberta)
    session_3.findById("wnd[1]").Close
    ' Desativa o tratamento de erros
    On Error GoTo 0
    ' Clica no bot�o de gravar
    session_3.findById("wnd[0]/tbar[0]/btn[11]").press
    ' Se a barra de status come�ar com "Base de desconto" novamente
    If Left(session_3.findById("wnd[0]/sbar").text, 16) = "Base de desconto" Then
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0
        ' Clica no bot�o de gravar novamente
        session_3.findById("wnd[0]/tbar[0]/btn[11]").press
    End If

End Sub

Sub ZFI156(ByVal abater_integrais As Boolean)

    contador_erro_zfi156 = 1

    ' Se o par�metro 'abater_integrais' for Verdadeiro (indica que o t�tulo foi abatido integralmente)
    If abater_integrais Then

        ' Entra na transa��o ZFI156 (transa��o customizada) na terceira sess�o do SAP
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N ZFI156"
        ' Simula a tecla Enter
        session_3.findById("wnd[0]").sendVKey 0

        ' *** ETAPA BAIXA DE TITULO QUE FOI ABATIDO INTEGRALMENTE ***

        ' Clica no bot�o "Baixa de T�tulo Ap�s Compensa��o"
        session_3.findById("wnd[0]/usr/btnBT_BX_TIT_APOS_COMPENSACAO").press
        ' Preenche o campo de c�digo da empresa com "BR10"
        session_3.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "BR10"
        ' Preenche o campo de cliente com o payer associado � OC
        session_3.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").text = payer_associado_OC
        ' Preenche o campo de atribui��o com "ABATIDO TOTAL"
        session_3.findById("wnd[0]/usr/txtS_ZUONR-LOW").text = "ABATIDO TOTAL"
        ' Clica no bot�o de m�ltipla sele��o para o tipo de documento
        session_3.findById("wnd[0]/usr/btn%_S_BLART_%_APP_%-VALU_PUSH").press
        ' Preenche os campos de sele��o com "R1" (Fatura) e "RV" (Nota de Cr�dito)
        session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "R1"
        session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "RV"
        ' Clica no bot�o de executar na janela de m�ltipla sele��o
        session_3.findById("wnd[1]/tbar[0]/btn[8]").press
        ' Clica no bot�o de executar
        session_3.findById("wnd[0]/tbar[1]/btn[8]").press
        ' Seleciona todas as linhas na grid
        session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
        ' Clica no bot�o de "Baixar" (ou fun��o similar para baixa de t�tulos)
        session_3.findById("wnd[0]/tbar[1]/btn[13]").press

    End If

    ' *** ETAPA ABATIMENTO DE TITULO QUE FOI ABATIDO PARCIALMENTE ***
processar_novamente:
    ' Entra na transa��o ZFI156 novamente
    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N ZFI156"
    ' Simula a tecla Enter
    session_3.findById("wnd[0]").sendVKey 0
    ' Clica no bot�o "Abatimento"
    session_3.findById("wnd[0]/usr/btnBT_ABATIMENTO").press
    ' Preenche o campo de c�digo da empresa com "BR10"
    session_3.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "BR10"
    ' Preenche o campo de cliente com o payer associado � OC
    session_3.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").text = payer_associado_OC
    ' Preenche o campo de atribui��o com "ABATIDO PARCIAL"
    session_3.findById("wnd[0]/usr/txtS_ZUONR-LOW").text = "ABATIDO PARCIAL"
    ' Clica no bot�o de m�ltipla sele��o para o tipo de documento
    session_3.findById("wnd[0]/usr/btn%_S_BLART_%_APP_%-VALU_PUSH").press
    ' Preenche o primeiro tipo de documento com "RV" (Nota de Cr�dito)
    session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "RV"
    ' Se o abatimento foi integral, busca tamb�m por "AB" (Documento de Abatimento)
    If abater_integrais Then
        session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "AB"
    ' Sen�o (abatimento parcial), busca por "R1" (Fatura)
    Else
        session_3.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "R1"
    End If

    ' Clica no bot�o de executar na janela de m�ltipla sele��o
    session_3.findById("wnd[1]/tbar[0]/btn[8]").press
    ' Clica no bot�o de executar
    session_3.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Se o abatimento foi integral
    If abater_integrais Then
        ' Define o objeto da tabela grid
        Set elemento_tabela = session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        ' Define a primeira linha como a linha atual
        elemento_tabela.currentCellRow = 0
        ' Loop para percorrer at� 1000 linhas da tabela
        For i5 = 1 To 1000
            On Error Resume Next
            elemento_tabela.setCurrentCell i5, "WRBTR"
            If Err.number <> 0 Then
                On Error GoTo 0
                If contador_erro_zfi156 < 20 Then
                    GoTo processar_novamente
                Else
                    erro_zfi156 = True
                    Call AlimentarDicionario_Relatorio_Processamento("Chamado com OCs em condi��o de abatimento que apresentaram erro no Cockpit/ZFI156: ", chamado)
                    Exit Sub
                End If
            End If
            
            ' Se o �ltimo caractere do valor for "-", indica um valor negativo (cr�dito)
            If VBA.Right(elemento_tabela.GetCellValue(i5, "WRBTR"), 1) = "-" Then
                ' Seleciona a linha atual e a linha 0
                elemento_tabela.selectedRows = "0," & i5
                ' Sai do loop
                Exit For
            End If
        Next i5
    ' Sen�o (abatimento parcial)
    Else
        ' Seleciona todas as linhas na grid
        session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    End If

    ' Clica no bot�o de "Processar" (ou fun��o similar para realizar o abatimento)
    session_3.findById("wnd[0]/tbar[1]/btn[13]").press

    ' Se os �ltimos 4 caracteres da barra de status n�o forem "BR10" (indicando poss�vel erro ou necessidade de nova tentativa)
    If Right(session_3.findById("wnd[0]/sbar").text, 4) <> "BR10" Then
        ' Volta para a linha 'processar_novamente' para tentar o processamento novamente
        GoTo processar_novamente
    ' Sen�o (processamento bem-sucedido)
    Else
        ' Extrai o n�mero do documento de compensa��o de abatimento da barra de status
        doc_compensacao_abatimento = Mid(session_3.findById("wnd[0]/sbar").text, 11, 9)
    End If

End Sub


