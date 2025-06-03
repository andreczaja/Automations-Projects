Attribute VB_Name = "PASSO_1_TRATATIVAS_ABAT_REEMB"
Public array_payers_com_dados_bancarios(), array_linhas_detalhe_abatimento() As Variant
Public condicao_payer, doc_compensacao_abatimento As String
Public condicao, soma_debito_AR, soma_cred_dev, valor As Single
Public reembolsos_com_dados_bancarios_processados, reembolsos_sem_dados_bancarios_processados, abatimentos_processados As Integer
Public linha_fim_payers_com_dados_bancarios As Long

' sub que irá categorizar todos os payers presentes na base e realizar as devidas tratativas
' ou seja, verifica se é condicao de abatimento, reembolso com dados bancarios ou reembolso sem dados bancarios
Public Sub ProcessarAbatimentoOuReembolso()
    
    
    soma_cred_dev = 0
    soma_debito_AR = 0
    ' iterando sobre o array de linhas abertas para obter o valor de crédito devolução total que o cliente possui
    For i = LBound(array_geral_linhas_abertas_FBL5N) To UBound(array_geral_linhas_abertas_FBL5N)
        If IsNumeric(array_geral_linhas_abertas_FBL5N(i)(7)) Then
            soma_cred_dev = soma_cred_dev + array_geral_linhas_abertas_FBL5N(i)(7)
        End If
    Next i
    soma_cred_dev = Round(soma_cred_dev, 2)
    
    session_2.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
    session_2.findById("wnd[0]").sendVKey 0
    session_2.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    session_2.findById("wnd[1]/usr/txtV-LOW").text = "id328"
    session_2.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session_2.findById("wnd[1]/tbar[0]/btn[8]").press
    session_2.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = payer_associado_OC
    session_2.findById("wnd[0]/tbar[1]/btn[16]").press
    session_2.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN006-LOW").text = "RV"
    ' filtrando apenas linhas com chave de ref3 preenchida
    session_2.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode "         88"
    session_2.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "         81"
    session_2.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode "         88"
    session_2.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN012_%_APP_%-VALU_PUSH").press
    i2 = 0
    For i = 0 To 9
        If i2 > 9 Then
            Exit For
        End If
        On Error Resume Next
        session_2.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1," & i & "]").SetFocus
        If Err.number <> 0 Then
            session_2.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = i
            i = 1
        End If
        session_2.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1," & i & "]").text = "*" & i2 & "*"
        i2 = i2 + 1
        On Error GoTo 0
    Next i
    session_2.findById("wnd[1]/tbar[0]/btn[8]").press
    session_2.findById("wnd[0]/usr/ctxtPA_STIDA").text = Format(Date + 5, tipo_data_sap)
    session_2.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = Format(Date + 10, tipo_data_sap)
    session_2.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = Format(Date + 500, tipo_data_sap)
    session_2.findById("wnd[0]/tbar[1]/btn[8]").press
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "São exibidas (.*?) partidas"
    regex.IgnoreCase = True
    regex.Global = True
    
    texto_sbar = session_2.findById("wnd[0]/sbar").text
    
    Set ocorrencia = regex.Execute(texto_sbar)
    ' Se houver um resultado, armazena na variável
    If ocorrencia.Count > 0 Then
        texto_sbar = ocorrencia(0).SubMatches(0)
        quantidade_linhas = CInt(texto_sbar)
    Else
        ' se não for encontrada nenhuma linha no AR do cliente, processará o chamado com condição de reembolso diretamente
        condicao = soma_cred_dev
        GoTo Processar_Reembolso
    End If
    
    linhas_visiveis = VerificarQuantidadeLinhasVisiveis(session_2, 4, 500, "wnd[0]/usr/lbl[")
    
    qtde_page_downs = 0
    If quantidade_linhas > linhas_visiveis - 3 Then
        qtde_page_downs = Application.WorksheetFunction.Floor(quantidade_linhas / linhas_visiveis, 1)
    End If
    
    
    ' calculando o total de linhas no AR do cliente
    ' faz-se também a verificação se o campo chave de ref 3 está preenchido para que seja contabilizado no calculo
    i = 0
    linhas_visiveis = 0
    soma_debito_AR = 0
    Do Until i > qtde_page_downs
        linhas_visiveis = VerificarQuantidadeLinhasVisiveis(session_2, 4, 100, "wnd[0]/usr/lbl[")
        For i2 = 4 To linhas_visiveis
            Dim valor_string As String
            
            valor_string = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_montante & "," & i2 & "]").text)
            valor_string = Replace(valor_string, "-", "")
            valor_string = Replace(valor_string, ".", "")
            valor_string = Replace(valor_string, ",", ".")
            valor = CSng(valor_string)
            soma_debito_AR = soma_debito_AR + valor
            If soma_debito_AR + soma_cred_dev > 0 Then
                GoTo Processar_Abatimento
            End If
        Next i2
        session_2.findById("wnd[0]").sendVKey 82
        i = i + 1
    Loop
    
    session_2.findById("wnd[0]").sendVKey 80
    
    condicao = soma_debito_AR + soma_cred_dev
    If condicao < 0 Then
Processar_Reembolso:
        condicao_payer = "reembolsados"
        ' CLIENTE NÃO ESTÁ NA BASE DE CLIENTES CADASTRADOS
        If VerificarDadosBancarios Then
            condicao_OCs_reembolso = True
            Call Processamento_Reembolso_SAP
        Else
            Call AlterarAtribuicao(session, "PDTE DADOS BANC")
            Call AlterarAtribuicao(session_2, "PDTE DADOS BANC")
            condicao_cliente_sem_dados_bancarios = True
        End If
    Else
Processar_Abatimento:
        condicao_payer = "abatidos"
        Call Processamento_Abatimentos_SAP
    End If

End Sub


Private Function VerificarDadosBancarios() As Boolean

    
    VerificarDadosBancarios = False
    ' verificao inicial se o cliente está dentro do array se sim, irá sair da função
    If UBound(VBA.Filter(array_payers_com_dados_bancarios, payer_associado_OC)) >= 0 Then
        VerificarDadosBancarios = True
        Exit Function
    End If
    
    Dim texto_msgbox_sap, chave_banco, conta_bancaria, titular_conta As String

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^_*$"
    regex.IgnoreCase = True
    regex.Global = False
    

    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N FD03"
    session_3.findById("wnd[0]").sendVKey 0
    
    
    
    session_3.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = payer_associado_OC
    session_3.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = "BR10"
    session_3.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error Resume Next
    session_3.findById("wnd[2]/tbar[0]/btn[0]").press
    session_3.findById("wnd[0]/tbar[1]/btn[25]").press
    On Error GoTo 0
    session_3.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").Select
    
    chave_banco = session_3.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7131/tblSAPMF02DTCTRL_ZAHLUNGSVERKEHR/ctxtKNBK-BANKL[1,0]").text
    conta_bancaria = session_3.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7131/tblSAPMF02DTCTRL_ZAHLUNGSVERKEHR/txtKNBK-BANKN[2,0]").text
    titular_conta = session_3.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7131/tblSAPMF02DTCTRL_ZAHLUNGSVERKEHR/txtKNBK-KOINH[3,0]").text
    
    ' aqui verifica os elementos que carregam as infos dos dados bancarios, se forem diferentes de vazio, ou no caso,
    ' diferente de uma cadeira variavel de caracteres "_" considerará que o cliente agora tem dados bancarios e portanto será adicionado
    ' a lista na aba de clientes com dados bancarios e também ao array de clientes com dados bancarios
    ' caso contrário, as informações como código, valor pendente de reembolso e analista responsavel serão armazenada na aba dados bancarios
    ' para posterior envio e notificação ao analista e ao cliente
    
    If Not regex.Test(chave_banco) And Not regex.Test(conta_bancaria) And Not regex.Test(titular_conta) Then
        array_payers_com_dados_bancarios = Add_ao_Array(array_payers_com_dados_bancarios, payer_associado_OC)
        linha_fim_payers_com_dados_bancarios = aba_dados_bancarios.Range("A1048576").End(xlUp).Offset(1, 0).Row
        aba_dados_bancarios.Range("A" & linha_fim_payers_com_dados_bancarios).Value = payer_associado_OC
        VerificarDadosBancarios = True
    Else
        Call AlimentarDicionario_Relatorio_Processamento("Chamados associados a clientes em condição de reembolsos sem dados bancários cadastrados: ", chamado)
    End If

End Function

