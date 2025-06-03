Attribute VB_Name = "PASSO_2_TRATATIVAS_ABAT_REEMB"
Public array_payers_reembolsos_com_dados_bancarios(), array_payers_reembolsos_sem_dados_bancarios(), array_payers_abatimento(), array_payers_com_dados_bancarios() As Variant
Public payer_atual, condicao_payer, analista_responsavel As String
Public condicao, soma_debito_AR, soma_cred_dev, soma_debito_AR_chave_ref_3_vazia As Single
Public qtde_linhas_totais_reembolsos, qtde_linhas_nao_passiveis_reembolso, reembolsos_com_dados_bancarios_processados, _
    reembolsos_sem_dados_bancarios_processados, abatimentos_processados As Integer
' sub que irá categorizar todos os payers presentes na base e realizar as devidas tratativas
' ou seja, verifica se é condicao de abatimento, reembolso com dados bancarios ou reembolso sem dados bancarios
Sub AbatimentoOuReembolso()

    Call declaracao_vars
    
    array_payers_com_dados_bancarios = Array()
    array_payers_reembolsos_com_dados_bancarios = Array()
    array_payers_reembolsos_sem_dados_bancarios = Array()
    array_payers_abatimento = Array()
    Call PreencherArrayPayersCOMDadosBancarios
    
    Call LimparFiltros(tabela_aba_fbl5n_credito_devolucao)
    reembolsos_sem_dados_bancarios_processados = 0
    reembolsos_com_dados_bancarios_processados = 0
    abatimentos_processados = 0
    
    aba_dados_bancarios.Range("B2:D1048576").ClearContents
    linha_fim = aba_fbl5n_credito_devolucao.Range("A1048576").End(xlUp).Row
    For linha = 2 To linha_fim
        payer_atual = aba_fbl5n_credito_devolucao.Range("C" & linha).Value
        
        If Not PayerDuplicado Then

            soma_cred_dev = Application.WorksheetFunction.SumIf(aba_fbl5n_credito_devolucao.Columns("C:C"), payer_atual, aba_fbl5n_credito_devolucao.Columns("P:P"))
            soma_debito_AR_chave_ref_3_vazia = Application.WorksheetFunction.SumIfs(aba_fbl5n_AR.Columns("P:P"), aba_fbl5n_AR.Columns("C:C"), payer_atual, aba_fbl5n_AR.Columns("AB:AB"), "")
            soma_debito_AR = Application.WorksheetFunction.SumIf(aba_fbl5n_AR.Columns("C:C"), payer_atual, aba_fbl5n_AR.Columns("P:P")) - soma_debito_AR_chave_ref_3_vazia
            condicao = soma_debito_AR + soma_cred_dev
            qtde_linhas_totais_reembolsos = Application.WorksheetFunction.CountIf(aba_fbl5n_AR.Columns("C:C"), payer_atual)
            qtde_linhas_nao_passiveis_reembolso = Application.WorksheetFunction.CountIfs(aba_fbl5n_AR.Columns("C:C"), payer_atual, aba_fbl5n_AR.Columns("AB:AB"), "")
            Debug.Print condicao
            ' CASO DE REEMBOLSO
            If condicao < 0 Then
                ' payer não tem dados bancarios
                If Not VerificarDadosBancarios Then
                    Call Add_ao_Array(array_payers_reembolsos_sem_dados_bancarios, payer_atual)
                Else
                ' payer tem dados bancarios
                    Call Add_ao_Array(array_payers_reembolsos_com_dados_bancarios, payer_atual)
                End If
            ElseIf condicao > 0 And qtde_linhas_totais_reembolsos > qtde_linhas_nao_passiveis_reembolso Then
            ' CASO DE ABATIMENTO
                Call Add_ao_Array(array_payers_abatimento, payer_atual)
            End If
        End If
    Next linha
    
    Call VerificarEixoXColunasSession
    Call VerificarEixoXColunasSession_2
    ' chama as subs para processar os payers armazenados em cada array
    If Form_SAP.opt_ambos Then
        Call Processamento_Abatimentos_SAP
        Call Processamento_Reembolso_SAP
    ElseIf Form_SAP.opt_apenas_abatimentos Then
        Call Processamento_Abatimentos_SAP
    ElseIf Form_SAP.opt_apenas_reembolsos Then
        Call Processamento_Reembolso_SAP
    End If
    
    
    MsgBox "Foram processados: " & vbNewLine & abatimentos_processados & " abatimentos" & vbNewLine & _
        reembolsos_com_dados_bancarios_processados & " reembolsos de clientes com dados bancários" & vbNewLine & _
            "Além disso, foram verificados e sinalizados no SAP " & reembolsos_sem_dados_bancarios_processados & " clientes em condição de reembolsos que não possuem dados bancários."

End Sub
' preenche no array todos os clientes que tem dados bancarios
Private Function PreencherArrayPayersCOMDadosBancarios()
    
    linha_fim = aba_dados_bancarios.Range("A1048576").End(xlUp).Row
    
    If linha_fim = 2 And aba_dados_bancarios.Range("A2").Value = "" Then
        Exit Function
    End If
    
    For linha = 2 To linha_fim
        payer_atual = aba_dados_bancarios.Range("A" & linha).Value
        If Not UBound(VBA.Filter(array_payers_com_dados_bancarios, CLng(payer_atual))) >= 0 Then
            Call Add_ao_Array(array_payers_com_dados_bancarios, payer_atual)
        End If
    Next linha
    
End Function

Private Function VerificarDadosBancarios() As Boolean

    
    VerificarDadosBancarios = False
    ' verificao inicial se o cliente está dentro do array se sim, irá sair da função
    If UBound(VBA.Filter(array_payers_com_dados_bancarios, CLng(payer_atual))) >= 0 Then
        VerificarDadosBancarios = True
        Exit Function
    End If
    
    Dim regex As Object
    Dim texto_msgbox_sap, chave_banco, conta_bancaria, titular_conta As String

    Set regex = New RegExp
    regex.Pattern = "^_*$"
    regex.IgnoreCase = True
    regex.Global = False
    
    ' cria a sessao 3 que irá fazer as verificacoes dos payers que até a ultima rodada ainda não possuiam dados bancarios
    If session_3 Is Nothing Then
        Call InteracaoTelasSAP(session_3, 3, "FD03")
    Else
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N FD03"
        session_3.findById("wnd[0]").sendVKey 0
    End If
    
    
    
    session_3.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = payer_atual
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
        If aba_dados_bancarios.Range("A2").Value = "" Then
            aba_dados_bancarios.Range("A2").Value = payer_atual
        ElseIf aba_dados_bancarios.Range("A2").Value <> "" Then
            aba_dados_bancarios.Range("A1048576").End(xlUp).Offset(1, 0).Value = payer_atual
        End If
        Call Add_ao_Array(array_payers_com_dados_bancarios, payer_atual)
        VerificarDadosBancarios = True
    Else
        On Error Resume Next
        analista_responsavel = Application.VLookup(payer_atual, aba_plan_distribuicao.Columns("A:E"), 5, False)
        If Err.number <> 0 Then
            analista_responsavel = "Cliente sem Analista Mapeado"
        End If
        On Error GoTo 0
        valor = Abs(Application.WorksheetFunction.SumIf(aba_fbl5n_credito_devolucao.Columns("C:C"), payer_atual, aba_fbl5n_credito_devolucao.Columns("P:P")))
        If aba_dados_bancarios.Range("B2").Value = "" Then
            aba_dados_bancarios.Range("B2").Value = payer_atual
            aba_dados_bancarios.Range("C2").Value = valor
            aba_dados_bancarios.Range("D2").Value = analista_responsavel
            
        ElseIf aba_dados_bancarios.Range("B2").Value <> "" Then
            aba_dados_bancarios.Range("B1048576").End(xlUp).Offset(1, 0).Value = payer_atual
            aba_dados_bancarios.Range("C1048576").End(xlUp).Offset(1, 0).Value = valor
            aba_dados_bancarios.Range("D1048576").End(xlUp).Offset(1, 0).Value = analista_responsavel
            
        End If
        reembolsos_sem_dados_bancarios_processados = reembolsos_sem_dados_bancarios_processados + 1
    End If

End Function

