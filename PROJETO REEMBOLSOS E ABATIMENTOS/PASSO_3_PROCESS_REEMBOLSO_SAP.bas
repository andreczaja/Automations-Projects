Attribute VB_Name = "PASSO_3_PROCESS_REEMBOLSO_SAP"
Public linha_fim_aba_reembolsos_pendentes As Integer
Public doc_f65, status_SBWP As String
Public array_docs_F65() As Variant
Public payer_repetido, documento_f65_repetido As Boolean
Public elemento_tabela_SBWP As Object
Public data_criacao As Date
Sub Processamento_Reembolso_SAP()
    
    If IsEmpty(array_payers_reembolsos_com_dados_bancarios) Then
        Exit Sub
    End If

    Call declaracao_vars
    
    ' limpando linhas vazias da base reembolsos ag aprovacao
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
    
    For linha = linha_fim_aba_reembolsos_pendentes To 2 Step -1
        If linha <> 2 And aba_reembolsos_pendentes.Range("A" & linha).Value = "" Then
            aba_reembolsos_pendentes.Rows(linha).EntireRow.Delete
        End If
    Next linha
    
    ' transferindo linhas de payers em condição de reembolso e que possuem dados bancarios registrados da aba FBL5N Cred Dev para Reembolsos Ag Aprovação

    linha_fim = aba_fbl5n_credito_devolucao.Range("A1048576").End(xlUp).Row
    
    data_agrupado_pagamento = aba_reembolsos_aprovados.Range("BC1").Value
    If data_agrupado_pagamento = ".." Or data_agrupado_pagamento = "" Then
        data_agrupado_pagamento = InputBox("A data do agrupado de pagamento não foi encontrada. Por favor digite-a abaixo no formato 'DD/MM/AAAA'")
        aba_reembolsos_aprovados.Range("BC1").Value = data_agrupado_pagamento
    End If
    For i = LBound(array_payers_reembolsos_com_dados_bancarios) To UBound(array_payers_reembolsos_com_dados_bancarios)
        
        payer_atual = array_payers_reembolsos_com_dados_bancarios(i)
        soma_cred_dev = Application.WorksheetFunction.SumIf(aba_fbl5n_credito_devolucao.Columns("C:C"), payer_atual, aba_fbl5n_credito_devolucao.Columns("P:P"))
        Call LimparFiltros(tabela_aba_fbl5n_credito_devolucao)
        Call LimparFiltros(tabela_aba_reembolsos_pendentes)
        linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
                       
        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N F-65"
        session_3.findById("wnd[0]").sendVKey 0
        session_3.findById("wnd[0]/usr/ctxtBKPF-BLDAT").text = Format(Date, tipo_data_sap)
        session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = Format(Date, tipo_data_sap)
        session_3.findById("wnd[0]/usr/ctxtBKPF-BLART").text = "ZD"
        session_3.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BRL"
        session_3.findById("wnd[0]/usr/txtBKPF-XBLNR").text = "REEMB AUTOMACAO"
        session_3.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "BR10"
        
        session_3.findById("wnd[0]/usr/txtBKPF-MONAT").text = Month(Date)
        session_3.findById("wnd[0]/usr/txtBKPF-BKTXT").text = Replace(UCase(VBA.Environ("USERPROFILE")), "C:\USERS\", "")
        session_3.findById("wnd[0]/usr/ctxtRF05V-NEWBS").text = "02"
        session_3.findById("wnd[0]/usr/ctxtRF05V-NEWKO").text = payer_atual
        session_3.findById("wnd[0]").sendVKey 0
        
        
        If VBA.Right(session_3.findById("wnd[0]/sbar").text, 29) <> "bloqueada para contabilização" Then
            
            tabela_aba_fbl5n_credito_devolucao.Range.AutoFilter Field:=3, Criteria1:=payer_atual
            aba_fbl5n_credito_devolucao.Range("A2:AB" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            
            If aba_reembolsos_pendentes.Range("A" & linha_fim_aba_reembolsos_pendentes).Value = "" Then
                aba_reembolsos_pendentes.Range("A" & linha_fim_aba_reembolsos_pendentes).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            ElseIf aba_reembolsos_pendentes.Range("A" & linha_fim_aba_reembolsos_pendentes).Value <> "" Then
                aba_reembolsos_pendentes.Range("A" & linha_fim_aba_reembolsos_pendentes).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            End If
        
            session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Abs(soma_cred_dev)
            session_3.findById("wnd[0]/usr/txtBSEG-ZUONR").text = "REEMB AUT " & VBA.Format(data_agrupado_pagamento, "dd.mm.yy")
            session_3.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "Processo automático de reembolso de devolução"
            session_3.findById("wnd[0]/usr/ctxtRF05V-NEWBS").text = "1D"
            session_3.findById("wnd[0]/usr/ctxtRF05V-NEWKO").text = payer_atual
            session_3.findById("wnd[0]").sendVKey 0
            If Left(session_3.findById("wnd[0]/sbar").text, 19) = "Entrada só na forma" Then
                If InStr(1, session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text, ".") Then
                    session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Replace(session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text, ".", ",")
                    session_3.findById("wnd[0]").sendVKey 0
                End If
            End If
            
            session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Abs(soma_cred_dev)
            session_3.findById("wnd[0]/usr/txtBSEG-ZUONR").text = "AUTOMACAO DEV"
            session_3.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "Processo automático de reembolso de devolução"
            session_3.findById("wnd[0]/usr/ctxtBSEG-ZLSCH").text = "T"
            session_3.findById("wnd[0]/tbar[1]/btn[7]").press
            If Left(session_3.findById("wnd[0]/sbar").text, 19) = "Entrada só na forma" Then
                If InStr(1, session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text, ".") Then
                    session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Replace(session_3.findById("wnd[0]/usr/txtBSEG-WRBTR").text, ".", ",")
                    session_3.findById("wnd[0]").sendVKey 0
                End If
            End If
            session_3.findById("wnd[0]/usr/txtBSEG-XREF2").text = "AUTOMACAO"
            session_3.findById("wnd[0]/tbar[1]/btn[7]").press
            
            session_3.findById("wnd[0]/usr/txtBSEG-WSKTO").SetFocus
            session_3.findById("wnd[0]/mbar/menu[0]/menu[4]").Select
            doc_f65 = Mid(session_3.findById("wnd[0]/sbar").text, 11, 10)
            
            Call LimparFiltros(tabela_aba_reembolsos_pendentes)
            
            
            linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
            For i2 = 2 To linha_fim_aba_reembolsos_pendentes
                If CStr(aba_reembolsos_pendentes.Range("C" & i2).Value) = payer_atual And aba_reembolsos_pendentes.Range("AC" & i2).Value = "" And aba_reembolsos_pendentes.Range("AD" & i2).Value = "" And aba_reembolsos_pendentes.Range("AE" & i2).Value = "" Then
                    aba_reembolsos_pendentes.Range("AC" & i2).Value = doc_f65
                    aba_reembolsos_pendentes.Range("AD" & i2).Value = Date
                    aba_reembolsos_pendentes.Range("AE" & i2).Value = "Não Solicitada Aprovação"
                End If
            Next i2
            
            Call LimparFiltros(tabela_aba_reembolsos_pendentes)
        Else
            session.findById("wnd[0]/usr/lbl[" & x_payer_session & ",2]").SetFocus
            session.findById("wnd[0]").sendVKey 2
            session.findById("wnd[0]/usr/lbl[" & x_num_doc_session & ",2]").SetFocus
            session.findById("wnd[0]").sendVKey 2
            session.findById("wnd[0]/tbar[1]/btn[38]").press
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = payer_atual
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press
            session.findById("wnd[2]/tbar[0]/btn[16]").press
            session.findById("wnd[2]/tbar[0]/btn[8]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[0]").sendVKey 5
            session.findById("wnd[0]/tbar[1]/btn[45]").press
            session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "CTA BLOQUEADA"
            session.findById("wnd[0]").sendVKey 0
            On Error Resume Next
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            On Error GoTo 0
        End If
    Next i
    
    Application.Wait (Now + TimeValue("00:02:00"))
    
    Call VerificarLinhasSBWP(session_3)
    ' verificando se existem linhas que foram criadas hoje e feito o processo de envio para aprovação na SBWP
    Call alterar_campos_linhas_R1

End Sub

' Sub que altera as linhas as quais já foi enviado para a aprovação via SBWP o campo atribuicao para "REEMB AUT " & VBA.Format(data_agrupado_pagamento, "dd.mm.yy")
Sub alterar_campos_linhas_R1()

    ''''''''''''''''''  ETAPA ALTERACAO DE LINHAS DE REEMBOLSOS ENVIADOS PARA APROVAÇÃO'''''''''''''''''''''
    session.findById("wnd[0]/usr/lbl[" & x_payer_session & ",2]").SetFocus
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/tbar[1]/btn[38]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
    
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
    
    Call LimparFiltros(tabela_aba_reembolsos_pendentes)
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=30, Criteria1:=VBA.Format(Date, "dd/mm/yyyy")
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=31, Criteria1:="Aguardando Aprovação"
    aba_reembolsos_pendentes.Range("C2:C" & linha_fim_aba_reembolsos_pendentes).SpecialCells(xlCellTypeVisible).Copy
    
    session.findById("wnd[2]/tbar[0]/btn[16]").press
    session.findById("wnd[2]/tbar[0]/btn[24]").press
    session.findById("wnd[2]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    ' COLOCAR AQUI NA ASIGNACION A DATA DO AGRUPADO DE PAGAMENTO
    session.findById("wnd[0]").sendVKey 5
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    If session.findById("wnd[0]/sbar").text = "Marcar pelo menos uma partida" Then
        Exit Sub
    End If
    session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "REEMB AUT " & VBA.Format(data_agrupado_pagamento, "dd.mm.yy")
    session.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    
    Call LimparFiltros(tabela_aba_reembolsos_pendentes)
    
    ''''''''''''''''''  ETAPA ALTERACAO DE LINHAS DE REEMBOLSOS AINDA NÃO ENVIADOS PARA APROVAÇÃO'''''''''''''''''''''
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=30, Criteria1:=Date
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=31, Criteria1:="Não Solicitada Aprovação"
    
    On Error Resume Next
    aba_reembolsos_pendentes.Range("C2:C" & linha_fim_aba_reembolsos_pendentes).SpecialCells(xlCellTypeVisible).Copy
    If Err.number <> 1004 Then
        On Error GoTo 0
        session.findById("wnd[2]/tbar[0]/btn[16]").press
        session.findById("wnd[2]/tbar[0]/btn[24]").press
        session.findById("wnd[2]/tbar[0]/btn[8]").press
    
    
        ' COLOCAR AQUI NA ASIGNACION A DATA DO AGRUPADO DE PAGAMENTO
        session.findById("wnd[0]").sendVKey 5
        session.findById("wnd[0]/tbar[1]/btn[45]").press
        If session.findById("wnd[0]/sbar").text = "Marcar pelo menos uma partida" Then
            Exit Sub
        End If
        session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "AG PROCESS SBWP"
        session.findById("wnd[0]").sendVKey 0
        On Error Resume Next
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        On Error GoTo 0
    Else
        Call LimparFiltros(tabela_aba_reembolsos_pendentes)
    End If
    ''''''''''''''''''  ETAPA ALTERACAO DE LINHAS DE CLIENTES SEM DADOS BANCARIOS'''''''''''''''''''''
    linha_fim = aba_dados_bancarios.Range("B1048576").End(xlUp).Row
    
    session.findById("wnd[0]/usr/lbl[" & x_payer_session & ",2]").SetFocus
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/tbar[1]/btn[38]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press

    aba_dados_bancarios.Range("B2:B" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
    
    session.findById("wnd[2]/tbar[0]/btn[16]").press
    session.findById("wnd[2]/tbar[0]/btn[24]").press
    session.findById("wnd[2]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    ' COLOCAR AQUI NA ASIGNACION A DATA DO AGRUPADO DE PAGAMENTO
    session.findById("wnd[0]").sendVKey 5
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    If session.findById("wnd[0]/sbar").text = "Marcar pelo menos uma partida" Then
        Exit Sub
    End If
    session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "PDTE DADOS BANC"
    session.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    
    Call LimparFiltros(tabela_aba_reembolsos_pendentes)

End Sub


