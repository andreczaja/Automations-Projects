Attribute VB_Name = "PASSO_3_PROCESS_ABATIMENTO_SAP"
Public num_doc, item, num_doc_item, referencia, vencimento, ultimo_num_doc, ultimo_item, ultima_referencia, ultimo_vencimento, ultimo_num_doc_item As String
Public Dict_AR, Dict_Cred_Dev, Dict_Analista As Object
Private array_num_doc_item_referencia_valor_vencimento
Public x_payer_session, x_num_doc_session, x_item_session, x_tipo_doc_session, x_payer_session_2, x_num_doc_session_2, x_item_session_2, x_tipo_doc_session_2, _
    linha_aba_titulos_a_abater, qtde_linhas, qtde_pags_r1_payer_atual As Integer
Public linha_fim_base_historica As Integer
Public valor_restante_apos_compensacao, valor_residual_boleto_abatido_parcial, valor, ultimo_valor As Single
Public condicao_sem_linhas_RV, conta_bloqueada As Boolean

Sub Processamento_Abatimentos_SAP()

    Call declaracao_vars
    Dim chave As Variant, valor As Variant
    
    If IsEmpty(array_payers_abatimento) Then
        Exit Sub
    End If
    
    aba_titulos_a_abater.Range("A2:H1048576").ClearContents
    tabela_titulos_a_abater.Resize Range("$A$1:$H$2")


    linha_aba_titulos_a_abater = 2
    For i = LBound(array_payers_abatimento) To UBound(array_payers_abatimento)
        payer_atual = array_payers_abatimento(i)
        conta_bloqueada = False
        
        Call VerificarContaBloqueada
        If Not conta_bloqueada Then
        
            Set Dict_Cred_Dev = CreateObject("Scripting.Dictionary")
            Set Dict_AR = CreateObject("Scripting.Dictionary")
            soma_debito_AR = 0
            soma_cred_dev = Application.WorksheetFunction.SumIf(aba_fbl5n_credito_devolucao.Columns("C:C"), payer_atual, aba_fbl5n_credito_devolucao.Columns("P:P"))
            
            ' armazenando no dicionario de credito de devolucao as infos pertinentes para serem descarregadas posteriormente na aba titulos a abater
            linha_fim = aba_fbl5n_credito_devolucao.Range("A1048576").End(xlUp).Row
            For linha = 2 To linha_fim
                If aba_fbl5n_credito_devolucao.Range("C" & linha).Value = payer_atual Then
                    referencia = aba_fbl5n_credito_devolucao.Range("F" & linha).Value
                    num_doc = aba_fbl5n_credito_devolucao.Range("G" & linha).Value
                    item = aba_fbl5n_credito_devolucao.Range("H" & linha).Value
                    vencimento = aba_fbl5n_credito_devolucao.Range("O" & linha).Value
                    valor = aba_fbl5n_credito_devolucao.Range("P" & linha).Value
                    num_doc_item = num_doc & "-" & item
                    array_num_doc_item_referencia_valor_vencimento = Array(num_doc, item, referencia, valor, vencimento)
                    Dict_Cred_Dev.Add num_doc_item, array_num_doc_item_referencia_valor_vencimento
                End If
            Next linha
            
            
             ' Criar o dicionário antes de iterar pelas linhas
            Call LimparFiltros(tabela_titulos_a_abater)
            
            linha_fim = aba_fbl5n_AR.Range("A1048576").End(xlUp).Row
            For linha = 2 To linha_fim
                referencia = aba_fbl5n_AR.Range("F" & linha).Value
                num_doc = aba_fbl5n_AR.Range("G" & linha).Value
                item = aba_fbl5n_AR.Range("H" & linha).Value
                num_doc_item = num_doc & "-" & item
                vencimento = aba_fbl5n_AR.Range("O" & linha).Value
                valor = aba_fbl5n_AR.Range("P" & linha).Value
                chave_de_ref_3 = aba_fbl5n_AR.Range("AB" & linha).Value
                
                If chave_de_ref_3 <> "" Then
                    
                    
                    If aba_fbl5n_AR.Range("C" & linha).Value = payer_atual And _
                       (soma_cred_dev + soma_debito_AR + aba_fbl5n_AR.Range("P" & linha).Value) < 0 Then
                       
                        soma_debito_AR = soma_debito_AR + aba_fbl5n_AR.Range("P" & linha).Value
                        array_num_doc_item_referencia_valor_vencimento = Array(num_doc, item, referencia, valor, vencimento)
                        Dict_AR.Add num_doc_item, array_num_doc_item_referencia_valor_vencimento
                        valor_residual_boleto_abatido_parcial = 0
                        
                        Call PreencherAbaTitulosaAbater(payer_atual, num_doc, item, referencia, valor, vencimento, "Boleto Abatido Integralmente", valor_residual_boleto_abatido_parcial)
                        
                    ElseIf aba_fbl5n_AR.Range("C" & linha).Value = payer_atual And _
                           (soma_cred_dev + soma_debito_AR + aba_fbl5n_AR.Range("P" & linha).Value) > 0 Then
                        
                        ultimo_num_doc = num_doc
                        ultimo_item = item
                        ultima_referencia = referencia
                        ultimo_valor = valor
                        ultimo_vencimento = vencimento
                        valor_residual_boleto_abatido_parcial = VBA.Round(aba_fbl5n_AR.Range("P" & linha).Value + soma_debito_AR + soma_cred_dev, 2)
                        
                        Call PreencherAbaTitulosaAbater(payer_atual, num_doc, item, referencia, valor, vencimento, "Boleto Abatido Parcialmente", valor_residual_boleto_abatido_parcial)
                        Exit For
                    End If
                End If
            Next linha
            
            Call LimparFiltros(tabela_titulos_a_abater)
            
            condicao_sem_linhas_RV = False
            
            If Application.WorksheetFunction.CountIfs(aba_titulos_a_abater.Columns("A:A"), payer_atual, aba_titulos_a_abater.Columns("G:G"), "Boleto Abatido Integralmente") > 0 Then
                Call FBL5N(True, condicao_sem_linhas_RV)
                If Not condicao_sem_linhas_RV Then
                    Call F32
                    If Not conta_bloqueada Then
                        Call ZFI156(True)
                    End If
                End If
            Else
                Call FBL5N(False, condicao_sem_linhas_RV)
                If Not condicao_sem_linhas_RV Then
                    Call ZFI156(False)
                End If
            End If
            abatimentos_processados = abatimentos_processados + 1
        End If
    Next i
    
End Sub

Sub PreencherAbaTitulosaAbater(ByVal payer_atual As String, ByVal num_doc As String, ByVal item As String, ByVal referencia As String, ByVal valor As Single, ByVal vencimento As String, tipo_abatimento As String, ByVal valor_residual As Single)

    aba_titulos_a_abater.Range("A" & linha_aba_titulos_a_abater).Value = payer_atual
    aba_titulos_a_abater.Range("B" & linha_aba_titulos_a_abater).Value = num_doc
    aba_titulos_a_abater.Range("C" & linha_aba_titulos_a_abater).Value = item
    aba_titulos_a_abater.Range("D" & linha_aba_titulos_a_abater).Value = referencia
    aba_titulos_a_abater.Range("E" & linha_aba_titulos_a_abater).Value = valor
    aba_titulos_a_abater.Range("F" & linha_aba_titulos_a_abater).Value = vencimento
    aba_titulos_a_abater.Range("G" & linha_aba_titulos_a_abater).Value = tipo_abatimento
    aba_titulos_a_abater.Range("H" & linha_aba_titulos_a_abater).Value = valor_residual
    
    linha_fim_base_historica = aba_base_historica.Range("A1048576").End(xlUp).Row
    ' ADICIONANDO LINHA DE ABATIMENTO TOTAL NA BASE HISTORICA
    aba_fbl5n_AR.Range("A" & linha & ":AB" & linha).Copy
    If aba_base_historica.Range("A" & linha_fim_base_historica).Value <> "" Then
        linha_fim_base_historica = linha_fim_base_historica + 1
    End If
    aba_base_historica.Range("A" & linha_fim_base_historica).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    If tipo_abatimento = "Boleto Abatido Integralmente" Then
        aba_base_historica.Range("AC" & linha_fim_base_historica).Value = "Abatimento Integral"
    ElseIf tipo_abatimento = "Boleto Abatido Parcialmente" Then
        aba_base_historica.Range("AC" & linha_fim_base_historica).Value = "Abatimento Parcial"
    End If
    
    aba_base_historica.Range("AD" & linha_fim_base_historica).Value = Date
    linha_aba_titulos_a_abater = linha_aba_titulos_a_abater + 1

End Sub

Sub FBL5N(abater_integrais As Boolean, ByVal condicao_sem_linhas_RV As Boolean)

    Dim qtde_partidas_compensacao As Integer
    
    session.findById("wnd[0]/usr/lbl[" & x_payer_session & ",2]").SetFocus
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/tbar[1]/btn[38]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = payer_atual
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]").sendVKey 5
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    If session.findById("wnd[0]/sbar").text = "Marcar pelo menos uma partida" Then
        Exit Sub
    End If
    If abater_integrais Then
        session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "ABATIDO TOTAL"
    Else
        session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "ABATIDO PARCIAL"
    End If
    session.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    

    Call LimparFiltros(tabela_titulos_a_abater)

    If Not abater_integrais Then
        
        session_2.findById("wnd[0]/usr/lbl[" & x_payer_session_2 & ",2]").SetFocus
        session_2.findById("wnd[0]").sendVKey 2
        session_2.findById("wnd[0]/usr/lbl[" & x_num_doc_session_2 & ",2]").SetFocus
        session_2.findById("wnd[0]").sendVKey 2
        session_2.findById("wnd[0]/tbar[1]/btn[38]").press
        session_2.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = payer_atual
        session_2.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press
        i2 = 2
        i_fim = aba_titulos_a_abater.Range("A1048576").End(xlUp).Row
        tabela_titulos_a_abater.Range.AutoFilter Field:=1, Criteria1:=payer_atual
        tabela_titulos_a_abater.Range.AutoFilter Field:=7, Criteria1:="Boleto Abatido Parcialmente"
        aba_titulos_a_abater.Range("B2:B" & i_fim).SpecialCells(xlCellTypeVisible).Copy
        session_2.findById("wnd[2]/tbar[0]/btn[16]").press
        session_2.findById("wnd[2]/tbar[0]/btn[24]").press
        session_2.findById("wnd[2]/tbar[0]/btn[8]").press
        session_2.findById("wnd[1]/tbar[0]/btn[0]").press
        
        qtde_linhas = VerificarQuantidadeLinhas(session_2, x_payer_session_2)
        qtde_pags_r1_payer_atual = VerificarQuantidadePaginas(session_2, x_num_doc_session_2, x_item_session_2)
        
        For i2 = 4 To qtde_linhas
            num_doc = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_num_doc_session_2 & "," & i2 & "]").text)
            item = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_item_session_2 & "," & i2 & "]").text)
            If ultimo_num_doc = num_doc And ultimo_item = item Then
                session_2.findById("wnd[0]/usr/chk[1," & i2 & "]").Selected = True
                session_2.findById("wnd[0]/tbar[1]/btn[45]").press
                session_2.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "ABATIDO PARCIAL"
                session_2.findById("wnd[0]").sendVKey 0
                On Error Resume Next
                session_2.findById("wnd[1]/tbar[0]/btn[0]").press
                On Error GoTo 0
                Exit For
            End If
        Next i2
        session_2.findById("wnd[0]").sendVKey 80
        
    Else
    

        session_2.findById("wnd[0]/usr/lbl[" & x_payer_session_2 & ",2]").SetFocus
        session_2.findById("wnd[0]").sendVKey 2
        session_2.findById("wnd[0]/usr/lbl[" & x_num_doc_session_2 & ",2]").SetFocus
        session_2.findById("wnd[0]").sendVKey 2
        session_2.findById("wnd[0]/tbar[1]/btn[38]").press
        session_2.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = payer_atual
        session_2.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press
        Call LimparFiltros(tabela_titulos_a_abater)
        i_fim = aba_titulos_a_abater.Range("A1048576").End(xlUp).Row
        tabela_titulos_a_abater.Range.AutoFilter Field:=1, Criteria1:=payer_atual
        aba_titulos_a_abater.Range("B2:B" & i_fim).SpecialCells(xlCellTypeVisible).Copy
        
        session_2.findById("wnd[2]/tbar[0]/btn[16]").press
        session_2.findById("wnd[2]/tbar[0]/btn[24]").press
        session_2.findById("wnd[2]/tbar[0]/btn[8]").press
        session_2.findById("wnd[1]/tbar[0]/btn[0]").press
        
        qtde_linhas = VerificarQuantidadeLinhas(session_2, x_payer_session_2)
        qtde_pags_r1_payer_atual = VerificarQuantidadePaginas(session_2, x_num_doc_session_2, x_item_session_2)
        
        For i2 = 1 To qtde_pags_r1_payer_atual
            For i3 = 4 To qtde_linhas
                num_doc = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_num_doc_session_2 & "," & i3 & "]").text)
                item = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_item_session_2 & "," & i3 & "]").text)
                For Each chave In Dict_AR.Keys
                    If Dict_AR(chave)(0) = num_doc And Dict_AR(chave)(1) = item Then
                        session_2.findById("wnd[0]/usr/chk[1," & i3 & "]").Selected = True
                        Exit For
                    End If
                Next chave
            Next i3
            session_2.findById("wnd[0]").sendVKey 82
        Next i2
        
        session_2.findById("wnd[0]/tbar[1]/btn[45]").press
        If session_2.findById("wnd[0]/sbar").text = "Marcar pelo menos uma partida" Then
            condicao_sem_linhas_RV = True
            Exit Sub
        End If
        
        session_2.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "ABATIDO TOTAL"
        session_2.findById("wnd[0]").sendVKey 0
        On Error Resume Next
        session_2.findById("wnd[1]/tbar[0]/btn[0]").press
        On Error GoTo 0
        
        session_2.findById("wnd[0]").sendVKey 80
        
        For i2 = 1 To qtde_pags_r1_payer_atual
            For i3 = 4 To qtde_linhas
                num_doc = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_num_doc_session_2 & "," & i3 & "]").text)
                item = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_item_session_2 & "," & i3 & "]").text)
                If ultimo_num_doc = num_doc And ultimo_item = item Then
                    session_2.findById("wnd[0]/usr/chk[1," & i3 & "]").Selected = True
                    session_2.findById("wnd[0]/tbar[1]/btn[45]").press
                    session_2.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "ABATIDO PARCIAL"
                    session_2.findById("wnd[0]").sendVKey 0
                    On Error Resume Next
                    session_2.findById("wnd[1]/tbar[0]/btn[0]").press
                    On Error GoTo 0
                    Exit For
                End If
            Next i3
        Next i2
        session_2.findById("wnd[0]").sendVKey 80
    End If
    

End Sub

Sub F32()

    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N F-32"
    session_3.findById("wnd[0]").sendVKey 0
    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[3,0]").Select
    session_3.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = payer_atual
    tipo_data_sap = VerificarFormatoDatas(session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text)
    session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = Format(Date, tipo_data_sap)
    session_3.findById("wnd[0]/usr/txtBKPF-MONAT").text = Month(Date)
    session_3.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "BR10"
    session_3.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BRL"
    session_3.findById("wnd[0]/tbar[1]/btn[16]").press
    
    If session_3.findById("wnd[0]/sbar").text <> "" Then
        conta_bloqueada = True
        Exit Sub
    End If
    

    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]").text = "ABATIDO TOTAL"
    session_3.findById("wnd[0]").sendVKey 0
    
    session_3.findById("wnd[0]/tbar[1]/btn[16]").press
    
    valor_restante_apos_compensacao = session_3.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-DIFFB").text
    
    session_3.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
    session_3.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press
    session_3.findById("wnd[0]/usr/tabsTS/tabpREST").Select
    
    
    qtde_partidas = session_3.findById("wnd[0]/usr/tabsTS/tabpREST/ssubPAGE:SAPDF05X:6106/txtRF05A-ANZPO").text
    
    For i2 = 0 To qtde_partidas
        Debug.Print session_3.findById("wnd[0]/usr/tabsTS/tabpREST/ssubPAGE:SAPDF05X:6106/tblSAPDF05XTC_6106/txtRFOPS_DK-BLART[3," & i2 & "]").text
        If session_3.findById("wnd[0]/usr/tabsTS/tabpREST/ssubPAGE:SAPDF05X:6106/tblSAPDF05XTC_6106/txtRFOPS_DK-BLART[3," & i2 & "]").text = "R1" Then
            session_3.findById("wnd[0]/usr/tabsTS/tabpREST/ssubPAGE:SAPDF05X:6106/tblSAPDF05XTC_6106/txtDF05B-PSDIF[8," & i2 & "]").SetFocus
            session_3.findById("wnd[0]").sendVKey 2
            Exit For
        End If
    Next i2
    
    session_3.findById("wnd[0]/mbar/menu[0]/menu[1]").Select
    session_3.findById("wnd[0]").sendVKey 21
    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0700/txtRF05A-AZEI1[0,0]").SetFocus
    session_3.findById("wnd[0]").sendVKey 2
    session_3.findById("wnd[0]/usr/txtBSEG-ZUONR").text = "ABATIDO PARCIAL"
    session_3.findById("wnd[0]").sendVKey 2
    If Left(session_3.findById("wnd[0]/sbar").text, 16) = "Base de desconto" Then
        session_3.findById("wnd[0]").sendVKey 0
    End If
    On Error Resume Next
    session_3.findById("wnd[1]").Close
    On Error GoTo 0
    session_3.findById("wnd[0]/tbar[0]/btn[11]").press
    If Left(session_3.findById("wnd[0]/sbar").text, 16) = "Base de desconto" Then
        session_3.findById("wnd[0]").sendVKey 0
        session_3.findById("wnd[0]/tbar[0]/btn[11]").press
    End If

End Sub

Sub ZFI156(baixar_integrais As Boolean)

    Dim i3, i_fim As Integer
    
    If baixar_integrais Then

        session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N ZFI156"
        session_3.findById("wnd[0]").sendVKey 0
        
        ' ETAPA BAIXA DE TITULO QUE FOI ABATIDO INTEGRALMENTE
        
        session_3.findById("wnd[0]/usr/btnBT_BX_TIT_APOS_COMPENSACAO").press
        session_3.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "BR10"
        session_3.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").text = payer_atual
        session_3.findById("wnd[0]/usr/txtS_ZUONR-LOW").text = "ABATIDO TOTAL"
        session_3.findById("wnd[0]/tbar[1]/btn[8]").press
        session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
        session_3.findById("wnd[0]/tbar[1]/btn[13]").press
        
    End If
    
    ' ETAPA ABATIMENTO DE TITULO QUE FOI ABATIDO PARCIALMENTE
processar_novamente:
    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N ZFI156"
    session_3.findById("wnd[0]").sendVKey 0
    session_3.findById("wnd[0]/usr/btnBT_ABATIMENTO").press
    session_3.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "BR10"
    session_3.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").text = payer_atual
    session_3.findById("wnd[0]/usr/txtS_ZUONR-LOW").text = "ABATIDO PARCIAL"
    session_3.findById("wnd[0]/tbar[1]/btn[8]").press
    session_3.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    session_3.findById("wnd[0]/tbar[1]/btn[13]").press
    
    If Right(session_3.findById("wnd[0]/sbar").text, 4) <> "BR10" Then
        GoTo processar_novamente
    End If
    
    Call LimparFiltros(tabela_titulos_a_abater)



End Sub


Sub VerificarContaBloqueada()

    session_3.findById("wnd[0]/tbar[0]/okcd").text = "/N F-32"
    session_3.findById("wnd[0]").sendVKey 0
    session_3.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[2,0]").Select
    session_3.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = payer_atual
    tipo_data_sap = VerificarFormatoDatas(session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text)
    session_3.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = Format(Date, tipo_data_sap)
    session_3.findById("wnd[0]/usr/txtBKPF-MONAT").text = Month(Date)
    session_3.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "BR10"
    session_3.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "BRL"
    session_3.findById("wnd[0]/tbar[1]/btn[16]").press
    
    If session_3.findById("wnd[0]/sbar").text <> "" Then
        conta_bloqueada = True
        session_2.findById("wnd[0]/usr/lbl[" & x_payer_session_2 & ",2]").SetFocus
        session_2.findById("wnd[0]").sendVKey 2
        session_2.findById("wnd[0]/tbar[1]/btn[38]").press
        session_2.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = payer_atual
        session_2.findById("wnd[1]/tbar[0]/btn[0]").press
        qtde_linhas = VerificarQuantidadeLinhas(session_2, x_payer_session_2)
        qtde_pags_r1_payer_atual = VerificarQuantidadePaginas(session_2, x_num_doc_session_2, x_item_session_2)
        
        For i2 = 1 To qtde_pags_r1_payer_atual
            For i3 = 4 To qtde_linhas
                num_doc = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_num_doc_session_2 & "," & i3 & "]").text)
                item = VBA.Trim(session_2.findById("wnd[0]/usr/lbl[" & x_item_session_2 & "," & i3 & "]").text)
                For Each chave In Dict_AR.Keys
                    If (Dict_AR(chave)(0) = num_doc And Dict_AR(chave)(1) = item) Or (ultimo_num_doc = num_doc And ultimo_item = item) Then
                        session_2.findById("wnd[0]/usr/chk[1," & i3 & "]").Selected = True
                        Exit For
                    End If
                Next chave
                
            Next i3
            session_2.findById("wnd[0]").sendVKey 82
        Next i2
        
        session_2.findById("wnd[0]/tbar[1]/btn[45]").press
        session_2.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "CTA BLOQUEADA"
        session_2.findById("wnd[0]").sendVKey 0
        On Error Resume Next
        session_2.findById("wnd[1]/tbar[0]/btn[0]").press
        On Error GoTo 0
        
        session_2.findById("wnd[0]").sendVKey 80
    End If
    
End Sub

