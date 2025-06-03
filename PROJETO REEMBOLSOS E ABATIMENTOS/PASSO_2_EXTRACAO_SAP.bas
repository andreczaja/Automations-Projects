Attribute VB_Name = "PASSO_2_EXTRACAO_SAP"
Public SapGui As Object, session As Object, session_2 As Object, session_3 As Object, connection As Object, Applic As Object, Window As Object
Public linha, linha_fim, i2, i3, i_fim, janelas_sap, OpenWin As Integer
Public data_agrupado_pagamento, dia_agrupado_pagamento, mes_agrupado_pagamento, ano_agrupado_pagamento, usuario_biz_aprovacao_reembolso, tipo_data_sap, _
     ultima_linha_preenchida_atribuicoes_exclusas, chave_de_ref_3 As String
Private nome_coluna  As String
Public i As Long
Public Folder As String

Sub extracao_sap_bases_novas()

    Call declaracao_vars
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    Set connection = Applic.Connections(0)
    Set session = connection.Children(0)
    Call VerificarFormatoPadraoSAP
    Form_SAP.Show
    dia_agrupado_pagamento = Left(Form_SAP.txt_box_data_agrupado_pgto_SAP, 2)
    mes_agrupado_pagamento = Mid(Form_SAP.txt_box_data_agrupado_pgto_SAP, 4, 2)
    ano_agrupado_pagamento = Right(Form_SAP.txt_box_data_agrupado_pgto_SAP, 2)
    data_agrupado_pagamento = dia_agrupado_pagamento & "." & mes_agrupado_pagamento & "." & ano_agrupado_pagamento
    aba_reembolsos_aprovados.Range("BC1").Value = data_agrupado_pagamento
    Call LimparFiltros(tabela_aba_fbl5n_AR)
    Call LimparFiltros(tabela_aba_fbl5n_credito_devolucao)
    Call LimparFiltros(tabela_aba_plan_distribuicao)
    Call LimparFiltros(tabela_titulos_a_abater)
    On Error GoTo 0
    
    ' verificando se estão preenchidas as celulas que carregam o numero do chamado criado na ellevo e se a base nao está vazia
    ' nesse caso irá chamar a sub que alterará a atribuição dessas linhas para "ELLEVO [Núm Chamado]"

    
    If aba_reembolsos_aprovados.Range("BB1").Value <> "" And aba_reembolsos_aprovados.Range("E2").Value <> "" Then
        Call alterar_linhas_reembolsos_aprovados
    End If
    
    Folder = BuscarPasta("", True)
    
    tabela_aba_plan_distribuicao.QueryTable.BackgroundQuery = False
    tabela_aba_plan_distribuicao.QueryTable.Refresh False
    
    Call VerificarAtribuicaoRVs
      
      
    ' inicio da etapa de extracao a qual acontece pela FBL5N, puxa todos os clientes da plan de distribuicao filtrando por apenas os cobraveis
    ' exclui linhas com atribuicoes especificas
    ' apenas documento R1
    session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
    session.findById("wnd[0]").sendVKey 0
    
    tipo_data_sap = VerificarFormatoDatas(session.findById("wnd[0]/usr/ctxtPA_STIDA").text)
    
    linha_fim = aba_plan_distribuicao.Range("A1048576").End(xlUp).Row
    aba_plan_distribuicao.Range("A2:A" & linha_fim).Copy
    session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/chkX_SHBV").Selected = True
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "BR10"
    
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/ABATREEMB"
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN011_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,0]").text = "PROCESSADO AUTOMAC"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,1]").text = "ELLEVO*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,2]").text = "*REEMBOLSO*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,3]").text = "*UTILIZAR*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,4]").text = "REEMB AUT*"
    session.findById("wnd[0]").sendVKey 0
    ultima_linha_preenchida_atribuicoes_exclusas = Replace(Replace(VBA.Right(session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").text, 3), "(", ""), ")", "")
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.Position = ultima_linha_preenchida_payers_exclusos + 2
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,3]").text = "AUTOMACAO DEV"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,4]").text = "AG PROCESS SBWP"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,5]").text = "ABATIDO PARCIAL"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/txtRSCSEL_255-SLOW_E[1,6]").text = "ABATIDO TOTAL"
    
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN016-LOW").text = "R1"
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    
    For i = 1 To 10000
        On Error Resume Next
        session.findById("wnd[0]/usr/lbl[" & i & ",2]").SetFocus
        nome_coluna = session.findById("wnd[0]/usr/lbl[" & i & ",2]").text
        On Error GoTo 0
        If nome_coluna = "Dt.lçto." Then
            session.findById("wnd[0]").sendVKey 2
            Exit For
        End If
    Next i
    
    session.findById("wnd[0]/tbar[1]/btn[38]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").text = Format(Date - 5, tipo_data_sap)
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    
    
    Call SalvarArquivo(session, "FBL5N-R1.txt")
    linha_fim = aba_fbl5n_credito_devolucao.Range("A1048576").End(xlUp).Row
    Call AtualizarBase(aba_fbl5n_credito_devolucao, tabela_aba_fbl5n_credito_devolucao, linha_fim)
    Call InteracaoTelasSAP(session_2, 2, "FBL5N")
    
    
    linha_fim = aba_fbl5n_credito_devolucao.Range("C1048576").End(xlUp).Row
    aba_fbl5n_credito_devolucao.Range("C2:C" & linha_fim).Copy
    
    ' EXTRACAO DA BASE DE RVs com variante AUT. DEVOLUCAO
    ' PARTIDAS ABERTA D+5
    ' VCTO LIQUIDO HOJE +10 NO MINIMO
    session_2.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    session_2.findById("wnd[1]/usr/txtV-LOW").text = "AUT.DEVOLUCAO"
    session_2.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session_2.findById("wnd[1]/tbar[0]/btn[8]").press
    session_2.findById("wnd[0]/usr/ctxtPA_VARI").text = "/ABATREEMB"
    session_2.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
    session_2.findById("wnd[1]/tbar[0]/btn[16]").press
    session_2.findById("wnd[1]/tbar[0]/btn[24]").press
    session_2.findById("wnd[1]/tbar[0]/btn[8]").press
    session_2.findById("wnd[0]/usr/ctxtPA_STIDA").text = Format(Date + 5, tipo_data_sap)
    session_2.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = Format(Date + 10, tipo_data_sap)
    session_2.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = Format(Date + 500, tipo_data_sap)
    session_2.findById("wnd[0]/tbar[1]/btn[8]").press
    
    Call SalvarArquivo(session_2, "FBL5N-AR.txt")
    
    linha_fim = aba_fbl5n_AR.Range("A1048576").End(xlUp).Row
    Call AtualizarBase(aba_fbl5n_AR, tabela_aba_fbl5n_AR, linha_fim)
 
    
    Call AbatimentoOuReembolso
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub VerificarAtribuicaoRVs()

    session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
    session.findById("wnd[0]").sendVKey 0
    linha_fim = aba_plan_distribuicao.Range("A1048576").End(xlUp).Row
    aba_plan_distribuicao.Range("A2:A" & linha_fim).Copy
    session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "BR10"
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN016-LOW").text = "RV"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN011_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").text = "ABATIDO TOTAL"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    If Left(session.findById("wnd[0]/sbar").text, 12) <> "São exibidas" Then
        Exit Sub
    End If
    session.findById("wnd[0]").sendVKey 5
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    If session.findById("wnd[0]/sbar").text = "Marcar pelo menos uma partida" Then
        Exit Sub
    End If
    session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = "-"
    session.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    

End Sub

Sub alterar_linhas_reembolsos_aprovados()

    session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
    session.findById("wnd[0]").sendVKey 0
    
    
    tipo_data_sap = VerificarFormatoDatas(session.findById("wnd[0]/usr/ctxtPA_STIDA").text)
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "BR10"
    ' ESCOLHENDO VARIANTE "REEMBO. AUTOMA" E DEMAIS FILTROS PARA GERAÇÃO DA BASE
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    session.findById("wnd[1]/usr/txtV-LOW").text = "REEMBO. AUTOMA"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    session.findById("wnd[0]/usr/ctxtPA_STIDA").text = VBA.Format(Date, tipo_data_sap)
    ' ALTERANDO LINHAS COM ATRIBUICAO IGUAL A PROCESSADO AUTOMAC PARA "ELLEVO [NUM CHAMADO ELLEVO]"
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN011_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").text = "PROCESSADO AUTOMAC"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    
    linha_fim = aba_reembolsos_aprovados.Range("A1048576").End(xlUp).Row
    aba_reembolsos_aprovados.Range("C2:C" & linha_fim).Copy
    
    session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    If Left(session.findById("wnd[0]/sbar").text, 12) <> "São exibidas" Then
        Exit Sub
    End If
    
    chamado_criado = aba_reembolsos_aprovados.Range("BB1").Value
    
    session.findById("wnd[0]").sendVKey 5
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = chamado_criado
    session.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    
    aba_reembolsos_aprovados.Range("BB1").ClearContents

End Sub




