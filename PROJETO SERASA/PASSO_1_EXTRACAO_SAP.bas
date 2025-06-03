Attribute VB_Name = "PASSO_1_EXTRACAO_SAP"
Option Explicit
Public caminho_pasta, tipo_processo, payer_atual, tipo_data_sap As String
Public aba_fbl5h_base_geral, aba_fbl5h_base_compensados_serasa, aba_numero_remessa, aba_base_historica As Worksheet
Public tabela_aba_fbl5h_base_geral, tabela_aba_fbl5h_base_compensados_serasa, tabela_aba_base_historica As ListObject
Public array_motivo_baixa
Public numero_partidas As Integer
Private SapGui As Object, session As Object, Connection As Object, Applic As Object
Public condicao_ocorrencia_encontrada, browser_open, payer_repetido As Boolean
Private array_payers_fbl5h
Sub FBL5H_compensados()

    Application.DisplayAlerts = False

    If SapGui Is Nothing Then
        On Error Resume Next
        Set SapGui = GetObject("SAPGUI")
        Set Applic = SapGui.GetScriptingEngine
        Set Connection = Applic.Connections(0)
        Set session = Connection.Children(0)
        On Error GoTo 0
    End If
    

    Dim linhas_visiveis As Range
    Dim total_linhas_visiveis As Long

    tabela_aba_base_historica.Range.AutoFilter Field:=31, Criteria1:="", Operator:=xlAnd

    On Error Resume Next
    aba_base_historica.ShowAllData
    Set linhas_visiveis = tabela_aba_base_historica.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not linhas_visiveis Is Nothing Then
        total_linhas_visiveis = linhas_visiveis.Rows.Count
    Else
        total_linhas_visiveis = 0
    End If
    
    i_fim = aba_base_historica.Range("A1").End(xlDown).Row
    
    If (aba_base_historica.Range("A2").Value = "" And i_fim = 2) Or total_linhas_visiveis = 0 Then
        caminho_pasta = BuscarPasta("", True)
        FileCopy caminho_pasta & "\ZMD50 - BASE VAZIA.xls", caminho_pasta & "\ZMD50 - BASE COMPENSADOS SERASA.xls"
        FileCopy caminho_pasta & "\FBL5H - BASE VAZIA.xls", caminho_pasta & "\FBL5H - BASE COMPENSADOS SERASA.xls"
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N FBL5H"
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "SERASA COMP"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' referencia
    session.findById("wnd[0]/usr/btnITEM_SEL").press
    aba_base_historica.Range("E2:E" & i_fim).SpecialCells(xlCellTypeVisible).Copy
    session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/btnPUSH[4,12]").SetFocus
    session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/btnPUSH[4,12]").press
    session.findById("wnd[2]/tbar[0]/btn[16]").press
    session.findById("wnd[2]/tbar[0]/btn[24]").press
    session.findById("wnd[2]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    ' numero do payer
    aba_base_historica.Range("B2:B" & i_fim).SpecialCells(xlCellTypeVisible).Copy
    session.findById("wnd[0]/usr/btn%_S_CUST_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    
    session.findById("wnd[0]/usr/ctxtP_DFILE").Text = caminho_pasta & "/FBL5H - BASE COMPENSADOS SERASA.xls"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    On Error Resume Next
    aba_base_historica.ShowAllData
    On Error GoTo 0
    
    If session.findById("wnd[0]/sbar").Text = "Linhas exibidas: 0" Then
        MsgBox "Nenhum documento da base histórica foi baixado do último dia útil para hoje. Nenhum documento txt de exclusão será criado.", vbOKOnly
        aba_fbl5h_base_compensados_serasa.Range("A2:X1048576").ClearContents
    Else
        aba_base_historica.Range("BB2:BB1048576").ClearContents
        Call ListaPayers(aba_base_historica)
        Call ZMD50_compensados
        i_fim = aba_fbl5h_base_compensados_serasa.Range("A1048576").End(xlUp).Row
        Call AtualizarBase(aba_fbl5h_base_compensados_serasa, tabela_aba_fbl5h_base_compensados_serasa, i_fim)
    End If

End Sub

Sub ZMD50_compensados()

    On Error Resume Next
    tabela_aba_base_historica.AutoFilter.ShowAllData
    On Error GoTo 0


    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N ZMD50"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtS_VKORG-LOW").Text = "BR10"
    
    
    linha = aba_base_historica.Range("BB1048576").End(xlUp).Row
    aba_base_historica.Range("BB2:BB" & linha).Copy
    
    
    session.findById("wnd[0]/usr/btn%_S_KUNNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtPVARIANT").Text = "/SERASA"
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    On Error GoTo 0
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = caminho_pasta
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZMD50 - BASE COMPENSADOS SERASA.xls"

    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    ' apagar selecao de payers sem duplicata na coluna BB
    aba_base_historica.Range("BB2:BB1048576").ClearContents
    
    ' apagando status de automação antigos
    aba_fbl5h_base_compensados_serasa.Range("AD2:AD1048576").ClearContents
    

End Sub

Sub FBL5H_base_geral()

    If SapGui Is Nothing Then
        On Error Resume Next
        Set SapGui = GetObject("SAPGUI")
        Set Applic = SapGui.GetScriptingEngine
        Set Connection = Applic.Connections(0)
        Set session = Connection.Children(0)
        On Error GoTo 0
    End If

    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N FBL5H"
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtV-LOW").Text = "SERASA"
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    tipo_data_sap = VerificarFormatoDatas(session.findById("wnd[0]/usr/ctxtP_KEYDO").Text)
    
    session.findById("wnd[0]/usr/ctxtP_KEYDO").Text = VBA.Format(VBA.Date + 5, tipo_data_sap)
    
    session.findById("wnd[0]/usr/btnITEM_SEL").press

    session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC").verticalScrollbar.Position = 0
    session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC").verticalScrollbar.Position = 50
    i = 0
    Do Until session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/txtGS_MULTI_OR-SCRTEXT_L[0,1]").Text = "Vencimento líquido"
        session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC").verticalScrollbar.Position = 50 + i
        i = i + 1
    Loop
    
    ' PRIMEIRO DIA DE EXTRAÇÃO ALTERAR DE DATA INICIAL D-11 ATÉ D-21 E DEPOIS ALTERAR DEFINITIVAMENTE DE D-20 ATÉ D-25
    session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/ctxtGS_MULTI_OR-LOW[2,1]").Text = VBA.Format(VBA.Date - 5000, tipo_data_sap)
    session.findById("wnd[1]/usr/tblSAPLSE16NMULTI_OR_TC/ctxtGS_MULTI_OR-HIGH[3,1]").Text = VBA.Format(VBA.Date - 20, tipo_data_sap)
    session.findById("wnd[1]/tbar[0]/btn[8]").press

    

    session.findById("wnd[0]/usr/ctxtP_DFILE").Text = caminho_pasta & "/FBL5H - BASE GERAL.xls"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    
    If session.findById("wnd[0]/sbar").Text = "Linhas exibidas: 0" Then
        MsgBox "Nenhum documento da base histórica foi baixado do último dia útil para hoje. Nenhum documento txt de exclusão será criado.", vbOKOnly
        aba_fbl5h_base_geral.Range("A2:X1048576").ClearContents
    Else
        aba_fbl5h_base_geral.Range("BB2:BB1048576").ClearContents
        Call ListaPayers(aba_fbl5h_base_geral)
        Call ZMD50_base_geral
        i_fim = aba_fbl5h_base_geral.Range("A1048576").End(xlUp).Row
        Call AtualizarBase(aba_fbl5h_base_geral, tabela_aba_fbl5h_base_geral, i_fim)
    End If
    
End Sub



Sub ZMD50_base_geral()

    On Error Resume Next
    tabela_aba_fbl5h_base_geral.AutoFilter.ShowAllData
    On Error GoTo 0

    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N ZMD50"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtS_VKORG-LOW").Text = "BR10"
    session.findById("wnd[0]/usr/btn%_S_KUNNR_%_APP_%-VALU_PUSH").press
    
    linha = aba_fbl5h_base_geral.Range("BB1048576").End(xlUp).Row
    aba_fbl5h_base_geral.Range("BB2:BB" & linha).Copy
 
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtPVARIANT").Text = "/SERASA"
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    On Error GoTo 0
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = caminho_pasta
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZMD50 - BASE GERAL.xls"

    session.findById("wnd[1]/tbar[0]/btn[11]").press
    
    ' apagando status de automação antigos
    aba_fbl5h_base_geral.Range("AD2:AD1048576").ClearContents
    
    
    Application.DisplayAlerts = True

End Sub
Sub VerificarFormatoPadraoSAP()

    If SapGui Is Nothing Then
        On Error Resume Next
        Set SapGui = GetObject("SAPGUI")
        Set Applic = SapGui.GetScriptingEngine
        Set Connection = Applic.Connections(0)
        Set session = Connection.Children(0)
        On Error GoTo 0
    End If

    session.findById("wnd[0]").SetFocus
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N SU3"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").Select
    
    If session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Key <> "" Then
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Key = ""
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        
        i = Connection.Children.Count - 1
        Do Until Connection.Children(CInt(i)) Is Nothing
            Set session = Connection.Children(CInt(i))
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/N"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]").Close
            On Error Resume Next
            session.findById("wnd[1]").SetFocus
            If Err.Number = 0 Then
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                Set SapGui = GetObject("SAPGUI")
                Set Applic = SapGui.GetScriptingEngine
                Set Connection = Applic.OpenConnection("002. P1L - SAP ECC Latin America (Single Sign On)", True)
                Set session = Connection.Children(0)
                Exit Do
            End If
            On Error GoTo 0
            i = i - 1
        Loop
        
        
    End If
End Sub

Private Function ListaPayers(ByVal aba_correspondente As Worksheet)

Dim partidas_geradas As String
Dim x_payer, partidas_visiveis, i2, comprimento_sbar As Integer

    comprimento_sbar = VBA.Trim(Len(session.findById("wnd[0]/sbar").Text))
    If comprimento_sbar = 18 Then
        i2 = 1
    ElseIf comprimento_sbar = 19 Then
        i2 = 2
    ElseIf comprimento_sbar = 20 Then
        i2 = 3
    ElseIf comprimento_sbar = 21 Then
        i2 = 4
    ElseIf comprimento_sbar = 22 Then
        i2 = 5
    ElseIf comprimento_sbar = 23 Then
        i2 = 6
    ElseIf comprimento_sbar = 24 Then
        i2 = 7
    ElseIf comprimento_sbar = 25 Then
        i2 = 8
    ElseIf comprimento_sbar = 26 Then
        i2 = 9
    ElseIf comprimento_sbar = 27 Then
        i2 = 10
    End If

    partidas_geradas = Replace(Replace(VBA.Trim(VBA.Right(session.findById("wnd[0]/sbar").Text, i2)), ".", ""), ",", "")
    numero_partidas = CInt(partidas_geradas)
    
    i = 1
    array_payers_fbl5h = Array(session.findById("wnd[0]/shellcont/shell").getcellvalue(i, "KUNNR"))
    
    
    For i = 1 To numero_partidas
        Do
            payer_atual = session.findById("wnd[0]/shellcont/shell").getcellvalue(i, "KUNNR")
            If payer_atual = "" Then
                session.findById("wnd[0]/shellcont/shell").firstVisibleRow = i
            Else
                Exit Do
            End If
        Loop
        If Not UBound(VBA.Filter(array_payers_fbl5h, payer_atual)) >= 0 Then
            ReDim Preserve array_payers_fbl5h(LBound(array_payers_fbl5h) To UBound(array_payers_fbl5h) + 1)
            array_payers_fbl5h(UBound(array_payers_fbl5h)) = payer_atual
        End If
    Next i


    i2 = 2
    For i = LBound(array_payers_fbl5h) To UBound(array_payers_fbl5h)
        aba_correspondente.Range("BB" & i2).Value = array_payers_fbl5h(i)
        i2 = i2 + 1
    Next i
    
End Function

