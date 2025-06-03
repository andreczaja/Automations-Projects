Attribute VB_Name = "PASSO0_extracao_SAP"
'variavel que possibilita que o codigo encerre caso se clique no botao de Cancelar no meio
' da execucao do codigo
Public condicao_sem_cobranca_preventiva As Boolean
Public SapGui As Object, WSHShell As Object, session As Object, Applic As Object, connection As Object, SapGuiWB As Object



Sub SAP_Login()

    frm_passos.Hide
    
    
    Folder = BuscarPasta("")

    ' Verifica se o SAP está aberto
    On Error Resume Next
    Set SapGui = GetObject("SAPGUI")
    On Error GoTo 0
            
    If Not SapGui Is Nothing Then
        
        ' SAP ABERTO
        Set SapGui = GetObject("SAPGUI")
        Set Applic = SapGui.GetScriptingEngine
        
        On Error Resume Next
        Set connection = Applic.Connections(0)
        
        If connection Is Nothing Then
            MsgBox "Favor iniciar a automatização apenas com o SAP aberto e logado!", vbOKOnly, "Electrolux Group"
            Exit Sub
        End If
        Set session = connection.Children(0)
        
    Else:
      
      ' SAP FECHADO
        MsgBox "Favor iniciar a automatização apenas com o SAP aberto e logado!", vbOKOnly, "Electrolux Group"
        End

    End If


'clica no botão para entrar no SAP

    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").Maximize
'voltar para a pagina maior
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nBPMDG/UTL_BROWSER"
    session.findById("wnd[0]").sendVKey 0

'troca de móduo (transação fbl5n)
    Call fbl5n
    
    
    Set SapGui = Nothing
    Set WSHShell = Nothing
    Set session = Nothing
    Set Applic = Nothing
    Set connection = Nothing
    Set SapGuiWB = Nothing
    
    Dim status_bloqueio_linha_fim_analista, status_bloqueio_linha_fim_check_bloqueios As Integer
    Dim analistas_sem_bloqueio, status_bloqueio As String
    Dim planilha As Workbook
    Dim aba, aba_export_sap As Worksheet
    Dim tbl As ListObject
    
    Set planilha = ThisWorkbook
    Set aba = planilha.Sheets("Controle Diário")
    Set aba_export_sap = planilha.Sheets("Export SAP")
    Set tbl = aba_export_sap.ListObjects("Export_FBL5N___Cobráveis")
    
    Call atualizar
    
    If condicao_sem_cobranca_preventiva Then
        aba.Range("BB1").ClearContents
        FileCopy Folder & "\FBL5N - BASE VAZIA.txt", Folder & "\FBL5N-C.txt"
        tbl.QueryTable.Refresh False
        MsgBox "A base está atualizada.", vbOKOnly
        Exit Sub
    End If
    
    
    
    aba.Range("BB1").Value = "Cobrança Preventiva"
    
    analistas_sem_bloqueio = ""
        
    status_bloqueio_linha_fim_analista = aba.Range("A1").End(xlDown).Row
    
    On Error Resume Next
    For status_bloqueio_linha_fim_check_bloqueios = 1 To status_bloqueio_linha_fim_analista
        If aba.Range("B" & status_bloqueio_linha_fim_check_bloqueios).Value = "" Then
            analistas_sem_bloqueio = analistas_sem_bloqueio & " - " & aba.Range("A" & status_bloqueio_linha_fim_check_bloqueios).Value
        End If
    Next status_bloqueio_linha_fim_check_bloqueios
    
    If analistas_sem_bloqueio <> "" Then
    analistas_sem_bloqueio = Mid(analistas_sem_bloqueio, 4)
        MsgBox ("Base de faturas cobráveis atualizada!" & vbNewLine & vbNewLine & _
            "Os analistas " & analistas_sem_bloqueio & " não inseriram dos bloqueios referente ao dia atual." & vbNewLine & _
            "Favor realizar a cobrança e executar os passos seguintes SOMENTE quando for corrigido."), vbOKOnly
        GoTo fim
    End If
    
    Sheets("Export SAP").Range("a6").Activate
    ActiveSheet.ShowAllData
    
    MsgBox "Base de faturas cobráveis atualizada", vbOKOnly

fim:
    
End Sub



Sub fbl5n()

Dim a As Integer
Dim fecha, variante As String

variante = "COLECCIÓN"
condicao_sem_cobranca_preventiva = False

0:
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n FBL5N"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/tbar[0]/btn[16]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").Text = "tc04"
        session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
        session.findById("wnd[1]/usr/txtENAME-LOW").Text = "BR_FI_NSILVA"
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        
        If variante = "COLECCIÓN" Then
            For a = 0 To 30
                If session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").getCellValue(a, "TEXT") = variante Then
                    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = a
                    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = a
                    Exit For
                End If
            Next a
        ElseIf variante = "/construtoras" Then
            For a = 0 To 30
                If session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").getCellValue(a, "TEXT") = variante Then
                    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = a
                    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = a
                    Exit For
                End If
            Next a
        End If
            session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
            session.findById("wnd[0]/tbar[1]/btn[8]").press
        
        If session.findById("wnd[0]/sbar").Text = "No se ha seleccionado ninguna partida (véase texto explicativo)" Then
            condicao_sem_cobranca_preventiva = True
            Exit Sub
        End If
        session.findById("wnd[0]/mbar/menu[5]/menu[6]/menu[0]").Select
        On Error Resume Next
        fecha = ""
        
        If variante = "COLECCIÓN" Then
            For a = 105 To 600
            
            fecha = session.findById("wnd[0]/usr/lbl[" & a & ",7]").Text
                Debug.Print fecha
                If fecha = "Fecha pago" Then
                
                    Exit For
                End If
            Next a
             
            session.findById("wnd[0]/usr/lbl[" & a & ",7]").SetFocus
            session.findById("wnd[0]").sendVKey 2
            session.findById("wnd[0]/tbar[1]/btn[38]").press
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "01.01.2016"
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = Format(Date, "dd.mm.yyyy")
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").SetFocus
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").caretPosition = 10
            session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").verticalScrollbar.Position = 12
            session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST/txtGT_WRITE_LIST-OUTPUTLEN[2,4]").Text = "20"
            session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST/txtGT_WRITE_LIST-OUTPUTLEN[2,5]").Text = "20"
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        End If
        session.findById("wnd[0]/tbar[1]/btn[32]").press
        session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST/txtGT_WRITE_LIST-OUTPUTLEN[2,11]").Text = "10"
        session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST").verticalScrollbar.Position = 12
        session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST/txtGT_WRITE_LIST-OUTPUTLEN[2,4]").Text = "20"
        session.findById("wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST/txtGT_WRITE_LIST-OUTPUTLEN[2,5]").Text = "20"
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Folder
        If variante = "COLECCIÓN" Then
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "FBL5N.txt"
        Else
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "FBL5N-C.txt"
        End If
        session.findById("wnd[1]/tbar[0]/btn[11]").press
        
        If variante = "COLECCIÓN" Then
            variante = "/construtoras"
            GoTo 0
        End If

End Sub


Sub atualizar()

    Dim contador As Integer
    Dim planilha As Workbook
    Dim aba_export_sap As Worksheet
    Dim tbl As ListObject
    
    Set planilha = ThisWorkbook
    Set aba_export_sap = planilha.Sheets("Export SAP")
    Set tbl = aba_export_sap.ListObjects("Export_FBL5N___Cobráveis")
    
    linha_fim = aba_export_sap.Range("A1048576").End(xlUp).Row
    
    tbl.QueryTable.Refresh False
    If linha_fim = aba_export_sap.Range("A1048576").End(xlUp).Row Then
        contador = 1
        Do Until contador = 5 Or linha_fim <> aba_export_sap.Range("A1048576").End(xlUp).Row
            tbl.QueryTable.Refresh False
            contador = contador + 1
        Loop
    End If
    
    Sheets("Base E-mails").ListObjects("Base_E_mails").QueryTable.Refresh False
    Sheets("Controle Diário").ListObjects("Status_Bloqueios_Diários_Analistas").QueryTable.Refresh False
    
    


End Sub
Public Function BuscarPasta(ByRef caminho_pasta As String) As String
    Dim fso As Object
    Dim pastaGeral As Object
    Dim pastaCHILE As Object
    Dim pastaCOBRANCA As Object
    Dim pastaBICOBRANCACHILE As Object
    
    ' Cria o objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Define a pasta inicial
    On Error Resume Next
    Set pastaGeral = fso.getfolder(VBA.Environ("USERPROFILE") & "\OneDrive - Electrolux\")
    On Error GoTo 0
    ' verificação pasta geral com onedrive - electrolux
    If Not VerificarPasta(pastaGeral) Then
        Set pastaGeral = fso.getfolder(VBA.Environ("USERPROFILE") & "\OneDrive\")
    End If
        
    ' verificação pasta geral só com onedrive
    If Not VerificarPasta(pastaGeral) Then GoTo pedir_caminho_manualmente
    
    On Error Resume Next
    ' verificação pasta Excelencia
    Set pastaCHILE = fso.getfolder(pastaGeral.Path & "\CHILE\")
    If Not VerificarPasta(pastaCHILE) Then
        Set pastaCOBRANCA = fso.getfolder(pastaGeral.Path & "\Cobrança\")
    Else
        Set pastaCOBRANCA = fso.getfolder(pastaCHILE.Path & "\Cobrança\")
    End If
    If Not VerificarPasta(pastaCOBRANCA) Then
        Set pastaBICOBRANCACHILE = fso.getfolder(pastaGeral.Path & "\BI - Cobrança Chile\")
    Else
        Set pastaBICOBRANCACHILE = fso.getfolder(pastaCOBRANCA.Path & "\BI - Cobrança Chile\")
    End If
    On Error GoTo 0
    If Not VerificarPasta(pastaBICOBRANCACHILE) Then
        GoTo pedir_caminho_manualmente
    Else
        caminho_pasta = pastaBICOBRANCACHILE.Path
        GoTo fim
    End If
    
pedir_caminho_manualmente:
    MsgBox "Favor escolher o caminho no seu Computador a ser descarregado o arquivo baixado do site Transbank. Se não possuir a pasta, crie o atalho no Sharepoint e execute novamente a automação." & _
        "(Escolha a pasta no seu computador equivalente à pasta do Sharepoint:" & _
            "Documentos > CHILE > Cobrança > BI - Cobrança Chile)"
     
    'Seleção da pasta para salvar o arquivo
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usuário selecionou uma pasta
            caminho_pasta = .SelectedItems(1) & "\"
        Else
             'O usuário cancelou a seleção da pasta
            MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
            End
        End If
    End With
    
fim:

    BuscarPasta = caminho_pasta
    ' Libera os objetos
    Set pastaGeral = Nothing
    Set pastaCHILE = Nothing
    Set fso = Nothing
    Set pastaCOBRANCA = Nothing
    Set pastaBICOBRANCACHILE = Nothing
    
End Function

Public Function VerificarPasta(ByVal pasta As Object) As Boolean
    VerificarPasta = True
    If pasta Is Nothing Then
        VerificarPasta = False
    End If
End Function

