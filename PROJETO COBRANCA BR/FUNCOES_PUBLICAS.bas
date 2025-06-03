Attribute VB_Name = "FUNCOES_PUBLICAS"
Sub VerificarFormatoPadraoSAP()

    session.findById("wnd[0]").SetFocus
    session.findById("wnd[0]/tbar[0]/okcd").text = "/N SU3"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").Select
    
    If session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Key <> "" Then
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Key = ""
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        
        i = connection.Children.Count - 1
        Do Until connection.Children(CInt(i)) Is Nothing
            Set session = connection.Children(CInt(i))
            session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]").Close
            On Error Resume Next
            session.findById("wnd[1]").SetFocus
            If Err.Number = 0 Then
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                Set SapGui = GetObject("SAPGUI")
                Set Applic = SapGui.GetScriptingEngine
                Set connection = Applic.OpenConnection("002. P1L - SAP ECC Latin America (Single Sign On)", True)
                Set session = connection.Children(0)
                Exit Do
            End If
            On Error GoTo 0
            i = i - 1
        Loop
        
        
    End If
End Sub
Public Function VerificarFormatoDatas(data As String) As String

Dim tipo, formato_data_tipo_1, formato_data_tipo_2, formato_data_tipo_3, formato_data_tipo_4 As String

    formato_data_tipo_1 = "yyyy-mm-dd"
    formato_data_tipo_2 = "dd.mm.yyyy"
    formato_data_tipo_3 = "yyyy.mm.dd"
    formato_data_tipo_4 = "yyyy/mm/dd"
    
    If data = VBA.Format(VBA.Date, formato_data_tipo_1) Then
        tipo = formato_data_tipo_1
           
    ElseIf data = VBA.Format(VBA.Date, formato_data_tipo_2) Then
        tipo = formato_data_tipo_2
        
    ElseIf data = VBA.Format(VBA.Date, formato_data_tipo_3) Then
        tipo = formato_data_tipo_3
           
    ElseIf data = VBA.Format(VBA.Date, formato_data_tipo_4) Then
        tipo = formato_data_tipo_4
        
    End If
    
    VerificarFormatoDatas = tipo
End Function

Public Function LimparFiltros()
    On Error Resume Next
    aba_export_sap.Range("A6").Activate
    aba_export_sap.ShowAllData
    On Error GoTo 0
End Function

' funcao criada para criar de forma dinamica a pasta a ser descarregado o arquivo do SAP
Public Function BuscarPasta(ByRef caminho_pasta As String, Citrix As Boolean) As String

    Dim fso As Object
    Dim pastaGeral As Object
    Dim pastaAUTOMATIZACOESBISRPA As Object
    Dim pastaMacroCobranca As Object
    Dim pastaArquivoTXTSAP As Object
    
    ' Cria o objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Define a pasta inicial
    If Citrix Then
        On Error Resume Next
        Set pastaGeral = fso.GetFolder(Replace(VBA.Environ("USERPROFILE"), "C:\", "\\Client\C$\") & "\OneDrive - Electrolux\")
        On Error GoTo 0
        ' verificação pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.GetFolder(Replace(VBA.Environ("USERPROFILE"), "C:\", "\\Client\C$\") & "\OneDrive\")
        End If
    ElseIf Not Citrix Then
        On Error Resume Next
        Set pastaGeral = fso.GetFolder(VBA.Environ("USERPROFILE") & "\OneDrive - Electrolux\")
        On Error GoTo 0
        ' verificação pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.GetFolder(VBA.Environ("USERPROFILE") & "\OneDrive\")
        End If
    End If
        
    ' verificação pasta geral só com onedrive
    If Not VerificarPasta(pastaGeral) Then GoTo pedir_caminho_manualmente
    
    On Error Resume Next
    ' verificação pasta Excelencia
    Set pastaAUTOMATIZACOESBISRPA = fso.GetFolder(pastaGeral.Path & "\AUTOMATIZAÇÕES, BIs & RPAs\")
    If Not VerificarPasta(pastaAUTOMATIZACOESBISRPA) Then
        Set pastaMacroCobranca = fso.GetFolder(pastaGeral.Path & "\Macro de Cobrança\")
    Else
        Set pastaMacroCobranca = fso.GetFolder(pastaAUTOMATIZACOESBISRPA.Path & "\Macro de Cobrança\")
    End If
    If Not VerificarPasta(pastaMacroCobranca) Then
        Set pastaArquivoTXTSAP = fso.GetFolder(pastaGeral.Path & "\Arquivo TXT COBRANCA SAP\")
    Else
        Set pastaArquivoTXTSAP = fso.GetFolder(pastaMacroCobranca.Path & "\Arquivo TXT COBRANCA SAP\")
    End If
    
    If Not VerificarPasta(pastaArquivoTXTSAP) Then
        GoTo pedir_caminho_manualmente
    Else
        caminho_pasta = pastaArquivoTXTSAP.Path
        GoTo fim
    End If
    On Error GoTo 0
    ' se não encontrar a pasta, irá pedir para o usuário selecionar o caminho manualmente com a msoFileDialogFolderPicker
pedir_caminho_manualmente:
    MsgBox "Favor escolher o caminho no seu Computador a ser descarregado o arquivo baixado do site Transbank. Se não possuir a pasta, crie o atalho no Sharepoint e execute novamente a automação." & _
        "(Escolha a pasta no seu computador equivalente à pasta do Sharepoint:" & _
            "Documentos > AUTOMATIZAÇÕES, BIs & RPAs > Macro de Cobrança > Arquivo TXT SAP)"
     
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
    Set pastaAUTOMATIZACOESBISRPA = Nothing
    Set pastaArquivoTXTSAP = Nothing
    Set pastaMacroCobranca = Nothing
    Set fso = Nothing

End Function

Public Function VerificarPasta(ByVal pasta As Object) As Boolean
    ' função simples que apenas verifica quando é chamada na função BuscarPasta se determinada pasta existe no diretório do usuário
    VerificarPasta = True
    If pasta Is Nothing Then
        VerificarPasta = False
    End If
End Function



