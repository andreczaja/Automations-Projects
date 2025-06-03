Attribute VB_Name = "FUNCOES_PUBLICAS"
' funcao que busca de forma dinamica a pasta presente no caminho equivalente a Documentos > AUTOMATIZAÇÕES, BIs & RPAs
'> Macro Reembolsos e Adiantamentos > Arquivos SAP Macro Reembolsos e Adiantamentos
Public Function BuscarPasta(ByRef caminho_pasta As String, Citrix As Boolean) As String
    Dim fso As Object
    Dim pastaGeral As Object
    Dim pastaAutomatizacoesBIRPA As Object
    Dim pastaMacroReembolsosAdiantamentos As Object
    Dim pastaArquivosSAP As Object
    
    ' Cria o objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Define a pasta inicial
    If Citrix Then
        On Error Resume Next
        Set pastaGeral = fso.getfolder(Replace(VBA.Environ("USERPROFILE"), "C:\", "\\Client\C$\") & "\OneDrive - Electrolux\")
        On Error GoTo 0
        ' verificação pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.getfolder(Replace(VBA.Environ("USERPROFILE"), "C:\", "\\Client\C$\") & "\OneDrive\")
        End If
    ElseIf Not Citrix Then
        On Error Resume Next
        Set pastaGeral = fso.getfolder(VBA.Environ("USERPROFILE") & "\OneDrive - Electrolux\")
        On Error GoTo 0
        ' verificação pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.getfolder(VBA.Environ("USERPROFILE") & "\OneDrive\")
        End If
    End If
        
    ' verificação pasta geral só com onedrive
    If Not VerificarPasta(pastaGeral) Then GoTo pedir_caminho_manualmente
    
    On Error Resume Next
    ' verificação pasta Excelencia
    Set pastaAutomatizacoesBIRPA = fso.getfolder(pastaGeral.Path & "\AUTOMATIZAÇÕES, BIs & RPAs\")
    If Not VerificarPasta(pastaAutomatizacoesBIRPA) Then
        Set pastaMacroReembolsosAdiantamentos = fso.getfolder(pastaGeral.Path & "\Macro Reembolsos e Adiantamentos\")
    Else
        Set pastaMacroReembolsosAdiantamentos = fso.getfolder(pastaAutomatizacoesBIRPA.Path & "\Macro Reembolsos e Adiantamentos\")
    End If
    If Not VerificarPasta(pastaMacroReembolsosAdiantamentos) Then
        Set pastaArquivosSAP = fso.getfolder(pastaGeral.Path & "\Arquivos SAP Macro Reembolsos e Adiantamentos\")
    Else
        Set pastaArquivosSAP = fso.getfolder(pastaMacroReembolsosAdiantamentos.Path & "\Arquivos SAP Macro Reembolsos e Adiantamentos\")
    End If
    On Error GoTo 0
    If Not VerificarPasta(pastaArquivosSAP) Then
        GoTo pedir_caminho_manualmente
    Else
        caminho_pasta = pastaArquivosSAP.Path
        GoTo fim
    End If
    
pedir_caminho_manualmente:
    MsgBox "Favor escolher o caminho no seu Computador a ser descarregado o arquivo baixado do site Transbank. Se não possuir a pasta, crie o atalho no Sharepoint e execute novamente a automação." & _
        "(Escolha a pasta no seu computador equivalente à pasta do Sharepoint:" & _
            "Documentos > AUTOMATIZAÇÕES, BIs & RPAs > Macro Reembolsos e Adiantamentos > Arquivos SAP Macro Reembolsos e Adiantamentos)"
     
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
    Set pastaAutomacoes = Nothing
    Set fso = Nothing
    Set pastaMacroReembolsosAdiantamentos = Nothing
    Set pastaArquivosSAP = Nothing
    
End Function
' funcao complementar da funcao BuscarPasta - apenas verifica se a pasta em questão existe
Public Function VerificarPasta(ByVal pasta As Object) As Boolean
    VerificarPasta = True
    If pasta Is Nothing Then
        VerificarPasta = False
    End If
End Function
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
            If Err.number = 0 Then
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                Set SapGui = GetObject("SAPGUI")
                Set App = SapGui.GetScriptingEngine
                Set connection = App.OpenConnection("002. P1L - SAP ECC Latin America (Single Sign On)", True)
                Set session = connection.Children(0)
                Exit Do
            End If
            On Error GoTo 0
            i = i - 1
        Loop
    End If
End Sub

' funcao que verifica o formato de data usuario de forma dinamica para inputar dados da forma correta
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

' funcao que assegura a atualizacao de base para trabalhar com dados corretos
Public Function AtualizarBase(ByVal aba As Worksheet, ByVal tabela As ListObject, ByVal linha_fim As Integer)

    Dim contador_atualizacoes As Integer
    
    contador_atualizacoes = 1
    Do Until contador_atualizacoes = 10
        tabela.QueryTable.BackgroundQuery = False
        tabela.QueryTable.Refresh
        contador_atualizacoes = contador_atualizacoes + 1
    Loop
    
    If linha_fim = aba.Range("A1048576").End(xlUp).Row Then
        contador_atualizacoes = 1
        Do Until contador_atualizacoes = 10 Or linha_fim <> aba.Range("A1048576").End(xlUp).Row
            tabela.QueryTable.BackgroundQuery = False
            tabela.QueryTable.Refresh
            contador_atualizacoes = contador_atualizacoes + 1
        Loop
    End If
End Function

' funcao que limpa filtros da tabela chamada pela sub
Public Function LimparFiltros(ByVal tabela As ListObject)

    On Error Resume Next
    tabela.AutoFilter.ShowAllData
    On Error GoTo 0

End Function

' funcao importante que cria e/ou define a sessao_2 e sessao_3 conforme etapa do código para interacao com SAP
Public Function InteracaoTelasSAP(ByRef session_number As Object, number As Integer, transacao As String)

    Debug.Print connection.Children.Count
    If connection.Children.Count < number Then

        session.createsession
        
        Do Until connection.Children.Count = number
            Application.Wait (Now + TimeValue("00:00:05"))
        Loop
        
        Debug.Print connection.Children.Count
        For i = 0 To connection.Children.Count - 1
            If connection.Children(CInt(i)).Info.Transaction = "SESSION_MANAGER" Then
                Set session_number = connection.Children(CInt(i))
            End If
        Next i
        session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N " & transacao
        session_number.findById("wnd[0]").sendVKey 0
    
    Else
        Debug.Print connection.Children.Count - 1
        For i = 0 To connection.Children.Count - 1
            Debug.Print connection.Children(CInt(i)).Info.Transaction
            If connection.Children(CInt(i)).Info.Transaction = "FBL5N" Then
                Set session = connection.Children(CInt(i))
                On Error Resume Next
                Set session_number = connection.Children(CInt(i + number - 1))
                If session_number Is Nothing Then
                    Set session_number = connection.Children(CInt(i + number - 2))
                End If
                On Error GoTo 0
                session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N " & transacao
                session_number.findById("wnd[0]").sendVKey 0
                Exit For
            End If
         Next i
    End If


End Function
' funcao que adiciona determinado item a determinado array
Public Function Add_ao_Array(ByRef array_() As Variant, ByVal item As String)
   
    ReDim Preserve array_(LBound(array_) To UBound(array_) + 1)
    array_(UBound(array_)) = item

End Function

' verifica os eixos x da coluna payer (cliente), num doc, item e tipo doc ESPECIFICAMENTE da session
Sub VerificarEixoXColunasSession()

    x_payer_session = 0
    x_num_doc_session = 0
    x_item_session = 0
    x_tipo_doc_session = 0
    For i2 = 1 To 500
        If x_payer_session <> 0 And x_num_doc_session <> 0 And x_item_session <> 0 And x_tipo_doc_session <> 0 Then
            Exit For
        End If
        On Error Resume Next
        session.findById("wnd[0]/usr/lbl[" & i2 & ",2]").SetFocus
        
        If Err.number = 0 Then
            Debug.Print session.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text
            If session.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text = "Cliente" Then
                x_payer_session = i2
                
            ElseIf session.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text = "Nº doc." Then
                x_num_doc_session = i2
                
            ElseIf session.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text = "Itm" Then
                x_item_session = i2
                
            ElseIf session.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text = "Tip" Then
                x_tipo_doc_session = i2
                
            End If
        End If
        On Error GoTo 0
    Next i2
    
End Sub
' verifica os eixos x da coluna payer (cliente), num doc, item e tipo doc ESPECIFICAMENTE da session_2
Sub VerificarEixoXColunasSession_2()
    x_payer_session_2 = 0
    x_num_doc_session_2 = 0
    x_item_session_2 = 0
    x_tipo_doc_session_2 = 0
    For i2 = 1 To 500
        If x_payer_session_2 <> 0 And x_num_doc_session_2 <> 0 And x_item_session_2 <> 0 And x_tipo_doc_session_2 <> 0 Then
            Exit For
        End If
        On Error Resume Next
        session_2.findById("wnd[0]/usr/lbl[" & i2 & ",2]").SetFocus
        
        If Err.number = 0 Then
            Debug.Print session_2.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text
            If session_2.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text = "Cliente" Then
                x_payer_session_2 = i2
                
            ElseIf session_2.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text = "Nº doc." Then
                x_num_doc_session_2 = i2
            ElseIf session_2.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text = "Itm" Then
                x_item_session_2 = i2
            ElseIf session_2.findById("wnd[0]/usr/lbl[" & i2 & ",2]").text = "Tip" Then
                x_tipo_doc_session_2 = i2
            End If
        End If
        On Error GoTo 0
    Next i2

End Sub

' salva o arquivo da FBL5N na pasta setada pela funcao BuscarPasta
Public Function SalvarArquivo(ByRef session_number As Object, nome_arquivo As String)

    session_number.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    session_number.findById("wnd[1]/tbar[0]/btn[0]").press
    session_number.findById("wnd[1]/usr/ctxtDY_PATH").text = Folder
    session_number.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
    session_number.findById("wnd[1]/tbar[0]/btn[11]").press

End Function

' verifica se o payer em questao já foi enquadrada em alguns dos 3 arrays (de abatimento, reembolsos com dados bancarios e reembolso sem dados bancarios)
' isso tudo para não haver duplicidade de tratamento nesses clientes
Public Function PayerDuplicado() As Boolean


    PayerDuplicado = False
    
    If UBound(Filter(array_payers_reembolsos_com_dados_bancarios, CLng(payer_atual))) >= 0 Or _
        UBound(Filter(array_payers_reembolsos_sem_dados_bancarios, CLng(payer_atual))) >= 0 Or _
            UBound(Filter(array_payers_abatimento, CLng(payer_atual))) >= 0 Then
        PayerDuplicado = True
        Exit Function
    End If
    
End Function


Public Function VerificarQuantidadeLinhas(ByRef session_number As Object, ByVal x_payer As Integer) As Integer

    qtde_linhas = 0
    For i2 = 4 To 100
        On Error Resume Next
        session_number.findById("wnd[0]/usr/lbl[" & x_payer & "," & i2 & "]").SetFocus
        If Not Err.number = 0 Then
            qtde_linhas = i2 - 1
            Exit For
        End If
        On Error GoTo 0
    Next i2
    VerificarQuantidadeLinhas = qtde_linhas
End Function
Public Function VerificarQuantidadePaginas(ByRef session_number As Object, ByVal x_num_doc As Integer, ByVal x_item As Integer) As Integer

    'verificando se as linhas encontradas de rv no cliente são maiores que a quantidades de linhas visiveis
    ' se sim, verificando quantas paginas possui e itera sobre elas para marcar todas as linhas
    Dim condicao_qtde_paginas_encontradas As Boolean
    Dim qtde_pags As Integer
    
    condicao_qtde_paginas_encontradas = False
    qtde_pags = 1
    i2 = 1

    Do Until condicao_qtde_paginas_encontradas
        num_doc = session_number.findById("wnd[0]/usr/lbl[" & x_num_doc & "," & qtde_linhas & "]").text
        item = session_number.findById("wnd[0]/usr/lbl[" & x_item & "," & qtde_linhas & "]").text
        session_number.findById("wnd[0]").sendVKey 82
        Debug.Print num_doc & item
        Debug.Print session_number.findById("wnd[0]/usr/lbl[" & x_num_doc & "," & qtde_linhas & "]").text & session_number.findById("wnd[0]/usr/lbl[" & x_item & "," & qtde_linhas & "]").text
        If num_doc & item <> session_number.findById("wnd[0]/usr/lbl[" & x_num_doc & "," & qtde_linhas & "]").text & session_number.findById("wnd[0]/usr/lbl[" & x_item & "," & qtde_linhas & "]").text Then
            qtde_pags = qtde_pags + 1
        Else
            condicao_qtde_paginas_encontradas = True
        End If
    Loop
    
    session_number.findById("wnd[0]").sendVKey 80
    
    VerificarQuantidadePaginas = qtde_pags
                
End Function

Public Function VerificarLinhasSBWP(ByRef session_number As Object) As Boolean
    
    VerificarLinhasSBWP = True
    ' ETAPA DE MANDAR PARA O APROVADOR AS SOLICITAÇÃO DE REEMBOLSO NA SBWP
    array_docs_F65 = Array()
    
    
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For linha = 2 To linha_fim_aba_reembolsos_pendentes
        doc_f65 = aba_reembolsos_pendentes.Range("AC" & linha).Value
        data_criacao = aba_reembolsos_pendentes.Range("AD" & linha).Value
        status_SBWP = aba_reembolsos_pendentes.Range("AE" & linha).Value
        If status_SBWP = "Não Solicitada Aprovação" Then
            If Not UBound(VBA.Filter(array_docs_F65, doc_f65)) >= 0 Then
                Call Add_ao_Array(array_docs_F65, doc_f65)
            End If
        End If
    Next linha
    
    If UBound(array_docs_F65) = -1 Then
        VerificarLinhasSBWP = False
        Exit Function
    End If
    
    
    session_number.findById("wnd[0]/tbar[0]/okcd").text = "/N SBWP"
    session_number.findById("wnd[0]").sendVKey 0
    
    
    session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "          2"
    session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "          5"
    session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").topNode = "          1"
    session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "          5"
    Set elemento_tabela_SBWP = session_number.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell")
    elemento_tabela_SBWP.selectColumn "WI_CD"
    elemento_tabela_SBWP.pressColumnHeader "WI_CD"
    elemento_tabela_SBWP.selectColumn "WI_CT"
    elemento_tabela_SBWP.pressColumnHeader "WI_CT"
    elemento_tabela_SBWP.selectColumn "WI_CT"
    elemento_tabela_SBWP.pressColumnHeader "WI_CT"
    contador = 1
recomecar_busca:
    For i = 0 To 1000000
        On Error Resume Next
        elemento_tabela_SBWP.setCurrentCell i, "WI_TEXT"
        elemento_tabela_SBWP.selectedRows = i
        If Err.number <> 0 Then
            linhas_SBWP = i
            Exit For
        End If
        On Error GoTo 0
    Next i
    
    If linhas_SBWP = 0 Then
        VerificarLinhasSBWP = False
        Exit Function
    End If

    For i = 0 To linhas_SBWP - 1
        doc_f65 = VBA.Right(elemento_tabela_SBWP.GetCellValue(i, "WI_TEXT"), 10)
        If UBound(VBA.Filter(array_docs_F65, doc_f65)) >= 0 Then
            Call CriarPastaDiariaeArquivo
            elemento_tabela_SBWP.currentCellRow = i
            elemento_tabela_SBWP.selectedRows = i
            elemento_tabela_SBWP.doubleClickCurrentCell
            session_number.findById("wnd[0]/usr/txtBKPF-BKTXT").text = session_number.findById("wnd[0]/usr/txtBKPF-BKTXT").text & "-"
            session_number.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
            session_number.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_PCATTA_CREA"
            session_number.findById("wnd[1]/usr/ctxtDY_PATH").text = Folder & "\Anexos Reembolsos\" & VBA.Format(Date, "dd.mm.yyyy")
            session_number.findById("wnd[1]/usr/ctxtDY_FILENAME").text = doc_f65 & ".xlsx"
            session_number.findById("wnd[1]/tbar[0]/btn[0]").press
            session_number.findById("wnd[0]/mbar/menu[0]/menu[6]").Select
            session_number.findById("wnd[1]/usr/ctxtG_INPUT").text = Form_SAP.approver
            session_number.findById("wnd[1]/usr/btnG_OK").press
            linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
            Call LimparFiltros(tabela_aba_reembolsos_pendentes)
            tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=29, Criteria1:=doc_f65
            aba_reembolsos_pendentes.Range("AD2:AD" & linha_fim_aba_reembolsos_pendentes).SpecialCells(xlCellTypeVisible).Value = Date
            aba_reembolsos_pendentes.Range("AE2:AE" & linha_fim_aba_reembolsos_pendentes).SpecialCells(xlCellTypeVisible).Value = "Aguardando Aprovação"
            Call LimparFiltros(tabela_aba_reembolsos_pendentes)
            reembolsos_com_dados_bancarios_processados = reembolsos_com_dados_bancarios_processados + 1
            elemento_tabela_SBWP.pressToolbarButton "EREF"
            elemento_tabela_SBWP.currentCellRow = 0
            GoTo recomecar_busca
        End If
    Next i
    
    Do Until contador > 2
        Application.Wait (Now + TimeValue("00:00:02"))
        contador = contador + 1
        elemento_tabela_SBWP.pressToolbarButton "EREF"
        GoTo recomecar_busca
    Loop

End Function

Sub CriarPastaDiariaeArquivo()

    Application.ScreenUpdating = False

    Dim pasta_diaria As String
    Dim novo_arquivo As Workbook
    pasta_diaria = Folder & "\Anexos Reembolsos\" & VBA.Format(Date, "dd.mm.yyyy")
    If Dir(pasta_diaria, vbDirectory) = "" Then
        MkDir pasta_diaria
    End If
    
    Call LimparFiltros(tabela_aba_reembolsos_pendentes)
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=29, Criteria1:=doc_f65
    aba_reembolsos_pendentes.Range("A1:AB" & linha_fim_aba_reembolsos_pendentes).SpecialCells(xlCellTypeVisible).Copy
    Workbooks.Add
    Set novo_arquivo = ActiveWorkbook
    With novo_arquivo.Sheets(1)
        .Range("A1").PasteSpecial
        .Columns("A:AB").AutoFit
    End With
    novo_arquivo.SaveAs pasta_diaria & "\" & doc_f65 & ".xlsx"
    novo_arquivo.Close
    Call LimparFiltros(tabela_aba_reembolsos_pendentes)
    
End Sub

        
