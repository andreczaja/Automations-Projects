Attribute VB_Name = "PASSO3_criar_arquivos_e_emails"
Private arquivocriado As Workbook
Private emails_clientes, email_analista, emailBody1, emailBody2, emailBody3, emailBody4, emailBody5, emailBody6, emailBody7, emailBody8, emailBody9, emailBody10, _
    Destinatario, Copia, nome_cliente, iniciais_analista As String
Private OutlookApp As Object
Private range_facturas As Range
Private OutlookMail, objFSO, ObjPasta, arquivo, ns, caixa_do_b2b, MailItem As Object
Private aba_modelos_email As Worksheet



Sub criar_arquivos_de_clientes_()
Attribute criar_arquivos_de_clientes_.VB_ProcData.VB_Invoke_Func = " \n14"


    Application.DisplayAlerts = False
    
    Call SetVarsEmails
    Call LimparFiltros
    
    Set aba_modelos_email = ThisWorkbook.Sheets("Modelos de Email")

' VERIFICACAO E ENVIO DE E-MAILS A ANALISTAS COM CLIENTES SEM EMAILS CADASTRADOS
    For i = LBound(array_analistas) To UBound(array_analistas)
        analista = array_analistas(i)
        If analista = "" Then
            Exit For
        End If
        tabela_aba_export_sap.Range.AutoFilter Field:=32, Criteria1:=analista, Operator:=xlAnd
        tabela_aba_export_sap.Range.AutoFilter Field:=34, Criteria1:="", Operator:=xlAnd
        If Application.WorksheetFunction.Subtotal(103, tabela_aba_export_sap.ListColumns(1).DataBodyRange) > 0 And Dir(Pasta_Diaria & "\Clientes de " & analista & " sem e-mail cadastrado " & CStr(VBA.Format(VBA.Date, "dd.mm.yyyy")) & ".xlsx") = "" Then
            aba_export_sap.Range("A6:AL" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            Workbooks.Add
            Set arquivocriado = ActiveWorkbook
            arquivocriado.Sheets(1).Range("A1").PasteSpecial
            arquivocriado.Sheets(1).Columns("A:AL").AutoFit
            On Error Resume Next
            arquivocriado.SaveAs Pasta_Diaria & "\Clientes de " & analista & " sem e-mail cadastrado " & CStr(VBA.Format(VBA.Date, "dd.mm.yyyy")) & ".xlsx"
            arquivocriado.Close
            On Error GoTo 0
        End If
        If Application.WorksheetFunction.Subtotal(103, tabela_aba_export_sap.ListColumns(1).DataBodyRange) > 0 Then
'email_analista_sem_email_cadastrado
            Call email_para_analista_clientes_sem_email
        End If
        Call LimparFiltros
    Next i

' VERIFICACAO E ENVIO DE E-MAILS A TODOS OS ANALISTAS COM CLIENTES SEM MAPEIO
    Call LimparFiltros
    tabela_aba_export_sap.Range.AutoFilter Field:=32, Criteria1:="", Operator:=xlAnd
    
    If Application.WorksheetFunction.Subtotal(103, tabela_aba_export_sap.ListColumns(1).DataBodyRange) > 0 And Dir(Pasta_Diaria & "\Clientes sem analista mapeado " & VBA.Format(VBA.Date, "dd.mm.yyyy") & ".xlsx") = "" Then
        aba_export_sap.Range("A6:AL" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
        Workbooks.Add
        Set arquivocriado = ActiveWorkbook
        arquivocriado.Sheets(1).Range("A1").PasteSpecial
        arquivocriado.Sheets(1).Columns("A:AL").AutoFit
        On Error Resume Next
        arquivocriado.SaveAs Pasta_Diaria & "\Clientes sem analista mapeado " & CStr(VBA.Format(VBA.Date, "dd.mm.yyyy")) & ".xlsx"
        arquivocriado.Close
        On Error GoTo 0
    End If
    If Application.WorksheetFunction.Subtotal(103, tabela_aba_export_sap.ListColumns(1).DataBodyRange) > 0 Then
        Call email_para_analistas_clientes_sem_mapeio
    End If
    
    
' VERIFICACAO E ENVIO DE E-MAILS DE COBRANCA AOS CLIENTES MAPEADOS E COM EMAILS
    Call LimparFiltros
    caminho_manual_cliente = BuscarPasta("", False) & "\Manual do Cliente.pdf"
    For i = LBound(array_payers) To UBound(array_payers)
        cod_cliente = array_payers(i)
        If cod_cliente = "" Then
            Exit For
        End If
        nome_cliente = Application.WorksheetFunction.VLookup(cod_cliente, _
            aba_export_sap.Range("C7:D" & linha_fim), 2, False)
        iniciais_analista = Application.WorksheetFunction.VLookup(cod_cliente, _
            aba_export_sap.Range("C7:AF" & linha_fim), 30, False)
        iniciais_analista = ObterIniciais(VBA.Trim(iniciais_analista))
        
        Call LimparFiltros
        tabela_aba_export_sap.Range.AutoFilter Field:=3, Criteria1:=cod_cliente, Operator:=xlAnd
        tabela_aba_export_sap.Range.AutoFilter Field:=40, Criteria1:="SIM", Operator:=xlAnd
        
        If Application.WorksheetFunction.Subtotal(103, tabela_aba_export_sap.ListColumns(1).DataBodyRange) > 0 Then
            Set range_facturas = aba_export_sap.Range("A6:AL" & linha_fim).SpecialCells(xlCellTypeVisible)
            
            Call email_de_cobranca
        End If
        Call LimparFiltros
    Next i
    
    aba_export_sap.Activate
    aba_export_sap.Range("A6").Activate
    
    MsgBox "Processo Finalizado, os e-mails foram enviados"
    
    Application.DisplayAlerts = True

    
End Sub

Sub email_para_analista_clientes_sem_email()

    Call SetVarsEmails

    Destinatario = Application.WorksheetFunction.VLookup(analista, _
        aba_export_sap.Range("AF7:AG" & linha_fim), 2, False)
            
    If Destinatario = "" Then
        Exit Sub
    Else
    
        emailBody1 = aba_modelos_email.Range("B2").Value & analista & "<br><br>"
        emailBody2 = aba_modelos_email.Range("B3").Value & "<br><br>"
        emailBody3 = aba_modelos_email.Range("B4").Value & "<br><br>"
        emailBody4 = aba_modelos_email.Range("B5").Value & "<br>"
        emailBody5 = aba_modelos_email.Range("B6").Value
            
        On Error Resume Next
        Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
        On Error GoTo 0
    
        With MailItem
            .SentOnBehalfOfName = "CobrancaBR_B2B@electrolux.com"
            '.display
            .To = Destinatario
            .subject = "[IMPOSSÍVEL COBRAR] Clientes sem E-mail Cadastrado - " & analista
            .HTMLBody = emailBody1 & emailBody2 & emailBody3 & .HTMLBody
            .Attachments.Add Pasta_Diaria & "\Clientes de " & analista & " sem e-mail cadastrado " & VBA.Format(VBA.Date, "dd.mm.yyyy") & ".xlsx"
            .send
        End With
     End If
    
End Sub
 
Sub email_para_analistas_clientes_sem_mapeio()
 
    For i = LBound(array_analistas) To UBound(array_analistas)
       Destinatario = Destinatario & ";" & array_analistas(i)
    Next i
 
    emailBody1 = aba_modelos_email.Range("E2").Value & analista & "<br><br>"
    emailBody2 = aba_modelos_email.Range("E3").Value & "<br><br>"
    emailBody3 = aba_modelos_email.Range("E4").Value & "<br><br>"
    emailBody4 = aba_modelos_email.Range("E5").Value & "<br>"
    emailBody5 = aba_modelos_email.Range("E6").Value
    
    
    On Error Resume Next
    Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
    On Error GoTo 0
    
    With MailItem
        .SentOnBehalfOfName = "CobrancaBR_B2B@electrolux.com"
        '.display
        .To = Destinatario
        .subject = "[IMPOSSÍVEL COBRAR] Clientes não mapeados"
        .HTMLBody = emailBody1 & emailBody2 & emailBody3 & .HTMLBody
        .Attachments.Add Pasta_Diaria & "\Clientes sem analista mapeado " & VBA.Format(VBA.Date, "dd.mm.yyyy") & ".xlsx"
        .send
    End With
 
 
End Sub
 
Sub email_de_cobranca()

    Destinatario = Application.WorksheetFunction.VLookup(cod_cliente, _
        aba_export_sap.Range("C7:AH" & linha_fim), 32, False) & ";" & _
        Application.WorksheetFunction.VLookup(cod_cliente, _
        aba_export_sap.Range("C7:AI" & linha_fim), 33, False) & ";" & _
        Application.WorksheetFunction.VLookup(cod_cliente, _
        aba_export_sap.Range("C7:AJ" & linha_fim), 34, False) & ";" & _
        Application.WorksheetFunction.VLookup(cod_cliente, _
        aba_export_sap.Range("C7:AK" & linha_fim), 35, False) & ";" & _
        Application.WorksheetFunction.VLookup(cod_cliente, _
        aba_export_sap.Range("C7:AL" & linha_fim), 36, False) & ";"
    
    If Destinatario = ";;;;;" Then
        Exit Sub
    Else
        emailBody1 = "<b>" & aba_modelos_email.Range("H2").Value & "</b>"
        emailBody2 = "<b>" & aba_modelos_email.Range("H3").Value & "</b>"
        emailBody3 = "<b>Beneficiário: </b><br> ELECTROLUX DO BRASIL - CNPJ: 76.487.032.0001/25."
        emailBody4 = aba_modelos_email.Range("H5").Value
                        
        emailBody5 = "Caso você ainda não tenha acesso ao " & "<b>PORTAL DO CLIENTE</b>" & " segue em anexo o manual com o passo a passo para se cadastrar."
        emailBody6 = "<b> Link do PORTAL DO CLIENTE: </b><a href='URL' target=""_blank"">https://electrolux.simplificamais.com.br/</a>"

        emailBody7 = "<b>" & aba_modelos_email.Range("H8").Value & "</b><br>"
        
        On Error Resume Next
        Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
        On Error GoTo 0
        
        With MailItem
            .SentOnBehalfOfName = "CobrancaBR_B2B@electrolux.com"
            '.display
            .To = Destinatario
            .subject = "Títulos em Atraso Electrolux | " & cod_cliente & " | " & nome_cliente & "| Analista " & iniciais_analista
            .Attachments.Add caminho_manual_cliente
            .HTMLBody = emailBody1 & "<br>" & emailBody2 & "<br><br>" & emailBody3 & RangeToHTML(range_facturas) & "<br><br>" & emailBody4 & emailBody5 & emailBody6 & "<br><br>" & _
                emailBody7 & .HTMLBody
            .send
        End With

    End If
 
End Sub


Private Function SetVarsEmails()

  ' Crie uma instância do Outlook
    On Error Resume Next
    Set OutlookApp = GetObject("Outlook.Application")
    On Error GoTo 0
    
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If

    ' Obtenha o Namespace do Outlook
    Set ns = OutlookApp.GetNamespace("MAPI")

    ' Iterando através das pastas e encontre a pasta de entrada da caixa de correio compartilhada
    For Each caixa_do_b2b In ns.Folders
        Debug.Print caixa_do_b2b.Name
        If caixa_do_b2b.Name = "Contas a receber Brasil Electrolux" Or caixa_do_b2b.Name = "CobrancaBR_B2B@electrolux.com" Then
            Set caixa_do_b2b = caixa_do_b2b.Folders("Inbox")
            Exit For
        End If
        Set caixa_do_b2b = Nothing
    Next caixa_do_b2b
    
    If caixa_do_b2b Is Nothing Then
        MsgBox "Vocï¿½ nï¿½o tem acesso ao e-mail b2b.chile.otc@electrolux.com, por favor solicite ao TI o acesso e tente novamente.", vbOKOnly
        End
    End If
    

    
End Function

Function RangeToHTML(rng As Range) As String
    Dim fso As Object
    Dim ts As Object
    Dim TempFile, HTMLContent As String
    Dim TempWB As Workbook

    ' Cria um novo arquivo temporário
    TempFile = VBA.Environ$("temp") & "\TempHTMLFile.htm"

    ' Cria um novo workbook temporário
    Set TempWB = Workbooks.Add(1)
    rng.Copy
    With TempWB.Sheets(1)
        .Cells(1, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Cells(1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Columns("A:B").Delete
        .Columns("C:C").Delete
        .Columns("D:D").Delete
        .Columns("E:J").Delete
        .Columns("G:BB").Delete
        .Columns("F:F").NumberFormat = "$#,###.##"
        .Cells(1, 1).Value = "Código Cliente"
        .Cells(1, 2).Value = "Nome Cliente"
        .Cells(1, 3).Value = "NF"
        .Cells(1, 4).Value = "Parcela"
        .Cells(1, 5).Value = "Vencimento"
        .Cells(1, 6).Value = "Montante"
        .Columns.AutoFit
    End With

    ' Salva o workbook temporário como um arquivo HTML
    With TempWB.PublishObjects.Add(xlSourceRange, TempFile, TempWB.Sheets(1).Name, TempWB.Sheets(1).UsedRange.Address, xlHtmlStatic)
        .Publish (True)
    End With

    ' Fecha o workbook temporário
    TempWB.Close SaveChanges:=False

    ' Abre o arquivo HTML e lê o conteúdo
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    HTMLContent = ts.ReadAll
    ts.Close

    HTMLContent = Replace(HTMLContent, "align=center", "align=left")

    ' Retorna o HTML modificado
    RangeToHTML = HTMLContent

    ' Remove o arquivo temporário
    Kill TempFile

    ' Limpeza
    Set ts = Nothing
    Set fso = Nothing
End Function
Function ObterIniciais(nome As String) As String
    Dim partes() As String
    Dim iniciais As String
    Dim i As Integer
    
    partes = Split(nome, " ")
    
    ' Percorre cada parte do nome e pega a primeira letra
    For i = LBound(partes) To UBound(partes)
        If partes(i) <> "" Then ' Evita espaços extras
            iniciais = iniciais & UCase(Left(partes(i), 1))
        End If
    Next i
    
    ' Retorna as iniciais
    ObterIniciais = iniciais
End Function

