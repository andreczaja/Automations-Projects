Attribute VB_Name = "PASSO3_enviar_email_cobrança"
Option Explicit


    Private OutlookApp As Object
    Private OutlookMail, ObjFSO, ObjPasta, Arquivo, ns, caixa_do_b2b, MailItem As Object
    Private qtde_docs, qtde_faturas As Integer
    Private fileName, emailBody1, emailBody2, emailBody3, emailBody4, Destinatario, cod_cliente_, nome_cliente_, _
        clientes_sem_email, clientes_nao_mapeados, condicao_numero_faturas, faturas_pendentes, analista_e_kam, cliente_nao_mapeado_telefone, cliente_nao_mapeado_dicom_equifax As String
    Private aba_base_emails_, aba_cobravel_hoje_, aba_export_sap_, aba_facturas_email_formatado As Worksheet
    Private base_cobraveis As Workbook
    Private tbl As ListObject
    Private tb2 As ListObject
    Private range_facturas As Range

Sub enviar_email_cobranca_()



    frm_passos.Hide
    
    Set base_cobraveis = ThisWorkbook
    Set aba_export_sap_ = base_cobraveis.Sheets("Export SAP")
    Set tbl = aba_export_sap_.ListObjects("Export_FBL5N___Cobráveis")
    Set aba_cobravel_hoje_ = base_cobraveis.Sheets("Cobraveis HOJE")
    Set tbl2 = aba_cobravel_hoje_.ListObjects("Tabela_Cobraveis_HOJE")
    Set aba_base_emails_ = base_cobraveis.Sheets("Base E-mails")
    Set aba_facturas_email_formatado = base_cobraveis.Sheets("Facturas e-mail formatado")
    
    
    Application.ScreenUpdating = False

    ' Crie uma instância do Outlook
    On Error Resume Next
        Set OutlookApp = GetObject("Outlook.Application")
    On Error GoTo 0
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If

    ' Obtenha o Namespace do Outlook
    Set ns = OutlookApp.GetNamespace("MAPI")

    ' Itere através das pastas e encontre a pasta de entrada da caixa de correio compartilhada
    For Each caixa_do_b2b In ns.Folders
        If caixa_do_b2b.Name = "b2b.chile.otc@electrolux.com" Then
            Set caixa_do_b2b = caixa_do_b2b.Folders("Caixa de Entrada") ' Substitua "Caixa de Entrada" pelo nome da pasta de entrada em seu Outlook
            Exit For
        End If
    Next caixa_do_b2b
    
    clientes_sem_email = ""
    cliente_nao_mapeado_telefone = ""
    cliente_nao_mapeado_dicom_equifax = ""
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''ETAPA ENVIO DE E-MAILS PARA COBRANÇA POR E-MAIL''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'setando variaveis de intervalo de quantidade de arquivos
    aba_cobravel_hoje_.Activate
    aba_cobravel_hoje_.Range("a1").Activate
    On Error Resume Next
    aba_cobravel_hoje_.ShowAllData
    
    qtde_docs = aba_cobravel_hoje_.Range("BB9999").End(xlUp).Row
    
    linha = 2
    
Cobranca_Email:
    Do Until linha > qtde_docs
     
        ' Crie um novo item de e-mail na pasta especificada
        On Error Resume Next
        Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
        On Error GoTo sem_acesso_email_b2b
        
        faturas_pendentes = ""
        cod_cliente_ = aba_cobravel_hoje_.Range("BB" & linha).Value
        condicao_numero_faturas = aba_cobravel_hoje_.Range("BC" & linha).Value
        nome_cliente_ = aba_cobravel_hoje_.Range("BD" & linha).Value
        
        'enviando e-mails de acordo com a situacao do cliente, se for mais de 10 faturas, envia o anexo,
        ' senao, envia a tabela com as faturas no corpo do e-mail
        
        
       If nome_cliente_ = "Cliente não mapeado" Or nome_cliente_ = "-" Or nome_cliente_ = "" Then
            clientes_nao_mapeados = clientes_nao_mapeados & " - " & cod_cliente_
            linha = linha + 1
            GoTo Cobranca_Email
        End If
        
        If condicao_numero_faturas = "Mais de 10 faturas" Then
        
            aba_cobravel_hoje_.Activate
            aba_cobravel_hoje_.Range("a1").Activate
            On Error Resume Next
            aba_cobravel_hoje_.ShowAllData
            
            tbl2.Range.AutoFilter Field:=2, Criteria1:=cod_cliente_
            tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por E-mail"
            
            linha_fim = aba_cobravel_hoje_.Range("A99999").End(xlUp).Row
            aba_facturas_email_formatado.Range("A:AA").ClearContents
            aba_cobravel_hoje_.Range("D1:Q" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            aba_facturas_email_formatado.Range("A1").PasteSpecial
            aba_facturas_email_formatado.Columns("A:E").Replace What:="FAE0", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            aba_facturas_email_formatado.Columns("A:E").Replace What:="NCE000", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            aba_facturas_email_formatado.Columns("A:E").Replace What:="NCE00", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            
            faturas_pendentes = ""
            qtde_faturas = aba_facturas_email_formatado.Range("A99999").End(xlUp).Row
            For i = 2 To qtde_faturas
                faturas_pendentes = faturas_pendentes & " - " & aba_facturas_email_formatado.Range("B" & i).Value
            Next i
                
            On Error Resume Next
            Destinatario = Application.WorksheetFunction.VLookup(cod_cliente_, _
                aba_base_emails_.Range("A2:E999999"), 4, False) & ";" & _
                Application.WorksheetFunction.VLookup(cod_cliente_, _
                aba_base_emails_.Range("A2:E999999"), 5, False)
            analista_e_kam = Application.WorksheetFunction.VLookup(cod_cliente_, _
                aba_base_emails_.Range("A2:G999999"), 6, False) & ";" & _
                Application.WorksheetFunction.VLookup(cod_cliente_, _
                aba_base_emails_.Range("A2:G999999"), 7, False)
            On Error GoTo Erro_Falta_Email_Cobranca_Email
                    
            If Destinatario = ";;" Or Destinatario = "" Or Destinatario = ";" Or Destinatario = "-" Or Destinatario = "--" Then
                GoTo Erro_Falta_Email_Cobranca_Email
            Else
            
                emailBody1 = "Estimados de " & nome_cliente_ & ", Buenas Tardes! " & "<br><br>" & _
                        "Esperamos encontrarles bien." & "<br>"
                emailBody2 = "Según lo verificado, la(s) factura(s) " & faturas_pendentes & _
                        " se encuentran vencidas en nuestro sistema."
                emailBody3 = "Adjunto archivo con el listado de las facturas mencionadas." & "<br><br>" & _
                "Solicitamos amablemente su respuesta, con el estado de pago del título vencido." & "<br><br>" & "Desde ya la gracias,"
                
                With MailItem
                    MailItem.display
                    .To = Destinatario
                    .CC = analista_e_kam
                    .subject = "DEMOSTRATIVO ELECTROLUX DE CHILE | " & cod_cliente_ & " - " & nome_cliente_
                    .HTMLBody = emailBody1 & "<br><br>" & emailBody2 & "<br><br>" & emailBody3 & .HTMLBody
                    .Attachments.Add Folder & cod_cliente_ & " - Facturas pendientes " & nome_cliente_ & " " & Day(Date) & "." & Month(Date) & "." & Year(Date) & ".xlsx"
                    '.send
                End With
                
                linha = linha + 1
                aba_facturas_email_formatado.Range("A:AA").ClearContents
                
            End If
                    
        ElseIf condicao_numero_faturas = "Menos de 10 faturas" Then
        
            aba_cobravel_hoje_.Activate
            aba_cobravel_hoje_.Range("a1").Activate
            On Error Resume Next
            aba_cobravel_hoje_.ShowAllData
        
            tbl2.Range.AutoFilter Field:=2, Criteria1:=cod_cliente_
            tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por E-mail"
            
            linha_fim = aba_cobravel_hoje_.Range("A99999").End(xlUp).Row
            aba_facturas_email_formatado.Range("A:AA").ClearContents
            aba_cobravel_hoje_.Range("D1:Q" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            aba_facturas_email_formatado.Range("a1").PasteSpecial
            linha_fim = aba_facturas_email_formatado.Range("A99999").End(xlUp).Row
            aba_facturas_email_formatado.Columns("C:E").EntireColumn.Delete
            aba_facturas_email_formatado.Columns("E:J").EntireColumn.Delete
            aba_facturas_email_formatado.Columns("A:E").Replace What:="FAE0", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            aba_facturas_email_formatado.Columns("A:E").Replace What:="NCE000", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            aba_facturas_email_formatado.Columns("A:E").Replace What:="NCE00", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

            
            faturas_pendentes = ""
            qtde_faturas = aba_facturas_email_formatado.Range("A99999").End(xlUp).Row
            For i = 2 To qtde_faturas
                faturas_pendentes = faturas_pendentes & " - " & aba_facturas_email_formatado.Range("B" & i).Value
            Next i
             
            Set range_facturas = aba_facturas_email_formatado.Range("A1:E" & linha_fim)
            
            
            
                On Error Resume Next
                Destinatario = Application.WorksheetFunction.VLookup(cod_cliente_, _
                    aba_base_emails_.Range("A2:G999999"), 4, False) & ";" & _
                    Application.WorksheetFunction.VLookup(cod_cliente_, _
                    aba_base_emails_.Range("A2:G999999"), 5, False)
                analista_e_kam = Application.WorksheetFunction.VLookup(cod_cliente_, _
                    aba_base_emails_.Range("A2:G999999"), 6, False) & ";" & _
                    Application.WorksheetFunction.VLookup(cod_cliente_, _
                    aba_base_emails_.Range("A2:G999999"), 7, False)
                On Error GoTo Erro_Falta_Email_Cobranca_Email
                
                If Destinatario = ";;" Or Destinatario = "" Or Destinatario = ";" Or Destinatario = "-" Or Destinatario = "--" Then
                    GoTo Erro_Falta_Email_Cobranca_Email
                Else
                    emailBody1 = "Estimados de " & nome_cliente_ & ", Buenas Tardes! " & "<br><br>" & _
                            "Esperamos encontrarles bien." & "<br>"
                    emailBody2 = "Según lo verificado, la(s) factura(s) " & faturas_pendentes & _
                            " se encuentran vencidas en nuestro sistema." & "<br>"
                    emailBody3 = "Solicitamos amablemente su respuesta, con el estado de pago del título vencido." & "<br><br>" & "Desde ya la gracias,"
                    
                    With MailItem
                    
                        MailItem.display
                        .To = Destinatario
                        .CC = analista_e_kam
                        .subject = "DEMOSTRATIVO ELECTROLUX DE CHILE | " & cod_cliente_ & " - " & nome_cliente_
                        .HTMLBody = emailBody1 & "<br><br>" & emailBody2 & "<br><br>" & RangeToHTML(range_facturas) & "<br><br>" & emailBody3 & .HTMLBody
                        '.send
                    End With
                    
                End If
                linha = linha + 1
                
        End If
        
    Loop
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''ETAPA ENVIO DE E-MAILS PARA COBRANÇA POR TELEFONE''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    aba_cobravel_hoje_.Activate
    aba_cobravel_hoje_.Range("a1").Activate
    On Error Resume Next
    aba_cobravel_hoje_.ShowAllData
    
    qtde_docs = aba_cobravel_hoje_.Range("BE9999").End(xlUp).Row

    linha = 2
Cobranca_Telefone:
    Do Until linha > qtde_docs

        ' Crie um novo item de e-mail na pasta especificada
        On Error Resume Next
        Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
        On Error GoTo sem_acesso_email_b2b
        
        analista_responsavel_cobranca_telefone = aba_cobravel_hoje_.Range("BE" & linha).Value
        
        
        
        If analista_responsavel_cobranca_telefone <> "-" Then
                On Error Resume Next
                Destinatario = Application.WorksheetFunction.VLookup(analista_responsavel_cobranca_telefone, _
                    aba_base_emails_.Range("C2:G999999"), 5, False)

                    emailBody1 = "Olá, " & analista_responsavel_cobranca_telefone & "<br><br>" & _
                            "Segue arquivo em anexo com as cobranças por realizar via contato telefônico referentes ao dia de hoje (" & Day(Date) & "." & Month(Date) & "." & Year(Date) & ")." & _
                            "<br>"
                    emailBody2 = "Atenciosamente,"
                    
                    With MailItem
                    
                        MailItem.display
                        .To = Destinatario
                        .subject = "Faturas por Cobrar - Electrolux de Chile | " & Day(Date) & "." & Month(Date) & "." & Year(Date)
                        .HTMLBody = emailBody1 & "<br>" & emailBody2 & .HTMLBody
                        .Attachments.Add Folder & "Facturas por cobrar - " & analista_responsavel_cobranca_telefone & " - " & Day(Date) & "." & Month(Date) & "." & Year(Date) & ".xlsx"
                        '.send
                    End With
                    
        End If
                
            linha = linha + 1

    Loop
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''ETAPA ENVIO DE E-MAILS DICOM/EQUIFAX''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
        'modulo que ira enviar os emails para os analistas referente a faturas em estado de
    'dicom ou equifax
    

     'setando variaveis de intervalo de quantidade de arquivos
     
    aba_cobravel_hoje_.Activate
    aba_cobravel_hoje_.Range("a1").Activate
    On Error Resume Next
    aba_cobravel_hoje_.ShowAllData

    
    qtde_docs = aba_cobravel_hoje_.Range("BH9999").End(xlUp).Row

    linha = 2
    
Dicom_Equifax:
    Do Until linha > qtde_docs

        ' Crie um novo item de e-mail na pasta especificada
        On Error Resume Next
        Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
        On Error GoTo sem_acesso_email_b2b
        
        analista_responsavel_dicom_equifax = aba_cobravel_hoje_.Range("BH" & linha).Value
        
        
        
        If analista_responsavel_cobranca_telefone <> "-" Then
                On Error Resume Next
                Destinatario = Application.WorksheetFunction.VLookup(analista_responsavel_dicom_equifax, _
                    aba_base_emails_.Range("C2:G999999"), 5, False)

                    emailBody1 = "Olá, " & analista_responsavel_dicom_equifax & "<br><br>" & _
                            "Segue arquivo em anexo com faturas que hoje (" & Day(Date) & "." & Month(Date) & "." & Year(Date) & ") " & _
                            "estão em condição de Dicom/Equifax, ou seja, já passou-se mais de 60 dias da data de vencimento, e mesmo após as cobranças " & _
                            "o pagamento não foi realizado." & vbNewLine & vbNewLine & "Favor entrar em contato com o responsável para publicar a dívida do cliente no" & _
                            " portal do Dicom/Equifax." & _
                            "<br>"
                    emailBody2 = "Atenciosamente,"
                    
                    With MailItem
                    
                        MailItem.display
                        .To = Destinatario
                        .subject = "Facturas em estado de Dicom/Equifax - Electrolux de Chile | " & Day(Date) & "." & Month(Date) & "." & Year(Date)
                        .HTMLBody = emailBody1 & "<br>" & emailBody2 & .HTMLBody
                        .Attachments.Add Folder & "Facturas en condicion de Dicom.Equifax - cliente(s) de " & analista_responsavel_dicom_equifax & " - " & Day(Date) & "." & Month(Date) & "." & Year(Date) & ".xlsx"
                        '.send
                    End With
                    
        End If
                
            linha = linha + 1

    Loop
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''ETAPA ENVIO DE E-MAILS COBRANÇA PREVENTIVA''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      'setando variaveis de intervalo de quantidade de arquivos
    
    If aba_cobravel_hoje_.Range("BK2").Value = "" Then
        GoTo fim
    End If
    qtde_docs = aba_cobravel_hoje_.Range("BK9999").End(xlUp).Row
    
    linha = 2
    
    
    
Cobranca_Preventiva:
    Do Until linha > qtde_docs
    
                        
            On Error Resume Next
            Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
            On Error GoTo sem_acesso_email_b2b

            aba_cobravel_hoje_.Activate
            aba_cobravel_hoje_.Range("a1").Activate
            On Error Resume Next
            aba_cobravel_hoje_.ShowAllData
            
            cod_cliente_ = aba_cobravel_hoje_.Range("BK" & linha).Value
            nome_cliente_ = aba_cobravel_hoje_.Range("BL" & linha).Value
        
            tbl2.Range.AutoFilter Field:=2, Criteria1:=cod_cliente_
            tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança Preventiva - Constructoras"
            
            linha_fim = aba_cobravel_hoje_.Range("A99999").End(xlUp).Row
            aba_facturas_email_formatado.Range("A:AA").ClearContents
            aba_cobravel_hoje_.Range("D1:Q" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            aba_facturas_email_formatado.Range("a1").PasteSpecial
            linha_fim = aba_facturas_email_formatado.Range("A99999").End(xlUp).Row
            aba_facturas_email_formatado.Columns("C:E").EntireColumn.Delete
            aba_facturas_email_formatado.Columns("E:J").EntireColumn.Delete
            aba_facturas_email_formatado.Columns("A:E").Replace What:="FAE0", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            aba_facturas_email_formatado.Columns("A:E").Replace What:="NCE000", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            aba_facturas_email_formatado.Columns("A:E").Replace What:="NCE00", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            
            faturas_pendentes = ""
            
            qtde_faturas = aba_facturas_email_formatado.Range("A99999").End(xlUp).Row
            For i = 2 To qtde_faturas
                faturas_pendentes = faturas_pendentes & " - " & aba_facturas_email_formatado.Range("B" & i).Value
            Next i
             
            Set range_facturas = aba_facturas_email_formatado.Range("A1:E" & linha_fim)
            
            
            
                On Error Resume Next
                Destinatario = Application.WorksheetFunction.VLookup(cod_cliente_, _
                    aba_base_emails_.Range("A2:G999999"), 4, False) & ";" & _
                    Application.WorksheetFunction.VLookup(cod_cliente_, _
                    aba_base_emails_.Range("A2:G999999"), 5, False)
                analista_e_kam = Application.WorksheetFunction.VLookup(cod_cliente_, _
                    aba_base_emails_.Range("A2:G999999"), 6, False) & ";" & _
                    Application.WorksheetFunction.VLookup(cod_cliente_, _
                    aba_base_emails_.Range("A2:G999999"), 7, False)
                On Error GoTo Erro_Falta_Email_Cobranca_Preventiva
                
                If Destinatario = ";;" Or Destinatario = "" Or Destinatario = ";" Or Destinatario = "-" Or Destinatario = "--" Then
                    GoTo Erro_Falta_Email_Cobranca_Preventiva
                Else
                    emailBody1 = "Estimados de " & nome_cliente_ & ", Buenas Tardes! " & "<br>" & _
                            "Esperamos encontrarles bien." & "<br>"
                    emailBody2 = "Este correo electrónico es para informarle que su(s) factura(s) " & faturas_pendentes & " vence(n) luego:"
                    emailBody3 = "En caso de dudas o problemas para realizar el pago, comuníquese con nosotros responda a este correo electrónico, que estaremos encantados de atenderle."
                    emailBody4 = "<p><b>COMUNICADO IMPORTANTE:</b>.</p>" & "Para evitar incumplimientos y bloqueos en las liberaciones de pedidos, por favor, siga con el pago en la fecha indicada." & "<br><br>" & "Saludos,"

                    
                    With MailItem
                    
                        MailItem.display
                        .To = Destinatario
                        .CC = analista_e_kam
                        .subject = "DEMOSTRATIVO ELECTROLUX DE CHILE | " & cod_cliente_ & " - " & nome_cliente_
                        .HTMLBody = emailBody1 & "<br>" & emailBody2 & "<br>" & RangeToHTML(range_facturas) & "<br>" & emailBody3 & "<br>" & emailBody4 & .HTMLBody
                        '.send
                    End With
                    
                End If
                linha = linha + 1
                
        
    Loop
    
    GoTo fim
    

Erro_Falta_Email_Cobranca_Email:
    clientes_sem_email = " - " & clientes_sem_email & " (" & cod_cliente_ & "-" & nome_cliente_ & ")"
    linha = linha + 1
GoTo Cobranca_Email

Erro_Falta_Email_Cobranca_Telefone:
    clientes_sem_email = " - " & clientes_sem_email & " (" & cod_cliente_ & "-" & nome_cliente_ & ")"
    linha = linha + 1
GoTo Cobranca_Telefone

Erro_Falta_Email_Dicom_Equifax:
    clientes_sem_email = " - " & clientes_sem_email & " (" & cod_cliente_ & "-" & nome_cliente_ & ")"
    linha = linha + 1
GoTo Dicom_Equifax

Erro_Falta_Email_Cobranca_Preventiva:
    clientes_sem_email = " - " & clientes_sem_email & " (" & cod_cliente_ & "-" & nome_cliente_ & ")"
    linha = linha + 1
GoTo Cobranca_Preventiva

sem_acesso_email_b2b:

MsgBox "Você não tem acesso ao e-mail b2b.chile.otc@electrolux.com, por favor solicite ao TI o acesso e tente novamente.", vbOKOnly


fim:

    linha = 2
    linha_fim = aba_cobravel_hoje_.Range("BF99999").End(xlUp).Row
    
    If linha_fim <> 1 Then
        Do Until linha > linha_fim
            cliente_nao_mapeado_telefone = " - " & cliente_nao_mapeado_telefone & "(" & aba_cobravel_hoje_.Range("BF" & linha).Value & "-" & aba_cobravel_hoje_.Range("BG" & linha).Value & ")"
            linha = linha + 1
        Loop
        cliente_nao_mapeado_telefone = Mid(cliente_nao_mapeado_telefone, 3)
    End If
        
    linha = 2
    linha_fim = aba_cobravel_hoje_.Range("BI99999").End(xlUp).Row
    
    cliente_nao_mapeado_dicom_equifax = ""
    
    If linha_fim <> 1 Then
        Do Until linha > linha_fim
            cliente_nao_mapeado_dicom_equifax = cliente_nao_mapeado_dicom_equifax & "(" & aba_cobravel_hoje_.Range("BI" & linha).Value & "-" & aba_cobravel_hoje_.Range("BJ" & linha).Value & ")"
            linha = linha + 1
        Loop
        cliente_nao_mapeado_dicom_equifax = Mid(cliente_nao_mapeado_dicom_equifax, 3)
    End If
    
    
    
        'diferentes msgbox dependendo das condicoes percorridas no codigo
        
    If clientes_sem_email <> "" Then
        clientes_sem_email = Mid(clientes_sem_email, 4)
    End If

Dim mensagem_ As String

mensagem_ = "Processo Concluído!" & vbNewLine & vbNewLine

If clientes_sem_email <> "" Then
    mensagem_ = mensagem_ & "Não foram encontrados os e-mails do(s) cliente(s) " & clientes_sem_email & " na aba Base E-mails." & vbNewLine & vbNewLine
End If

If cliente_nao_mapeado_telefone <> "" Then
    mensagem_ = mensagem_ & "Os clientes com faturas a cobrar por telefone não possuem analistas responsáveis: " & cliente_nao_mapeado_telefone & vbNewLine & vbNewLine
End If

If cliente_nao_mapeado_dicom_equifax <> "" Then
    mensagem_ = mensagem_ & "Existem faturas a serem publicadas em Dicom/Equifax sem analistas responsáveis: " & cliente_nao_mapeado_dicom_equifax & vbNewLine & vbNewLine
End If

If Right(mensagem_, 4) = vbNewLine & vbNewLine Then
    mensagem_ = Left(mensagem_, Len(mensagem_) - 4)
End If

MsgBox mensagem_, vbOKOnly


    ' Limpeza
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    Set ObjFSO = Nothing
    Set ObjPasta = Nothing
    Set Arquivo = Nothing
    Set caixa_do_b2b = Nothing
    Set MailItem = Nothing
    Set ns = Nothing
    
    
End Sub

Function RangeToHTML(rng As Range) As String
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    ' Cria um novo arquivo temporário
    TempFile = Environ$("temp") & "\TempHTMLFile.htm"

    ' Cria um novo workbook temporário
    Set TempWB = Workbooks.Add(1)
    rng.Copy
    With TempWB.Sheets(1)
        .Cells(1, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        .Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Cells(1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Columns.AutoFit
        .Columns("C:C").NumberFormat = "dd.mm.yyyy"
        .Columns("D:D").NumberFormat = "#,##0"
        .Cells(1, 4).Value = "Monto"
    End With

    qtde_facturas = TempWB.Sheets(1).Range("A1").End(xlDown).Row

    For i = 2 To qtde_facturas
        If TempWB.Sheets(1).Range("E" & i).Value Like "CLT*" Then
            TempWB.Sheets(1).Range("E" & i).ClearContents
        End If
    Next i

    ' Salva o workbook temporário como um arquivo HTML
    With TempWB.PublishObjects.Add(xlSourceRange, TempFile, TempWB.Sheets(1).Name, TempWB.Sheets(1).UsedRange.Address, xlHtmlStatic)
        .Publish (True)
    End With

    ' Fecha o workbook temporário
    TempWB.Close SaveChanges:=False

    ' Abre o arquivo HTML e lê o conteúdo
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangeToHTML = ts.ReadAll
    ts.Close
    RangeToHTML = Replace(RangeToHTML, "align=center", "align=left")
    Set ts = Nothing
    Set fso = Nothing

    ' Remove o arquivo temporário
    Kill TempFile
End Function
