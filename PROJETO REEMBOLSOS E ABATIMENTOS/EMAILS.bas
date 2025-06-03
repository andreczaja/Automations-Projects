Attribute VB_Name = "EMAILS"
Private arquivocriado As Workbook
Private emails_clientes, email_analista, emailbody1, emailBody2, emailBody3, emailBody4, emailBody5, emailBody6, emailBody7, _
emailBody8, emailBody9, emailBody10, Destinatario, Copia, nome_cliente, status As String
Private OutlookApp As Object
Private range_linhas As Range
Private OutlookMail, objFSO, ObjPasta, arquivo, ns, caixa_do_b2b, MailItem As Object
Public array_analistas() As Variant
Private qtde_linhas_filtradas As Long


Sub emails_etapa_1()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' declaracao comum e set das vars exclusivas para manipulação do outlook
    Call declaracao_vars

    Call SetVarsEmails
    
    ' atualizacao para assegurar que as infos que serão utilizadas no e-mail estão efetivamente atualizadas
    tabela_aba_plan_distribuicao.QueryTable.Refresh False
    
    Call LimparFiltros(tabela_aba_plan_distribuicao)
    
    
    ' verifica se a data do agrupado de pagamento que foi preenchida na etapa ellevo via form está na celula BC1 aba reembolsos aprovados, se não,
    ' abre uma input box para o usuario preencher manualmente (ISSO EXISTE APENAS POR PRECAUÇÃO)
    If aba_reembolsos_aprovados.Range("BC1").Value = "" Or aba_reembolsos_aprovados.Range("BC1").Value = ".." Then
        data_agrupado_pagamento = InputBox("A data do agrupado de pagamento não foi encontrada. Por favor digite-a abaixo no formato 'DD/MM/AAAA'")
    Else
        data_agrupado_pagamento = aba_reembolsos_aprovados.Range("BC1").Value
    End If
    
    ' verificando se a base de reembolsos aprovados está vazia, se sim, não irá seguir a automação
    linha_fim = aba_reembolsos_aprovados.Range("A1048576").End(xlUp).Row
    If linha_fim = 2 And aba_reembolsos_aprovados.Range("A2").Value = "" Then
        MsgBox "Nenhum e-mail de notificação de reembolso a ser enviado.", vbOKOnly
        Exit Sub
    End If
    
    array_payers_reembolsos_com_dados_bancarios = Array()
    array_payers_reembolsos_sem_dados_bancarios = Array()
    array_payers_abatimento = Array()
    linha_fim_aba_reembolsos_pendentes = aba_reembolsos_pendentes.Range("A1048576").End(xlUp).Row
    ' preenchendo o array reembolsos com dados bancarios
    For linha = 2 To linha_fim_aba_reembolsos_pendentes
        payer_atual = aba_reembolsos_pendentes.Range("C" & linha).Value
        status = aba_reembolsos_pendentes.Range("AE" & linha).Value
        
        If payer_atual <> "" And status = "Ellevo Criado" And Not PayerDuplicado Then
            ' CASO DE REEMBOLSO de payer com dados bancarios
            Call Add_ao_Array(array_payers_reembolsos_com_dados_bancarios, payer_atual)
        End If
    Next linha
    
    ' itera pelo array e envia um e-mail para cada um dos clientes desse array
    For i = LBound(array_payers_reembolsos_com_dados_bancarios) To UBound(array_payers_reembolsos_com_dados_bancarios)
        payer_atual = array_payers_reembolsos_com_dados_bancarios(i)
        Call LimparFiltros(tabela_reembolsos_aprovados)
        Call email_reembolsos
    Next i
    
    Call limpar_base_reembolsos_pendentes
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "E-mails enviados"
    
End Sub
Sub email_reembolsos()

    ' junta numa string com separador ';' todas as 5 colunas de e-mail de determinado cliente conforme plan distribuicao
    Destinatario = Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:L"), 12, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:M"), 13, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:N"), 14, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:O"), 15, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:P"), 16, False)
    
    ' se não encontrar destinatario, pula pro proximo cliente
    If Destinatario = ";;;;" Then
        Destinatario = Application.WorksheetFunction.VLookup(payer_atual, aba_plan_distribuicao.Columns("A:Q"), 17, False)
    End If
    
    ' valor total de reembolsos de determinado cliente conforme plan de reemb aprovados
    Call LimparFiltros(tabela_aba_fbl5n_credito_devolucao)
    valor = Round(Application.WorksheetFunction.SumIf(aba_reembolsos_pendentes.Columns("C:C"), payer_atual, aba_reembolsos_pendentes.Columns("P:P")) * -1, 2)
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=3, Criteria1:=payer_atual
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=31, Criteria1:="Ellevo Criado"
    ' filtra pelo cliente e pega o range, pois é esse range que será colado formatado no corpo do e-mail do cliente
    Set range_linhas = aba_reembolsos_pendentes.Range("A1:AB" & linha_fim_aba_reembolsos_pendentes).SpecialCells(xlCellTypeVisible)
    nome_cliente = Application.WorksheetFunction.VLookup(payer_atual, aba_plan_distribuicao.Columns("A:B"), 2, False)
    
    emailbody1 = aba_modelos_de_emails.Range("H3").Value & "<br><br>"
    emailBody2 = aba_modelos_de_emails.Range("H4").Value
    emailBody3 = aba_modelos_de_emails.Range("H5").Value & "<br><br>"
    emailBody4 = aba_modelos_de_emails.Range("H6").Value
    emailBody5 = aba_modelos_de_emails.Range("H7").Value & "<br><br>"
    emailBody6 = aba_modelos_de_emails.Range("H8").Value & "<br>"
    emailBody7 = aba_modelos_de_emails.Range("H9").Value
        
    On Error Resume Next
    Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
    On Error GoTo 0

    With MailItem
        ' forçando o envio da caixa de Cobranca Brasil Electrolux
        .SentOnBehalfOfName = "CobrancaBR_B2B@electrolux.com"
        .HTMLBody = ""
        .To = Destinatario
        .subject = aba_modelos_de_emails.Range("H2").Value & " - " & payer_atual & " - " & nome_cliente
        .HTMLBody = emailbody1 & emailBody2 & "<b>" & valor & "</b>" & emailBody3 & RangeToHTML(range_linhas, "REEMBOLSO") & "<br><br>" & emailBody4 & _
            data_agrupado_pagamento & emailBody5 & emailBody6 & emailBody7 & .HTMLBody
        .send
    End With
    
End Sub
Sub limpar_base_reembolsos_pendentes()

    Call LimparFiltros(tabela_aba_reembolsos_pendentes)
    tabela_aba_reembolsos_pendentes.Range.AutoFilter Field:=31, Criteria1:="Ellevo Criado"
    aba_reembolsos_pendentes.Range("A2:AE" & linha_fim_aba_reembolsos_pendentes).Select
    aba_reembolsos_pendentes.Range("A2:AE" & linha_fim_aba_reembolsos_pendentes).SpecialCells(xlCellTypeVisible).Delete
    Call LimparFiltros(tabela_aba_reembolsos_pendentes)

End Sub
Sub emails_etapa_2()
Attribute emails_etapa_2.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Form_Emails.Show
    Call declaracao_vars
    Call SetVarsEmails
    
    tabela_aba_plan_distribuicao.QueryTable.Refresh False
    
    Call LimparFiltros(tabela_aba_plan_distribuicao)
    array_analistas = Array()
    linha_fim = aba_plan_distribuicao.Range("A1048576").End(xlUp).Row
    
    ' preenchendo o array de analistas
    For linha = 2 To linha_fim
        analista = aba_plan_distribuicao.Range("Q" & linha).Value
        If analista <> "" Then
            If Not UBound(VBA.Filter(array_analistas, analista)) >= 0 Then
                Call Add_ao_Array(array_analistas, analista)
            End If
        End If
    Next linha
    
    array_payers_reembolsos_sem_dados_bancarios = Array()
    array_payers_reembolsos_com_dados_bancarios = Array()
    array_payers_abatimento = Array()
    
    ' preenchendo o array reembolsos com dados bancarios, reembolsos sem dados bancarios e abatimentos
    linha_fim = aba_fbl5n_credito_devolucao.Range("A1048576").End(xlUp).Row
    For linha = 2 To linha_fim
        payer_atual = aba_fbl5n_credito_devolucao.Range("C" & linha).Value
        
        If Not PayerDuplicado Then
            soma_cred_dev = Application.WorksheetFunction.SumIf(aba_fbl5n_credito_devolucao.Columns("C:C"), payer_atual, aba_fbl5n_credito_devolucao.Columns("P:P"))
            soma_debito_AR_chave_ref_3_vazia = Application.WorksheetFunction.SumIfs(aba_fbl5n_AR.Columns("P:P"), aba_fbl5n_AR.Columns("C:C"), payer_atual, aba_fbl5n_AR.Columns("AB:AB"), "")
            soma_debito_AR = Application.WorksheetFunction.SumIf(aba_fbl5n_AR.Columns("C:C"), payer_atual, aba_fbl5n_AR.Columns("P:P")) - soma_debito_AR_chave_ref_3_vazia
            condicao = soma_debito_AR + soma_cred_dev
  
            If condicao < 0 And Application.WorksheetFunction.CountIf(aba_dados_bancarios.Columns("B:B"), payer_atual) > 0 Then
                ' Condição de reembolsos de payer sem dados bancarios
                Call Add_ao_Array(array_payers_reembolsos_sem_dados_bancarios, payer_atual)
            ElseIf condicao > 0 And Application.WorksheetFunction.CountIf(aba_titulos_a_abater.Columns("A:A"), payer_atual) > 0 Then
                ' CASO DE ABATIMENTO
                Call Add_ao_Array(array_payers_abatimento, payer_atual)
            End If
        End If
    Next linha
    
    ' VERIFICACAO E ENVIO DE E-MAILS A ANALISTAS COM REEMBOLSOS DE CLIENTES PENDENTES POR CONTA DE FALTA DE DADOS BANCARIOS CADASTRADOS NO SAP
    If Form_Emails.opt_ambos Or Form_Emails.opt_apenas_reembolsos Then
        For i = LBound(array_analistas) To UBound(array_analistas)
            email_analista = array_analistas(i)
            On Error Resume Next
            aba_dados_bancarios.ShowAllData
            On Error GoTo 0
            linha_fim = aba_dados_bancarios.Range("B1048576").End(xlUp).Row
            Debug.Print aba_dados_bancarios.Range("D" & i2).Value
            For i2 = 2 To linha_fim
                Debug.Print Application.WorksheetFunction.VLookup(aba_dados_bancarios.Range("D" & i2).Value, aba_plan_distribuicao.Columns("E:Q"), 13, False)
                If Application.WorksheetFunction.VLookup(aba_dados_bancarios.Range("D" & i2).Value, aba_plan_distribuicao.Columns("E:Q"), 13, False) = email_analista Then
                    analista_responsavel = aba_dados_bancarios.Range("D" & i2).Value
                    Exit For
                End If
            Next i2
            
            'analista_responsavel = Application.WorksheetFunction.VLookup(aba_dados_bancarios.Columns("D:D"), analista_responsavel)
            If Application.WorksheetFunction.CountIf(aba_dados_bancarios.Columns("D:D"), analista_responsavel) > 0 Then
                Call email_para_analistas_clientes_com_reembolsos_pendentes
            End If
        Next i
        
        On Error Resume Next
        aba_dados_bancarios.ShowAllData
        On Error GoTo 0
        
        For i = LBound(array_payers_reembolsos_sem_dados_bancarios) To UBound(array_payers_reembolsos_sem_dados_bancarios)
            payer_atual = array_payers_reembolsos_sem_dados_bancarios(i)
            Call LimparFiltros(tabela_aba_fbl5n_credito_devolucao)
            
            If Application.WorksheetFunction.CountIf(aba_fbl5n_credito_devolucao.Columns("C:C"), payer_atual) > 0 Then
                Call email_para_clientes_sem_dados_bancarios
            End If
            
        Next i
    End If
    
    If Form_Emails.opt_ambos Or Form_Emails.opt_apenas_abatimentos Then
        Call LimparFiltros(tabela_aba_fbl5n_credito_devolucao)
        linha_aba_titulos_a_abater = aba_titulos_a_abater.Range("A1048576").End(xlUp).Row
        
        For i = LBound(array_payers_abatimento) To UBound(array_payers_abatimento)
            payer_atual = array_payers_abatimento(i)
            Call LimparFiltros(tabela_titulos_a_abater)
            
            If Application.WorksheetFunction.CountIf(aba_titulos_a_abater.Columns("A:A"), payer_atual) > 0 Then
                Call email_abatimento
            End If
            
        Next i
    End If

    Call LimparFiltros(tabela_aba_fbl5n_credito_devolucao)
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "E-mails enviados"

    
End Sub

Sub email_para_analistas_clientes_com_reembolsos_pendentes()

    Call declaracao_vars
    Destinatario = Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Range("A:Q"), 17, False)
    Destinatario = email_analista
    
    If Destinatario = "" Then
        Exit Sub
    Else
        linha_fim = aba_dados_bancarios.Range("B1048576").End(xlUp).Row
        aba_dados_bancarios.Range("A:D").AutoFilter Field:=4, Criteria1:=analista_responsavel
        Set range_linhas = aba_dados_bancarios.Range("B1:D" & linha_fim).SpecialCells(xlCellTypeVisible)
        emailbody1 = aba_modelos_de_emails.Range("B3").Value & "<br><br>"
        emailBody2 = aba_modelos_de_emails.Range("B4").Value & "<br>"
        emailBody3 = aba_modelos_de_emails.Range("B5").Value & "<br>"
        emailBody4 = aba_modelos_de_emails.Range("B6").Value & "<br>"
        emailBody5 = aba_modelos_de_emails.Range("B7").Value
            
        On Error Resume Next
        Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
        On Error GoTo 0
    
        With MailItem
            .SentOnBehalfOfName = "CobrancaBR_B2B@electrolux.com"
            .HTMLBody = ""
            .To = Destinatario
            .subject = aba_modelos_de_emails.Range("B2").Value
            .HTMLBody = emailbody1 & emailBody2 & emailBody3 & RangeToHTML(range_linhas, "PARA_ANALISTA_COM_CLIENTE_SEM_DADOS_BANCARIOS") & "<br>" & emailBody4 & emailBody5 & .HTMLBody
            .send
        End With
     End If
    
End Sub

Sub email_para_clientes_sem_dados_bancarios()

    Destinatario = Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:L"), 12, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:M"), 13, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:N"), 14, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:O"), 15, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:P"), 16, False)
    
    If Destinatario = ";;;;" Then
        Destinatario = Application.WorksheetFunction.VLookup(payer_atual, aba_plan_distribuicao.Columns("A:Q"), 17, False)
    End If
    
    linha_fim = aba_fbl5n_credito_devolucao.Range("A1048576").End(xlUp).Row
    tabela_aba_fbl5n_credito_devolucao.Range.AutoFilter Field:=3, Criteria1:=payer_atual
    Set range_linhas = aba_fbl5n_credito_devolucao.Range("A1:AB" & linha_fim).SpecialCells(xlCellTypeVisible)
    valor = Round(Application.WorksheetFunction.SumIf(aba_fbl5n_credito_devolucao.Columns("C:C"), payer_atual, aba_fbl5n_credito_devolucao.Columns("P:P")) * -1, 2)
    email_analista = Application.WorksheetFunction.VLookup(payer_atual, aba_plan_distribuicao.Columns("A:Q"), 17, False)
    nome_cliente = Application.WorksheetFunction.VLookup(payer_atual, aba_plan_distribuicao.Columns("A:B"), 2, False)
    
    emailbody1 = aba_modelos_de_emails.Range("E3").Value & "<br><br>"
    emailBody2 = aba_modelos_de_emails.Range("E4").Value & " <b>" & valor & "</b>"
    emailBody3 = aba_modelos_de_emails.Range("E5").Value & "<br><br>"
    emailBody4 = aba_modelos_de_emails.Range("E6").Value & "<br><br>"
    emailBody5 = aba_modelos_de_emails.Range("E7").Value & "<br>"
    emailBody6 = aba_modelos_de_emails.Range("E8").Value
        
    On Error Resume Next
    Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
    On Error GoTo 0

    With MailItem
        .SentOnBehalfOfName = "CobrancaBR_B2B@electrolux.com"
        .HTMLBody = ""
        .To = Destinatario
        .CC = email_analista
        .subject = aba_modelos_de_emails.Range("E2").Value & " - " & payer_atual & " - " & nome_cliente
        .HTMLBody = emailbody1 & emailBody2 & emailBody3 & RangeToHTML(range_linhas, "PARA_CLIENTE_SEM_DADOS_BANCARIOS") & "<br><br>" & emailBody4 & emailBody5 & emailBody6 & .HTMLBody
        .send
    End With

End Sub

Sub email_abatimento()

    Destinatario = Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:L"), 12, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:M"), 13, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:N"), 14, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:O"), 15, False) & ";" & _
        Application.WorksheetFunction.VLookup(payer_atual, _
        aba_plan_distribuicao.Columns("A:P"), 16, False)
        
    If Destinatario = ";;;;" Then
        Destinatario = Application.WorksheetFunction.VLookup(payer_atual, aba_plan_distribuicao.Columns("A:Q"), 17, False)
    End If
    linha_fim = aba_titulos_a_abater.Range("A1048576").End(xlUp).Row
    valor = Round(Application.WorksheetFunction.SumIf(aba_titulos_a_abater.Columns("A:A"), payer_atual, aba_titulos_a_abater.Columns("H:H")), 2)
    tabela_titulos_a_abater.Range.AutoFilter Field:=1, Criteria1:=payer_atual
    Set range_linhas = aba_titulos_a_abater.Range("A1:H" & linha_fim).SpecialCells(xlCellTypeVisible)
    email_analista = Application.WorksheetFunction.VLookup(payer_atual, aba_plan_distribuicao.Columns("A:Q"), 17, False)
    nome_cliente = Application.WorksheetFunction.VLookup(payer_atual, aba_plan_distribuicao.Columns("A:B"), 2, False)
    
    emailbody1 = aba_modelos_de_emails.Range("K3").Value & "<br><br>"
    emailBody2 = aba_modelos_de_emails.Range("K4").Value
    emailBody3 = aba_modelos_de_emails.Range("K5").Value & "<br><br>"
    emailBody4 = aba_modelos_de_emails.Range("K6").Value & "<br><br>"
    emailBody5 = aba_modelos_de_emails.Range("K7").Value & "<br>"
    emailBody6 = aba_modelos_de_emails.Range("K8").Value
        
    On Error Resume Next
    Set MailItem = caixa_do_b2b.Items.Add(olMailItem)
    On Error GoTo 0

    With MailItem
        .SentOnBehalfOfName = "CobrancaBR_B2B@electrolux.com"
        .HTMLBody = ""
        .To = Destinatario
        .CC = email_analista
        .subject = aba_modelos_de_emails.Range("K2").Value & " - " & nome_cliente
        .HTMLBody = emailbody1 & emailBody2 & valor & emailBody3 & RangeToHTML(range_linhas, "ABATIMENTO") & "<br>" & emailBody4 & emailBody5 & emailBody6 & .HTMLBody
        .send
    End With
    
End Sub

Public Function RangeToHTML(rng As Range, tipo_email As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim TempFile, HTMLContent As String
    Dim TempWB As Workbook

    ' Cria um novo arquivo temporário
    TempFile = Environ$("temp") & "\TempHTMLFile.htm"

    ' Cria um novo workbook temporário
    Set TempWB = Workbooks.Add(1)
    rng.Copy
    
    If tipo_email = "PARA_ANALISTA_COM_CLIENTE_SEM_DADOS_BANCARIOS" Then
        With TempWB.Sheets(1)
            .Cells(1, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            .Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Cells(1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Cells(1, 1).Value = "Código"
            .Cells(1, 2).Value = "Valor"
            .Cells(1, 3).Value = "Analista"
            .Columns("B:B").NumberFormat = "$#,##0.##"
            .Columns("A:C").AutoFit
        End With
    ElseIf tipo_email = "PARA_CLIENTE_SEM_DADOS_BANCARIOS" Or tipo_email = "REEMBOLSO" Then
        With TempWB.Sheets(1)
            .Cells(1, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            .Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Cells(1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Columns("A:B").Delete
            .Columns("C:C").Delete
            .Columns("D:E").Delete
            .Columns("E:K").Delete
            .Columns("F:G").Delete
            .Columns("G:Z").Delete
            .Cells(1, 1).Value = "Código"
            .Cells(1, 2).Value = "Nome"
            .Cells(1, 3).Value = "NFD"
            .Cells(1, 4).Value = "Tipo Documento"
            .Cells(1, 5).Value = "Valor"
            .Cells(1, 6).Value = "Informação Adicional"
            .Columns("E:E").NumberFormat = "$#,##0.##"
            .Columns("A:F").AutoFit
        End With
        i2 = 2
        ' tratativa para deixar todas as linhas de reembolso positivas para enviar ao cliente
        Do While TempWB.Sheets(1).Cells(i2, 4).Value <> ""
            TempWB.Sheets(1).Cells(i2, 5).Value = TempWB.Sheets(1).Cells(i2, 5).Value * -1
            i2 = i2 + 1
        Loop
        
    ElseIf tipo_email = "ABATIMENTO" Then
        With TempWB.Sheets(1)
            .Cells(1, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            .Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Cells(1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Columns("A:H").AutoFit
            i_fim = TempWB.Sheets(1).Range("H1048576").End(xlUp).Row
            TempWB.Sheets(1).Range("H" & i_fim).Font.Color = -16776961
            TempWB.Sheets(1).Range("H" & i_fim).Font.Bold = True
        End With
    End If


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
