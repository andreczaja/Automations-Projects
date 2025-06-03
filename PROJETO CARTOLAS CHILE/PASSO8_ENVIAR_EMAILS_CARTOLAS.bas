Attribute VB_Name = "PASSO8_ENVIAR_EMAILS_CARTOLAS"
Option Explicit

Sub enviar_email_()

    Dim OutlookApp As Object
    Dim OutlookMail, ObjFSO, ObjPasta, Arquivo As Object
    Dim base_emails As Worksheet
    Dim linha, qtde_docs As Integer
    Dim fileName, emailBody1, emailBody2, signature, folder, payer, destinatario, payers_sem_email As String
    Dim condicao_payer_sem_email As Boolean
    
    
    Set base_emails = ThisWorkbook.Sheets("Base E-mails")

    MsgBox "Por favor, selecione a pasta onde os arquivos estão salvos.", vbInformation, "Aviso"

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usuário selecionou uma pasta
            folder = .selectedItems(1) & "\"
            Set ObjFSO = CreateObject("Scripting.FileSystemObject")
            Set ObjPasta = ObjFSO.GetFolder(folder)
        Else
            ' O usuário cancelou a seleção da pasta
            MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False

    ' Crie uma instância do Outlook
    On Error Resume Next
        'Set OutlookApp = GetObject("Outlook.Application")
    On Error GoTo 0
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    
    ' Carregue o conteúdo do arquivo HTML com a assinatura
   ' signature = ("C:\Users\CardoAnd03\OneDrive - Electrolux\Pictures\assinatura email.png")

    'setando variaveis de intervalo de quantidade de arquivos
    linha = 1
    qtde_docs = 0
    For Each Arquivo In ObjPasta.Files
        qtde_docs = qtde_docs + 1
        base_emails.Range("AH" & linha).Value = Mid(Arquivo.Name, 7, 8)
        base_emails.Range("AI" & linha).Value = Mid(Arquivo.Name, 25, 20)
        linha = linha + 1
    Next Arquivo
    Application.Wait (Now + TimeValue("00:00:01"))
    base_emails.Activate
    Columns("AI:AI").Select
    Selection.Replace What:=".xlsx", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    condicao_payer_sem_email = False

    linha = 1
    
1:
    Do Until linha > qtde_docs
    
        Set OutlookMail = OutlookApp.CreateItem(0)
        payer = base_emails.Range("AH" & linha).Value
        
        If Application.WorksheetFunction.CountIf(base_emails.Columns("A:A"), payer) = 0 Then
            GoTo Erro_Falta_Email
        Else
        destinatario = Application.WorksheetFunction.VLookup(base_emails.Range("AH" & linha), _
            base_emails.Range("A2:J999999"), 5, False) & ";" & _
            Application.WorksheetFunction.VLookup(base_emails.Range("AH" & linha), _
            base_emails.Range("A2:J999999"), 6, False) & ";" & _
            Application.WorksheetFunction.VLookup(base_emails.Range("AH" & linha), _
            base_emails.Range("A2:J999999"), 7, False) & ";" & _
            Application.WorksheetFunction.VLookup(base_emails.Range("AH" & linha), _
            base_emails.Range("A2:J999999"), 8, False) & ";" & _
            Application.WorksheetFunction.VLookup(base_emails.Range("AH" & linha), _
            base_emails.Range("A2:J999999"), 9, False) & ";" & _
            Application.WorksheetFunction.VLookup(base_emails.Range("AH" & linha), _
            base_emails.Range("A2:J999999"), 10, False)
        End If
        If destinatario = ";;;;;" Then
            GoTo Erro_Falta_Email
        Else
        emailBody1 = "Olá,"
        emailBody2 = "Segue anexo composição de pagamento de verbas comerciais depositadas em " & _
                    base_emails.Range("AI" & linha).Value & "." & "<br><br>" & "Atenciosamente,"
            
        With OutlookMail

            OutlookMail.display
            .to = destinatario
            .Subject = "Composição de Depósito de Verbas Comerciais - " & base_emails.Range("AI" & linha).Value
            .htmlbody = emailBody1 & "<br><br>" & emailBody2 & .htmlbody
            .Attachments.Add folder & "Payer " & payer & " extracao " & base_emails.Range("AI" & linha).Value & ".xlsx"
            
        End With
        End If
    linha = linha + 1
    
    Loop
    
    Application.ScreenUpdating = True
    
    GoTo fim

Erro_Falta_Email:
    payers_sem_email = payers_sem_email & " " & payer & " "
    linha = linha + 1
    condicao_payer_sem_email = True
    
GoTo 1

fim:
    If condicao_payer_sem_email = True Then
        MsgBox "Processo Concluído! Não foram encontrados os e-mails do(s) payer(s) " & payers_sem_email & " na aba Base E-mails. Favor revisar.", vbOKOnly
    Else
        MsgBox "Processo Concluído!", vbOKOnly
    End If
    
    base_emails.Range("AH:AI").ClearContents
    ' Limpeza
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing


End Sub

