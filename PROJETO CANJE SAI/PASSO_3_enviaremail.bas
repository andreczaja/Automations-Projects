Attribute VB_Name = "PASSO_3_enviaremail"
Option Explicit

Sub enviar_email_()
    
    Dim OutlookApp As Object
    Dim OutlookMail, ObjFSO, ObjPasta, Arquivo As Object
    Dim linha As Integer
    Dim linha_fim As Integer
    Dim qtde_sais As Integer
    Dim qtde_docs As Integer
    Dim fileName As String
    Dim emailBody1 As String
    Dim emailBody2 As String
    Dim signature As String
    Dim folder As String
    Dim condicao_recebedora As Boolean
    Dim condicao_devedora As Boolean
    
    On Error GoTo Erro
    nome_mes_anterior = Choose(Month(Date) - 1, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")

1:
    ano_mes_anterior = Year(Date - 20)
    Set canje = ThisWorkbook.Sheets(1)
    linha = 22
    linha_fim = canje.Range("D22").End(xlDown).Row
    qtde_sais = canje.Range("D22").End(xlDown).Row - 21
    
    
    MsgBox "Por favor, selecione a pasta onde os arquivos estão salvos.", vbInformation, "Aviso"

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usuário selecionou uma pasta
            folder = .SelectedItems(1) & "\"
            Set ObjFSO = CreateObject("Scripting.FileSystemObject")
            Set ObjPasta = ObjFSO.GetFolder(folder)
            
            For Each Arquivo In ObjPasta.Files
                ' Verifica se o nome do arquivo começa com "CANJE"
                If Left(Arquivo.Name, 5) = "CANJE" Then
                qtde_docs = qtde_docs + 1
                End If
            Next Arquivo
        Else
            ' O usuário cancelou a seleção da pasta
            MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
            Exit Sub
        End If
    End With
    
    

    ' Crie uma instância do Outlook
    On Error Resume Next
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    
    Do Until linha = linha_fim + 1
    Set OutlookMail = OutlookApp.CreateItem(0)

    
    If canje.Range("E" & linha).Value <> 0 And canje.Range("R" & linha).Value = "Normal (devemos pagar à eles)" And canje.Range("S" & linha).Value = "" Then
        condicao_recebedora = True
        condicao_devedora = False
    Else
        condicao_recebedora = False
    End If
    
    
    If canje.Range("E" & linha).Value <> 0 And canje.Range("R" & linha).Value = "Devedora (devemos cobrar deles)" And canje.Range("S" & linha).Value = "" Then
        condicao_devedora = True
        condicao_recebedora = False
    Else
        condicao_devedora = False
    End If
    
    With OutlookMail
    
        If condicao_recebedora = True Then
            OutlookMail.display
            .To = canje.Range("Q" & linha) ' Substitua pelo endereço de e-mail do destinatário
            .Subject = "Canje " & nome_mes_anterior & " - " & ano_mes_anterior & " - " & canje.Range("d" & linha) ' Substitua pelo assunto desejado
            .CC = "pablo.ruz@electrolux.com ; aron.gonzalez@electrolux.com"
            ' Adicione o corpo do e-mail
            emailBody1 = "Estimados," _
            & "<br><br>" & "Adjunto archivo con el detalle de las facturas que serán aplicadas en el Canje de " & _
            nome_mes_anterior & " " & ano_mes_anterior ' Substitua pelo texto desejado
            emailBody2 = "Saludos,"
            .htmlbody = emailBody1 _
            & "<br><br>" _
            & emailBody2 & _
            "<signature>" _
            & .htmlbody
            .Attachments.Add folder & "CANJE " & nome_mes_anterior & " - " & ano_mes_anterior & " - " & canje.Range("D" & linha) & ".xlsx"   ' Anexa o arquivo ao e-mail
        
        ElseIf condicao_devedora = True Then
            OutlookMail.display
            .To = canje.Range("Q" & linha) ' Substitua pelo endereço de e-mail do destinatário
            .Subject = "Canje " & nome_mes_anterior & " - " & ano_mes_anterior & " - " & canje.Range("D" & linha) ' Substitua pelo assunto desejado
            .CC = " pablo.ruz@electrolux.com ; aron.gonzalez@electrolux.com"
            ' Adicione o corpo do e-mail
            emailBody1 = "Estimado," & "<br><br>" _
            & "Adjunto archivo con el detalle de las facturas que serán aplicadas en el Canje de " & _
            nome_mes_anterior & " " & ano_mes_anterior ' Substitua pelo texto desejado
            emailBody2 = "Informo que presenta un saldo en contra de $" & Format(canje.Range("l" & linha).Value, "#,###,##0") * -1 & " CLP. Favor indicar fecha de pago." & _
              "<br><br>" & _
             "Saludos"
            .htmlbody = emailBody1 _
            & "<br><br>" _
            & emailBody2 & _
            "<signature>" _
            & .htmlbody
            .Attachments.Add folder & "CANJE " & nome_mes_anterior & " - " & ano_mes_anterior & " - " & canje.Range("D" & linha) & ".xlsx"
        End If
        
    End With
    
    linha = linha + 1
    
    Loop
    
        GoTo 2

Erro:

nome_mes_anterior = "Diciembre"

GoTo 1

2:

    ' Limpeza
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing


End Sub

