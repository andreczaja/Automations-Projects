Attribute VB_Name = "extracao_SAP"
'variavel que possibilita que o codigo encerre caso se clique no botao de Cancelar no meio
' da execucao do codigo

Public SapGui As Object, WSHShell As Object, session As Object, Applic As Object, connection As Object, SapGuiWB, Window As Object
Public linha, linha_fim, i, OpenWin As Integer
Public Planilha As Workbook
Public aba_geral As Worksheet
Public sessao_f_32 As Boolean


Sub SAP_Login()

       
    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    
    On Error Resume Next
    Set connection = Applic.Connections(0)
    
    If connection Is Nothing Then
        MsgBox "Feche o SAP no gerenciador de tarefas!", vbOKOnly, "Electrolux Group"
        Exit Sub
    End If
    Set session = connection.Children(0)
 
 

    OpenWin = 0

    For i = 0 To connection.Children.Count - 1
        Set session = connection.Children(CInt(i))
        If session.Info.Transaction <> "F-32" Then
            Set Window = session.ActiveWindow
            If Window.iconic = False Then
                OpenWin = OpenWin + 1
            End If
        Else
            sessao_f_32 = True
            Exit For
        End If
    Next i
    
    If Not sessao_f_32 Then
        MsgBox "Para continuar o script você deve estar no meio da compensação do pagamento na tela 'COMPENSAR DEUDOR VISUALIZAR RESUMEN'.", vbOKOnly, "Electrolux Group"
        End
    End If

'clica no botão para entrar no SAP

        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").maximize

'troca de móduo (transação fbl5n)
f32

    Set connection = Nothing
    Set Applic = Nothing
    Set SapGui = Nothing
    Set SapGuiWB = Nothing
    Set session = Nothing
    
    
    End
    
    
End Sub



Sub f32()



        Set Planilha = ThisWorkbook
        Set aba_geral = Planilha.Sheets("Automatização")
    
    i = 0
    linha = 2
    linha_fim = aba_geral.Range("A1").End(xlDown).Row
    

    
    ' Se a transação F-32 não estiver ativa na primeira janela, verifique outras janelas abertas
    If sessao_f_32 Then
            
                session.findById("wnd[0]").maximize
                session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = "24"
                session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = aba_geral.Range("A" & linha).Value
                session.findById("wnd[0]").sendVKey 0
                
                Do Until linha = linha_fim
                
                 
                        session.findById("wnd[0]/usr/txtBSEG-WRBTR").Text = aba_geral.Range("C" & linha).Value
                        session.findById("wnd[0]/usr/txtBSEG-SKFBT").Text = aba_geral.Range("C" & linha).Value
                        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = aba_geral.Range("E" & linha).Value
                        session.findById("wnd[0]/tbar[1]/btn[25]").press
                
                        linha = linha + 1
                Loop
                
                        session.findById("wnd[0]/usr/txtBSEG-WRBTR").Text = aba_geral.Range("C" & linha).Value
                        session.findById("wnd[0]/usr/txtBSEG-SKFBT").Text = aba_geral.Range("C" & linha).Value
                        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = aba_geral.Range("E" & linha).Value
                        session.findById("wnd[0]/tbar[1]/btn[14]").press
            
                            sessao_f_32 = True
    End If

    
    ' Se a transação F-32 não estiver ativa em nenhuma janela, exiba uma mensagem
    If Not sessao_f_32 Then
        MsgBox "Nenhuma janela SAP com a transação F-32 ativa foi encontrada.", vbInformation
    End If
    
    
    MsgBox "Processo Concluído.", vbOKOnly

End Sub
