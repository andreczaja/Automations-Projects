Attribute VB_Name = "PASSO6_CONTABILIZACAO_SAP"
'variavel que possibilita que o codigo encerre caso se clique no botao de Cancelar no meio
' da execucao do codigo

Public btnCancelClicked As Boolean
Dim SapGui As Object, WSHShell As Object, session As Object, Applic As Object, connection As Object, SapGuiWB As Object
Public data_inicio As String, data_fim As String, folder As String
Dim i As Integer
Dim i_fim As Integer


Sub SAP_Login()



    MsgBox "Por favor, escolha a pasta 'Refacturacion' no Sharepoint para salvar o relatório da FBL5N (Documentos > CHILE > Melhorias & Automações > Refacturacion)" & _
    ". OBS: Se não possuir o atalho para a pasta, por favor, crie.", vbOKCancel, "Aviso"



    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usuário selecionou uma pasta
            folder = .selectedItems(1) & "\"
        Else
            ' O usuário cancelou a seleção da pasta
            MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
            Exit Sub
        End If
    End With

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
            MsgBox "Feche o SAP no gerenciador de tarefas!", vbOKOnly, "Electrolux Group"
            Exit Sub
            End If
            Set session = connection.Children(0)
            
        Else:
          
          ' SAP FECHADO
            Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus

            'inicia a variável com o objeto SAP
            Set WSHShell = CreateObject("WScript.Shell")

            Do Until WSHShell.AppActivate("SAP Logon ")
            Application.Wait Now + TimeValue("0:00:01")
            Loop

            Set WSHShell = Nothing

            Set SapGui = GetObject("SAPGUI")
    
            Set Applic = SapGui.GetScriptingEngine
            
            On Error Resume Next
            Set connection = Applic.OpenConnection("SAP Electrolux Chile Prod", True)
            
            If connection Is Nothing Then
            MsgBox "Verifique se o SAP está fora do ar!", vbOKOnly, "Electrolux Group"
            Exit Sub
            SapGui.CloseSession
            End If

            
            Set session = connection.Children(0)
        
'DADOS PARA FAZER O LOGIN NO SISTEMA
            On Error Resume Next
            session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "300" 'client do sistema
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = fmLogin.txtLogin 'usuario
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = fmLogin.txtSenha
            session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "ES"  'idioma do sistema

        End If

'clica no botão para entrar no SAP

        session.findById("wnd[0]").SendVKey 0
        session.findById("wnd[0]").Maximize
'voltar para a pagina maior
        On Error Resume Next
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nBPMDG/UTL_BROWSER"
        session.findById("wnd[0]").SendVKey 0

'troca de móduo (transação fbl5n)
fbl5n
    
    Sheets("Export SAP").ListObjects("Consulta1").QueryTable.Refresh False
    
    
End Sub



Sub fbl5n()

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nBPMDG/UTL_BROWSER"
        session.findById("wnd[0]").SendVKey 0
        session.findById("wnd[0]/tbar[0]/okcd").text = "FBL5N"
        session.findById("wnd[0]").SendVKey 0
        session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/tbar[0]/btn[16]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "tc04"
        session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/CL_COMPLETE"
        session.findById("wnd[0]/tbar[1]/btn[16]").press
        session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN013_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").Select
        session.findById("wnd[1]/tbar[0]/btn[16]").press
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,0]").text = Format(Date - 15, "dd.mm.yyyy")
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").text = Format(Date - 7, "dd.mm.yyyy")
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").SetFocus
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").caretPosition = 10
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN015_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/tbar[0]/btn[16]").press
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "ea"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "eb"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 2
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN015-LOW").SetFocus
        session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN015-LOW").caretPosition = 2
        session.findById("wnd[0]/usr/chkX_NORM").Selected = True
        session.findById("wnd[0]/usr/chkX_SHBV").Selected = False
        session.findById("wnd[0]/usr/chkX_MERK").Selected = False
        session.findById("wnd[0]/usr/chkX_PARK").Selected = False
        session.findById("wnd[0]/usr/chkX_APAR").Selected = False
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "FBL5N.txt"
        session.findById("wnd[1]/tbar[0]/btn[11]").press

End Sub
