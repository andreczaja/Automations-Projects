Attribute VB_Name = "PASSO_3_ALTERACAO_CAMPOS_SAP"
Sub alteracao_campos_sap()


    Application.DisplayAlerts = False

    On Error Resume Next
    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    Set Connection = Applic.Connections(0)
    Set session = Connection.Children(0)
    On Error GoTo 0
    Call declaracao_variaveis
    
End Sub
