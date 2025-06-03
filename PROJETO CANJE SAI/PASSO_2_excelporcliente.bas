Attribute VB_Name = "PASSO_2_excelporcliente"
Public nome_mes_anterior As String
Public ano_mes_anterior As Integer


Sub divisao_sais()

Dim sai As String
Dim linha As Integer
Dim linha_fim_sais As Long
Dim linha_fim_duplicata As Integer
Dim export_geral As Workbook
Dim fim_mes_anterior As Date
Dim fim_mes_anterior_numerc As Double
Dim resultado As Double
Dim tbl As ListObject
Dim aba_fac_deudor, aba_canje As Worksheet
Dim arquivo_criado As Workbook
Dim cod_cliente As String

    Application.ScreenUpdating = False
    
    Set export_geral = ThisWorkbook
    Set aba_fac_deudor = export_geral.Sheets("1.1 FAC Deudor")
    Set aba_canje = export_geral.Sheets("0. CANJE")
    Set tbl = aba_fac_deudor.ListObjects("FBL5N__FAC_Deudor")
    
    aba_fac_deudor.Range("aa:ab").ClearContents
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0
    
    ano_mes_anterior = Year(Date - 20)
    fim_mes_anterior = WorksheetFunction.EoMonth((Date), -1)
    fim_mes_anterior_numerc = VBA.Format(fim_mes_anterior, "0")
    linha_fim_sais = aba_fac_deudor.Range("d2").End(xlDown).Row
    
    
    aba_fac_deudor.Range("c2:c" & linha_fim_sais).Copy
    aba_fac_deudor.Range("aa2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    aba_fac_deudor.Range("aa2:aa" & linha_fim_sais).RemoveDuplicates Columns:=1, Header:=xlNo
    linha_fim_duplicata = aba_fac_deudor.Range("aa2").End(xlDown).Row
    
    On Error GoTo Erro
    nome_mes_anterior = Choose(Month(Date) - 1, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    On Error GoTo 0

1:



    MsgBox "Por favor, escolha a pasta onde os arquivos serão salvos.", vbInformation, "Aviso"

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usuário selecionou uma pasta
            CaminhoPasta = .SelectedItems(1) & "\"
        Else
            ' O usuário cancelou a seleção da pasta
            MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
            Exit Sub
        End If
    End With
    
    

    For linha = 2 To linha_fim_duplicata
        cod_cliente = aba_fac_deudor.Range("aa" & linha).Value
        
        sai = Application.WorksheetFunction.VLookup(CLng(cod_cliente), aba_canje.Range("A21:D200"), 4, False)
        If Application.WorksheetFunction.VLookup(CLng(cod_cliente), aba_canje.Range("A21:E200"), 5, False) = 0 Then
            GoTo proximo_sai
        End If
        tbl.Range.AutoFilter Field:=3, Criteria1:=cod_cliente
        tbl.Range.AutoFilter Field:=9, Criteria1:=">=" & fim_mes_anterior_numerc - 100
        tbl.Range.SpecialCells(xlCellTypeVisible).Copy
        Workbooks.Add
        Set arquivo_criado = ActiveWorkbook
        arquivo_criado.Sheets(1).Range("A1").PasteSpecial
        arquivo_criado.Sheets(1).Columns("A:D").Delete Shift:=xlToLeft
        arquivo_criado.Sheets(1).Columns("B:C").Delete Shift:=xlToLeft
        arquivo_criado.Sheets(1).Columns("F:G").Delete Shift:=xlToLeft
        arquivo_criado.Sheets(1).Columns("A:E").AutoFit
        arquivo_criado.Sheets(1).Name = "Estado de Cuenta"
        arquivo_criado.Sheets(1).Range("a1").End(xlToLeft).AutoFilter
        Columns("A:A").Replace What:="FAE0", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        Columns("A:A").Replace What:="NCE00", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        arquivo_criado.Sheets(1).Range("A1:E" & linha_fim_sais).AutoFilter Field:=3, Criteria1:="<=" & fim_mes_anterior_numerc, Operator:=xlAnd
        arquivo_criado.Sheets(1).Range("A1:E" & linha_fim_sais).SpecialCells(xlCellTypeVisible).Copy
        Sheets.Add After:=ActiveSheet
        arquivo_criado.Sheets(2).Name = "Canje " & nome_mes_anterior & " " & ano_mes_anterior
        arquivo_criado.Sheets(2).PasteSpecial
        resultado = WorksheetFunction.Sum(arquivo_criado.Sheets(2).Range("E:E"))
        arquivo_criado.Sheets(2).Range("E1").End(xlDown).Offset(1, 0).Value = resultado
        arquivo_criado.Sheets(2).Range("E1").End(xlDown).NumberFormat = "#,##0"
        arquivo_criado.Sheets(2).Range("E1").End(xlDown).Font.Bold = True
        arquivo_criado.Sheets(2).Range("D1").End(xlDown).Offset(1, 0).Value = "TOTAL CANJE " & nome_mes_anterior
        arquivo_criado.Sheets(2).Columns("A:E").AutoFit
        arquivo_criado.Sheets(1).Activate
        resultado = WorksheetFunction.Sum(Range("E:E"))
        Range("q5").Value = resultado
        Range("q4").Value = "TOTAL PENDIENTE"
        Range("q5").NumberFormat = "#,##0"
        Range("q4:Q5").Font.Bold = True
        Range("q4:Q5").AutoFilter
        ActiveWorkbook.SaveAs CaminhoPasta & "CANJE " & nome_mes_anterior & " - " & ano_mes_anterior & " - " & sai & ".xlsx"
        ActiveWorkbook.Close
proximo_sai:
    Next linha
    
    GoTo 2

Erro:

nome_mes_anterior = "Diciembre"

GoTo 1

2:

MsgBox "CANJE por cliente realizado, verifique os montantes e envie os e-mails", vbOKOnly, "CANJE" & "/ " & nome_mes_anterior & "/" & ano_mes_anterior

Application.ScreenUpdating = True

End Sub



