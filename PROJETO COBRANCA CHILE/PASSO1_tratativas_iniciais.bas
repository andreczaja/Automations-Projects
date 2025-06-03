Attribute VB_Name = "PASSO1_tratativas_iniciais"
    Public base_cobraveis As Workbook
    Public aba_export_sap, aba_cobravel_hoje, aba_base_emails, aba_controle_diario As Worksheet
    Public tbl As ListObject
    Public tbl2 As ListObject
    Public tbl3 As ListObject
    
    Public cod_cliente, nome_cliente, Folder, clientes_nao_mapeados, analista_responsavel_cobranca_telefone, analista_responsavel_dicom_equifax, _
    analistas_sem_bloqueio, tipo_da_cobranca, analista As String
    
    Public linha, linha_fim, linha_fim_aba_cobravel_hoje_selecao_de_linhas, contador_arquivos_criados_email_mais_de_10_faturas, _
    contador_cobranca_email_menos_de_10_faturas, contador_cobranca_telefone, contador_dicom_equifax, qtde_facturas, i As Integer
    
    Public clientes_nao_cobraveis
    Public arquivocriado As Workbook
    

Sub tratativas_iniciais_()


    frm_passos.Hide

    
    Set base_cobraveis = ThisWorkbook
    Set aba_export_sap = base_cobraveis.Sheets("Export SAP")
    Set tbl = aba_export_sap.ListObjects("Export_FBL5N___Cobráveis")
    Set aba_cobravel_hoje = base_cobraveis.Sheets("Cobraveis HOJE")
    Set aba_base_emails = base_cobraveis.Sheets("Base E-mails")
    Set tbl2 = aba_cobravel_hoje.ListObjects("Tabela_Cobraveis_HOJE")
    Set aba_controle_diario = base_cobraveis.Sheets("Controle Diário")
    Set tbl3 = aba_controle_diario.ListObjects("Status_Bloqueios_Diários_Analistas")
    
    'definindo array de clientes que não entram no circuito de cobrança (key accounts)
    clientes_nao_cobraveis = Array(232280, 234482, 237141, 232651, 272186, 234920, 200070)
    
    tbl3.QueryTable.Refresh False
    
    analistas_sem_bloqueio = ""
    cliente_nao_mapeado_telefone = ""
    cliente_nao_mapeado_dicom_equifax = ""
    
    
    If aba_export_sap.Range("BB1").Value = "" Then
        Folder = BuscarPasta("")
        FileCopy Folder & "\FBL5N - BASE VAZIA.txt", Folder & "\FBL5N-C.txt"
        tbl.QueryTable.Refresh False
    End If
    
    tbl.AutoFilter.ShowAllData
    
    
    ' estrutura que verifica se existem analistas que não preencheram o bloqueio do dia, se não preencheu, a estrutura não continua
    
    For i = 2 To aba_controle_diario.Range("A1").End(xlDown).Row
        If aba_controle_diario.Range("B" & i).Value = "" Then
            analistas_sem_bloqueio = analistas_sem_bloqueio & " - " & aba_controle_diario.Range("A" & i).Value
        End If
    Next i
    
    If analistas_sem_bloqueio <> "" Then
        analistas_sem_bloqueio = Mid(analistas_sem_bloqueio, 3)
        MsgBox "Os analistas : " & analistas_sem_bloqueio & " não preencheram os bloqueios diários. Impossível continuar!", vbOKOnly
        End
    End If
    
    Application.ScreenUpdating = False
    
    'apaga clientes inadimplentes de envios anteriores
'    aba_cobravel_hoje.Activate
    aba_export_sap.Range("a6").Activate
    On Error Resume Next
    aba_export_sap.ShowAllData

    aba_cobravel_hoje.Range("a1").Activate
    On Error Resume Next
    aba_cobravel_hoje.ShowAllData
    aba_cobravel_hoje.Range("a2:am99999").ClearContents
    aba_cobravel_hoje.Range("bb2:bz99999").ClearContents

        

    'verifica condições de filtros
    On Error Resume Next
    aba_export_sap.ShowAllData
    
    
        
    'FILTRANDO POR "Cobrança por E-mail", "Cobrança por Telefone", "Cobrança por E-mail-Dicom/Equifax", "Cobrança Preventiva - Constructoras"
    tbl.Range.AutoFilter Field:=39, Criteria1:=Array("Cobrança por E-mail", "Cobrança por Telefone", "Cobrança por E-mail-Dicom/Equifax", _
    "Cobrança Preventiva - Constructoras"), Operator:=xlFilterValues
    
    'excluindo linhas com chave de referencia 3 preenchidas
     tbl.Range.AutoFilter Field:=20, Criteria1:=""
     
    'excluindo linhas com chave de reclamacion preenchidas
     tbl.Range.AutoFilter Field:=28, Criteria1:=""
     
    If Application.WorksheetFunction.Subtotal(103, tbl.ListColumns(1).DataBodyRange) = 0 Then
        MsgBox "Não foram encontrados títulos a cobrar na data atual", vbOKOnly
        End
    End If
    
    
    ' PASSANDO LINHAS PARA A ABA COBRAVEL HOJE
    linha_fim = aba_export_sap.Range("A6").End(xlDown).Row
    aba_export_sap.Range("A6:AM" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
    aba_cobravel_hoje.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    For i = 1 To 1000
        If aba_cobravel_hoje.Range("A" & i).Value = "" Then
            linha_fim = aba_cobravel_hoje.Range("A" & i - 1).Row
            Exit For
        End If
    Next i
    Sheets("Cobraveis HOJE").ListObjects("Tabela_Cobraveis_HOJE").Resize Range("A1:AN" & linha_fim)
    aba_cobravel_hoje.Range("A" & linha_fim + 1 & ":AN999999").ClearContents

        linha = 2
        
        tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por E-mail"
        If Application.WorksheetFunction.Subtotal(103, tbl2.ListColumns(1).DataBodyRange) > 0 Then
            With aba_cobravel_hoje
                .Range("B" & linha & ":B" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
                .Range("BB2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                .Range("BB2").RemoveDuplicates Columns:=1, Header:=xlNo
            End With
                
                
                On Error Resume Next
                aba_cobravel_hoje.Range("a1").Activate
                aba_cobravel_hoje.ShowAllData
                
            
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ETAPA DE PREENCHIMENTO COLUNAS BB, BC E BD REFERENTE A COBRANÇA POR E-MAIL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            ' determina a linha final de código de clientes com cobrança por e-mail
            If aba_cobravel_hoje.Range("BB2").Value = "" Then
                MsgBox "Nenhum documento cobrável encontrado", vbOKOnly
                End
            ElseIf aba_cobravel_hoje.Range("BB3").Value = "" Then
                linha_fim = 2
            Else
                linha_fim = aba_cobravel_hoje.Range("BB1").End(xlDown).Row
            End If
        
        
            linha = 2
            
            Do Until linha > linha_fim
                cod_cliente = aba_cobravel_hoje.Range("BB" & linha).Value
                nome_cliente = aba_cobravel_hoje.Range("BD" & linha).Value
                
                If Application.WorksheetFunction.CountIfs(aba_cobravel_hoje.Range("B:B"), cod_cliente, aba_cobravel_hoje.Range("AM:AM"), "Cobrança por E-mail") >= 10 Then
                    aba_cobravel_hoje.Range("BC" & linha).Value = "Mais de 10 faturas"
                Else
                    aba_cobravel_hoje.Range("BC" & linha).Value = "Menos de 10 faturas"
                End If
                
                If Application.WorksheetFunction.CountIf(aba_base_emails.Columns("A:A"), cod_cliente) = 0 Then
                    aba_cobravel_hoje.Range("BD" & linha).Value = "Cliente não mapeado"
                Else
                    aba_cobravel_hoje.Range("BD" & linha).Value = Application.WorksheetFunction.VLookup(cod_cliente, aba_base_emails.Columns("A:B"), 2, False)
                End If
                
                linha = linha + 1
            Loop
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ETAPA DE PREENCHIMENTO COLUNAS BE, BF E BG REFERENTE A COBRANÇA POR TELEFONE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    linha_fim = aba_cobravel_hoje.Range("A99999").End(xlUp).Row
    
    tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por Telefone"
    
    If Application.WorksheetFunction.Subtotal(103, tbl2.ListColumns(1).DataBodyRange) = 0 Then
        GoTo etapa_cobranca_dicom_equifax
    End If
    
    With aba_cobravel_hoje
            .Range("AN2:AN" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            .Range("BE2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            .Range("BE:BE").RemoveDuplicates Columns:=1, Header:=xlNo
    End With
    
    
    aba_cobravel_hoje.Range("a1").Activate
    On Error Resume Next
    aba_cobravel_hoje.ShowAllData
    
    tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por Telefone"
    tbl2.Range.AutoFilter Field:=40, Criteria1:=Array("-", ""), Operator:=xlFilterValues
    

    If Application.WorksheetFunction.Subtotal(103, tbl2.ListColumns(1).DataBodyRange) > 0 Then
        With aba_cobravel_hoje
            .Range("B2:B" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            .Range("BF2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            .Range("BF:BF").RemoveDuplicates Columns:=1, Header:=xlNo
        End With
        
        If aba_cobravel_hoje.Range("BF3").Value = "" Then
            linha_fim = 2
        Else
            linha_fim = aba_cobravel_hoje.Range("BF1").End(xlDown).Row
        End If
            
        linha = 2
            
        Do Until linha > linha_fim
            cod_cliente = aba_cobravel_hoje.Range("BF" & linha).Value
            aba_cobravel_hoje.Range("BG" & linha).Value = Application.WorksheetFunction.VLookup(cod_cliente, aba_base_emails.Range("A:B"), 2, False)
            linha = linha + 1
        Loop
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ETAPA DE PREENCHIMENTO COLUNAS BH, BI E BJ REFERENTE A DICOM/EQUIFAX
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
etapa_cobranca_dicom_equifax:

    aba_cobravel_hoje.Range("a1").Activate
    On Error Resume Next
    aba_cobravel_hoje.ShowAllData

    linha = 2
    linha_fim = aba_cobravel_hoje.Range("A99999").End(xlUp).Row
    
    tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por E-mail-Dicom/Equifax"
    
    If Application.WorksheetFunction.Subtotal(103, tbl2.ListColumns(1).DataBodyRange) = 0 Then
        GoTo etapa_cobranca_preventiva
    End If
    
    With aba_cobravel_hoje
            .Range("AN2:AN" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            .Range("BH2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            .Range("BH:BH").RemoveDuplicates Columns:=1, Header:=xlNo
    End With
    
    On Error Resume Next
    aba_cobravel_hoje.Range("a1").Activate
    aba_cobravel_hoje.ShowAllData
    
    tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por E-mail-Dicom/Equifax"
    tbl2.Range.AutoFilter Field:=40, Criteria1:=Array("-", ""), Operator:=xlFilterValues
    

    If Application.WorksheetFunction.Subtotal(103, tbl2.ListColumns(1).DataBodyRange) > 0 Then
        With aba_cobravel_hoje
            .Range("B2:B" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            .Range("BI2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            .Range("BI:BI").RemoveDuplicates Columns:=1, Header:=xlNo
        End With
        
    On Error Resume Next
    aba_cobravel_hoje.Range("a1").Activate
    aba_cobravel_hoje.ShowAllData
        
        If aba_cobravel_hoje.Range("BI3").Value = "" Then
            linha_fim = 2
        Else
            linha_fim = aba_cobravel_hoje.Range("BI99999").End(xlUp).Row
        End If
            
        linha = 2
            
        Do Until linha > linha_fim
            cod_cliente = aba_cobravel_hoje.Range("BI" & linha).Value
            aba_cobravel_hoje.Range("BJ" & linha).Value = Application.WorksheetFunction.VLookup(cod_cliente, aba_base_emails.Range("A:B"), 2, False)
            linha = linha + 1
        Loop
    End If



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ETAPA DE PREENCHIMENTO COLUNAS BK E BL REFERENTE A COBRANÇA PREVENTIVA''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
etapa_cobranca_preventiva:


    

    aba_cobravel_hoje.Range("a1").Activate
    On Error Resume Next
    aba_cobravel_hoje.ShowAllData

    linha = 2
    linha_fim = aba_cobravel_hoje.Range("A99999").End(xlUp).Row
    
    tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança Preventiva - Constructoras"
    
    If Application.WorksheetFunction.Subtotal(103, tbl2.ListColumns(1).DataBodyRange) = 0 Then
        GoTo fim
    End If
    
    With aba_cobravel_hoje
            .Range("B2:B" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
            .Range("BK2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            .Range("BK:BK").RemoveDuplicates Columns:=1, Header:=xlNo
    End With
    
    On Error Resume Next
    aba_cobravel_hoje.Range("a1").Activate
    aba_cobravel_hoje.ShowAllData
    
        
    If aba_cobravel_hoje.Range("BK3").Value = "" Then
        linha_fim = 2
    Else
        linha_fim = aba_cobravel_hoje.Range("BK99999").End(xlUp).Row
    End If
            
    linha = 2
            
    Do Until linha > linha_fim
        cod_cliente = aba_cobravel_hoje.Range("BK" & linha).Value
        aba_cobravel_hoje.Range("BL" & linha).Value = Application.WorksheetFunction.VLookup(cod_cliente, aba_base_emails.Range("A:B"), 2, False)
        linha = linha + 1
    Loop

fim:
    criar_arquivos_de_clientes_

End Sub
