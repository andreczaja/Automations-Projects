Attribute VB_Name = "PASSO2_criar_arquivos"
Sub criar_arquivos_de_clientes_()
Attribute criar_arquivos_de_clientes_.VB_ProcData.VB_Invoke_Func = " \n14"

    MsgBox "Por favor, escolha a pasta onde os arquivos por cliente serão salvos. Para assegurar que o sistema funcione corretamente, escolha uma pasta vazia ou crie uma nova.", vbInformation, "Aviso"

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usuï¿½rio selecionou uma pasta
            Folder = .SelectedItems(1) & "\"
        Else
            ' O usuï¿½rio cancelou a seleï¿½ï¿½o da pasta
            MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
            End
        End If
    End With

contador_arquivos_criados_email_mais_de_10_faturas = 0
contador_cobranca_email_menos_de_10_faturas = 0
contador_cobranca_telefone = 0
contador_dicom_equifax = 0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''ETAPA CRIAÇÃO DE ARQUIVOS POR CLIENTE PARA COBRANÇA POR E-MAIL''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
1:
    aba_cobravel_hoje.Activate
    aba_cobravel_hoje.Range("A1").Activate
    On Error Resume Next
    aba_cobravel_hoje.ShowAllData

    linha_fim = aba_cobravel_hoje.Range("BB200").End(xlUp).Row
    
    If linha_fim = 1 Then
        GoTo Cobranca_Telefone
    End If
    
    For linha = 2 To linha_fim
    
            If aba_cobravel_hoje.Range("BC" & linha).Value = "Mais de 10 faturas" Then
            
                cod_cliente = aba_cobravel_hoje.Range("BB" & linha).Value
                nome_cliente = aba_cobravel_hoje.Range("BD" & linha).Value
                
            
                tbl2.Range.AutoFilter Field:=2, Criteria1:=cod_cliente, Operator:=xlAnd
                tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por E-mail", Operator:=xlAnd
                
                linha_fim_aba_cobravel_hoje_selecao_de_linhas = aba_cobravel_hoje.Range("A1").End(xlDown).Row
                
                aba_cobravel_hoje.Range("D1:Q" & linha_fim_aba_cobravel_hoje_selecao_de_linhas).Select
                Selection.SpecialCells(xlCellTypeVisible).Copy
                Workbooks.Add
                Set arquivocriado = ActiveWorkbook
                arquivocriado.Sheets(1).Range("A1").PasteSpecial
                Columns("A:N").AutoFit
                Columns("C:E").EntireColumn.Delete
                Columns("E:J").EntireColumn.Delete
                Columns("C:C").NumberFormat = "dd/mm/yyyy"
                arquivocriado.Sheets(1).Range("D1").Value = "Monto"
                Columns("D:D").NumberFormat = "#,##0"
                arquivocriado.Sheets(1).Range("D1").End(xlDown).Offset(1, -1).Value = "TOTAL"
                
                arquivocriado.Sheets(1).Range("D1").End(xlDown).Offset(1, 0).FormulaR1C1 = WorksheetFunction.Sum(Range("D:D"))
                arquivocriado.Sheets(1).Range("D1").End(xlDown).Font.Bold = True
                arquivocriado.Sheets(1).Range("D1").End(xlDown).Offset(0, -1).Font.Bold = True
                Range("B:B").Select
                arquivocriado.Sheets(1).Columns("A:E").Replace What:="FAE0", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                arquivocriado.Sheets(1).Columns("A:E").Replace What:="NCE000", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                arquivocriado.Sheets(1).Columns("A:E").Replace What:="NCE00", Replacement:="", LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                qtde_facturas = ActiveSheet.Range("b99999").End(xlUp).Row
                
                For i = 2 To qtde_facturas
                    If arquivocriado.Sheets(1).Range("E" & i).Value Like "CLT*" Then
                        arquivocriado.Sheets(1).Range("E" & i).ClearContents
                    End If
                Next i
                On Error Resume Next
                ActiveSheet.Name = "Facturas pendientes"
                contador_arquivos_criados_email_mais_de_10_faturas = contador_arquivos_criados_email_mais_de_10_faturas + 1
                On Error GoTo error_save_as_arquivo_cobranca_email
                
salvar_novamente_arquivo_cobranca_email:

                arquivocriado.SaveAs Folder & cod_cliente & " - Facturas pendientes " & nome_cliente & " " & Day(Date) & "." & Month(Date) & "." & Year(Date) & ".xlsx"
                ActiveWorkbook.Close
                
                aba_cobravel_hoje.Range("a1").Activate
                On Error Resume Next
                aba_cobravel_hoje.ShowAllData
        ElseIf aba_cobravel_hoje.Range("BC" & linha).Value = "Menos de 10 faturas" Then
            contador_cobranca_email_menos_de_10_faturas = contador_cobranca_email_menos_de_10_faturas + 1
        End If
            
    Next linha

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''ETAPA CRIAÇÃO DE ARQUIVOS POR CLIENTE PARA COBRANÇA POR TELEFONE''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    aba_cobravel_hoje.Range("a1").Activate
    On Error Resume Next
    aba_cobravel_hoje.ShowAllData

Cobranca_Telefone:

    linha_fim = aba_cobravel_hoje.Range("BE20").End(xlUp).Row
    
    If linha_fim = 1 Then
        GoTo cobranca_dicom_equifax
    End If
    
    For linha = 2 To linha_fim
        
            analista = aba_cobravel_hoje.Range("BE" & linha).Value
            tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por Telefone", Operator:=xlAnd
            tbl2.Range.AutoFilter Field:=40, Criteria1:=analista, Operator:=xlAnd
            
            linha_fim_aba_cobravel_hoje_selecao_de_linhas = aba_cobravel_hoje.Range("A1").End(xlDown).Row
            
            aba_cobravel_hoje.Range("B1:Q" & linha_fim_aba_cobravel_hoje_selecao_de_linhas).Select
            Selection.SpecialCells(xlCellTypeVisible).Copy
            Workbooks.Add
            Set arquivocriado = ActiveWorkbook
            arquivocriado.Sheets(1).Range("A1").PasteSpecial
            Columns("A:Z").AutoFit
            On Error Resume Next
            ActiveSheet.Name = "Facturas por cobrar"
            
            ' verificacao para nomear arquivo caso sejam titulos de clientes em que o analista nao foi mapeado
            If analista = "-" Then
                analista = "Analista responsável não mapeado"
            End If
            
            contador_cobranca_telefone = contador_cobranca_telefone + 1
            On Error GoTo error_save_as_arquivo_cobranca_telefone
                
salvar_novamente_arquivo_cobranca_telefone:

        arquivocriado.SaveAs Folder & "Facturas por cobrar - " & analista & " - " & Day(Date) & "." & Month(Date) & "." & Year(Date) & ".xlsx"
        ActiveWorkbook.Close
        
        aba_cobravel_hoje.Activate
        On Error Resume Next
        aba_cobravel_hoje.ShowAllData

            
    Next linha
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''ETAPA CRIAÇÃO DE ARQUIVOS POR CLIENTE COM CONDIÇÃO DICOM/EQUIFAX''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
cobranca_dicom_equifax:

    aba_cobravel_hoje.Range("a1").Activate
    On Error Resume Next
    aba_cobravel_hoje.ShowAllData


    linha_fim = aba_cobravel_hoje.Range("BH200").End(xlUp).Row
    
    If linha_fim = 1 Then
        GoTo fim
    End If
    
    For linha = 2 To linha_fim
        
            analista_responsavel_dicom_equifax = aba_cobravel_hoje.Range("BH" & linha).Value
            tbl2.Range.AutoFilter Field:=39, Criteria1:="Cobrança por E-mail-Dicom/Equifax", Operator:=xlAnd
            tbl2.Range.AutoFilter Field:=40, Criteria1:=analista_responsavel_dicom_equifax, Operator:=xlAnd
            
            linha_fim_aba_cobravel_hoje_selecao_de_linhas = aba_cobravel_hoje.Range("A1").End(xlDown).Row
            
            aba_cobravel_hoje.Range("B1:Q" & linha_fim_aba_cobravel_hoje_selecao_de_linhas).Select
            Selection.SpecialCells(xlCellTypeVisible).Copy
            Workbooks.Add
            Set arquivocriado = ActiveWorkbook
            arquivocriado.Sheets(1).Range("A1").PasteSpecial
            Columns("A:Z").AutoFit
            On Error Resume Next
            ActiveSheet.Name = "Facturas en Dicom.Equifax"
            
            ' verificacao para nomear arquivo caso sejam titulos de clientes em que o analista nao foi mapeado
            If analista_responsavel_dicom_equifax = "-" Then
                analista_responsavel_dicom_equifax = "Analista responsável não mapeado"
            End If
            
            contador_dicom_equifax = contador_dicom_equifax + 1
            On Error GoTo error_save_as_arquivo_dicom_equifax
                
salvar_novamente_arquivo_dicom_equifax:

        arquivocriado.SaveAs Folder & "Facturas en condicion de Dicom.Equifax - cliente(s) de " & analista_responsavel_dicom_equifax & " - " & Day(Date) & "." & Month(Date) & "." & Year(Date) & ".xlsx"
        ActiveWorkbook.Close
        
        aba_cobravel_hoje.Activate
        On Error Resume Next
        aba_cobravel_hoje.ShowAllData

            
    Next linha
    
fim:
    linha_fim = aba_cobravel_hoje.Range("BF200").End(xlUp).Row
    
    For linha = 2 To linha_fim
    
        If aba_cobravel_hoje.Range("BF" & linha).Value <> "" Then
            clientes_nao_mapeados = clientes_nao_mapeados & vbNewLine & "(" & aba_cobravel_hoje.Range("BF" & linha).Value & " - " & aba_cobravel_hoje.Range("BG" & linha).Value & ")"
        End If
    Next linha
    
    linha_fim = aba_cobravel_hoje.Range("BI200").End(xlUp).Row
    
    For linha = 2 To linha_fim
    
        If aba_cobravel_hoje.Range("BI" & linha).Value <> "" Then
            clientes_nao_mapeados = clientes_nao_mapeados & vbNewLine & "(" & aba_cobravel_hoje.Range("BI" & linha).Value & " - " & aba_cobravel_hoje.Range("BJ" & linha).Value & ")"
        End If
    Next linha

        'diferentes msgbox dependendo das condicoes percorridas no codigo
    
 If contador_arquivos_criados_email_mais_de_10_faturas = 0 And contador_cobranca_email_menos_de_10_faturas = 0 And clientes_nao_mapeados = "" And contador_cobranca_telefone = 0 And contador_dicom_equifax = 0 Then
        MsgBox "Nenhum arquivo criado pois não houveram títulos a cobrar, verifique se existem cobranças preventivas às Construtoras. Se sim, prossiga para o envio de e-mails.", vbOKOnly
        aba_export_sap.Activate
        
    Else
        Dim mensagem As String
        mensagem = "Processo Concluído."
        
        If contador_arquivos_criados_email_mais_de_10_faturas > 0 Then
            mensagem = mensagem & vbNewLine & "Foram criados " & contador_arquivos_criados_email_mais_de_10_faturas & " arquivos em " & Folder & " para clientes com mais de 10 faturas."
        End If
        
        If contador_cobranca_email_menos_de_10_faturas > 0 Then
            mensagem = mensagem & vbNewLine & "Serão enviados " & contador_cobranca_email_menos_de_10_faturas & " e-mails de cobrança para clientes com menos de 10 faturas."
        End If
        
        If clientes_nao_mapeados <> "" Then
            mensagem = mensagem & vbNewLine & "Os seguintes clientes não foram mapeados na base de e-mails: " & clientes_nao_mapeados
        End If
        
        If contador_cobranca_telefone > 0 Then
            mensagem = mensagem & vbNewLine & "Foram criados " & contador_cobranca_telefone & " arquivos de cobrança por telefone."
        End If
        
        If contador_dicom_equifax > 0 Then
            mensagem = mensagem & vbNewLine & "Foram criados " & contador_dicom_equifax & " arquivos de dívida a ser publicada no boletim comercial."
        End If
        
        MsgBox mensagem, vbOKOnly
    End If
         
    aba_export_sap.Activate
    
    enviar_email_cobranca_
    
    End
    

    
error_save_as_arquivo_cobranca_email:
    Application.Wait Now + TimeValue("0:00:05")
    GoTo salvar_novamente_arquivo_cobranca_email
    
error_save_as_arquivo_cobranca_telefone:
     Application.Wait Now + TimeValue("0:00:05")
    GoTo salvar_novamente_arquivo_cobranca_telefone
    
    
error_save_as_arquivo_dicom_equifax:
     Application.Wait Now + TimeValue("0:00:05")
    GoTo salvar_novamente_arquivo_dicom_equifax


End Sub
