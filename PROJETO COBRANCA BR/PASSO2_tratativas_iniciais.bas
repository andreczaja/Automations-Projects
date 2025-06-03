Attribute VB_Name = "PASSO2_tratativas_iniciais"
Public cod_cliente, nf, clientes_nao_mapeados, analista, caminho_manual_cliente As String
Public array_payers, array_analistas

    

Sub tratativas_iniciais_()

    
    Application.ScreenUpdating = False

    Set aba_export_sap = ThisWorkbook.Sheets("Export SAP")
    Set tabela_aba_export_sap = aba_export_sap.ListObjects("Export_FBL5N___Cobráveis")
    Set aba_nao_cobraveis = ThisWorkbook.Sheets("Payers Não Cobraveis")
    Set tabela_aba_nao_cobraveis = aba_nao_cobraveis.ListObjects("Plan_Distr_Não_Cobrar")
    Set aba_relatorio_portal_devolucoes = ThisWorkbook.Sheets("Relatório Portal de Devoluções")
    Set tabela_aba_relatorio_portal_devolucoes = aba_relatorio_portal_devolucoes.ListObjects("Tabela_Relatório_Portal_de_Devoluções")
    
    
    Folder = Replace(BuscarPasta("", False), "Arquivo TXT COBRANCA SAP", "Arquivos de Cobrança")
    
    Pasta_Diaria = Folder & "\" & CStr(VBA.Format(VBA.Date, "dd.mm.yyyy"))

    If Dir(Pasta_Diaria, vbDirectory) = "" Then
        MkDir Pasta_Diaria
        Application.Wait (Now + TimeValue("00:00:10"))
    End If
    
    Do Until Dir(Pasta_Diaria, vbDirectory) <> ""
        Application.Wait (Now + TimeValue("00:00:01"))
    Loop
    
    'tabela_aba_export_sap.QueryTable.Refresh False
    
 
    Call LimparFiltros

    
    linha_fim = aba_export_sap.Range("A6").End(xlDown).Row

    
    ' INSERINDO ARRAY DE CLIENTES QUE POSSUEM NFS A COBRAR
    array_payers = Array(0)
    array_analistas = Array(0)

    For linha = 7 To linha_fim
        cod_cliente = aba_export_sap.Range("C" & linha).Value
        If linha = 7 Then
            array_payers(0) = cod_cliente
        Else
            ' Exibir os itens do array
            For i = LBound(array_payers) To UBound(array_payers)
                If array_payers(i) = cod_cliente Then
                    GoTo proxima_linha_array_payers
                End If
            Next i
            ' Redimensionar o array para acomodar o novo item
            ReDim Preserve array_payers(0 To UBound(array_payers) + 1)
            
            ' Adicionar o novo item ao final do array
            array_payers(UBound(array_payers)) = cod_cliente
        End If
proxima_linha_array_payers:
        analista = aba_export_sap.Range("AF" & linha).Value
        If linha = 7 Then
            array_analistas(0) = analista
        Else
            ' Exibir os itens do array
            For i = LBound(array_analistas) To UBound(array_analistas)
                If array_analistas(i) = analista Then
                    GoTo proxima_linha_array_analistas
                End If
            Next i
            ' Redimensionar o array para acomodar o novo item
            ReDim Preserve array_analistas(0 To UBound(array_analistas) + 1)
            
            ' Adicionar o novo item ao final do array
            array_analistas(UBound(array_analistas)) = analista
        End If
proxima_linha_array_analistas:
        nf = aba_export_sap.Range("F" & linha).Value
        If Not IsError(Application.VLookup(nf, aba_relatorio_portal_devolucoes.Columns("A:A"), 1, False)) Then
            aba_export_sap.Range("AM" & linha).Value = "NF com Ocorrência em aberto no Portal de Devoluções"
        End If
    Next linha
    
    Call criar_arquivos_de_clientes_
    
    
End Sub
