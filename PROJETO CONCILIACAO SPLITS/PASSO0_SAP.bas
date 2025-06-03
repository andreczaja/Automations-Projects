Attribute VB_Name = "PASSO0_SAP"
'variavel que possibilita que o codigo encerre caso se clique no botao de Cancelar no meio
' da execucao do codigo

Public SapGui As Object, WSHShell As Object, session As Object, Applic As Object, connection As Object, SapGuiWB, Window As Object
Public linha, linha_fim, linha_fim_pacotes_pagamento, linha_fim_contas_devedoras, i, i2, x_referencia, x_parcela, x_fecha_pago, y, numero_de_faturas_compensacao, numero_de_pagamentos_compensacao, faturas_selecionadas_base, _
faturas_visiveis_base, soma_todas_as_linhas, contagem_codigos_devedores, contagem_codigos_acreedores As Integer
Public tipo_de_documento, conta_do_cliente, conta_devedora, conta_acreedora, folio, parcela, valor, fecha, texto_chave_referencia_2, pacote_de_pagamento As String
Public aba_geral, aba_linhas_dz As Worksheet
Public tabela_dz As ListObject
Public array_pacotes_pagamento
Public sessao_f32, presenca_de_linha_de_pagamento As Boolean


Sub SAP_Login()

    ' Verifica se o SAP está aberto
    On Error Resume Next
    Set SapGui = GetObject("SAPGUI")
    On Error GoTo 0
    
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
    Set aba_geral = ThisWorkbook.Sheets("Composição_Cliente")

    On Error Resume Next
    aba_geral.ShowAllData
    
    ' pegando pacotes de pagamento
    linha_fim = aba_geral.Range("A999999").End(xlUp).Row
    
    '  preenchendo o array com todos os pacotes de pagamento inseridos
    array_pacotes_pagamento = Array()
    
    For linha = 4 To linha_fim
        conta_do_cliente = aba_geral.Range("A" & linha).Value
        folio = aba_geral.Range("B" & linha).Value
        valor = aba_geral.Range("C" & linha).Value
        parcela = aba_geral.Range("E" & linha).Value
        tipo_de_documento = aba_geral.Range("F" & linha).Value
        pacote_de_pagamento = aba_geral.Range("G" & linha).Value
        
        If tipo_de_documento = "Pagamento" And folio = "" Then
            folio = "-"
        End If
        
        ' Verificar se há campos vazios
        If conta_do_cliente = "" Or folio = "" Or valor = "" Or parcela = "" Or tipo_de_documento = "" Or pacote_de_pagamento = "" Then
            MsgBox "Por favor, preencha todas as colunas obrigatórias da linha " & linha & _
            ", não podem existir linhas com a coluna A, B, C, E, F e G vazias." & vbNewLine & _
            "OBS: Caso seja um cliente que não paga em parcelas, preencha a coluna E com '1'.", vbOKOnly
            End
        End If
        
        ' Se o array estiver vazio, adicione o primeiro pacote de pagamento
        If UBound(array_pacotes_pagamento) = -1 Then
            ReDim Preserve array_pacotes_pagamento(0)
            array_pacotes_pagamento(0) = pacote_de_pagamento
            GoTo proxima_linha_a_verificar_pacote_de_pagamento
        End If
        
        ' Verificar se o pacote já está no array
        For i = LBound(array_pacotes_pagamento) To UBound(array_pacotes_pagamento)
            If array_pacotes_pagamento(i) = pacote_de_pagamento Then
                GoTo proxima_linha_a_verificar_pacote_de_pagamento
            End If
        Next i
        
        ' Redimensionar o array para adicionar o novo pacote de pagamento
        ReDim Preserve array_pacotes_pagamento(0 To UBound(array_pacotes_pagamento) + 1)
        array_pacotes_pagamento(UBound(array_pacotes_pagamento)) = pacote_de_pagamento ' Corrigido aqui
    
proxima_linha_a_verificar_pacote_de_pagamento:
    Next linha

    ' verificando se os pacotes de pagamento estão dentro dos limites para que sejam compensados
    For i = 0 To UBound(array_pacotes_pagamento)
        pacote_de_pagamento = array_pacotes_pagamento(i)
        soma_todas_as_linhas = Application.WorksheetFunction.SumIfs(aba_geral.Range("C4:C" & linha_fim), aba_geral.Range("F4:F" & linha_fim), "Factura Electrolux", aba_geral.Range("G4:G" & linha_fim), pacote_de_pagamento) - _
                                Application.WorksheetFunction.SumIfs(aba_geral.Range("C4:C" & linha_fim), aba_geral.Range("F4:F" & linha_fim), "Factura Acreedora", aba_geral.Range("G4:G" & linha_fim), pacote_de_pagamento) - _
                                Application.WorksheetFunction.SumIfs(aba_geral.Range("C4:C" & linha_fim), aba_geral.Range("F4:F" & linha_fim), "Nota de Crédito", aba_geral.Range("G4:G" & linha_fim), pacote_de_pagamento) - _
                                Application.WorksheetFunction.SumIfs(aba_geral.Range("C4:C" & linha_fim), aba_geral.Range("F4:F" & linha_fim), "Pagamento", aba_geral.Range("G4:G" & linha_fim), pacote_de_pagamento)
        If soma_todas_as_linhas <= -100 Or soma_todas_as_linhas >= 100 Then
            MsgBox "Por favor, verifique as linhas inseridas no pacote de pagamento " & pacote_de_pagamento & ". A soma dos valores (Factura Electrolux - Factura Acreedora - Nota de Crédito - Pagamento) não está entre -100 e 100", vbOKOnly
            End
        End If
    Next i
    
    
    ' verificando se existe mais de um devedor ou mais de um acreedor por pacote de pagamento
    For i = 0 To UBound(array_pacotes_pagamento)
        contagem_codigos_devedores = 0
        conta_devedora = ""
        contagem_codigos_acreedores = 0
        conta_acreedora = ""
        For linha = 4 To linha_fim
            If Left(aba_geral.Range("A" & linha).Value, 1) = 2 And aba_geral.Range("G" & linha).Value = array_pacotes_pagamento(i) Then
                If conta_devedora <> aba_geral.Range("A" & linha).Value Then
                    conta_devedora = aba_geral.Range("A" & linha).Value
                    contagem_codigos_devedores = contagem_codigos_devedores + 1
                End If
            ElseIf Left(aba_geral.Range("A" & linha).Value, 1) = 3 And aba_geral.Range("G" & linha).Value = array_pacotes_pagamento(i) Then
                If conta_acreedora <> aba_geral.Range("A" & linha).Value Then
                    conta_acreedora = aba_geral.Range("A" & linha).Value
                    contagem_codigos_acreedores = contagem_codigos_acreedores + 1
                End If
            End If
            If contagem_codigos_devedores > 1 Or contagem_codigos_acreedores > 1 Then
                MsgBox "Por favor preencha apenas uma conta devedora e uma conta acreedora por pacote de pagamento", vbOKOnly
                End
            End If
        Next linha
    Next i
    
    
    ' tratativas gerais da base
    For linha = 4 To linha_fim
        conta_do_cliente = aba_geral.Range("A" & linha).Value
        parcela = aba_geral.Range("E" & linha).Value
        tipo_de_documento = aba_geral.Range("F" & linha).Value
        ' verificando se foi inserida conta acreedora em parte devedora e vice-versa
        If (tipo_de_documento = "Nota de Crédito" Or tipo_de_documento = "Factura Electrolux" Or tipo_de_documento = "Pagamento") And Left(conta_do_cliente, 1) = "3" Then
            MsgBox "A linha " & linha & " foi identificada como Nota de Crédito/Factura Electrolux/Pagamento." & _
            "Portanto, preencha na coluna 'A' a conta devedora do cliente para continuar.", vbOKOnly
            End
        ElseIf tipo_de_documento = "Factura Acreedora" And Left(conta_do_cliente, 1) = "2" Then
            MsgBox "A linha " & linha & " foi identificada como Factura Acreedora." & _
            "Portanto, preencha na coluna 'A' a conta acreedora do cliente para continuar.", vbOKOnly
            End
        End If
        ' verificando se existem valores negativos na base, se sim, corrige eles
        If (tipo_de_documento = "Pagamento" Or tipo_de_documento = "Nota de Crédito" Or tipo_de_documento = "Factura Acreedora") And aba_geral.Range("C" & linha).Value < 0 Then
            aba_geral.Range("C" & linha).Value = aba_geral.Range("C" & linha).Value * -1
        End If
    Next linha
    
    For i = 0 To UBound(array_pacotes_pagamento)
        presenca_de_linha_de_pagamento = False
        numero_de_faturas_compensacao = Application.WorksheetFunction.CountIfs(aba_geral.Range("F4:F" & linha_fim), "Factura Electrolux", aba_geral.Range("G4:G" & linha_fim), array_pacotes_pagamento(i))
        numero_de_pagamentos_compensacao = Application.WorksheetFunction.CountIfs(aba_geral.Range("F4:F" & linha_fim), "Pagamento", aba_geral.Range("G4:G" & linha_fim), array_pacotes_pagamento(i))
        
        If numero_de_pagamentos_compensacao > 0 Then
            presenca_de_linha_de_pagamento = True
        End If
        
        conta_devedora = ""
        conta_acreedora = ""
        For linha = 4 To linha_fim
            tipo_de_documento = aba_geral.Range("F" & linha).Value
            pacote_de_pagamento = aba_geral.Range("G" & linha).Value
            If tipo_de_documento = "Factura Electrolux" And pacote_de_pagamento = array_pacotes_pagamento(i) Then
                conta_devedora = aba_geral.Range("A" & linha).Value
                Exit For
            End If
        Next linha
        For linha = 4 To linha_fim
            tipo_de_documento = aba_geral.Range("F" & linha).Value
            pacote_de_pagamento = aba_geral.Range("G" & linha).Value
            If tipo_de_documento = "Factura Acreedora" And pacote_de_pagamento = array_pacotes_pagamento(i) Then
                conta_acreedora = aba_geral.Range("A" & linha).Value
                Exit For
            End If
        Next linha
        '''' preencher conta devedora do cliente
        
        Call FBL5N
        Call F32
    Next i
    
    


    Set connection = Nothing
    Set Applic = Nothing
    Set SapGui = Nothing
    Set SapGuiWB = Nothing
    Set session = Nothing
    
End Sub

Sub FBL5N()
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
    session.findById("wnd[0]").sendVKey 0
    ' copiando faturas electrolux para o sap
    On Error Resume Next
    aba_geral.ShowAllData
    aba_geral.Range("A3").AutoFilter Field:=6, Criteria1:="Factura Electrolux"
    aba_geral.Range("A3").AutoFilter Field:=7, Criteria1:=array_pacotes_pagamento(i)
    aba_geral.Range("B4:B" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode "         58"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "         54"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode "         58"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN015_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = conta_devedora
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "TC04"
    session.findById("wnd[0]/usr/chkX_NORM").Selected = True
    session.findById("wnd[0]/usr/chkX_SHBV").Selected = False
    session.findById("wnd[0]/usr/chkX_MERK").Selected = False
    session.findById("wnd[0]/usr/chkX_PARK").Selected = False
    session.findById("wnd[0]/usr/chkX_APAR").Selected = False
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/COBRANCA"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    If session.findById("wnd[0]/sbar").text = "Se visualizan " & CStr(numero_de_faturas_compensacao) & " partidas" Then
        session.findById("wnd[0]").sendVKey 5
    Else
        
        ' pegando coordenadas reais da coluna referencia conforme monitor do usuario
        For x_referencia = 45 To 100
            'Debug.Print session.findById("wnd[0]/usr/lbl[" & x_referencia & ",4]").Text
            On Error Resume Next
            If session.findById("wnd[0]/usr/lbl[" & x_referencia & ",4]").text = "" Then
            ElseIf session.findById("wnd[0]/usr/lbl[" & x_referencia & ",4]").text = "Referencia" Then
                Exit For
            End If
        Next x_referencia
        
        ' pegando coordenadas reais da coluna posição/parcela conforme monitor do usuario
        For x_parcela = 100 To 150
            On Error Resume Next
            'Debug.Print session.findById("wnd[0]/usr/lbl[" & x_parcela & ",4]").Text
            If session.findById("wnd[0]/usr/lbl[" & x_parcela & ",4]").text = "" Then
            ElseIf session.findById("wnd[0]/usr/lbl[" & x_parcela & ",4]").text = "Pos" Then
                Exit For
            End If
        Next x_parcela
        
        ' pegando quantidade de faturas visiseis conforme monitor do usuario
        For y = 6 To 100
            If session.findById("wnd[0]/usr/lbl[" & x_referencia & "," & y & "]").text = "" Then
                faturas_visiveis_base = y - 6
                Exit For
            End If
        Next y
        
        faturas_selecionadas_base = 0
        
    
procurar_na_proxima_pagina:
        ' verificando novamente a quantidade faturas visiveis na base
        For y = 6 To 100
            If session.findById("wnd[0]/usr/lbl[" & x_referencia & "," & y & "]").text = "" Then
                faturas_visiveis_base = y - 6
                Exit For
            End If
        Next y
        ' estrutura que verifica fatura e parcela e marca a partida para que seja troca o campo CLAVE DE REFERENCIA 2
        For y = 6 To faturas_visiveis_base + 5
            folio = session.findById("wnd[0]/usr/lbl[" & x_referencia & "," & y & "]").text
            parcela = Trim(session.findById("wnd[0]/usr/lbl[" & x_parcela & "," & y & "]").text)
            
            If Application.WorksheetFunction.CountIfs(aba_geral.Range("B:B"), folio, aba_geral.Range("E:E"), parcela) <> 0 Then
                session.findById("wnd[0]/usr/chk[1," & y & "]").Selected = True
                faturas_selecionadas_base = faturas_selecionadas_base + 1
            End If
        Next y
        
        ' verificando se todas as linhas de faturas foram selecionadas
        If faturas_selecionadas_base <> numero_de_faturas_compensacao Then
            session.findById("wnd[0]").sendVKey 82
            GoTo procurar_na_proxima_pagina
        End If
    End If
   
    ' CÓDIGO DE ALTERAÇÃO MASSIVA DO CAMPO CHAVE DE REFERENCIA 2
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    session.findById("wnd[1]/usr/txt*BSEG-XREF2").SetFocus
    session.findById("wnd[1]/usr/txt*BSEG-XREF2").text = frm_dados.txt_box_clave_ref_2
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").SetFocus
    While session.ActiveWindow.text = "Información"
        If contador < 60 Then
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            contador = contador + 1
        Else
            MsgBox "A FBL5N demorou demais para carregar, vefique.", vbOKOnly
            End
        End If
    Wend
    
    ' BUSCANDO A(S) LINHA(S) DE PAGAMENTO(S)
    If presenca_de_linha_de_pagamento Then
        faturas_selecionadas_base = 0
        session.findById("wnd[0]/tbar[0]/okcd").text = "/N FBL5N"
        session.findById("wnd[0]").sendVKey 0
        ' copiando faturas electrolux para o sap
        On Error Resume Next
        aba_geral.ShowAllData
        aba_geral.Range("A3").AutoFilter Field:=6, Criteria1:="Pagamento"
        aba_geral.Range("A3").AutoFilter Field:=7, Criteria1:=array_pacotes_pagamento(i)
        aba_geral.Range("C4:C" & linha_fim).SpecialCells(xlCellTypeVisible).Copy
        session.findById("wnd[0]/tbar[1]/btn[16]").press
        session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN018_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/tbar[0]/btn[16]").press
        session.findById("wnd[1]/tbar[0]/btn[24]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/tbar[0]/btn[16]").press
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = conta_devedora
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "TC04"
        session.findById("wnd[0]/usr/chkX_NORM").Selected = True
        session.findById("wnd[0]/usr/chkX_SHBV").Selected = False
        session.findById("wnd[0]/usr/chkX_MERK").Selected = False
        session.findById("wnd[0]/usr/chkX_PARK").Selected = False
        session.findById("wnd[0]/usr/chkX_APAR").Selected = False
        session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/COBRANCA"
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        ' CÓDIGO DE ALTERAÇÃO MASSIVA DO CAMPO CHAVE DE REFERENCIA 2
        If session.findById("wnd[0]/sbar").text = "Se visualizan " & CStr(numero_de_pagamentos_compensacao) & " partidas" Then
            session.findById("wnd[0]").sendVKey 5
        Else
            ' pegando coordenadas reais da coluna posição/parcela conforme monitor do usuario
            For x_fecha_pago = 100 To 150
                On Error Resume Next
                If session.findById("wnd[0]/usr/lbl[" & x_fecha_pago & ",4]").text = "" Then
                ElseIf session.findById("wnd[0]/usr/lbl[" & x_fecha_pago & ",4]").text = "Fecha Pago" Then
                    Exit For
                End If
            Next x_fecha_pago
            For y = 6 To numero_de_pagamentos_compensacao + 5
                If session.findById("wnd[0]/usr/lbl[" & x_fecha_pago & "," & y & "]").text = frm_dados.txt_box_date_insert Then
                    session.findById("wnd[0]/usr/chk[1," & y & "]").Selected = True
                    faturas_selecionadas_base = faturas_selecionadas_base + 1
                End If
            Next y
            If numero_de_pagamentos_compensacao <> faturas_selecionadas_base Then
                MsgBox "Não foram encontradas as linhas de pagamento na base da FBL5N do cliente conforme detalhado na aba da planilha, favor analisar.", vbOKOnly
                End
            End If
        End If
        session.findById("wnd[0]/tbar[1]/btn[45]").press
        session.findById("wnd[1]/usr/txt*BSEG-XREF2").SetFocus
        session.findById("wnd[1]/usr/txt*BSEG-XREF2").text = frm_dados.txt_box_clave_ref_2
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        On Error Resume Next
        session.findById("wnd[1]/tbar[0]/btn[0]").SetFocus
        Debug.Print session.ActiveWindow.text
        While session.ActiveWindow.text = "Información"
            If contador < 60 Then
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                contador = contador + 1
            Else
                MsgBox "A FBL5N demorou demais para carregar, vefique.", vbOKOnly
                End
            End If
        Wend
    End If

End Sub

Sub F32()

Dim criar_nova_linha As Boolean

    session.findById("wnd[0]/tbar[0]/okcd").text = "/N F-32"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRF05A-AGKON").text = conta_devedora
    session.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = frm_dados.txt_box_date_insert
    session.findById("wnd[0]/usr/txtBKPF-MONAT").text = Mid(frm_dados.txt_box_date_insert, 4, 2)
    session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "TC04"
    session.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "CLP"
    session.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[8,0]").SetFocus
    session.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[8,0]").Select
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[0,0]").text = frm_dados.txt_box_clave_ref_2
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/tbar[1]/btn[14]").press

    
    
    criar_nova_linha = True
    For linha = 4 To linha_fim
        conta_acreedora = aba_geral.Range("A" & linha).Value
        folio = aba_geral.Range("B" & linha).Value
        valor = aba_geral.Range("C" & linha).Value
        parcela = aba_geral.Range("E" & linha).Value
        tipo_de_documento = aba_geral.Range("F" & linha).Value
        pacote_de_pagamento = aba_geral.Range("G" & linha).Value
        
        
        If Application.WorksheetFunction.CountIfs(aba_geral.Range("F4:F" & linha_fim), "Factura Acreedora", aba_geral.Range("G4:G" & linha_fim), array_pacotes_pagamento(i)) + _
        Application.WorksheetFunction.CountIfs(aba_geral.Range("F4:F" & linha_fim), "Nota de Crédito", aba_geral.Range("G4:G" & linha_fim), array_pacotes_pagamento(i)) = 0 Then
            GoTo fim
        End If
        If pacote_de_pagamento <> array_pacotes_pagamento(i) Or tipo_de_documento = "Factura Electrolux" Or tipo_de_documento = "Pagamento" Then
            GoTo proxima_linha
        End If
        
        If Not criar_nova_linha Then
            GoTo passos_seguintes_linha_criada
        End If
    
criando_nova_linha:

        If tipo_de_documento = "Nota de Crédito" Then
            session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "14"
            session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = conta_devedora
        ElseIf tipo_de_documento = "Factura Acreedora" And pacote_de_pagamento = array_pacotes_pagamento(i) Then
            session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "24"
            session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = conta_acreedora
        End If
        session.findById("wnd[0]").sendVKey 0
passos_seguintes_linha_criada:
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = valor
        session.findById("wnd[0]/usr/txtBSEG-SKFBT").text = valor
        If texto = "" Then
            session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = folio
        Else
            session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = texto
        End If
        If tipo_de_documento = "Nota de Crédito" Then
            session.findById("wnd[0]/tbar[1]/btn[7]").press
            session.findById("wnd[0]/usr/ctxtBSEG-XREF1").text = folio
        ElseIf tipo_de_documento = "Factura Acreedora" Then
            session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = folio
        End If
        If linha < linha_fim And aba_geral.Range("A" & linha + 1).Value = tipo_de_documento Then
            ' caso linha igual a que já foi inserida anteriormente
            ' botao copiar posicion
            session.findById("wnd[0]/tbar[1]/btn[25]").press
            criar_nova_linha = False
        ElseIf linha < linha_fim And aba_geral.Range("A" & linha + 1).Value <> tipo_de_documento Then
            criar_nova_linha = True
        End If
proxima_linha:
    Next linha
fim:
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    On Error Resume Next
    aba_geral.ShowAllData
    For linha = 4 To linha_fim
        pacote_de_pagamento = aba_geral.Range("G" & linha).Value
        If pacote_de_pagamento = array_pacotes_pagamento(i) Then
            aba_geral.Range("H" & linha).Value = Mid(session.findById("wnd[0]/sbar").text, 5, 9)
        End If
    Next linha

End Sub


