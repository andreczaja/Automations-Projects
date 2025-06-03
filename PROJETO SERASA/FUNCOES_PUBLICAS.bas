Attribute VB_Name = "FUNCOES_PUBLICAS"
Public Function VerificarPlanilhaDistribuicao(ByVal payer As String, ByVal aba_correspondente As Worksheet, ByVal tipo_processo As String) As Boolean

    Dim resultado As String
    On Error Resume Next
    resultado = Application.VLookup(payer, aba_correspondente.Columns("B:AB"), 27, False)
    On Error GoTo 0
    
    If (resultado = "Sim" And tipo_processo = "I") Or (resultado <> "Sim" And tipo_processo = "E") Then
        VerificarPlanilhaDistribuicao = True
    Else
        VerificarPlanilhaDistribuicao = False
    End If
    
End Function


Public Function TratativasReferencia(ByRef referencia As String) As String

    Dim resultado As String
    Dim char_init As String
    Dim i As Integer
    Dim startNum As Integer
    Dim endNum As Integer
    
    ' Inicializa variáveis
    resultado = ""
    startNum = 0
    endNum = 0

    ' Verifica se a string de referência não está vazia
    If Len(referencia) = 0 Then
        TratativasReferencia = resultado
        Exit Function
    End If

    ' Localiza o índice inicial da sequência numérica
    For i = 1 To Len(referencia)
        char_init = Mid(referencia, i, 1)
        
        If IsNumeric(char_init) And Val(char_init) <> 0 Then
            startNum = i ' Marca o índice inicial dos números
            Exit For
        End If
    Next i

    ' Se nenhum número foi encontrado, retorna string vazia
    If startNum = 0 Then
        TratativasReferencia = resultado
        Exit Function
    End If

    ' Localiza o índice final da sequência numérica
    For i = startNum To Len(referencia)
        char_init = Mid(referencia, i, 1)
        
        If Not IsNumeric(char_init) Then
            endNum = i - 1 ' Marca o índice final dos números
            Exit For
        End If
    Next i

    ' Caso o final não tenha sido definido (números até o final da string)
    If endNum = 0 Then endNum = Len(referencia)

    ' Extrai a sequência numérica com base nos índices encontrados
    resultado = Mid(referencia, startNum, endNum - startNum + 1)

    ' Retorna o resultado
    TratativasReferencia = resultado

End Function



Public Function VerificarBloqueioAdvertencia(ByVal bloqueio As String, ByVal tipo_processo As String) As Boolean
    If tipo_processo = "E" Then
        If bloqueio <> "*" And bloqueio <> "" And UCase(bloqueio) <> "W" Then
            VerificarBloqueioAdvertencia = True
        Else
            VerificarBloqueioAdvertencia = False
        End If
    ElseIf tipo_processo = "I" Then
        If bloqueio = "*" Or bloqueio = "" Or UCase(bloqueio) = "W" Then
            VerificarBloqueioAdvertencia = True
        Else
            VerificarBloqueioAdvertencia = False
        End If
    End If
    
End Function

Public Function VerificarLinhaDuplicada(ByVal aba_base_historica As Worksheet, ByVal aba_correspondente As Worksheet, ByVal linha_aba_correspondente As Integer, ByVal tipo_processo As String) As Boolean
    Dim linha_existente As Boolean
    Dim concatenado_aba_correspondente, concatenado_aba_historica As String
    Dim i, linha_fim_base_historica As Integer
    
    linha_fim_base_historica = aba_base_historica.Range("A1").End(xlDown).Row
    concatenado_aba_correspondente = ConcatenarColunasPayerReferenciaNumDocItem(aba_correspondente, linha_aba_correspondente)

    For i = 2 To linha_fim_base_historica
        concatenado_aba_historica = ConcatenarColunasPayerReferenciaNumDocItem(aba_base_historica, i)
        If concatenado_aba_correspondente = concatenado_aba_historica Then
            If tipo_processo = "E" Then
                If aba_base_historica.Range("AE" & i).Value = "" Then
                    VerificarLinhaDuplicada = True
                Else
                    aba_correspondente.Range("AD" & linha_aba_correspondente).Value = "Linha referente a título incluído e excluído do Serasa"
                    VerificarLinhaDuplicada = False
                End If
                Exit Function
            End If
            VerificarLinhaDuplicada = False
            aba_correspondente.Range("AD" & linha_aba_correspondente).Value = "Linha já existente na base histórica"
            Exit Function
        End If
    Next i
    If tipo_processo = "E" Then
        aba_correspondente.Range("AD" & linha_aba_correspondente).Value = "Linha referente a título incluído e excluído do Serasa"
        VerificarLinhaDuplicada = False
    ElseIf tipo_processo = "I" Then
        VerificarLinhaDuplicada = True
    End If
        
End Function

Public Function ConcatenarColunasPayerReferenciaNumDocItem(ByVal aba As Worksheet, ByVal linha As Integer) As String
    ConcatenarColunasPayerReferenciaNumDocItem = aba.Range("B" & linha).Value & _
                            aba.Range("E" & linha).Value & _
                            aba.Range("I" & linha).Value & _
                            aba.Range("F" & linha).Value

End Function

Public Function ExcluirIncluirLinhaBaseHistorica(ByVal tipo_processo As String, ByVal linha_aba_correspondente As Integer, ByVal aba_correspondente As Worksheet, ByVal aba_base_historica As Worksheet)
    Dim i, linha_fim_base_historica As Integer
    Dim concatenado_aba_correspondente, concatenado_aba_historica As String
    Dim condicao_linha_encontrada As Boolean
    
    condicao_linha_encontrada = False
    
    If aba_base_historica.Range("A2").Value = "" Then
        linha_fim_base_historica = 2
    Else
        linha_fim_base_historica = aba_base_historica.Range("A1").End(xlDown).Row
    End If
    concatenado_aba_correspondente = ConcatenarColunasPayerReferenciaNumDocItem(aba_correspondente, linha_aba_correspondente)
        
        For i = 2 To linha_fim_base_historica
            concatenado_aba_historica = ConcatenarColunasPayerReferenciaNumDocItem(aba_base_historica, i)
            If concatenado_aba_correspondente = concatenado_aba_historica Then
                If tipo_processo = "I" Then
                    aba_correspondente.Range("AD" & linha_aba_correspondente).Value = "Linha já existente na base histórica"
                    Exit Function
                ElseIf tipo_processo = "E" Then
                    aba_base_historica.Range("AE" & i).Value = Date
                    aba_correspondente.Range("AD" & linha_aba_correspondente).Value = "Excluída dívida na base histórica"
                    condicao_linha_encontrada = True
                End If
                
                Exit Function
            End If
        Next i
        
        If tipo_processo = "I" Then
            aba_correspondente.Range("A" & linha_aba_correspondente & ":AC" & linha_aba_correspondente).Copy
            If aba_base_historica.Range("A2").Value = "" And linha_fim_base_historica = 2 Then
                aba_base_historica.Range("A" & linha_fim_base_historica).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                aba_base_historica.Range("AD" & linha_fim_base_historica).Value = Date
                aba_base_historica.Range("AF" & linha_fim_base_historica).Value = remessa_final
            Else
                aba_base_historica.Range("A" & linha_fim_base_historica).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                aba_base_historica.Range("AD" & linha_fim_base_historica).Offset(1, 0).Value = Date
                aba_base_historica.Range("AF" & linha_fim_base_historica).Offset(1, 0).Value = remessa_final
            End If
            
            aba_correspondente.Range("AD" & linha_aba_correspondente).Value = "Incluído na base histórica/Enviado ao SERASA"
            
        ElseIf tipo_processo = "E" And Not condicao_linha_encontrada Then
        
            aba_correspondente.Range("A" & linha_aba_correspondente & ":AC" & linha_aba_correspondente).Copy
            If aba_base_historica.Range("A2").Value = "" And linha_fim_base_historica = 2 Then
                aba_base_historica.Range("A" & linha_fim_base_historica).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                aba_base_historica.Range("AD" & linha_fim_base_historica).Value = "????"
                aba_base_historica.Range("AE" & linha_fim_base_historica).Value = Date
                aba_base_historica.Range("AG" & linha_fim_base_historica).Value = remessa_final
            Else
                aba_base_historica.Range("A" & linha_fim_base_historica).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                aba_base_historica.Range("AD" & linha_fim_base_historica).Offset(1, 0).Value = "????"
                aba_base_historica.Range("AE" & linha_fim_base_historica).Offset(1, 0).Value = Date
                aba_base_historica.Range("AG" & linha_fim_base_historica).Offset(1, 0).Value = remessa_final
            End If
            aba_correspondente.Range("AD" & linha_aba_correspondente).Value = "Excluído na base histórica/Enviado ao SERASA"
        End If
        

End Function


Public Function BuscarPasta(ByRef caminho_pasta As String, Citrix As Boolean) As String
    Dim fso As Object
    Dim pastaGeral As Object
    Dim pastaAutomatizacoesBIRPA As Object
    Dim pastaExcelencia As Object
    Dim pastaSERASA As Object
    Dim pastaArquivoTXTSERASASAP As Object

    
    ' Cria o objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Define a pasta inicial
    If Citrix Then
        On Error Resume Next
        Set pastaGeral = fso.getfolder(Replace(VBA.Environ("USERPROFILE"), "C:\", "\\Client\C$\") & "\OneDrive - Electrolux\")
        On Error GoTo 0
        ' verificação pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.getfolder(Replace(VBA.Environ("USERPROFILE"), "C:\", "\\Client\C$\") & "\OneDrive\")
        End If
    ElseIf Not Citrix Then
        On Error Resume Next
        Set pastaGeral = fso.getfolder(VBA.LCase(VBA.Environ("USERPROFILE")) & "\OneDrive - Electrolux\")
        On Error GoTo 0
        Debug.Print VBA.Environ("USERPROFILE") & "\OneDrive - Electrolux\"
        ' verificação pasta geral com onedrive - electrolux
        If Not VerificarPasta(pastaGeral) Then
            Set pastaGeral = fso.getfolder(VBA.Environ("USERPROFILE") & "\OneDrive\")
        End If
    End If
        
    ' verificação pasta geral só com onedrive
    If Not VerificarPasta(pastaGeral) Then GoTo pedir_caminho_manualmente
    
    On Error Resume Next
    Set pastaAutomatizacoesBIRPA = fso.getfolder(pastaGeral.Path & "\AUTOMATIZAÇÕES, BIs & RPAs\")
    If Not VerificarPasta(pastaAutomatizacoesBIRPA) Then
        Set pastaExcelencia = fso.getfolder(pastaGeral.Path & "\Excelencia\")
    Else
        Set pastaExcelencia = fso.getfolder(pastaAutomatizacoesBIRPA.Path & "\Excelencia\")
    End If
    If Not VerificarPasta(pastaExcelencia) Then
        Set pastaSERASA = fso.getfolder(pastaGeral.Path & "\SERASA\")
    Else
        Set pastaSERASA = fso.getfolder(pastaExcelencia.Path & "\SERASA\")
    End If
    If Not VerificarPasta(pastaSERASA) Then
        Set pastaArquivoTXTSERASASAP = fso.getfolder(pastaGeral.Path & "\Arquivo TXT SERASA SAP\")
    Else
        Set pastaArquivoTXTSERASASAP = fso.getfolder(pastaSERASA.Path & "\Arquivo TXT SERASA SAP\")
    End If
    On Error GoTo 0
    If Not VerificarPasta(pastaArquivoTXTSERASASAP) Then
        GoTo pedir_caminho_manualmente
    Else
        caminho_pasta = pastaArquivoTXTSERASASAP.Path
        GoTo fim
    End If
    
pedir_caminho_manualmente:
    MsgBox "Favor escolher o caminho no seu Computador a ser descarregado o arquivo baixado do site Transbank. Se não possuir a pasta, crie o atalho no Sharepoint e execute novamente a automação." & _
        "(Escolha a pasta no seu computador equivalente à pasta do Sharepoint:" & _
            "Documentos > AUTOMATIZAÇÕES, BIs & RPAs > Excelencia > SERASA > Arquivo TXT SERASA SAP)"
     
    'Seleção da pasta para salvar o arquivo
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' O usuário selecionou uma pasta
            caminho_pasta = .SelectedItems(1) & "\"
        Else
             'O usuário cancelou a seleção da pasta
            MsgBox "Nenhuma pasta selecionada. O processo foi cancelado."
            End
        End If
    End With
    
fim:

    BuscarPasta = caminho_pasta
    ' Libera os objetos
    Set pastaGeral = Nothing
    Set pastaExcelencia = Nothing
    Set fso = Nothing
    Set pastaAutomatizacoesBIRPA = Nothing
    Set pastaArquivoTXTSERASASAP = Nothing
    Set pastaSERASA = Nothing
    
End Function
 
Public Function VerificarPasta(ByVal pasta As Object) As Boolean
    VerificarPasta = True
    If pasta Is Nothing Then
        VerificarPasta = False
    End If
End Function

Public Function AtualizarBase(ByVal aba As Worksheet, ByVal tabela As ListObject, ByVal i_fim As Integer)

    Dim contador_atualizacoes As Integer
    
    Application.Wait (Now + TimeValue("00:00:05"))
    Application.ScreenUpdating = True

    contador_atualizacoes = 1
    Do While contador_atualizacoes < 5
        tabela.QueryTable.BackgroundQuery = False
        tabela.QueryTable.Refresh
        
        If i_fim <> aba.Range("A1048576").End(xlUp).Row Then Exit Do
        
        contador_atualizacoes = contador_atualizacoes + 1
    Loop

End Function


Public Function VerificarFormatoDatas(data As String) As String

Dim tipo, formato_data_tipo_1, formato_data_tipo_2, formato_data_tipo_3, formato_data_tipo_4 As String

formato_data_tipo_1 = "yyyy-mm-dd"
formato_data_tipo_2 = "dd.mm.yyyy"
formato_data_tipo_3 = "yyyy.mm.dd"
formato_data_tipo_4 = "yyyy/mm/dd"

If data = VBA.Format(VBA.DateSerial(2024, 11, 27), formato_data_tipo_1) Then
    tipo = formato_data_tipo_1
       
ElseIf data = VBA.Format(VBA.DateSerial(2024, 11, 27), formato_data_tipo_2) Then
    tipo = formato_data_tipo_2
    
ElseIf data = VBA.Format(VBA.DateSerial(2024, 11, 27), formato_data_tipo_3) Then
    tipo = formato_data_tipo_3
       
ElseIf data = VBA.Format(VBA.DateSerial(2024, 11, 27), formato_data_tipo_4) Then
    tipo = formato_data_tipo_4
    
End If

VerificarFormatoDatas = tipo
End Function



