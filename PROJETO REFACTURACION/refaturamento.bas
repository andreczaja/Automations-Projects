Attribute VB_Name = "refaturamento"
Sub INTERACAO_SAP_()

Dim clase_pedido, organizacion_ventas, canal_distribucion, sector, oficina_de_ventas, grupo_de_vendedores, cod_cliente, solicitante, num_ped_cliente, material, denominacion As String
Dim nova_asignacion, nova_nc_fac_gerada, CICd, importe, centro, asignacion, referencia, observaciones_electronica, observaciones_referencia_fatura As String
Dim array_construtoras
Dim condicao_cliente_construtora As Boolean
Dim session
Dim plan_rpa_refaturacao As Workbook
Dim aba_rpa_refaturacao, aba_lista_contrutoras As Worksheet
Dim i, i2, linha, linha_fim, contador_notas_de_credito_geradas, contador_faturas_geradas As Integer


    Set plan_rpa_refaturacao = ThisWorkbook
    Set aba_rpa_refaturacao = plan_rpa_refaturacao.Sheets("VA01")
    Set aba_lista_contrutoras = plan_rpa_refaturacao.Sheets("LISTA CONSTRUTORAS")
    
            contador_faturas_geradas = 0
            contador_notas_de_credito_geradas = 0

            linha_fim = aba_lista_contrutoras.Range("A2").End(xlDown).Row
            ' Redimensionando o array de clientes
            ReDim array_construtoras(1 To linha_fim - 1)
            
            ' Preenchendo o array com os códigos de clientes
            For linha = 2 To linha_fim ' Começando da linha 2
                array_construtoras(linha - 1) = aba_lista_contrutoras.Range("A" & linha)
            Next linha
            
    
            
            ' INICIO INTERACAO COM SAP
            
            If Not IsObject(App) Then
               Set SapGuiAuto = GetObject("SAPGUI")
               Set App = SapGuiAuto.GetScriptingEngine
            End If
            If Not IsObject(Connection) Then
               Set Connection = App.Children(0)
            End If
            If Not IsObject(session) Then
               Set session = Connection.Children(0)
            End If
            If IsObject(WScript) Then
               WScript.ConnectObject session, "on"
               WScript.ConnectObject App, "on"
            End If
    
            'COMEÇCANDO A INTERACAO COM O SAP E ENTRAR NA VA01
    
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nBPMDG/UTL_BROWSER"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[0]/okcd").Text = "VA01"
            session.findById("wnd[0]").sendVKey 0
            
            linha = 3
            
            If aba_rpa_refaturacao.Range("B4").Value = "" Then
                linha_fim = 3
            Else
                linha_fim = aba_rpa_refaturacao.Range("B3").End(xlDown).Row
            End If
    
            condicao_cliente_construtora = False
            
            For i2 = linha To linha_fim
                
                clase_pedido = aba_rpa_refaturacao.Range("B" & i2).Value
                organizacion_ventas = aba_rpa_refaturacao.Range("C" & i2).Value
                canal_distribucion = aba_rpa_refaturacao.Range("D" & i2).Value
                sector = aba_rpa_refaturacao.Range("E" & i2).Value
                oficina_de_ventas = aba_rpa_refaturacao.Range("F" & i2).Value
                solicitante = aba_rpa_refaturacao.Range("I" & i2).Value
                num_ped_cliente = aba_rpa_refaturacao.Range("J" & i2).Value
                material = aba_rpa_refaturacao.Range("N" & i2).Value
                denominacion = aba_rpa_refaturacao.Range("P" & i2).Value
                CICd = aba_rpa_refaturacao.Range("T" & i2).Value
                importe = aba_rpa_refaturacao.Range("U" & i2).Value
                centro = aba_rpa_refaturacao.Range("W" & i2).Value
                asignacion = aba_rpa_refaturacao.Range("Y" & i2).Value
                referencia = aba_rpa_refaturacao.Range("Z" & i2).Value
                observaciones_electronica = aba_rpa_refaturacao.Range("AB" & i2).Value
                observaciones_referencia_fatura = aba_rpa_refaturacao.Range("AC" & i2).Value
                nova_asignacion = aba_rpa_refaturacao.Range("AE" & i2).Value
                nova_nc_fac_gerada = aba_rpa_refaturacao.Range("AF" & i2).Value
                
                If clase_pedido = "" Or organizacion_ventas = "" Or canal_distribucion = "" Or sector = "" Or oficina_de_ventas = "" Or solicitante = "" Or _
                num_ped_cliente = "" Or material = "" Or denominacion = "" Or CICd = "" Or importe = "" Or centro = "" Or asignacion = "" Or referencia = "" Then
                    MsgBox ("Favor conferir a linha " & i2 & " há colunas obrigatórias não preenchidas."), vbOKOnly
                    End
                ElseIf clase_pedido <> "ZC80" And clase_pedido <> "ZCSV" Then
                    MsgBox ("Não foram parametrizadas linhas de refaturamento com clase de pedido diferente de ZSCV e ZC80, favor conferir a linha " & i2 & "."), vbOKOnly
                    End
                End If
                
            Next i2

    
    
    For i = linha To linha_fim
    
            clase_pedido = aba_rpa_refaturacao.Range("B" & i).Value
            organizacion_ventas = aba_rpa_refaturacao.Range("C" & i).Value
            canal_distribucion = aba_rpa_refaturacao.Range("D" & i).Value
            sector = aba_rpa_refaturacao.Range("E" & i).Value
            oficina_de_ventas = aba_rpa_refaturacao.Range("F" & i).Value
            solicitante = aba_rpa_refaturacao.Range("I" & i).Value
            num_ped_cliente = aba_rpa_refaturacao.Range("J" & i).Value
            material = aba_rpa_refaturacao.Range("N" & i).Value
            denominacion = aba_rpa_refaturacao.Range("P" & i).Value
            CICd = aba_rpa_refaturacao.Range("T" & i).Value
            importe = aba_rpa_refaturacao.Range("U" & i).Value
            centro = aba_rpa_refaturacao.Range("W" & i).Value
            asignacion = aba_rpa_refaturacao.Range("Y" & i).Value
            referencia = aba_rpa_refaturacao.Range("Z" & i).Value
            observaciones_electronica = aba_rpa_refaturacao.Range("AB" & i).Value
            observaciones_referencia_fatura = aba_rpa_refaturacao.Range("AC" & i).Value
            nova_asignacion = aba_rpa_refaturacao.Range("AE" & i).Value
            nova_nc_fac_gerada = aba_rpa_refaturacao.Range("AF" & i).Value
           
           

            
            ' PROCEDIMENTO DE EMISSAO DE NC
            session.findById("wnd[0]/usr/ctxtVBAK-AUART").Text = clase_pedido
            session.findById("wnd[0]/usr/ctxtVBAK-VKORG").Text = organizacion_ventas
            session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").Text = canal_distribucion
            session.findById("wnd[0]/usr/ctxtVBAK-SPART").Text = sector
            session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").Text = oficina_de_ventas
            
            ' REDIRECIONAMENTO CONFORME INFORMACOES PREENCHIDAS NAS LINHAS
            
            If nova_asignacion = "" And nova_nc_fac_gerada <> "" Then
                MsgBox "Favor verificar a linha " & i & " pois a coluna de nova asignacion está vazia e a de nova referência, não!", vbOKOnly
                End
            End If
            
            If nova_asignacion <> "" And nova_nc_fac_gerada = "" And clase_pedido = "ZC80" Then
                GoTo etapa_idcp_nc
            ElseIf nova_asignacion <> "" And nova_nc_fac_gerada = "" And clase_pedido = "ZCSV" Then
                GoTo etapa_idcp_factura
            ElseIf nova_asignacion <> "" And nova_nc_fac_gerada <> "" Then
                GoTo proxima_nota
            End If
            
            
            If UCase(clase_pedido) = "ZC80" Then
            
                    ' CONDICIONAL QUE PREENCHE G03 PARA RESTO DOS CLIENTES E G04 SOMENTE CONSTRUTORAS NO CAMPO
                    ' GRUPO DE VENDEDORES DA PRIMEIRA TELA
                    
                    For i2 = LBound(array_construtoras) To UBound(array_construtoras)
                        If array_construtoras(i2) = cod_cliente Then
                            condicao_cliente_construtora = True
                        End If
                    Next i2
            
                    If condicao_cliente_construtora Then
                        session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").Text = "G04"
                    Else
                        session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").Text = "G03"
                    End If
            
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = "FAE0" & num_ped_cliente ' COLUNA J FATURA
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").Text = Format(Date, "dd.mm.yyyy") ' DATA DE HOJE
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").Text = solicitante ' CODIGO SAP COLUNA I
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/ctxtVBKD-FKDAT").Text = Format(Date, "dd.mm.yyyy") ' DATA DE HOJE
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/cmbVBAK-FAKSK").Key = " "
                    session.findById("wnd[0]").sendVKey 0
                    On Error Resume Next
                    session.findById("wnd[1]/tbar[0]/btn[12]").press
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4414/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").Key = "Z03"
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4414/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").Text = material
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4414/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,0]").Text = "1"
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4414/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ARKTX[6,0]").Text = "Refacturacion FAE0" & num_ped_cliente
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4414/subSUBSCREEN_TC:SAPMV45A:4902/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_PKON").press
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4414/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST").getAbsoluteRow(0).Selected = True
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4414/subSUBSCREEN_TC:SAPMV45A:4902/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_PKON").press
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,3]").Text = CICd
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,3]").Text = importe
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\03").Select
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\03/ssubSUBSCREEN_BODY:SAPMV45A:4452/ctxtVBAP-WERKS").Text = centro
                    session.findById("wnd[0]/tbar[0]/btn[3]").press
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\06").Select
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-ZUONR").Text = asignacion
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text = referencia
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10").Select
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = observaciones_electronica
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "ZCT8", "Column1"
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "ZCT8", "Column1"
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = observaciones_referencia_fatura
                    session.findById("wnd[0]/mbar/menu[0]/menu[9]").Select
                    session.findById("wnd[0]/tbar[0]/btn[11]").press
                    
                    
                    ' colocando numero da asignacion na coluna AE da aba base
                    nova_asignacion = Mid(session.findById("wnd[0]/sbar").Text, 11, 10)
                    aba_rpa_refaturacao.Range("AE" & i).Value = nova_asignacion
                    
etapa_idcp_nc:
                    ' INICIO PROCEDIMENTO IDCP
                    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N IDCP"
                    session.findById("wnd[0]").sendVKey 0
                    
                    session.findById("wnd[0]/usr/ctxtVKORG").Text = "vct4"
                    session.findById("wnd[0]/usr/ctxtVTWEG").Text = ""
                    session.findById("wnd[0]/usr/ctxtLOTNO").Text = "ct04"
                    session.findById("wnd[0]/usr/ctxtBOKNO").Text = "01"
                    session.findById("wnd[0]/tbar[1]/btn[8]").press
                    session.findById("wnd[1]/usr/ctxtPR_NUM").Text = "LOCL"
                    session.findById("wnd[1]/usr/ctxtVBELN-LOW").Text = nova_asignacion
                    session.findById("wnd[1]/usr/ctxtMSG_TYPE").Text = "ZC02"
                    
                    nova_nc_fac_gerada = "NCE00" & session.findById("wnd[1]/usr/txtPR_LOW").Text
                    aba_rpa_refaturacao.Range("AF" & i).Value = nova_nc_fac_gerada
                    
                    session.findById("wnd[1]/tbar[0]/btn[8]").press
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
                    session.findById("wnd[0]/tbar[1]/btn[46]").press
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
                    
                    contador_notas_de_credito_geradas = contador_notas_de_credito_geradas + 1
                    
                    ' voltando na transacao para emitir a fatura
                    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N VA01"
                    session.findById("wnd[0]").sendVKey 0
                    
                    ' INICIO DO PROCEDIMENTO DA NOVA FATURA
            ElseIf UCase(clase_pedido) = "ZCSV" Then
            
            
                    ' CONDICIONAL QUE PREENCHE G03 PARA RESTO DOS CLIENTES E G04 SOMENTE CONSTRUTORAS NO CAMPO
                    ' GRUPO DE VENDEDORES DA PRIMEIRA TELA
                    
                    For i2 = LBound(array_construtoras) To UBound(array_construtoras)
                        If array_construtoras(i2) = cod_cliente Then
                            condicao_cliente_construtora = True
                        End If
                    Next i2
            
                    If condicao_cliente_construtora Then
                        session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").Text = "G04"
                    Else
                        session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").Text = "G03"
                    End If
            
                    session.findById("wnd[0]").sendVKey 0
                    
                    If Len(num_ped_cliente) = 7 Then
                        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = "FAE0" & num_ped_cliente ' COLUNA J FATURA
                    Else
                        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text = num_ped_cliente
                    End If
                    
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").Text = Format(Date, "dd.mm.yyyy") ' DATA DE HOJE
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").Text = solicitante ' CODIGO SAP COLUNA I
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-FAKSK").Key = " "
                    session.findById("wnd[0]").sendVKey 0
                    
                    On Error Resume Next
                    session.findById("wnd[1]/tbar[0]/btn[12]").press
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").Text = material
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").Text = "1"
                    
                    If Len(num_ped_cliente) = 7 Then
                        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[5,0]").Text = "Refacturacion FAE0" & num_ped_cliente
                    Else
                        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[5,0]").Text = "Refacturacion " & num_ped_cliente
                    End If
                    
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").getAbsoluteRow(0).Selected = True
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_PKON").press
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,3]").Text = CICd
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,3]").Text = importe
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\03").Select
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\03/ssubSUBSCREEN_BODY:SAPMV45A:4452/ctxtVBAP-WERKS").Text = centro
                    session.findById("wnd[0]/tbar[0]/btn[3]").press
                    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\05").Select
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\05/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-ZUONR").Text = asignacion
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").Text = referencia
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = observaciones_electronica
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "ZCT8", "Column1"
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "ZCT8", "Column1"
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = observaciones_referencia_fatura
                    session.findById("wnd[0]/mbar/menu[0]/menu[9]").Select
                    session.findById("wnd[0]/tbar[0]/btn[11]").press
                    
                     ' colocando numero da asignacion na coluna AE da aba base
                    nova_asignacion = Mid(session.findById("wnd[0]/sbar").Text, 11, 10)
                    aba_rpa_refaturacao.Range("AE" & i).Value = nova_asignacion
                    
etapa_idcp_factura:
                    ' INICIO PROCEDIMENTO IDCP
                    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N IDCP"
                    session.findById("wnd[0]").sendVKey 0
                    
                    session.findById("wnd[0]/usr/ctxtVKORG").Text = "vct4"
                    session.findById("wnd[0]/usr/ctxtVTWEG").Text = ""
                    session.findById("wnd[0]/usr/ctxtLOTNO").Text = "ct02"
                    session.findById("wnd[0]/usr/ctxtBOKNO").Text = "14"
                    session.findById("wnd[0]/tbar[1]/btn[8]").press
                    session.findById("wnd[1]/usr/ctxtPR_NUM").Text = "LOCL"
                    session.findById("wnd[1]/usr/ctxtVBELN-LOW").Text = nova_asignacion
                    session.findById("wnd[1]/usr/ctxtMSG_TYPE").Text = "ZC01"
                    
                    nova_nc_fac_gerada = "FAE0" & session.findById("wnd[1]/usr/txtPR_LOW").Text
                    aba_rpa_refaturacao.Range("AF" & i).Value = nova_nc_fac_gerada
                    
                    session.findById("wnd[1]/tbar[0]/btn[8]").press
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
                    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
                    session.findById("wnd[0]/tbar[1]/btn[46]").press
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
                    
                    contador_faturas_geradas = contador_faturas_geradas + 1
                    
                    ' voltando na transacao para emitir a fatura
                    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N VA01"
                    session.findById("wnd[0]").sendVKey 0
            End If
            
proxima_nota:
    Next i
            
            If contador_faturas_geradas > 0 And contador_notas_de_credito_geradas > 0 Then
                MsgBox ("Processo de refaturamento concluído!" & vbNewLine & "Foram geradas:" & vbNewLine & contador_notas_de_credito_geradas & _
                    " Nota(s) de Crédito " & vbNewLine & contador_faturas_geradas & " Fatura(s)."), vbOKOnly
            ElseIf contador_faturas_geradas = 0 And contador_notas_de_credito_geradas > 0 Then
                MsgBox ("Processo de refaturamento concluído!" & vbNewLine & "Foram geradas:" & vbNewLine & contador_notas_de_credito_geradas & _
                    " Nota(s) de Crédito "), vbOKOnly
            ElseIf contador_faturas_geradas > 0 And contador_notas_de_credito_geradas = 0 Then
                MsgBox ("Processo de refaturamento concluído!" & vbNewLine & "Foram geradas:" & vbNewLine & contador_faturas_geradas & " Fatura(s)."), vbOKOnly
            ElseIf contador_faturas_geradas = 0 And contador_notas_de_credito_geradas = 0 Then
                MsgBox "Nenhum documento foi emitido. Favor revisar", vbOKOnly
            End If
    End Sub
