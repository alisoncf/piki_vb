Attribute VB_Name = "pcp"
Sub P_ImprPcp2(sSeq_Pcp As String, obj As Object)
    Dim iQtd_pag As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
    
    
    
    strSQL = "select tp.cod_pro,tm.nome_mar,tt.nome_cat,tp.nome_pro,sum(ti.qfa_ipc) as soma "
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp, tab_itenspcp as ti, "
    strSQL = strSQL & TAB_MARCAS & " as tm, " & TAB_CATEGORIAS & "  as tt "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar"
    strSQL = strSQL & " and tt.cod_Cat=tp.cod_cat"
    strSQL = strSQL & " and ti.seq_pcp='" & sSeq_Pcp & "' "
    strSQL = strSQL & " and ti.sts_ipc='S' "
    
    strSQL = strSQL & " group by tp.cod_pro"
    rPesq.Open strSQL, conn
    If rPesq.RecordCount = 0 Then
        P_Status ""
        MsgBox "Não constam itens nesta compra!", vbExclamation
        Exit Sub
    Else
        If CCur(rPesq.RecordCount) < 30 Then
            iQtd_pag = 1
        ElseIf CCur(rPesq.RecordCount) Mod 30 = 0 Then
            iQtd_pag = CInt(rPesq.RecordCount) / 30
        Else
            iQtd_pag = rPesq.RecordCount / (rPesq.RecordCount) Mod 30
            iQtd_pag = iQtd_pag + 1
        End If
    End If
    
    P_Status "Imprimindo Cotação..."
    
    P_CabRelTimbr strNome_Emp, obj
    
    P_PrintXY obj, "COTAÇÃO: " & sSeq_Pcp, 170, 0 + 20, False
    P_PrintXY obj, "DATA: " & Date, 170, 3.5 + 20, False
    P_PrintXY obj, "HORA: " & Time, 170, 7 + 20, False
    obj.fontsize = 12
    obj.CurrentY = 25
    P_ImprX obj, "   Solicitamos envio de cotação dos abaixo relacionados: ", 0, False
    
    obj.fontsize = 10
    obj.Print
    obj.FontBold = True
    P_ImprX obj, "Produto", 0, , , 1
    P_ImprX obj, "Ref.:", 90
    P_ImprX obj, "Qtde.", 140, True, True, 1
    P_ImprX obj, "Unit. ", 160, True, True, 1
    P_ImprX obj, "Imp. ", 175, True, True, 1
    P_ImprX obj, "TOTAL", 185, True, , 1
    obj.FontBold = False
    obj.CurrentY = obj.CurrentY + 4
    obj.Line (0, obj.CurrentY)-(200, obj.CurrentY)
    With rPesq
        .MoveFirst
        Do While Not .EOF
            If obj.ScaleHeight <= obj.CurrentY + 50 Then
                'If !sts_ipc = "S" Then
                    If TypeName(obj) = "Printer" Then
                        P_Rodape strEmp_Nome & " - " & strEmp_Cidade & " /" & strEmp_UF, True, "Pag: " & obj.Page & " de " & iQtd_pag
                    End If
                    obj.FontBold = True
                    P_ImprX obj, "Mercadoria ", 0, , , 1
                    P_ImprX obj, "Ref.:", 90, , , 1
                    P_ImprX obj, "Qtde.", 140, True, True, 1
                    P_ImprX obj, "Unit. ", 160, True, True, 1
                    P_ImprX obj, "Imp. ", 175, True, True, 1
                    P_ImprX obj, "TOTAL", 185, True, , 1
                    obj.FontBold = False
                    obj.CurrentY = obj.CurrentY + 4
                    obj.Line (0, obj.CurrentY)-(200, obj.CurrentY)
                'End If
            End If
            'If !sts_ipc = "S" Then
                obj.CurrentY = obj.CurrentY + 2
                P_ImprX obj, !nome_pro, 0, , , 1
                P_ImprX obj, "", 90, , , 1
                P_ImprX obj, FormatVal(!soma) & " ", 140, True, True, 1
                P_ImprX obj, "", 160, True, True, 1
                    P_ImprX obj, "", 175, True, True, 1
                P_ImprX obj, "", 185, True, , 1
                obj.CurrentY = obj.CurrentY + 4
                obj.Line (0, obj.CurrentY)-(200, obj.CurrentY)
            'End If
            .MoveNext
        Loop
    End With
    obj.Print
    P_ImprX obj, "Condições de Pagamento: __________________________________________", 0, True
    P_ImprX obj, "Total:______________", 165, False
        On Error Resume Next
        P_FimRel obj, strNome_Emp & " - " & strCidade_Emp & " /" & strUF_Emp, True, "Pag: " & obj.Page & " de " & iQtd_pag
    
    P_Status ""

End Sub

Sub ImprimePcp2(sSeq_Pcp As String)
        P_ImprPcp sSeq_Pcp, frmVisualiza.Pic
        frmVisualiza.Show 1
        If frmVisualiza.tag <> "" Then
            P_ImprPcp sSeq_Pcp, Printer
            Printer.EndDoc
        End If
        Unload frmVisualiza
End Sub

