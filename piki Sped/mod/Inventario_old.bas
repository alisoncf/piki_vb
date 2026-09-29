Attribute VB_Name = "Inventario"
Function F_GravaInvent() As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    If Not AbreBancoAdo() Then
        Exit Function
    End If
    conn.BeginTrans
    strSQL = "insert into tab_invent("
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & Time & "',"
    strSQL = strSQL & "'" & strCod_Usu & "')"
    
    conn.Execute strSQL
    conn.CommitTrans
    F_GravaInvent = F_BuscaUmCampo("seq_inv", "tab_inventario", "")
    Exit Function
trata_Erro:
    
End Function
Sub P_CabInv(obj As Object, num_inv As String, Data As String, folha As String)
    obj.ScaleMode = Printer.ScaleMode
    obj.FontName = "courier new"
    obj.fontsize = 8
    P_ImprLinha obj, "-", 110
    
    P_CentraImpressao obj, "REGISTRO DE INVENTÁRIO Nº " & num_inv
    P_ImprLinha obj, "-", 110
    P_ImprX obj, "FIRMA..............: " & strNome_Emp, 5, False
    P_ImprX obj, "INSCRIÇÃO ESTADUAL : " & strIE_Emp, 5, True
    P_ImprX obj, "CNPJ...............: " & strCNPJ_Emp, 110, False
    P_ImprX obj, "FOLHA..............: " & folha, 5, True
    P_ImprX obj, "ESTOQUE REFERENTE A: " & Data, 110, False
    P_ImprLinha obj, "-", 110
    
End Sub

Sub P_CabInvProd(obj As Object, saldo_transp As String)
    
    obj.FontName = "courier new"
    obj.fontsize = 8
    Dim iX As Single
    Dim iY As Single
    Dim Fx As Single
    Dim Fy As Single
    P_ImprLinha obj, "-", 110
    
    
    iY = obj.CurrentY
    iX = obj.CurrentX
    P_ImprX obj, "CLS.FSC.", 3
    P_ImprX obj, "DISCRIMINAÇÃO: ", 55, True
    P_ImprX obj, "UNIDADE", 115
    P_ImprX obj, "QUANTIDADE", 150, True, 3
    P_ImprX obj, "UNITÁRIO", 172, True, 3
    P_ImprX obj, "TOTAL", 190, False, 3
    P_ImprLinha obj, "-", 110
    P_ImprX obj, "SALDO DE TRANSPORTE===========>: ", 20, True
    P_ImprX obj, saldo_transp, 198, False, 3
    P_ImprLinha obj, "-", 110
    Fy = obj.CurrentY
    Fx = obj.CurrentX
    
    P_Line obj, iX, iY - 1.5, iX, Fy - 1.5, vbDot
    P_Line obj, 19, iY - 1.5, 19, Fy - 1.5, vbDot
    P_Line obj, 114, iY - 1.5, 114, Fy - 1.5, vbDot
    P_Line obj, 129, iY - 1.5, 129, Fy - 1.5, vbDot
    P_Line obj, 154, iY - 1.5, 154, Fy - 1.5, vbDot
    P_Line obj, 176, iY - 1.5, 176, Fy - 1.5, vbDot
    P_Line obj, 202, iY - 1.5, 202, Fy - 1.5, vbDot
    
    
End Sub

Sub P_ItemInv(obj As Object, class As String, discrimina As String, _
    unidade As String, qtd As String, unit As String, total As String)
    
    obj.FontName = "courier new"
    
    obj.fontsize = 8
    
    
    P_ImprX obj, class, 5
    P_ImprX obj, discrimina, 20, True
    P_ImprX obj, unidade, 120, True
    P_ImprX obj, qtd, 150, True, 3
    P_ImprX obj, unit, 175, True, 3
    P_ImprX obj, total, 198, False, 3
    
    
    
    
    
End Sub
Sub P_RodaPeInvProd(obj As Object, saldo_transp As String)
    Dim Fy As Single
    Dim Fx As Single
    Dim iY As Single
    obj.FontName = "courier new"
    P_ImprLinha obj, "-", 110
    iY = 45
    
    
    Fy = obj.CurrentY
    P_ImprX obj, "SALDO A TRANSPORTAR ===========>: ", 20, True
    P_ImprX obj, saldo_transp, 198, False, 3
    P_ImprLinha obj, "-", 110
    Fx = obj.CurrentX
    
    
    
    P_Line obj, 19, iY - 1.5, 19, Fy - 1.5, vbDot
    P_Line obj, 114, iY - 1.5, 114, Fy - 1.5, vbDot
    P_Line obj, 129, iY - 1.5, 129, Fy - 1.5, vbDot
    P_Line obj, 154, iY - 1.5, 154, Fy - 1.5, vbDot
    P_Line obj, 176, iY - 1.5, 176, Fy - 1.5, vbDot
    P_Line obj, 202, iY - 1.5, 202, Fy - 1.5, vbDot
    On Error Resume Next
    'If TypeName(obj) = "printer" Then
        obj.NewPage
    'End If
    Err.Clear
End Sub
Sub P_RodaPeInvFinal(obj As Object, saldo_transp As String)
    
    obj.FontName = "courier new"
    P_ImprLinha obj, "-", 110
    
    
    P_ImprLinha obj, "-", 110
    
    P_ImprX obj, "SALDO DO INVENTÁRIO ===========>: ", 20, True
    P_ImprX obj, saldo_transp, 198, False, 3
    P_ImprLinha obj, "-", 110
    
    
    
    
End Sub


Sub P_CabImo(obj As Object)
    obj.ScaleMode = Printer.ScaleMode
    obj.FontName = "courier new"
    obj.fontsize = 8
 '   P_ImprLinha obj, "-", 110
    P_CentraImpressao obj, "A T I V O   I M O B I L I Z A D O"
'    P_ImprLinha obj, "-", 110
    
    
End Sub

