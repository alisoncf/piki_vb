Attribute VB_Name = "Consignado"


Public Function F_GravaConsign(seq_con As String, cod_pro As String, _
    qtd_con As String, cod_Rep As String, unit_con As String, _
    Optional MudaEst As Boolean = True, Optional retif As Boolean = False) As String

    Dim sRetif As String
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = seq_con
    If retif = True Then
        sRetif = "S"
    Else
        sRetif = "N"
    End If
    
    strSQL = "insert into tab_consignacoes(seq_con,cod_pro,data_con,qtd_con,reman_con,unit_con,"
    strSQL = strSQL & "cod_rep,cod_usu,retif_con)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & F_GravaDec(qtd_con) & ","
    strSQL = strSQL & F_GravaDec(qtd_con) & ","
    strSQL = strSQL & F_GravaDec(unit_con) & ","
    strSQL = strSQL & "'" & cod_Rep & "',"
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & sRetif & "')     "
    Conn.Execute strSQL
'    MsgBox "Passei"
If MudaEst Then
    strSQL = "update tab_produtos set er_pro=er_pro-" & F_GravaDec(qtd_con) & ","
    strSQL = strSQL & " r_pro=r_pro+" & F_GravaDec(qtd_con)
    strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
    Conn.Execute strSQL
    If F_GravaEstCons(cod_Rep, cod_pro, qtd_con) Then
        
        F_GravaConsign = sSeq
       
    Else
        MsgBox "2"
    End If


End If
    Exit Function
trata_Erro:
'    MsgBox "1"
    

End Function

Public Function F_DeletaConsign(seq_con As String, cod_pro As String, _
    qtd_con As String, cod_Rep As String) As Boolean


    Dim strSQL As String
    
    strSQL = "delete from tab_consignacoes where seq_con='" & seq_con & "' "
    
    Conn.Execute strSQL
    
    If F_GravaEstCons(cod_Rep, cod_pro, "-" & qtd_con) Then
    
        
        If F_AlteraEstRes(cod_pro, "-" & qtd) Then
            F_DeletaConsign = True
        End If
    End If
    
    Exit Function
trata_Erro:
    
    

End Function
Public Function F_GravaConsignAberta(cod_pro As String, _
    qtd_con As String, cod_Rep As String, unit_con As String, Optional retif As Boolean = False) As String


    Dim strSQL As String
    Dim sSeq As String
    Dim sRetif As String
    If retif = True Then
        sRetif = "S"
    Else
        sRetif = "N"
    End If
    On Error GoTo trata_Erro
    sSeq = MaxKeySQL("tab_consignabertas", "seq_con", "", "000000#")
    
    strSQL = "insert into tab_consignabertas(seq_con,cod_pro,data_con,qtd_con,unit_con,"
    strSQL = strSQL & "cod_rep,cod_usu,retif_con)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & F_GravaDec(qtd_con) & ","
    
    strSQL = strSQL & F_GravaDec(unit_con) & ","
    strSQL = strSQL & "'" & cod_Rep & "',"
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & sRetif & "')"
    Conn.Execute strSQL
    
    
    
        
        
    F_GravaConsignAberta = sSeq
        
    
    
    Exit Function
trata_Erro:
    
    

End Function
Public Function F_AumentaQtdConsignAberta(seq_con As String, qtd_con As String) As String
    Dim strSQL As String
    Dim sSeq As String
    Dim sRetif As String
    
    On Error GoTo trata_Erro
    
    
    strSQL = "update tab_consignabertas set qtd_con=qtd_con+" & F_GravaDec(qtd_con)
    strSQL = strSQL & " where seq_con='" & seq_con & "' "
    
    Conn.Execute strSQL
    
    
    
        
        
    F_AumentaQtdConsignAberta = True
        
    
    
    Exit Function
trata_Erro:
    
    

End Function

Public Sub P_AlteraQtdConsignAberta(cod_pro As String, _
    cod_Rep As String, qtd_con As String)


    Dim strSQL As String
    Dim sSeq As String
    
    
    strSQL = "update tab_consignabertas set qtd_con=" & F_GravaDec(qtd_con)
    
    
    strSQL = strSQL & "where cod_rep='" & cod_Rep & "' "
    strSQL = strSQL & " and cod_pro='" & cod_pro & "'"
    
    Conn.Execute strSQL
    
    
    
        
        
    
        
    
    
    Exit Sub
trata_Erro:
    
    

End Sub

Public Function F_DeletaConsignAberta(seq_con As String) As Boolean


    Dim strSQL As String
    
    strSQL = "delete from tab_consignabertas where seq_con='" & seq_con & "' "
    
    Conn.Execute strSQL
    
    F_DeletaConsignAberta = True
    
    Exit Function
trata_Erro:
    
    

End Function

Public Function F_GravaEstCons(cod_Rep As String, cod_pro As String, qtd As String) As Boolean
    Dim strSQL As String
    'já altera o estoque reservado
    On Error GoTo trata_Erro
    strSQL = F_BuscaUmCampo("estoque_pro", "tab_estcons", "cod_rep='" & cod_Rep & _
        "' and cod_pro='" & cod_pro & "'")
    
    If strSQL = "" Then
        strSQL = "insert into tab_estcons(cod_rep,cod_pro,estoque_pro)values("
        strSQL = strSQL & "'" & cod_Rep & "',"
        strSQL = strSQL & "'" & cod_pro & "',"
        strSQL = strSQL & F_GravaDec(qtd) & ") "
        Conn.Execute strSQL
    Else
        
        strSQL = "update tab_estcons set estoque_pro=estoque_pro+ " & F_GravaDec(qtd)
        strSQL = strSQL & " where cod_rep='" & cod_Rep & "' "
        strSQL = strSQL & " and cod_pro='" & cod_pro & "' "
        Conn.Execute strSQL
    End If
    F_GravaEstCons = True
   ' P_GeraLog strSQL, "tab_estcons"
    
    Exit Function
trata_Erro:

End Function

Function F_QtdProdutoLancadoDevAberta(cod_pro As String, cod_Rep, seq_con As String) As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select sum(qtd_Dev) as total from tab_devconsignabertas where cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " and cod_rep='" & cod_Rep & "' "
    strSQL = strSQL & " and seq_con='" & seq_con & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        
            F_QtdProdutoLancadoDevAberta = !total
        
        .Close
    End With
trata_Erro:
End Function
Sub imprimeDevConsign(seq_dev As String)
    P_ImprDevConsign frmVisualiza.Pic, seq_dev
    frmVisualiza.Show 1
    frmVisualiza.txtbxCopias.text = "2"
    If frmVisualiza.tag <> "" Then
        For iCount = 1 To Int(frmVisualiza.txtbxCopias.text)
            P_ImprDevConsign Printer, seq_dev
            Printer.EndDoc
        Next
    End If
    Unload frmVisualiza
    
End Sub
Sub P_ImprDevConsign(obj As Object, seq_dev As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim yIni As Single
    Dim cTotal As Currency
    strSQL = "select tr.cod_rep,tp.pv_pro,retif_dev,td.cod_usu,tc.unit_con,tr.nome_rep,tp.nome_pro,tp.ref_pro,td.data_dev,ti.*,tc.data_con "
    strSQL = strSQL & " from tab_representantes as tr,tab_consignacoes as tc, tab_produtos as tp,"
    strSQL = strSQL & " tab_devconsign as td, tab_itdevconsign as ti "
    strSQL = strSQL & " where td.seq_dev=ti.seq_dev "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    
    strSQL = strSQL & " and tr.cod_rep=td.cod_rep "
    strSQL = strSQL & " and tc.seq_con=ti.seq_con "
    strSQL = strSQL & " and td.seq_dev ='" & seq_dev & "' "
    strSQL = strSQL & " and tp.cod_pro=tc.cod_pro  "
    strSQL = strSQL & " order by seq_ite "
    
    abreObjetoAdo rPesq, strSQL
    If rPesq!retif_dev = "N" Then
        P_CabRelTimbr "Guia de Devolução " & seq_dev & " de " & Format(rPesq!data_dev, "dd/mm/yy"), obj, False
    Else
        P_CabRelTimbr "Devolução para Correção de Estoque " & seq_dev & " de " & Format(rPesq!data_dev, "dd/mm/yy"), obj, False
    
        obj.fontsize = 14
        P_CentraImpressao obj, "Correção de Estoque", False
        obj.Print
    End If
    obj.FontName = "times new roman"
    obj.fontsize = 11
    
    P_ImprX obj, "Número:" & seq_dev, 2
    P_ImprX obj, "Data:" & rPesq!data_dev, 30
    
    P_ImprX obj, "Funcionário:", 70
    obj.FontBold = True
    P_ImprX obj, F_BuscaUmCampo("guerra_fun", "tab_funcionarios", "cod_fun='" & rPesq!cod_usu & "'"), 95, False
    obj.FontBold = False
    P_ImprX obj, "Representante:", 2
    obj.FontBold = True
    P_ImprX obj, rPesq!cod_Rep & " - " & rPesq!nome_rep, 30, False
    obj.FontBold = False
    
    
    
    obj.FontBold = False
    obj.FontName = "arial"
    obj.fontsize = 10
    P_line1 obj
    
    yIni = obj.CurrentY
    
    P_ImprX obj, "Código", 6, True, 1
    P_ImprX obj, "Produto", 25, True, 1
    P_ImprX obj, "Qtd.", 125, True, direita
    P_ImprX obj, "Unit.", 140, True, direita
    P_ImprX obj, "Total.", 160, True, direita
    P_ImprX obj, "NºTransf.", 170, True, esquerda
    P_ImprX obj, "Retirada.", 185, False, esquerda
    P_line1 obj
    P_Grades obj, "0 5 24 116 126 146 169 184 200", yIni, obj.CurrentY, vertical, vbSolid
    yIni = obj.CurrentY
    obj.FontName = "arial"
    With rPesq
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, "Guia de Devolução " & seq_con
                P_AddPicVis obj
            End If
            If !unit_con <> !pv_pro Then
                obj.FontItalic = True
                obj.FontBold = True
            Else
                obj.FontItalic = False
                obj.FontBold = False
            End If
            
            P_ImprX obj, .AbsolutePosition + 1, 0, True, 1
            P_ImprX obj, !cod_pro, 6, True, 1
            obj.fontsize = 8
            P_ImprX obj, StrConv(!nome_pro, vbProperCase), 25, True, esquerda
            obj.fontsize = 10
            P_ImprX obj, !qtd_dev, 125, True, direita
            
            P_ImprX obj, FormatVal(!unit_con), 145, True, direita
            P_ImprX obj, FormatVal(!unit_con * !qtd_dev), 168, True, direita
            P_ImprX obj, !seq_con, 170, True, esquerda
            P_ImprX obj, Format(!data_con, "dd/mm/yy"), 185, False, esquerda
            
            cTotal = cTotal + (!unit_con * !qtd_dev)
            P_line1 obj, vbDot
            
            .MoveNext
        Loop
        .Close
    End With
  
    
    
    
    obj.fontsize = 14
    P_ImprX obj, "Total.................: " & FormatVal(cTotal), 190, False, direita
    'P_RodapeVis obj, "Não é Válido como Documento Fiscal nem como Garantia de Mercadoria"
      obj.FontItalic = False
    obj.Print ""
    obj.fontsize = 14
    P_line1 obj
    
    
    P_CentraImpressao obj, "Não é Válido como Documento Fiscal nem como Garantia de Mercadoria"
    
    P_line1 obj
    obj.fontsize = 10
    P_AddPicVis obj
    
    
    
End Sub
Sub P_ListaDevConsign(cod_Rep As String)
    frmListaDevConsign.lblRepresentante.tag = cod_Rep
    frmListaDevConsign.Show 1
    
    
    
    
End Sub
Sub P_VerificaErroEstConsignRemanDifEstoque(Optional cod_Rep As String = "")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select tr.cod_rep,tp.nome_pro as produto,"
    strSQL = strSQL & " tc.cod_pro as cod,"
    strSQL = strSQL & "tr.nome_rep as representante,"
    strSQL = strSQL & "sum(tc.reman_con) as Reman,"
    strSQL = strSQL & "te.estoque_pro as estoque"
    strSQL = strSQL & "from tab_produtos as tp, tab_Estcons as te,"
    strSQL = strSQL & "tab_consignacoes As tc, tab_representantes As tr"
    strSQL = strSQL & "where te.cod_pro = tp.cod_pro"
    strSQL = strSQL & "and tp.cod_pro=tc.cod_pro"
    strSQL = strSQL & "and tc.cod_rep=tr.cod_rep"
    strSQL = strSQL & "and te.cod_pro=tc.cod_pro"
    strSQL = strSQL & "group by tc.cod_pro"
    strSQL = strSQL & "Having Estoque <> reman"
    strSQL = strSQL & "order by tr.cod_rep,tc.cod_pro"

    
End Sub
Function F_GravaPosicaoEstoqueCons(data_fec As String, cod_Rep As String, cod_pro As String, qtd_pro As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_posicaoestoquecons(data_fec,cod_rep,cod_pro,qtd_pro)values("
    strSQL = strSQL & F_GravaData(data_fec) & ","
    strSQL = strSQL & "'" & cod_Rep & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd_pro) & ")"
    Conn.Execute strSQL
    F_GravaPosicaoEstoqueCons = True
    Exit Function
trata_Erro:
    
    
End Function
Function F_AlteraPosicaoEstoqueCons(data_fec As String, cod_Rep As String, cod_pro As String, _
    transferido_pro As String, devolvido_pro As String, acertado_pro As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_posicaoestoquecons set "
    strSQL = strSQL & " transferido_pro=" & F_GravaDec(transferido_pro) & ", "
    strSQL = strSQL & " devolvido_pro=" & F_GravaDec(devolvido_pro) & ", "
    strSQL = strSQL & " acertado_pro=" & F_GravaDec(acertado_pro) & "  "
    strSQL = strSQL & " where data_fec=" & F_GravaData(data_fec)
    strSQL = strSQL & " and cod_rep='" & cod_Rep & "' "
    strSQL = strSQL & " and cod_pro='" & cod_pro & "' "
    Conn.Execute strSQL
    F_AlteraPosicaoEstoqueCons = True
    Exit Function
trata_Erro:
    
    
End Function
Function F_FechaConsign(data_fec As String, qacerto As String, vacertos As String, qtransf As String, _
    vtransf As String, qdev As String, vdev As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_fechaconsign(data_fec,hora_fec, qacertos_fec,vacerto_fec, "
    strSQL = strSQL & "qtransf_fec,vtransf_fec,qdev_fec,vdev_fec,cod_usu)values("
    strSQL = strSQL & F_GravaData(data_fec) & ","
    strSQL = strSQL & "'" & Time & "',"
    strSQL = strSQL & F_GravaDec(qacerto) & ","
    strSQL = strSQL & F_GravaDec(vacerto) & ","
    strSQL = strSQL & F_GravaDec(qtransf) & ","
    strSQL = strSQL & F_GravaDec(vtransf) & ","
    strSQL = strSQL & F_GravaDec(qdev) & ","
    strSQL = strSQL & F_GravaDec(vdev) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    
    Conn.Execute strSQL
    F_FechaConsign = True
    Exit Function
trata_Erro:
    
    
End Function


Function RepresentanteImpressaoBloqueada(seq_con As String) As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select bloqimpr from tab_representantes as tr inner join tab_consignacoes as tc "
    strSQL = strSQL & " on tr.cod_rep=tc.cod_rep where seq_con='" & seq_con & "' "
    
    abreObjetoAdo rPesq, strSQL
    If rPesq!bloqimpr = "S" Then
        RepresentanteImpressaoBloqueada = True
    End If
trata_Erro:
End Function
