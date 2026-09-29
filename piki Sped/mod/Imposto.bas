Attribute VB_Name = "Imposto"
Type imposto
    cod_imp As String
    perc_imp As Currency
    desc_imp As String
    ecf_imp As String
    
End Type
Function F_GravaImp(perc_imp As String, desc_imp As String, ecf_imp As String) As String
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    sCod = MaxKeySQL("tab_impostos", "cod_sec", "", "00#")
    strSQL = "insert into tab_impostos(cod_imp,desc_imp,perc_imp,ecf_imp)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & desc_imp & "',"
    strSQL = strSQL & F_GravaDec(perc_imp) & ","
    strSQL = strSQL & "'" & ecf_imp & "') "
    Conn.Execute strSQL
    F_GravaImp = sCod
    Exit Function
trata_Erro:
End Function
Function F_AlteraImp(cod_imp As String, perc_imp As String, desc_imp As String, ecf_imp As String) As Boolean
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    F_AlteraImp = False
    
    strSQL = "update tab_impostos set "
    strSQL = strSQL & "desc_imp='" & desc_imp & "',"
    strSQL = strSQL & "perc_imp=" & F_GravaDec(perc_imp) & ","
    strSQL = strSQL & "ecf_imp='" & ecf_imp & "' "
    strSQL = strSQL & " where cod_imp='" & cod_imp & "' "
    Conn.Execute strSQL
    F_AlteraImp = True
    Exit Function
trata_Erro:
End Function
Function F_SelectImp(expr As String) As imposto.imposto
    
    
    frmSelect.sSelect = "select cod_imp as Cod,desc_imp as descrição,perc_imp as perc from tab_impostos where cod_imp<>'' "
        
    frmSelect.cCriterio.Add " and desc_imp like '%" & sExpr & "%'"
    
    frmSelect.sOrd = " order by desc_imp "
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.Col = 0
        F_SelectImp.cod_imp = frmSelect.msflxgrd
        frmSelect.msflxgrd.Col = 1
        F_SelectImp.desc_imp = frmSelect.msflxgrd
        frmSelect.msflxgrd.Col = 2
        F_SelectImp.perc_imp = frmSelect.msflxgrd
        
    End If
    Unload frmSelect
    

End Function
        

Sub P_PreencheImpostos(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    strSQL = "select * from tab_impostos "
    rPesq.Open strSQL, Conn
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!cod_imp & "-" & rPesq!desc_imp
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub
