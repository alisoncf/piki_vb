Attribute VB_Name = "Pacote"
Type pacote
    seq_pac As String
    num_ven As String
    data_pac As String
    enviou_pac As String
    despachou_pac As String
    status_pac As String
    
    

End Type
Function F_GravaPacote(num_ven As String) As String
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_erro
    
    sSeq = MaxKeySQL("tab_pacotes", "seq_pac", "", "0000#")
    
    
    strSQL = "insert into tab_pacotes(seq_pac,num_ven,data_pac,enviou_pac,"
    strSQL = strSQL & "status_pac)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'1')"
    conn.Execute strSQL
    
    Exit Function
trata_erro:
        
    
    
End Function

Function F_BaixaPac(seq_pac As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_erro
    F_BaixaPac = True
    strSQL = "update tab_pacotes set status_pac='2' where seq_pac='" & seq_pac & "'"
    conn.Execute strSQL
trata_erro:
    
    Exit Function
    F_BaixaPac = False
End Function
