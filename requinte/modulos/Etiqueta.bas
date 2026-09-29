Attribute VB_Name = "Etiqueta"
Type etiqueta
    cod As String
    desc As String
    Altura As Single
    distH As Single
    distV As Single
    Largura As Single
    mEsq As Single
    mSup As Single
    qtdCol As Integer
    qtdLinha As Integer
    fonte As String
    tamfonte As String
    tambar As Integer
    altbar As Integer
    
End Type

Function F_BuscaEtiqueta(sExpr As String) As etiqueta.etiqueta
        On Error GoTo trata_Erro
        frmEdiEti.tag = "x"
        frmEdiEti.Show 1
        If frmEdiEti.tag <> "" Then
            With frmEdiEti.ado
                F_BuscaEtiqueta.Altura = !altura_eti
                F_BuscaEtiqueta.Largura = !largura_eti
                F_BuscaEtiqueta.cod = !cod_eti
                F_BuscaEtiqueta.desc = !nome_eti
                F_BuscaEtiqueta.distH = !disth_eti
                F_BuscaEtiqueta.distV = !distv_eti
                F_BuscaEtiqueta.mEsq = !mesq_eti
                F_BuscaEtiqueta.mSup = !msup_eti
                F_BuscaEtiqueta.qtdCol = !qtdcol_eti
                F_BuscaEtiqueta.qtdLinha = !qtdlinha_eti
                F_BuscaEtiqueta.fonte = !fonte_eti
                F_BuscaEtiqueta.tamfonte = !tamfonte_eti
                F_BuscaEtiqueta.tambar = !tambar_eti
                F_BuscaEtiqueta.altbar = !altbar_eti
            End With
        End If
    Exit Function
trata_Erro:
        P_MostraErro
End Function



Function F_BuscaEtiquetaPadrao() As etiqueta.etiqueta
        On Error GoTo trata_Erro
        Dim strSQL As String
        Dim rPesq As New ADODB.Recordset
        strSQL = "select * from tab_etiquetas where padrao_eti='S'"
        abreObjetoAdo rPesq, strSQL
        
            With rPesq
                If .RecordCount > 0 Then
                    F_BuscaEtiquetaPadrao.Altura = !altura_eti
                    F_BuscaEtiquetaPadrao.Largura = !largura_eti
                    F_BuscaEtiquetaPadrao.cod = !cod_eti
                    F_BuscaEtiquetaPadrao.desc = !nome_eti
                    F_BuscaEtiquetaPadrao.distH = !disth_eti
                    F_BuscaEtiquetaPadrao.distV = !distv_eti
                    F_BuscaEtiquetaPadrao.mEsq = !mesq_eti
                    F_BuscaEtiquetaPadrao.mSup = !msup_eti
                    F_BuscaEtiquetaPadrao.qtdCol = !qtdcol_eti
                    F_BuscaEtiquetaPadrao.qtdLinha = !qtdlinha_eti
                    F_BuscaEtiquetaPadrao.fonte = !fonte_eti
                    F_BuscaEtiquetaPadrao.tamfonte = !tamfonte_eti
                    F_BuscaEtiquetaPadrao.tambar = !tambar_eti
                    F_BuscaEtiquetaPadrao.altbar = !altbar_eti
                End If
                .Close
            End With
        
    Exit Function
trata_Erro:
        P_MostraErro
End Function



Sub P_AlteraPadrao(cod_eti As String)
    Dim strSQL As String
    strSQL = "update tab_etiquetas set padrao_eti='N' "
    Conn.Execute strSQL
    
    strSQL = "update tab_etiquetas set padrao_eti='S' "
    strSQL = strSQL & " where cod_eti='" & cod_eti & "' "
    Conn.Execute strSQL
    
    
    
End Sub

Function F_GravaEti(nome_eti As String, altura_eti As String, _
    disth_eti As String, distv_eti As String, largura_eti As String, mesq_eti As String, _
    msup_eti As String, qtdcol_eti As String, qtdlinha_eti As String, fonte As String, _
    tamfonte As String, tambar As String, altbar As String) As String
    
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = MaxKeySQL("tab_etiquetas", "cod_eti", "", "")
    strSQL = "insert into tab_etiquetas(nome_eti,cod_eti,altura_eti,disth_eti,distv_eti,"
    strSQL = strSQL & "largura_eti,mesq_eti,msup_eti,qtdcol_eti,qtdlinha_eti,fonte_eti,tamfonte_eti,tambar_eti,altbar_eti)values("
    strSQL = strSQL & "'" & nome_eti & "',"
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaDec(altura_eti) & ","
    strSQL = strSQL & F_GravaDec(disth_eti) & ","
    strSQL = strSQL & F_GravaDec(distv_eti) & ","
    strSQL = strSQL & F_GravaDec(largura_eti) & ","
    strSQL = strSQL & F_GravaDec(mesq_eti) & ","
    strSQL = strSQL & F_GravaDec(msup_eti) & ","
    strSQL = strSQL & F_GravaDec(qtdcol_eti) & ","
    strSQL = strSQL & F_GravaDec(qtdlinha_eti) & ","
    strSQL = strSQL & "'" & fonte & "',"
    strSQL = strSQL & F_GravaDec(tamfonte) & ","
    strSQL = strSQL & F_GravaDec(tambar) & ","
    strSQL = strSQL & F_GravaDec(altbar) & ")"
    Conn.Execute strSQL
    F_GravaEti = sSeq
    Exit Function
    
trata_Erro:
    
End Function
    
Function F_AlteraEti(cod_eti As String, nome_eti As String, altura_eti As String, _
    disth_eti As String, distv_eti As String, largura_eti As String, mesq_eti As String, _
    msup_eti As String, qtdcol_eti As String, qtdlinha_eti As String, _
        fonte As String, tamfonte As String, tambar As String, altbar As String) As Boolean
    
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    
    strSQL = "update tab_etiquetas set nome_eti='" & nome_eti & "', "
    strSQL = strSQL & "altura_eti=" & F_GravaDec(altura_eti) & ","
    strSQL = strSQL & "disth_eti=" & F_GravaDec(disth_eti) & ","
    strSQL = strSQL & "distv_eti=" & F_GravaDec(distv_eti) & ","
    strSQL = strSQL & "largura_eti=" & F_GravaDec(largura_eti) & ","
    strSQL = strSQL & "mesq_eti=" & F_GravaDec(mesq_eti) & ","
    strSQL = strSQL & "msup_eti=" & F_GravaDec(msup_eti) & ","
    strSQL = strSQL & "qtdcol_eti=" & F_GravaDec(qtdcol_eti) & ","
    strSQL = strSQL & "qtdlinha_eti=" & F_GravaDec(qtdlinha_eti) & ", "
    strSQL = strSQL & "fonte_eti='" & fonte & "',"
    strSQL = strSQL & "tambar_eti=" & F_GravaDec(tambar) & ", "
    strSQL = strSQL & "altbar_eti=" & F_GravaDec(altbar) & ", "
    strSQL = strSQL & "tamfonte_eti=" & F_GravaDec(tamfonte) & " "
    strSQL = strSQL & " where cod_eti='" & cod_eti & "' "
    Conn.Execute strSQL
    If Banco.RecordsAffected > 0 Then
    F_AlteraEti = True
    End If
    Exit Function
    
trata_Erro:
    
End Function

