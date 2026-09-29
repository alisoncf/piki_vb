Attribute VB_Name = "FormasPag"
Public Type FormasPag
    cod_for As String
    nome_For As String
    desconto_for As Currency
    tipo_cai As String
    avista_For As String
    parcelas_for As String
    intervalo_for As String
    cod_tip As String
    nome_tip As String
    gerencial_for As String
End Type
Public Type tipopag
    codigo As String
    descricao As String


End Type

Function GravaFormasPag(cod_for As String, nome_For As String, desconto_for As String, _
    tipo_cai As String, avista_For As String, parcelas_for As String, intervalo_for As String, _
    cod_tip As String, gerencial_for As String, tPag As String) As Boolean
    
    On Error GoTo trata_Erro
    If cod_for = "" Then
        cod_for = MaxKeySQL(TAB_FORMASPAG, "cod_for", "", "0#")
    End If
    strSQL = "replace into " & TAB_FORMASPAG & " (cod_For,nome_for,desconto_for,tipo_Cai,avista_for,"
    strSQL = strSQL & " parcelas_for,intervalo_for,cod_tip,gerencial_for,tpag)values("
    strSQL = strSQL & "'" & cod_for & "',"
    strSQL = strSQL & "'" & nome_For & "',"
    strSQL = strSQL & F_GravaDec(desconto_for) & ","
    strSQL = strSQL & "'" & tipo_cai & "',"
    strSQL = strSQL & "'" & avista_For & "',"
    strSQL = strSQL & "" & parcelas_for & ","
    strSQL = strSQL & "" & intervalo_for & ","
    strSQL = strSQL & "'" & cod_tip & "',"
    strSQL = strSQL & "'" & gerencial_for & "',"
    strSQL = strSQL & "'" & tPag & "')"
    Conn.Execute strSQL
    
    GravaFormasPag = True
    Exit Function
trata_Erro:
    GravaFormasPag = False
    
    
    
End Function

Function BuscaTipoPag(desc As String) As tipopag
    frmEdiTipoPag.tag = "x"
    frmEdiTipoPag.txtBusca.Text = desc
    frmEdiTipoPag.Show 1
    
    If frmEdiTipoPag.tag <> "" Then
        BuscaTipoPag.descricao = frmEdiTipoPag.ado!nome_tip
        BuscaTipoPag.codigo = frmEdiTipoPag.ado!cod_tip
    End If
    Unload frmEdiEspCai
End Function

Function F_BuscaFormasPag(Optional cod As String = "", Optional expr As String = "", Optional gerencial As String = "N", Optional soAvista As Boolean = False) As FormasPag.FormasPag
    frmEdiFormasPag.tag = "x"
    On Error GoTo trata_Erro
    
    
    
    
    frmEdiFormasPag.txtCod.Text = cod
    frmEdiFormasPag.txtExpr.Text = expr
    frmEdiFormasPag.soAvista = soAvista
    frmEdiFormasPag.Show 1
    
    
If frmEdiFormasPag.tag <> "" Then
        With frmEdiFormasPag.adc
            If .RecordCount > 0 Then
                
                F_BuscaFormasPag.avista_For = !avista_For
                F_BuscaFormasPag.cod_for = !cod_for
                F_BuscaFormasPag.cod_tip = !cod_tip
                F_BuscaFormasPag.desconto_for = !desconto_for
                F_BuscaFormasPag.intervalo_for = !intervalo_for
                F_BuscaFormasPag.nome_For = !nome_For
                F_BuscaFormasPag.nome_tip = !nome_tip
                F_BuscaFormasPag.parcelas_for = !parcelas_for
                F_BuscaFormasPag.tipo_cai = !tipo_cai
                F_BuscaFormasPag.gerencial_for = !gerencial_for
                Unload frmEdiFormasPag
                
            End If
            
        End With
    End If
    Exit Function
trata_Erro:
    P_MostraErro
End Function

Function F_SelectFormasPag(sExpr As String, Optional avista_For As String = "", Optional sCodFor As String = "") As FormasPag.FormasPag
    
    
    Dim sCriterio As String
    Dim sAv As String
    
    F_SelectFormasPag = F_BuscaFormasPag(sCodFor, sExpr)
    Exit Function
    
    
    
    

End Function

Sub P_PreencheOpcoesPag(cmbx As Object, Optional tipo_pag As String = "P")
    P_Preenche cmbx, "tab_formaspag", "concat(cod_for,'-',nome_for)", "tipo_pag='" & tipo_pag & "' "

End Sub
Sub P_PreencheTipoCai(cmbx As Object)
    P_Preenche cmbx, "tab_formaspag", "distinct tipo_cai", "", "tipo_cai"

End Sub

Function AlteraFormasPag(cod_for As String, nome_For As String, desconto_for As String, tipo_cai As String, _
    avista_For As String, parcelas_for As String, intervalo_for As String, cod_tip As String, _
    gerencial_for As String) As Boolean

    Dim strSQL As String
    Dim sCodFor As String
   
    
    On Error GoTo trata_Erro
    
    strSQL = "update tab_formaspag set "
    strSQL = strSQL & "nome_for='" & nome_For & "',"
    strSQL = strSQL & "desconto_for=" & F_GravaDec(desconto_for) & ","
    strSQL = strSQL & "tipo_cai='" & tipo_cai & "',"
    strSQL = strSQL & "avista_for='" & avista_For & "',"
    strSQL = strSQL & "parcelas_for=" & parcelas_for & ","
    strSQL = strSQL & "intervalo_for=" & intervalo_for & ","
    strSQL = strSQL & "cod_tip='" & cod_tip & "',"
    strSQL = strSQL & "gerencial_for='" & gerencial_for & "' "
    strSQL = strSQL & " where cod_for='" & cod_for & "'"

    Conn.Execute strSQL
    AlteraFormasPag = True
    Exit Function
trata_Erro:
    

End Function
Function DesabilitaForma(cod_for As String, opcao As Boolean) As Boolean

    Dim strSQL As String
    Dim op As String
    If opcao = False Then
        op = "N"
    Else
    
        op = "S"
    End If
    
    strSQL = "update tab_formaspag set "
    strSQL = strSQL & " visivel_for='" & op & "' "
    strSQL = strSQL & " where cod_for='" & cod_for & "'"

    Conn.Execute strSQL
    DesabilitaForma = True
    Exit Function
trata_Erro:
    

End Function


Function P_GravaFormasPag(nome_For As String, desconto_for As String, tipo_cai As String, _
    avista_For As String, parcelas_for As String, intervalo_for As String, cod_tip As String, _
    gerencial_for As String) As Boolean

    Dim strSQL As String
    Dim sCodFor As String
    sCodFor = MaxKeySQL("tab_formaspag", "cod_for", "", "0#")
    
    On Error GoTo trata_Erro
    
    strSQL = "insert into tab_formaspag(cod_for,nome_for,desconto_for,tipo_cai,avista_for,"
    strSQL = strSQL & "parcelas_for,intervalo_for,cod_tip,gerencial_for)values("
    strSQL = strSQL & "'" & sCodFor & "',"
    strSQL = strSQL & "'" & nome_For & "',"
    strSQL = strSQL & F_GravaDec(desconto_for) & ","
    strSQL = strSQL & "'" & tipo_cai & "',"
    strSQL = strSQL & "'" & avista_For & "',"
    strSQL = strSQL & parcelas_for & ","
    strSQL = strSQL & intervalo_for & ","
    strSQL = strSQL & "'" & cod_tip & "',"
    strSQL = strSQL & "'" & gerencial_for & "')"
    Conn.Execute strSQL
    P_GravaFormasPag = True
    Exit Function
trata_Erro:
    

End Function

