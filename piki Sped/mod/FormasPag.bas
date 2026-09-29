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
End Type

Public Type tipopag
    codigo As String
    descricao As String


End Type

Function F_BuscaFormasPag(Optional cod As String = "", Optional expr As String = "") As FormasPag.FormasPag
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    If cod = "" Then
        frmEdiFormasPag.tag = "x"
        frmEdiFormasPag.txtCod.Text = cod
        frmEdiFormasPag.txtExpr.Text = expr
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
                    
                    Unload frmEdiFormasPag
                    
                End If
                
            End With
        End If
    Else
        strSQL = "select tf.*, tt.cod_tip,tt.nome_tip from tab_formaspag as tf LEFT join tab_tipopag as tt "
        strSQL = strSQL & " on tt.cod_tip=tf.cod_tip where 1> 0 "
        strSQL = strSQL & " and cod_for='" & cod & "' "
    
        abreObjetoAdo rPesq, strSQL
        
        With rPesq
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
                    
                                        
                End If
                .Close
        End With
    
    
    
    
    End If
    Exit Function
trata_Erro:
    MostraErro
End Function

Function F_SelectFormasPag(sExpr As String, Optional avista_For As String = "", Optional sCodFor As String = "") As FormasPag.FormasPag
    Dim sCriterio As String
    Dim sAv As String
    
    frmEdiFormasPag.txtExpr.Text = sExpr
    frmEdiFormasPag.txtCod.Text = sCodFor
    frmEdiFormasPag.tag = "x"
    frmEdiFormasPag.Show 1
    
    If frmEdiFormasPag.tag <> "" Then
       With frmEdiFormasPag.adc
            F_SelectFormasPag.cod_for = !cod_for
            F_SelectFormasPag.nome_For = !nome_For
            F_SelectFormasPag.desconto_for = !desconto_for
            F_SelectFormasPag.tipo_cai = !tipo_cai
            F_SelectFormasPag.parcelas_for = !parcelas_for
            F_SelectFormasPag.intervalo_for = !intervalo_for
            F_SelectFormasPag.nome_tip = !nome_tip
            F_SelectFormasPag.cod_tip = !cod_tip
            F_SelectFormasPag.avista_For = !avista_For
        End With
    End If
    
    Unload frmEdiFormasPag
    Exit Function
    If avista_For <> "" Then
        sAv = " and avista_for='" & avista_For & "' "
    End If
    frmSelect.sSelect = "select cod_for as cod,nome_for as nome,desconto_for as desconto,tipo_cai as caixa," _
    & "parcelas_for as parcelas,intervalo_for as intervalo,nome_tip as pagamento,tf.cod_tip as cod ,avista_for as avista " & _
    " from tab_formaspag  as tf,  tab_tipopag as tp " & _
    " where tp.cod_tip=tf.cod_tip " & sAv
    
    
    sCriterio = " and (nome_for like '" & sExpr & "%'"
    If F_SoNumero(sExpr) <> "" Then
        sCriterio = sCriterio & " or cod_for ='" & sExpr & "'"
    End If
    sCriterio = sCriterio & ") "
    frmSelect.cCriterio.Add sCriterio
    frmSelect.sOrd = " order by cod_for "
    frmSelect.txtbxExpr.Text = expr
    frmSelect.msflxgrd.ColAlignment(1) = 0
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0 'codfor
        
        F_SelectFormasPag.cod_for = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 1 'nomefor
        
        F_SelectFormasPag.nome_For = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 2 'descontofor
        
        F_SelectFormasPag.desconto_for = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 3 'tipocai
        
        F_SelectFormasPag.tipo_cai = frmSelect.msflxgrd
        
        frmSelect.msflxgrd.col = 4 'parcelas for
        F_SelectFormasPag.parcelas_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 5 'intervalor
        F_SelectFormasPag.intervalo_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 6
        F_SelectFormasPag.nome_tip = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 7
        F_SelectFormasPag.cod_tip = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 8
        F_SelectFormasPag.avista_For = frmSelect.msflxgrd
        
        
    End If
    Unload frmSelect
    

End Function

Sub P_PreencheOpcoesPag(cmbx As Object, Optional tipo_pag As String = "P")
    P_Preenche cmbx, "tab_formaspag", "concat(cod_for,'-',nome_for)", "tipo_pag='" & tipo_pag & "' "

End Sub
Sub P_PreencheTipoCai(cmbx As Object)
    P_Preenche cmbx, "tab_formaspag", "distinct tipo_cai", "", "tipo_cai"

End Sub


Function GravaTipoPag(codigo As String, descricao As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "insert into tab_tipopag(cod_tip,nome_tip)values("
    strSQL = strSQL & "'" & codigo & "',"
    strSQL = strSQL & "'" & descricao & "')"
    conn.Execute strSQL
    GravaTipoPag = True
    Exit Function
trata_Erro:
    P_MostraErro
    GravaTipoPag = False
End Function

Function AlteraTipoPag(codigo As String, descricao As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "update tab_tipopag set  "
    strSQL = strSQL & "nome_tip='" & descricao & "' "
    
    strSQL = strSQL & " where cod_tip='" & codigo & "' "
    
    conn.Execute strSQL
    AlteraTipoPag = True
    Exit Function
trata_Erro:
    
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
    conn.Execute strSQL
    
    GravaFormasPag = True
    Exit Function
trata_Erro:
    GravaFormasPag = False
    
    
    
End Function
