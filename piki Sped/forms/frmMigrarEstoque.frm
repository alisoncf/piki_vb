VERSION 5.00
Begin VB.Form frmMigrarEstoque 
   Caption         =   "Form2"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   LinkTopic       =   "Form2"
   ScaleHeight     =   7620
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVIncula 
      Caption         =   "vinculacao"
      Height          =   1215
      Left            =   10320
      TabIndex        =   18
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdAjustaITd 
      Caption         =   "Ajusta ITD"
      Height          =   735
      Left            =   2640
      TabIndex        =   17
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdDev 
      Caption         =   "Devoluções"
      Height          =   855
      Left            =   6360
      TabIndex        =   16
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdChaves 
      Caption         =   "Chaves"
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdCompras 
      Caption         =   "Compras"
      Height          =   1095
      Left            =   4200
      TabIndex        =   14
      Top             =   4440
      Width           =   3735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Financeiro"
      Height          =   1095
      Left            =   5640
      TabIndex        =   13
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Fornecedores"
      Height          =   1095
      Left            =   2280
      TabIndex        =   12
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Todos"
      Height          =   495
      Left            =   9360
      TabIndex        =   11
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Compras"
      Height          =   1335
      Left            =   7920
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Clientes"
      Height          =   1095
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdUsu 
      Caption         =   "Usuários"
      Height          =   855
      Left            =   4560
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Atualizar itvendas"
      Height          =   855
      Left            =   3240
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAtlzProd 
      Caption         =   "Atualizar produtos"
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Caixa"
      Height          =   1095
      Left            =   7920
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Docs"
      Height          =   1095
      Left            =   3120
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Vendas"
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Similares"
      Height          =   1335
      Left            =   5640
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subdivisões (Categorias e outros)"
      Height          =   1335
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdProd 
      Caption         =   "Produtos"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmMigrarEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bancoantigo As New ADODB.Connection
Dim data_padrao As String

Private Sub cmdAjustaITd_Click()
        Dim strSQL As String
    Dim cod_cid As String
    Dim b As database
    Dim rPesq As Recordset
    Set b = OpenDatabase("estoque_old")
    strSQL = "select * from tab_itdoc where reman_itd=0 order by seq_doc desc "
    Set rPesq = b.OpenRecordset(strSQL)
    
    'conn.Execute "delete from fornecedores"
    
    Dim i As Integer
    With rPesq
        Do While Not .EOF
            
            strSQL = "update itdoc set reman_itd=0 where seq_itd='" & !seq_itd & "' "
            conn.Execute strSQL
            cmdAjustaITd.caption = !seq_itd
            'cmdAjustaITd.Refresh
            .MoveNext
            
            'i = i + 1
        Loop
        
        .Close
    End With
    
End Sub

Private Sub cmdAtlzProd_Click()
    AlterarEstoque
End Sub

Private Sub cmdChaves_Click()
    Dim strSQL As String
    strSQL = "update tab_codigos set value_cod='" & MaxKeySQL(TAB_DOC, "seq_doc", "", "000000#") & "' "
    strSQL = strSQL & " where field_cod='seq_doc'"
'    conn.Execute strSQL
    
    strSQL = "update tab_codigos set value_cod='" & MaxKeySQL(TAB_FINANCEIRO, "seq_fin", "", "000000#") & "' "
    strSQL = strSQL & " where field_cod='seq_fin'"
 '   conn.Execute strSQL
    
    strSQL = "update tab_codigos set value_cod='" & MaxKeySQL(TAB_VENDAS, "num_ven", "", "000000#") & "' "
    strSQL = strSQL & " where field_cod='num_ven'"
  '  conn.Execute strSQL
    
    
    
    
    
    
    
End Sub

Private Sub cmdCompras_Click()
    Dim strSQL As String
    
    Dim b As database
    Dim rPesq As Recordset
    Dim sSeq As String
    Dim i As Integer
    Dim rItens As Recordset
    Set b = OpenDatabase("estoque_old")
    
    strSQL = "delete from sped_nf "
    'conn.Execute strSQL
    
    strSQL = "delete from sped_nf_itens "
    'conn.Execute strSQL
    
    
    
        
    strSQL = "select * from tab_cmp order by emissao_not"
    
    Set rPesq = b.OpenRecordset(strSQL)
    
    With rPesq
        Do While Not .EOF
            sSeq = SpedNF.GravaSpedNf("0", "1", F_BuscaUmCampo("codigo", "sped_participante", "cod_for='" & !cod_For & "'"), _
                 "55", "00", "1", !Num_not, "", Format(!emissao_not), Format(!entrada_not), CStr(!valor_not), "1", "0", _
                    "0", "0", "0", "0", "0", "0", CStr(!bc_not), CStr(!icms_not), "0", "0", CStr(!ipi_not), "0", "0", "0", "0", _
                    "4", Format(!entrada_not), "01")
                
                i = 0
            If sSeq <> "" Then
                strSQL = "select tp.*,ti.* from tab_itcmp as ti, tab_produtos as tp where tp.cod_pro=ti.cod_pro "
                strSQL = strSQL & " and ti.cod_for='" & !cod_For & "' "
                strSQL = strSQL & " and ti.num_not='" & !Num_not & "' "
                strSQL = strSQL & " and ti.serie_not='" & !serie_not & "' "
                 
                Set rItens = b.OpenRecordset(strSQL)
                With rItens
                    Do While Not .EOF
                        i = i + 1
                        If SpedNF.InsereItemSpedNF(sSeq, CStr(i), !cod_pro, !cod_pro, !nome_pro, CStr(!qtd_ite), _
                        !unidade_pro, "0", "", !unit_ite, "0", "0", "40", "5949", "", "0", "0", "0", "0", "0", "0", _
                        "0", "0", "40", "", "0", "0", !ipi_ite, "40", "0", "0", "0", "0", "0", "40", "0", "0", "0", "0", "0", "", _
                        !codbarra_pro, Format(!ncm_pro, "0000000#"), CStr(!qtd_ite), CStr(!custo_ite), "0", "N") = False Then
                        
                        
                            MsgBox "erro"
                        End If
                        
                        
                        
                                                
                        .MoveNext
                    Loop
                End With
                
            End If
            Me.caption = Format(!emissao_not)
            .MoveNext
        Loop
    End With
    MsgBox "Migrado"
End Sub

Private Sub cmdDev_Click()
    Dim strSQL As String
    Dim cod_cid As String
    Dim b As database
    Dim rPesq As Recordset
    Dim sSeq As String
    Dim rItens As Recordset
    Dim i As Integer
    Set b = OpenDatabase("estoque_old")
    strSQL = "select * from tab_devolucoes"
    Set rPesq = b.OpenRecordset(strSQL)
    
    'conn.Execute "delete from " & TAB_DEVOLUCOES
    'conn.Execute "delete from " & TAB_ITDEV
    
    
    With rPesq
        Do While Not .EOF
                    
            sSeq = GravaDevolucao(!data_dev, Right(!data_dev, 8), !cod_usu, !gerado_Dev, _
                !num_ven, "")
            
            If GravaItDev(sSeq, !cod_pro, !qtd_dev, !gerado_Dev, "01") Then
            End If
            
                    
            i = i + 1
            Me.caption = i
            Me.Refresh
            .MoveNext
        Loop
    End With
    
End Sub

Private Sub cmdProd_Click()
    Command1_Click
End Sub

Private Sub cmdUsu_Click()
    Dim strSQL As String
    
    Dim b As database
    Dim rPesq As Recordset
    Set b = OpenDatabase("estoque_old")
    
    
    

    strSQL = "delete from tab_funcionarios"
'    conn.Execute strSQL

    strSQL = "delete from tab_usuarios"
 '   conn.Execute strSQL
    
    strSQL = "select * from tab_funcionarios"
    Set rPesq = b.OpenRecordset(strSQL)
    
    
    With rPesq
        Do While Not .EOF
            strSQL = funcionario.F_GravaFun(!nome_fun, !guerra_fun, Left(!endereco_fun, 40), "", "ANAPOLIS", "GO", _
                "", "", "", !rg_fun, !cpf_fun, "", "", "", _
                "", "", "", !cod_Car, !telefone_fun)
                
            
                
            .MoveNext
'            Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
        Loop
        
        
        
        MsgBox "mIGRADO"
        .Close
    End With
    strSQL = "select * from tab_usuarios"
    Set rPesq = b.OpenRecordset(strSQL)
    
    
    With rPesq
        Do While Not .EOF
            strSQL = "insert into tab_usuarios(cod_fun,nivel_usu,status_usu,senha_usu)values("
            strSQL = strSQL & "'" & !cod_fun & "',"
            strSQL = strSQL & "'" & !nivel_usu & "',"
            strSQL = strSQL & "'" & !status_usu & "',"
            strSQL = strSQL & "'" & !senha_usu & "') "
            conn.Execute strSQL
            
                
            .MoveNext
'            Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
        Loop
        
        
        
        
        .Close
    End With
    
End Sub

Private Sub Command1_Click()
        Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim c As New ADODB.Connection
    Dim b As database
    Dim r As Recordset
    Dim sCodImp As String
    Dim i As Integer
    Dim cst As String
    Set b = OpenDatabase("estoque_old")
    
    strSQL = "delete from " & TAB_PRODUTOS
    'conn.Execute strSQL
    
    strSQL = "delete from estoque"
    'conn.Execute strSQL
    
    
    strSQL = "select * from tab_produtos order by cod_pro desc "
    Set r = b.OpenRecordset(strSQL)
    
    With r
        Do While Not .EOF
            
            If !codfab_pro = "334846" Then
                MsgBox "aqui"
            End If
            
            If !cod_imp = "I02" Then
                sCodImp = "12"
                cst = ""
            ElseIf !cod_imp = "I05" Then
                sCodImp = "FF"
                cst = "500"
            Else
                MsgBox "imposto errado"
            End If
        
        
            strSQL = Produtos.GravaProduto(!codbarra_pro, !nome_pro, !unidade_pro, !cod_cat, !cod_mar, !cod_sub, "0", !codfab_pro, _
                !lucro_pro, !endereco_pro, Format(!cod_For), !estmin_pro, Format(!detalhe_pro), !local_Pro, !status_pro, _
                Format(!ncm_pro, "0000000#"), !ef_pro, !er_pro, !cr_pro, !cr_pro, !pv_pro, !pv_pro, "N", sCodImp, IIf(!status_pro = "D", "N", "S"), "", "", !codbarra_pro, _
                 sCodImp, cst, cst, "", "", "99", "99", "99", "", !cod_pro)
                'Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
            
            
            
            
            If strSQL = "" Then
                MsgBox "erro"
            Else
                'If Not AdicionaEstoque("01", !cod_pro, !er_pro, !ef_pro) Then
'                    MsgBox "Erro"
                'End If
            End If
            
            .MoveNext
            i = i + 1
            Me.caption = i
        Loop
        .Close
    End With
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
End Sub
Sub migraCat()
    Dim strSQL As String
    Dim b As database
    Dim rPesq As Recordset
    Set b = OpenDatabase("estoque_old")
    
    
    

    strSQL = "delete from " & TAB_CATEGORIAS
    'conn.Execute strSQL


    
    strSQL = "select * from tab_categorias"
    Set rPesq = b.OpenRecordset(strSQL)
    
    
    With rPesq
        Do While Not .EOF
            strSQL = "insert into " & TAB_CATEGORIAS & "(cod_cat,nome_cat)values("
            strSQL = strSQL & "'" & !cod_cat & "',"
            strSQL = strSQL & "'" & !nome_cat & "')"
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub
Sub migraSubcat()
    Dim strSQL As String
    Dim b As database
    Dim rPesq As Recordset
    Set b = OpenDatabase("estoque_old")
    
    
    

    strSQL = "delete from " & TAB_SUBCAT
    'conn.Execute strSQL


    
    strSQL = "select * from tab_subcat"
    Set rPesq = b.OpenRecordset(strSQL)
    
    
    With rPesq
        Do While Not .EOF
            strSQL = "insert into " & TAB_SUBCAT & "(cod_sub,nome_sub)values("
            strSQL = strSQL & "'" & !cod_sub & "',"
            strSQL = strSQL & "'" & !nome_sub & "')"
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub
Sub migramarcas()
    Dim strSQL As String
    Dim b As database
    Dim rPesq As Recordset
    Set b = OpenDatabase("estoque_old")
    
    
    

    strSQL = "delete from " & TAB_MARCAS
    'conn.Execute strSQL


    
    strSQL = "select * from tab_marcas"
    Set rPesq = b.OpenRecordset(strSQL)
    
    
    With rPesq
        Do While Not .EOF
            strSQL = "insert into " & TAB_MARCAS & "(cod_mar,nome_mar)values("
            strSQL = strSQL & "'" & !cod_mar & "',"
            strSQL = strSQL & "'" & !nome_mar & "')"
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub cmdVIncula_Click()
        Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim c As New ADODB.Connection
    Dim b As database
    Dim r As Recordset
    Dim sCodImp As String
    Dim i As Integer
    Dim cst As String
    Dim l As Long
    Set b = OpenDatabase("estoque_old")
    
    strSQL = "delete from tab_vinc"
   ' conn.Execute strSQL
    
    
    
    strSQL = "select * from tab_vinc order by seq_vin desc"
    Set r = b.OpenRecordset(strSQL)
    With r
        Do While Not .EOF
            
            If F_GravaVinc(!seq_itd, !data_vin, !valor_vin, !acres_vin, !decres_vin, !venc_vin, !parc_vin, _
                !seq_doc, !seq_vin) <> "" Then
            
                
            
            End If
'            l = l + 1
            cmdVIncula.caption = !data_vin
            cmdVIncula.Refresh
            
            .MoveNext
        Loop
    End With

End Sub

Private Sub Command10_Click()
    Dim strSQL As String
    
    Dim b As database
    Dim rPesq As Recordset
    Dim rItens As Recordset
    Set b = OpenDatabase("estoque_old")
    
    
    
    strSQL = "select * from tab_cmp order by seq_not "
    Set rPesq = b.OpenRecordset(strSQL)
    
    
    With rPesq
    End With
    
    
    
    
    
    
    
    
    
    
End Sub

Private Sub Command11_Click()
    Command4_Click
    Command5_Click
    Command6_Click
    cmdProd_Click
    Command8_Click
    Command9_Click
  '  cmdUsu_Click
End Sub

Private Sub Command12_Click()
    Dim strSQL As String
    Dim cod_cid As String
    Dim b As database
    Dim rPesq As Recordset
    Set b = OpenDatabase("estoque_old")
    strSQL = "select * from tab_fornecedores"
    Set rPesq = b.OpenRecordset(strSQL)
    
    'conn.Execute "delete from fornecedores"
    
    
    With rPesq
        Do While Not .EOF
            
          '  If !cod_Cid = "" Then
          '      cod_Cid = "5201108"
          '  ElseIf IsNull(!cod_Cid) Then
          '      cod_Cid = "5201108"
          '  Else
            
          '      cod_Cid = !cod_Cid
          '  End If
           strSQL = fornecedor.F_GravaFor(!nome_For, !fantasia_for, !cnpj_for, !ie_for, !endereco_for, !bairro_for, !cidade_For, !uf_for, _
           !cep_for, !contato_for, !email_for, !telefone_for, !fax_for, Format(!obs_for), !cod_For)
            
            If F_BuscaUmCampo("cod_for", "sped_participante", "cnpj='" & F_SoNumero(!cnpj_for) & "'") = "" Then
        
                strSQL = GravaParticipante(!nome_For, "1058", F_SoNumero(!cnpj_for), "", F_SoNumero(!ie_for), _
                "", "", !endereco_for, "", "", !bairro_for, "", !cod_For, !cep_for, !uf_for)
            End If
        
        Me.caption = .RecordCount
        .MoveNext
        
    Loop
    MsgBox "Migrado"
End With
End Sub

Private Sub Command13_Click()
      Dim strSQL As String
    
    Dim b As database
    Dim rPesq As Recordset
    Dim rItens As Recordset
    Dim iItem As Integer
    Dim i As Integer
    Set b = OpenDatabase("estoque_old")
    
    
    strSQL = "delete from financeiro where data_fin>=" & F_GravaData(data_padrao) & " "
    'conn.Execute strSQL
    

    
    strSQL = "select * from tab_financeiro order by seq_fin desc "
    Set rPesq = b.OpenRecordset(strSQL)
    
    With rPesq
        Do While Not .EOF
            Me.caption = !data_fin
                strSQL = F_GravaFinanceiro(!tipo_fin, !data_fin, CStr(!valor_fin), !historico_fin, !cod_con, "", "000", !seq_fin)
                If strSQL = "" Then
                    MsgBox "Erro"
                    Exit Sub
                End If
                .MoveNext
            
            
                        
        Loop
                
        .Close
    End With
End Sub

Private Sub Command6_Click()
'

  Dim strSQL As String
    
    Dim b As database
    Dim rPesq As Recordset
    Dim rItens As Recordset
    Dim iItem As Integer
    Dim i As Integer
    Set b = OpenDatabase("estoque_old")
    
    
    strSQL = "delete from caixa where data_cai>=" & F_GravaData(data_padrao) & " "
    'conn.Execute strSQL
    

    
    strSQL = "select * from tab_caixa order by seq_cai desc "
    Set rPesq = b.OpenRecordset(strSQL)
    
    With rPesq
        Do While Not .EOF
            Me.caption = !data_cai
                strSQL = F_GravaCai(!tipo_cai, !data_cai, CStr(!valor_Cai), !historico_cai, !num_ven, !esp_cai, !cod_con, "0", "0", !cod_cai, "1", "000", "S", "0")
                .MoveNext
            
            
                        
        Loop
                
        .Close
    End With
End Sub

Private Sub AlterarEstoque()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim c As New ADODB.Connection
    Dim b As database
    Dim r As Recordset
    Dim sCodImp As String
    Dim i As Integer
    Dim cst As String
    Set b = OpenDatabase("estoque_old")
    
        
    
    strSQL = "select * from tab_composicao order by cod_pro"
    Set r = b.OpenRecordset(strSQL)
    
    With r
        
        Do While Not .EOF
        
            If Not Produtos.AlteraEstoqueProduto(!cod_pro, !e2_pro, !e1_pro) Then
                MsgBox "erro"
            End If
            If Not Produtos.AlteraEstoque("01", !cod_pro, !e2_pro, !e1_pro) Then
                MsgBox "erro"
            End If
            
            strSQL = "Update " & TAB_PRODUTOS & " set er_pro=" & F_GravaDec(!e2_pro) & ","
            strSQL = strSQL & " ef_pro=" & F_GravaDec(!e1_pro) & ","
            strSQL = strSQL & " cr_pro=" & F_GravaDec(!c2_pro) & ","
            strSQL = strSQL & " pv_pro=" & F_GravaDec(!p2_pro)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                'Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
            conn.Execute strSQL
            
            
            
            
            .MoveNext
            i = i + 1
            Me.caption = i
        Loop
        .Close
    End With
    
    
    
    
    
End Sub

Private Sub Command2_Click()

migramarcas
migraSubcat
migraCat
    
    
    


End Sub

Private Sub Command3_Click()
    Dim strSQL As String
    
    Dim b As database
    Dim rPesq As Recordset
    Set b = OpenDatabase("estoque_old")
    
    
    

    strSQL = "delete from tab_similares"
    'conn.Execute strSQL


    
    strSQL = "select * from tab_similares"
    Set rPesq = b.OpenRecordset(strSQL)
    
    
    With rPesq
        Do While Not .EOF
            strSQL = "insert into tab_similares(cod_pro,cod_sim,tipo_sim)values("
            strSQL = strSQL & "'" & !cod_pro & "',"
            strSQL = strSQL & "'" & !cod_sim & "',"
            strSQL = strSQL & "'" & !tipo_sim & "')"
            conn.Execute strSQL
            .MoveNext
'            Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
        Loop
        
        .Close
    End With
End Sub

Private Sub Command4_Click()
    Dim strSQL As String
    
    Dim b As database
    Dim rPesq As Recordset
    Dim rItens As Recordset
    Dim iItem As Integer
    Dim i As Integer
    Dim maxVen As String
    Dim cod_For As String
    Set b = OpenDatabase("estoque_old")
    
    
    
    



    
    
    'strSQL = "delete from " & TAB_VENFEC & " where num_ven in (select num_ven from " & TAB_VENDAS & " where dataefet_ven>=" & F_GravaData(data_padrao) & ")"
    'conn.Execute strSQL
    
    strSQL = "select max(num_ven) as maximo from tab_vendas where situacao_ven=4 and dataefet_ven<" & F_GravaData(data_padrao) & ""
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    
    
    maxVen = rPesq!maximo
    
    
    strSQL = "delete from itvendas where num_ven>" & maxVen
    'conn.Execute strSQL
    strSQL = "delete from vendas where num_ven>" & maxVen
    'conn.Execute strSQL
    strSQL = "delete from venfin where num_ven>" & maxVen
    'conn.Execute strSQL
    strSQL = "delete from venfec where num_ven>" & maxVen
    'conn.Execute strSQL
    'strSQL = "delete from " & TAB_VENFIN & " where num_ven in (select num_ven from " & TAB_VENDAS & " where dataefet_ven>=" & F_GravaData(data_padrao) & ")"
    'conn.Execute strSQL
    
  '  strSQL = "delete from " & TAB_ITVENDAS & " where num_ven in (select num_ven from " & TAB_VENDAS & " where dataefet_ven>=" & F_GravaData(data_padrao) & ")"
  '  conn.Execute strSQL
    
    
    
    strSQL = "delete from " & TAB_VENFEC & " where num_ven in (select num_ven from " & TAB_VENDAS & " where situacao_Ven=1) "
    'conn.Execute strSQL
    
    strSQL = "delete from " & TAB_VENFIN & " where num_ven in (select num_ven from " & TAB_VENDAS & " where situacao_Ven=1) "
    'conn.Execute strSQL
    
    strSQL = "delete from " & TAB_ITVENDAS & " where num_ven in (select num_ven from " & TAB_VENDAS & " where situacao_Ven=1) "
    'conn.Execute strSQL
    
    'strSQL = "delete from tab_vendas where dataefet_ven >=" & F_GravaData(data_padrao)
    'conn.Execute strSQL
    
    strSQL = "delete from tab_vendas where situacao_ven=1"
    'conn.Execute strSQL
    
    
    
    
    strSQL = "select * from tab_vendas where num_ven=614400 order by num_ven desc "
    Set rPesq = b.OpenRecordset(strSQL)
    
    
    With rPesq
        Do While Not .EOF
            On Error Resume Next
            cod_For = !cod_tip
            If cod_For = "02" Then
                cod_For = "06"
            ElseIf cod_For = "01" Then
                cod_For = "01"
            ElseIf cod_For = "03" Then
                cod_For = "07"
            ElseIf cod_For = "05" Then
                cod_For = "11"
            ElseIf cod_For = "04" Then
                cod_For = "08"
            Else
                cod_For = "01"
            'MsgBox ""
            End If
            
            If GravaVendaCompleta(!num_ven, !cod_cli, !data_ven, Format(!dataefet_ven), "127.0.0.1", !valor_ven, _
                !cod_fun, !cod_usu, !cod_cai, !situacao_ven, !desconto_ven, !cod_tip, !nomecli_ven, "", "0", "", "", "", "N", _
                !status_ven, "", "C", "", "", "", !valor_ven, !seqefet_ven) Then
                
                strSQL = "select * from tab_itvendas where num_ven='" & !num_ven & "' "
                Set rItens = b.OpenRecordset(strSQL)
                
                With rItens
                    iItem = 0
                    Do While Not .EOF
                        iItem = iItem + 1
                        strSQL = "insert into " & TAB_ITVENDAS & "(num_ven,seq_ite,cod_pro,qtd_ite,unit_ite,desc_ite,"
                        strSQL = strSQL & " desconto_ite,final_ite,pq_itd,cod_emp)values("
                        strSQL = strSQL & "'" & !num_ven & "',"
                        strSQL = strSQL & iItem & ","
                        strSQL = strSQL & "'" & !cod_pro & "',"
            
                        strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
                        strSQL = strSQL & F_GravaDec(!unit_ite) & ","
                        strSQL = strSQL & "'" & !descricao_ite & "',"
            
                        strSQL = strSQL & "0,"
                        strSQL = strSQL & F_GravaDec(!unit_ite * !qtd_ite) & ","
                        strSQL = strSQL & F_GravaDec(CStr(!qtd_ite * !unit_ite)) & ","
                        strSQL = strSQL & "'01') "
                        conn.Execute strSQL
                    .MoveNext
                Loop
                .Close
                End With
                
                
                strSQL = "select * from tab_venfec where num_ven='" & !num_ven & "' "
                Set rItens = b.OpenRecordset(strSQL)
                With rItens
                    Do While Not .EOF
                        strSQL = F_AdicionaVenFec(!num_ven, !avista_ven + !aprazo_ven, "01", "01", "1", "DUP")
                        .MoveNext
                    Loop
                    .Close
                End With
                
            
            
            End If
            .MoveNext
            
            Me.caption = Format(!dataefet_ven)
        Loop
        .Close
    End With

End Sub

Private Sub Command5_Click()
    Dim strSQL As String
    
    Dim b As database
    Dim rPesq As Recordset
    Dim rItens As Recordset
    Dim iItem As Integer
    Dim i As Integer
    Set b = OpenDatabase("estoque_old")
    
    
    strSQL = "delete from itdoc where seq_doc in (select seq_doc from doc where data_doc>=" & F_GravaData(data_padrao) & ") "
  '  conn.Execute strSQL
    strSQL = "delete from cheques where seq_doc in (select seq_doc from doc where data_doc>=" & F_GravaData(data_padrao) & ")"
   ' conn.Execute strSQL
    strSQL = "delete from doc where data_doc>=" & F_GravaData(data_padrao)
    'conn.Execute strSQL

'    strSQL = "delete from " & TAB_DOC
'    conn.Execute strSQL
'
'    strSQL = "delete from " & TAB_ITDOC
'    conn.Execute strSQL
'
'    strSQL = "delete from " & TAB_CHEQUES
'    conn.Execute strSQL
    
'    strSQL = "delete from baixaitd"
'    conn.Execute strSQL
'    strSQL = "delete from baixafin"
'    conn.Execute strSQL
'
    
    strSQL = "select * from tab_doc where seq_doc=0304101 order by seq_doc desc "
    Set rPesq = b.OpenRecordset(strSQL)
    
    With rPesq
        Do While Not .EOF
                
                If !num_ven = "0614471" Then
                    MsgBox ""
                End If
                
            
                    
                strSQL = F_GravaDoc(!historico_doc, !num_doc, !valor_doc, !data_doc, !cod_ter, !tipo_doc, !num_ven, !Parc_Doc, "V", "C", "001", !seq_doc)
                
                               
                If strSQL <> "" Then
                
                    strSQL = "select * from tab_itdoc where seq_doc='" & !seq_doc & "' "
                
                    Set rItens = b.OpenRecordset(strSQL, dbOpenSnapshot)
                    With rItens
                        Do While Not .EOF
                            strSQL = F_GravaItDoc(!seq_doc, !valor_itd, !venc_itd, !venc_itd, !Parc_Itd, a_receber_avulso, _
                            !boleto_itd, "1", "000", !seq_itd, !reman_itd)
                            .MoveNext
                        Loop
                        .Close
                    End With
                    
                    strSQL = "select * from tab_cheques where seq_doc='" & !seq_doc & "' "
                    Set rItens = b.OpenRecordset(strSQL, dbOpenSnapshot)
                    With rItens
                        Do While Not .EOF
                            If !tipo_che = "R" Then
                                strSQL = F_GravaCh(!seq_doc, !num_che, !cod_bco, !age_che, !cco_che, !nome_che, _
                                 !Valor_che, Format(!Venc_che), !cod_ter, !tipo_che, !num_ven, "", !Parc_Che, "00", "", !seq_che, Format(!dep_che))
                            Else
                                strSQL = F_GravaChePag(!seq_doc, !num_che, "", !Valor_che, !Venc_che, !cod_ter, !Parc_Che, _
                              !seq_che, Format(!dep_che))
                            End If
                            .MoveNext
                        Loop
                        .Close
                    End With
                Else
                    'MsgBox "erro"
                    'Exit Sub
                End If
            
            .MoveNext
            
            Me.caption = !data_doc
            Me.Refresh
            
        Loop
                
        .Close
    End With
End Sub



Private Sub MigraPRodutos()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim c As New ADODB.Connection
    Dim b As database
    Dim r As Recordset
    Dim sCodImp As String
    Dim i As Integer
    Dim cst As String
    Set b = OpenDatabase("estoque_old")
    
        
    
    strSQL = "select * from tab_produtos order by cod_pro"
    Set r = b.OpenRecordset(strSQL)
    
    With r
        
        Do While Not .EOF
        
            If Not Produtos.AlteraEstoqueProduto(!cod_pro, !er_pro, !ef_pro) Then
                MsgBox "erro"
            End If
            If Not Produtos.AlteraEstoque("01", !cod_pro, !er_pro, !ef_pro) Then
                MsgBox "erro"
            End If
            
                'Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
            
            
            
            
            
            .MoveNext
            i = i + 1
            Me.caption = i
        Loop
        .Close
    End With
    
    
    
    
    
    
    
    
    
End Sub

Private Sub Command8_Click()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    
    strSQL = "Update " & TAB_ITVENDAS & " set final_ite=unit_ite*qtd_ite"
    conn.Execute strSQL
    
End Sub

Private Sub Command9_Click()
    Dim strSQL As String
    Dim cod_cid As String
    Dim b As database
    Dim rPesq As Recordset
    Set b = OpenDatabase("estoque_old")
    strSQL = "select * from tab_clientes"
    Set rPesq = b.OpenRecordset(strSQL)
    
    'conn.Execute "delete from clientes"
    
    
    With rPesq
        Do While Not .EOF
            
            If !cod_cid = "" Then
                cod_cid = "5201108"
            ElseIf IsNull(!cod_cid) Then
                cod_cid = "5201108"
            Else
            
                cod_cid = !cod_cid
            End If
           strSQL = "insert into " & TAB_CLIENTES & "(cod_cli,nome_cli,cpf_cli,cnpj_cli,rg_cli, endereco_Cli, "
           strSQL = strSQL & "bairro_cli,cidade_Cli,uf_cli,cep_cli,telefone_Cli,fax_Cli,"
           strSQL = strSQL & "tipo_cli,cel_cli,obs_Cli,limite_cli,numero_cli,complemento_cli,ie_cli,cod_cid) values("
           strSQL = strSQL & "" & !cod_cli & ","
           strSQL = strSQL & "'" & !nome_Cli & "',"
           strSQL = strSQL & "'" & !cpf_cli & "',"
           strSQL = strSQL & "'" & !cpf_cli & "',"
           strSQL = strSQL & "'" & !rg_cli & "',"
           strSQL = strSQL & "'" & !endereco_cli & "',"
           strSQL = strSQL & "'" & !bairro_cli & "',"
           strSQL = strSQL & "'" & !cidade_cli & "',"
           strSQL = strSQL & "'" & Left(!uf_cli, 2) & "',"
           strSQL = strSQL & "'" & !cep_cli & "',"
           strSQL = strSQL & "'" & !telefone_cli & "',"
           strSQL = strSQL & "'" & !fax_cli & "',"
           If Len(!cpf_cli) > 14 Then
                strSQL = strSQL & "'J',"
           Else
                strSQL = strSQL & "'F',"
           End If
           
           
           strSQL = strSQL & "'" & !cel_cli & "',"
           strSQL = strSQL & "'" & !obs_cli & "',"
           strSQL = strSQL & F_GravaDec(!limite_cli) & ", "
           strSQL = strSQL & "'" & !NUMERO_cli & "',"
           strSQL = strSQL & "'" & !COMPLEMENTO_cli & "',"
           strSQL = strSQL & "'" & !ie_Cli & "',"
           strSQL = strSQL & "'" & cod_cid & "')"
        '   On Error Resume Next
    conn.Execute strSQL
    
        Me.caption = .RecordCount
        .MoveNext
        
    Loop
    MsgBox "Migrado"
End With
End Sub

Private Sub Form_Load()
    data_padrao = "01/12/2099"
End Sub
