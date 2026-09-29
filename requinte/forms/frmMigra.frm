VERSION 5.00
Begin VB.Form frmMigra 
   Caption         =   "Migração"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Fabricantes"
      Height          =   1095
      Left            =   2640
      TabIndex        =   7
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Principio Ativo"
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Contas Rec"
      Height          =   1215
      Left            =   2640
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cliente:"
      Height          =   1215
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Marcas"
      Height          =   1215
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fornecedores"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Produtos"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Categorias"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmMigra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As database

Private Sub Command1_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "delete from tab_categorias"
    Banco.Execute strSQL
    strSQL = "select * from arq_03"
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = "insert into tab_categorias(cod_cat,nome_cat)values("
            strSQL = strSQL & "'" & !Codigo & "',"
            strSQL = strSQL & "'" & !Descricao & "')"
            Banco.Execute strSQL
            
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim iCount As Integer
    Dim sCodSub As String
    Dim sCodPri As String
    strSQL = "delete from tab_produtos"
    
    Banco.Execute strSQL
    strSQL = "delete from tab_estoque"
    Banco.Execute strSQL
    strSQL = "select * from produto "
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            If !secao = "01" Then
                sCodSub = "003"
            ElseIf !secao = "02" Then
                sCodSub = "007"
            ElseIf !secao = "03" Then
                sCodSub = "005"
            ElseIf !secao = "04" Then
                sCodSub = "008"
            ElseIf !secao = "05" Then
                sCodSub = "004"
            ElseIf !secao = "06" Then
                sCodSub = "001"
            ElseIf !secao = "07" Then
                sCodSub = "006"
            End If
            
            sCodPri = F_BuscaUmCampo("cod_pri", "tab_principios", "nome_pri='" & !prsal & "'")
            If sCodPri = "" Then
                sCodPri = "01"
            End If
            
            strSQL = Produtos.F_GravaProduto(F_AjustaTexto(!prdesc), Format(!codbarra), sCodSub, _
                sCodPri, "", Format(!prfabr), Format(!prcons), "UN", "", 0, "", "000", "", "", "", "", "", "")
            If strSQL <> "" Then
                strSQL = Produtos.F_AddEstoque(strSQL, "", "", 0, 0, Format(!prfabr))
            End If
            
            
            
            
            
            Me.caption = .AbsolutePosition & "/" & .RecordCount
            Me.Refresh
            .MoveNext
        Loop
    End With
            
    
    
End Sub
Sub P_Grava(rPesq As Recordset)


    Dim strSQL As String
    Dim sCod As String
    Dim sMar As String
    Dim iCount As Integer
    With rPesq
        Do While Not .EOF
            sCod = F_BuscaUmCampo(" max(cod_pro)", "tab_produtos", "cod_pro like '" & rPesq!grupo & "%'")
            If sCod = "" Then
                sCod = "0001"
            Else
        
                sCod = Format(Int(Mid(sCod, 4, 4)) + 1, "000#")
            End If
            sCod = F_Ean8(rPesq!grupo & sCod)
    
            strSQL = "insert into tab_produtos(cod_pro,nome_pro,unidade_pro,"
            strSQL = strSQL & "cod_cat,cod_mar,codfab_pro,"
            strSQL = strSQL & "lucro_pro,"
            strSQL = strSQL & "endereco_pro,er_pro,ef_pro,st_pro,cr_pro,cf_pro,cm_pro,pv_pro,balanca_pro,cod_old,cod_imp)Values("
            strSQL = strSQL & "'" & sCod & "' ,"
            
            strSQL = strSQL & "'" & F_AjustaTexto(!Descricao) & "',"
            strSQL = strSQL & "'" & !unidade & "',"
            strSQL = strSQL & "'" & !grupo & "',"
            strSQL = strSQL & "'" & !subgrupo & "', "
            strSQL = strSQL & "'" & !Codigo & "',"
            strSQL = strSQL & F_GravaDec(!markup) & ","
            strSQL = strSQL & "'" & !endereco & "', "
            strSQL = strSQL & F_GravaDec(!sau) & ","
            strSQL = strSQL & F_GravaDec(!sau) & ","
            strSQL = strSQL & "'" & !sit_trib & "',"
            strSQL = strSQL & F_GravaDec(!p_custo) & ","
            strSQL = strSQL & F_GravaDec(!p_custo) & ","
            strSQL = strSQL & F_GravaDec(!p_cusmd) & ","
            strSQL = strSQL & F_GravaDec(!p_venda) & ","
            strSQL = strSQL & "'N',"
            strSQL = strSQL & "'" & !Codigo & "',"
            strSQL = strSQL & "'" & !icms_d & "') "
        
            
            Banco.Execute strSQL
            
    
    
    
        
            strSQL = "insert into tab_codbarras(cod_pro,codbarra_pro)values("
            strSQL = strSQL & "'" & sCod & "',"
            strSQL = strSQL & "'" & !cod_barra & "')"
            Banco.Execute strSQL
                
                
    
            strSQL = "insert into tab_embalagens(cod_pro,qtd_emb)values("
            strSQL = strSQL & "'" & sCod & "',"
            strSQL = strSQL & !embalagem & ")"
            Banco.Execute strSQL
                
                
            .MoveNext
            Me.caption = .AbsolutePosition
            Me.Refresh
        Loop
    End With
    Exit Sub

End Sub

Private Sub Command3_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "delete from tab_fornecedores"
    Banco.Execute strSQL
    strSQL = "delete from tab_telefones where tipo_ter='F'"
'    Banco.Execute strSQL
    
    strSQL = "select * from fornecedores"
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            'fornecedor.F_GravaFor(
            .MoveNext
        Loop
        .Close
    End With
End Sub
Sub gravafor(rPesq As Recordset)
    Dim strSQL As String
    Dim sCod As String
'    On Error GoTo trata_erro
    sCod = F_BuscaUmCampo(" max(cod_for)", "tab_fornecedores", "")
    If sCod = "" Then       'pega só 4 dígitos para compor o dígito verificador
        sCod = "0001"       'que é o último número do cod_cli
    Else
        sCod = Left(sCod, 4)
        sCod = Format(Int(sCod) + 1, "000#")
    End If
    sCod = F_GeraDig(sCod)
    
    With rPesq
    
    strSQL = "insert into tab_fornecedores(cod_for,nome_for,fantasia_for,cnpj_for,ie_for,endereco_for, "
    strSQL = strSQL & "bairro_for,cidade_for,uf_for,cep_for,telefone_for,fax_for,contato_for,email_for,codantigo_for,obs_for) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & !nome & "',"
    strSQL = strSQL & "'" & !fantasia & "',"
    strSQL = strSQL & "'" & !cgc_ & "',"
    strSQL = strSQL & "'" & !inscr & "',"
    strSQL = strSQL & "'" & !endereco & "',"
    strSQL = strSQL & "'" & !bairro & "',"
    strSQL = strSQL & "'" & F_AjustaTexto(Format(!cidade)) & "',"
    strSQL = strSQL & "'" & !Estado & "',"
    strSQL = strSQL & "'" & !cep & "',"
    strSQL = strSQL & "'" & !Telefone & "',"
    strSQL = strSQL & "'" & !fone_02 & "',"
    strSQL = strSQL & "'" & !contato & "',"
    strSQL = strSQL & "'" & !email & "',"
    strSQL = strSQL & "'" & !Codigo & "',"
    strSQL = strSQL & "'" & !obs1 & "') "
    
    Banco.Execute strSQL
    If Not IsNull(!Telefone) Then
        strSQL = "insert into tab_telfor(cod_For,num_tel,tipo_tel)values("
        strSQL = strSQL & "'" & sCod & "',"
        strSQL = strSQL & "'" & !Telefone & "','2')"
        Banco.Execute strSQL
    End If
    If Not IsNull(!fone_02) Then
        strSQL = "insert into tab_telfor(cod_For,num_tel,tipo_tel)values("
        strSQL = strSQL & "'" & sCod & "',"
        strSQL = strSQL & "'" & !fone_02 & "','2')"
        On Error Resume Next
        Banco.Execute strSQL
    End If
    End With
Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub Command4_Click()
        Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "delete from tab_marcas"
    Banco.Execute strSQL
    strSQL = "select * from arq_04"
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = "insert into tab_marcas(cod_mar,nome_mar)values("
            strSQL = strSQL & "'" & !Codigo & "',"
            strSQL = strSQL & "'" & !Descricao & "')"
            Banco.Execute strSQL
            
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command5_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "delete from tab_clientes"
    Banco.Execute strSQL
    
    strSQL = "delte from tab_telefones where tipo_ter='C' "
    
   
    
    strSQL = "delete from tab_autorizados"
    Banco.Execute strSQL
    strSQL = "delete from tab_referencias"
    Banco.Execute strSQL
     strSQL = "select * from cliente"
    
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            GravaCli rPesq
            .MoveNext
        Loop
        .Close
        MsgBox "Migrado"
    End With
End Sub
Sub GravaCli(rPesq As Recordset)
    Dim strSQL As String
    Dim sCod As String
    
    
    'scod = clientes.F_GravaCli(!clnome,!cldtnasc,!clcpf,!clrg,!clende,!clbai,!cldia,"GO",!clcep,!clobs
    
    sCod = MaxKeySQL("tab_clientes", "cod_cli", "", "0000#")
    With rPesq
    
    
    
    
    
    strSQL = "insert into tab_clientes(cod_cli,nome_cli,cpf_cli,rg_cli,endereco_Cli, "
    strSQL = strSQL & "bairro_cli,cidade_Cli,uf_cli,cep_cli,"
    strSQL = strSQL & "datanasc_cli,mae_cli,"
    strSQL = strSQL & "status_cli,tipo_cli,obs_Cli,cod_old) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & !clNome & "',"
    
        strSQL = strSQL & "'" & !clcpf & "',"
        
    
    strSQL = strSQL & "'" & !clrg & "',"
    strSQL = strSQL & "'" & !clEnde & "',"
    strSQL = strSQL & "'" & !clbairro & "',"
    strSQL = strSQL & "'" & !clcida & "',"
    strSQL = strSQL & "'GO',"
    strSQL = strSQL & "'" & !clcep & "',"
    
    strSQL = strSQL & F_GravaData(!cldtnasc) & ","
    strSQL = strSQL & "'" & !clmae & "',"
 
    
    strSQL = strSQL & "'1',"
    strSQL = strSQL & "'F',"
    
    strSQL = strSQL & "'" & !clobs & "', "
    strSQL = strSQL & "'" & !clcodi & "') "
    Banco.Execute strSQL
    
    
    If Not IsNull(!cltele) Then
        If Not Clientes.F_GravaTelefone(sCod, !cltele, "1") Then
        
        End If
    End If
    If Not IsNull(!clfonecom) Then
        If Not Clientes.F_GravaTelefone(sCod, !clfonecom, "2") Then
        
        End If
    End If
    
    If Not IsNull(!depend1) Then
        If Clientes.F_GravaAutorizados(sCod, !depend1, "00") <> "" Then
        End If
    End If
    If Not IsNull(!depend2) Then
        If Clientes.F_GravaAutorizados(sCod, !depend2, "00") <> "" Then
        End If
    End If
    If Not IsNull(!depend3) Then
        If Clientes.F_GravaAutorizados(sCod, !depend3, "00") <> "" Then
        End If
    End If
    If Not IsNull(!depend4) Then
        If Clientes.F_GravaAutorizados(sCod, !depend4, "00") <> "" Then
        End If
    End If
    End With
End Sub

Private Sub Command6_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
     '   strSQL = "delete from tab_doc"
    ' Banco.Execute strSQL
    'strSQL = "delete from tab_itdoc"
    'Banco.Execute strSQL
    Dim sDoc As String
    Dim sIt As String
    Dim sValor As Currency
    strSQL = "select * from arq_16 where status is null order by data_venc desc "
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = F_BuscaUmCampo("cod_cli", "tab_clientes", "cod_old='" & !Codigo & "'")
            If Not IsNull(!valor_baix) Then
                
            sDoc = F_GravaDoc(!nome & "(M)", !duplicata, !valor_dupl, !data_emiss, strSQL, "R", "", "1")
            If IsNull(!valor_baix) Then
                sValor = !valor_dupl
            Else
                
                sValor = !valor_dupl - !valor_baix
            End If
            'sIt = F_GravaItDoc(sDoc, CStr(sValor), !data_venc,  "1")
            End If
            Me.caption = .AbsolutePosition
            .MoveNext
        Loop
        .Close
        MsgBox "Migrado"
    End With
    
End Sub

Private Sub Command7_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "delete from tab_principios where cod_pri >'00001'"
    Banco.Execute strSQL
    strSQL = "select distinct prsal from produto where prsal<>''"
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    
    
    With rPesq
        Do While Not .EOF
            strSQL = principio.F_GravaPri(!prsal)
            .MoveNext
        Loop
    End With
    
    
    
    
    
End Sub

Private Sub Command8_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "delete from tab_fabricantes "
    Banco.Execute strSQL
    strSQL = "select distinct laborat from abcfarma "
    Set rPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = fabricante.F_GravaFabricante(F_AjustaTexto(!laborat), "", "", "", "", "", "", "")
            If strSQL = "" Then
                MsgBox "Erro!", vbCritical
            End If
            .MoveNext
        Loop
        .Close
    End With
    
    MsgBox "Migrado"
    
End Sub

Private Sub Form_Load()
    If Not F_AbreBancoODBC(b, "farm_old") Then
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    b.Close
End Sub
