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
   Begin VB.CommandButton Command7 
      Caption         =   "cUPONS"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
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
    Dim rPesq As New ADODB.Recordset
    strSQL = "delete from tab_categorias"
    conn.Execute strSQL
    strSQL = "select * from arq_03"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "insert into tab_categorias(cod_cat,nome_cat)values("
            strSQL = strSQL & "'" & !codigo & "',"
            strSQL = strSQL & "'" & !descricao & "')"
            conn.Execute strSQL
            
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "delete from " & TAB_PRODUTOS & ""
    conn.Execute strSQL
    strSQL = "delete from tab_codbarras"
    conn.Execute strSQL
    strSQL = "delete from tab_embalagens"
    conn.Execute strSQL
    strSQL = "select * from arq_07"
    rPesq.Open strSQL, conn_client
    
    'P_Grava rPesq
            
    
    
End Sub
Sub P_Grava(rPesq As ADODB.Recordset)


    Dim strSQL As String
    Dim sCod As String
    Dim sMar As String
    Dim iCount As Integer
    With rPesq
        Do While Not .EOF
            sCod = F_BuscaUmCampo(" max(cod_pro)", TAB_PRODUTOS, "cod_pro like '" & rPesq!grupo & "%'")
            If sCod = "" Then
                sCod = "0001"
            Else
        
                sCod = Format(Int(Mid(sCod, 4, 4)) + 1, "000#")
            End If
            sCod = F_Ean8(rPesq!grupo & sCod)
    
            strSQL = "insert into " & TAB_PRODUTOS & "(cod_pro,nome_pro,unidade_pro,"
            strSQL = strSQL & "cod_cat,cod_mar,codfab_pro,"
            strSQL = strSQL & "lucro_pro,"
            strSQL = strSQL & "endereco_pro,er_pro,ef_pro,st_pro,cr_pro,cf_pro,cm_pro,pv_pro,balanca_pro,cod_old,cod_imp)Values("
            strSQL = strSQL & "'" & sCod & "' ,"
            
            strSQL = strSQL & "'" & F_AjustaTexto(!descricao) & "',"
            strSQL = strSQL & "'" & !unidade & "',"
            strSQL = strSQL & "'" & !grupo & "',"
            strSQL = strSQL & "'" & !subgrupo & "', "
            strSQL = strSQL & "'" & !codigo & "',"
            strSQL = strSQL & F_GravaDec(!markup) & ","
            strSQL = strSQL & "'" & !Endereco & "', "
            strSQL = strSQL & F_GravaDec(!sau) & ","
            strSQL = strSQL & F_GravaDec(!sau) & ","
            strSQL = strSQL & "'" & !sit_trib & "',"
            strSQL = strSQL & F_GravaDec(!p_custo) & ","
            strSQL = strSQL & F_GravaDec(!p_custo) & ","
            strSQL = strSQL & F_GravaDec(!p_cusmd) & ","
            strSQL = strSQL & F_GravaDec(!p_venda) & ","
            strSQL = strSQL & "'N',"
            strSQL = strSQL & "'" & !codigo & "',"
            strSQL = strSQL & "'" & !icms_d & "') "
        
            
            conn.Execute strSQL
            
    
    
    
        
            strSQL = "insert into tab_codbarras(cod_pro,codbarra_pro)values("
            strSQL = strSQL & "'" & sCod & "',"
            strSQL = strSQL & "'" & !cod_barra & "')"
            conn.Execute strSQL
                
                
    
            strSQL = "insert into tab_embalagens(cod_pro,qtd_emb)values("
            strSQL = strSQL & "'" & sCod & "',"
            strSQL = strSQL & !embalagem & ")"
            conn.Execute strSQL
                
                
            .MoveNext
            Me.caption = .AbsolutePosition
            Me.Refresh
        Loop
    End With
    Exit Sub

End Sub

Private Sub Command3_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "delete from tab_fornecedores"
    conn.Execute strSQL
    strSQL = "delete from tab_telfor"
    conn.Execute strSQL
    
    strSQL = "select * from arq_02"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            gravafor rPesq
            .MoveNext
        Loop
        .Close
    End With
End Sub
Sub gravafor(rPesq As ADODB.Recordset)
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
    strSQL = strSQL & "'" & !Endereco & "',"
    strSQL = strSQL & "'" & !BAIRRO & "',"
    strSQL = strSQL & "'" & F_AjustaTexto(Format(!Cidade)) & "',"
    strSQL = strSQL & "'" & !Estado & "',"
    strSQL = strSQL & "'" & !cep & "',"
    strSQL = strSQL & "'" & !telefone & "',"
    strSQL = strSQL & "'" & !fone_02 & "',"
    strSQL = strSQL & "'" & !contato & "',"
    strSQL = strSQL & "'" & !email & "',"
    strSQL = strSQL & "'" & !codigo & "',"
    strSQL = strSQL & "'" & !obs1 & "') "
    
    conn.Execute strSQL
    If Not IsNull(!telefone) Then
        strSQL = "insert into tab_telfor(cod_For,num_tel,tipo_tel)values("
        strSQL = strSQL & "'" & sCod & "',"
        strSQL = strSQL & "'" & !telefone & "','2')"
        conn.Execute strSQL
    End If
    If Not IsNull(!fone_02) Then
        strSQL = "insert into tab_telfor(cod_For,num_tel,tipo_tel)values("
        strSQL = strSQL & "'" & sCod & "',"
        strSQL = strSQL & "'" & !fone_02 & "','2')"
        On Error Resume Next
        conn.Execute strSQL
    End If
    End With
Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub Command4_Click()
        Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "delete from tab_marcas"
    conn.Execute strSQL
    strSQL = "select * from arq_04"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "insert into tab_marcas(cod_mar,nome_mar)values("
            strSQL = strSQL & "'" & !codigo & "',"
            strSQL = strSQL & "'" & !descricao & "')"
            conn.Execute strSQL
            
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command5_Click()
       Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "delete from " & TAB_CLIENTES
    conn.Execute strSQL
    
    strSQL = "select * from arq_01"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            GravaCli rPesq
            .MoveNext
        Loop
        .Close
        MsgBox "Migrado"
    End With
End Sub
Sub GravaCli(rPesq As ADODB.Recordset)
    Dim strSQL As String
    Dim sCod As String
    sCod = F_BuscaUmCampo(" max(cod_Cli)", TAB_CLIENTES, "")
    If sCod = "" Then       'pega só 4 dígitos para compor o dígito verificador
        sCod = "0001"       'que é o último número do cod_cli
    Else
        sCod = Left(sCod, 4)
        sCod = Format(Int(sCod) + 1, "000#")
    End If
    sCod = F_GeraDig(sCod)
    With rPesq
    
    
    strSQL = "insert into " & TAB_CLIENTES & "(cod_cli,nome_cli,cpf_cli,rg_cli,endereco_Cli, "
    strSQL = strSQL & "bairro_cli,cidade_Cli,uf_cli,cep_cli,telefone_Cli,"
    strSQL = strSQL & "nasc_cli,mae_cli,"
    strSQL = strSQL & "status_cli,tipo_cli,cel_cli,obs_Cli,cod_old) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & !nome & "',"
    If !status = "J" Then
        strSQL = strSQL & "'" & !cgc_ & "',"
    Else
        strSQL = strSQL & "'" & !cpf_ & "',"
        
    End If
    strSQL = strSQL & "'" & !inscr & "',"
    strSQL = strSQL & "'" & !Endereco & "',"
    strSQL = strSQL & "'" & !BAIRRO & "',"
    strSQL = strSQL & "'" & !Cidade & "',"
    strSQL = strSQL & "'" & !Estado & "',"
    strSQL = strSQL & "'" & !cep & "',"
    strSQL = strSQL & "'" & !telefone & "',"
    strSQL = strSQL & F_GravaData("") & ","
   strSQL = strSQL & "'" & "',"
 
    
    strSQL = strSQL & "'1',"
        strSQL = strSQL & "'" & !status & "',"
    strSQL = strSQL & "'" & !fax & "',"
    strSQL = strSQL & "'" & !obsER1 & "', "
    strSQL = strSQL & "'" & !codigo & "') "
    conn.Execute strSQL
    End With
End Sub



Private Sub Command7_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rPesq2 As New ADODB.Recordset
    strSQL = "select num_ven from tab_saidas order by num_ven"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            P_GravaCupom !num_ven
            Me.caption = .AbsolutePosition + 1 & " de " & .RecordCount
            Me.Refresh
            .MoveNext
        Loop
        .Close
    End With
    
    
    
End Sub

Private Sub Form_Load()
    If Not F_AbreBancoODBC(b, "symac") Then
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    b.Close
End Sub
