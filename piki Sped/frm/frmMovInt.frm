VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMovInt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Movimento Interno"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnImprVenda 
      Caption         =   "Impr. Venda"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7920
      Picture         =   "frmMovInt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnConcluir 
      Caption         =   "&Concluir (F7)"
      Height          =   615
      Left            =   6840
      Picture         =   "frmMovInt.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6885
      Width           =   2175
   End
   Begin VB.TextBox txtbxDesc 
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   6840
      TabIndex        =   34
      Top             =   5880
      Width           =   4215
   End
   Begin VB.ComboBox cmbxTipo 
      Height          =   315
      ItemData        =   "frmMovInt.frx":0294
      Left            =   6840
      List            =   "frmMovInt.frx":029E
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   6600
         Width           =   6375
         Begin VB.TextBox txtbxProduto 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   720
            TabIndex        =   29
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   28
            Text            =   "1"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtbxUnit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5160
            TabIndex        =   27
            Text            =   "0,00"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Qtd:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   300
         End
         Begin VB.Label lblUltItem 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Produto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   720
            TabIndex        =   31
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Custo:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5160
            TabIndex        =   30
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.TextBox txtbxGrid 
         Height          =   285
         Left            =   960
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   6375
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   11245
         _Version        =   393216
         BackColorBkg    =   14868702
         GridColor       =   16777215
         GridColorFixed  =   12632256
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9120
      Picture         =   "frmMovInt.frx":02D6
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   9120
      Picture         =   "frmMovInt.frx":0420
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6885
      Width           =   2055
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Impr. Custo"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6840
      Picture         =   "frmMovInt.frx":056A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6240
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C6DADD&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   6840
      ScaleHeight     =   4905
      ScaleWidth      =   4305
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Custo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   870
         Left            =   0
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblCodBarr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2040
         TabIndex        =   18
         Top             =   3840
         Width           =   2325
      End
      Begin VB.Label lblCat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   3240
         Width           =   2325
      End
      Begin VB.Label lblCapPro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   15
         TabIndex        =   16
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label lblProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         TabIndex        =   15
         Top             =   2520
         Width           =   4845
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         Height          =   195
         Left            =   15
         TabIndex        =   14
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label lblMarca 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         TabIndex        =   13
         Top             =   3120
         Width           =   1845
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   195
         Left            =   2040
         TabIndex        =   12
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label lblCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   3840
         Width           =   1965
      End
      Begin VB.Label lblPreco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         TabIndex        =   10
         Top             =   4560
         Width           =   1725
      End
      Begin VB.Label lblRef 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   4560
         Width           =   1965
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência:"
         Height          =   195
         Left            =   2040
         TabIndex        =   8
         Top             =   4320
         Width           =   825
      End
      Begin VB.Label lblProd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Tag             =   "!"
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   870
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód:"
         Height          =   195
         Left            =   15
         TabIndex        =   3
         Top             =   3600
         Width           =   330
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. Barra:"
         Height          =   195
         Left            =   2040
         TabIndex        =   2
         Top             =   3600
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         Height          =   195
         Left            =   15
         TabIndex        =   1
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label lblCusto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   870
         Left            =   1680
         TabIndex        =   39
         Tag             =   "!"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblValor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   870
         Left            =   1680
         TabIndex        =   4
         Tag             =   "!"
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6840
      TabIndex        =   36
      Top             =   5040
      Width           =   360
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Descrição:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6840
      TabIndex        =   35
      Top             =   5640
      Width           =   765
   End
End
Attribute VB_Name = "frmMovInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Sub P_InsereProd()
    Dim strSQL As String
    Dim iCount As Integer
    On Error GoTo trata_erro
    strSQL = "delete from tab_itmov where num_mov='" & Me.tag & "' "
    conn.Execute strSQL
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            strSQL = "insert into tab_itmov(num_mov,cod_pro,custo_ite,qtd_ite,seq_ite)values("
        
            strSQL = strSQL & "'" & Me.tag & "', "
            .col = 5
            strSQL = strSQL & "'" & .Text & "',"
            .col = 3
            strSQL = strSQL & F_GravaDec(.Text) & ","
            .col = 2
            strSQL = strSQL & F_GravaDec(.Text) & ","
            strSQL = strSQL & iCount & ")"
            conn.Execute strSQL
        Next
    End With

    'RESERVANDO O ESTOQUE
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = False
    Exit Sub
trata_erro:
    MostraErro
End Sub
Sub P_Conclui()
    Dim strSQL As String
    Dim iCount As Integer
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            
            strSQL = "update tab_produtos set er_pro=er_pro-"
            .col = 2
            strSQL = strSQL & F_GravaDec(.Text) & "  "
            
            strSQL = strSQL & " where cod_pro="
            .col = 5
            strSQL = strSQL & "'" & .Text & "' "
            
            conn.Execute strSQL
        Next
        strSQL = "update tab_movint set situacao_mov=2 "
        strSQL = strSQL & " where num_mov='" & Me.tag & "'  "
        conn.Execute strSQL
        MsgBox "Movimento concluído!", vbInformation
        
    End With
    Unload Me
    
End Sub

Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "select tv.*,tp.*,tg.nome_cat,tg.nome_cat,"
    strSQL = strSQL & "te.nome_mar,te.cod_mar,ti.*,(ti.custo_ite*ti.qtd_ite) as total "
    strSQL = strSQL & " from tab_produtos as tp,tab_movint as tv, tab_categorias as tg,"
    strSQL = strSQL & " tab_marcas as te, tab_Itmov as ti   "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro and tp.cod_cat=tg.cod_cat "
    
    strSQL = strSQL & " and tp.cod_mar=te.cod_mar "
    strSQL = strSQL & " and ti.num_mov=tv.num_mov and tv.num_mov='" & Me.tag & "' "
    strSQL = strSQL & " order by seq_ite "
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    
            
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub insere()
    Dim iPos As Integer
    Dim i As Integer
    Dim p As Produto
    Dim strSQL As String
            
            
            
            
            p = F_BuscaProdSaida(txtbxProduto.Text)
            If p.cod_pro <> "" Then
                
                txtbxProduto.Text = p.nome_pro
                txtbxProduto.tag = p.cod_pro
                lblCat.caption = p.nome_cat
                lblCod.caption = p.cod_pro
                lblCodBarr.caption = p.codbarra_pro
                lblMarca.caption = p.nome_mar
                lblProduto.caption = p.nome_pro
                lblPreco.caption = p.cr_pro
                lblRef.caption = p.codfab_pro
                
                
                lblProduto.caption = p.nome_pro
                
                If p.cr_pro = 0 Then
                    MsgBox "Este produto encontra-se com o custo (c2) zerado!", vbCritical
                    Exit Sub
                End If
                
                
                If p.totalbalanca_pro <> 0 Then
                        If strBALANCA = "V" Then 'valor
                            txtbxQtd.Text = p.totalbalanca_pro / p.pv_pro
                        Else
                            txtbxQtd.Text = Round(p.totalbalanca_pro / p.pv_pro, 3)
                            txtbxProduto.Text = p.nome_pro
                            txtbxProduto.tag = p.cod_pro
                        End If
                        txtbxQtd_LostFocus
                        txtbxUnit.Text = p.pv_pro
                Else
                    If txtbxQtd.Text = "" Then
                        txtbxQtd.Text = "1"
                    End If
                    If p.balanca_pro <> "S" Then
                        txtbxQtd.Text = Int(txtbxQtd.Text)
                    End If
                    
                End If
                
                lblProduto.caption = p.nome_pro
                If strContrpre_usu = "S" Then
                    txtbxUnit.Enabled = True
                    txtbxUnit.Text = p.cr_pro
                    txtbxUnit.SetFocus
                    'cmndbtnInserir.Visible = True
                    Exit Sub
                End If
                If p.cod_pro <> "" Then
                    strSQL = msflxgrd.rows & Chr(9) & p.nome_pro & Chr(9) & txtbxQtd.Text & Chr(9)
                    strSQL = strSQL & FormatVal(p.cr_pro) & Chr(9)
                    strSQL = strSQL & FormatVal(p.cr_pro * CCur(txtbxQtd.Text)) & Chr(9) & p.cod_pro
                    msflxgrd.AddItem strSQL
                    
                    lblValor.caption = FormatVal(CCur(lblValor.caption) + p.cr_pro * CCur(txtbxQtd.Text))
'                    lblCusto.caption = FormatVal(CCur(lblCusto.caption) + p.c2_pro * CCur(txtbxQtd.Text))
                End If
                
                
                
                    
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
                    
                
        End If
        
        
    ZebraGrid msflxgrd, 2, &HD3E7EB
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = True
    lblProd.caption = msflxgrd.rows
End Sub

Private Sub cmndbtnConcluir_Click()
        If cmbxTipo.ListIndex = -1 Then
            MsgBox "Informe o tipo do movimento interno!", vbCritical
            Exit Sub
        End If
        cmndbtnSalvar_Click
        
        P_Conclui
End Sub

Private Sub cmndbtnImprimir_Click()
    ImprMovInt Me.tag
End Sub

Private Sub cmndbtnImprVenda_Click()
    ImprMovInt Me.tag
End Sub

Private Sub cmndbtnSair_Click()
    If cmndbtnSalvar.Enabled = True Then
        If MsgBox("Você fez alterações no orçamento, deseja Salvar?", vbYesNo + vbQuestion) = vbYes Then
            salva
        End If
    End If
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If msflxgrd.rows = 1 Then
        MsgBox "Não há o que ser salvo!", vbCritical
        Exit Sub
    End If
        If cmbxTipo.Text = "" Then
            MsgBox "Primeiro informe o tipo de Movimento Interno!", vbExclamation
            cmbxTipo.SetFocus
            Exit Sub
        End If
        salva

    
        
    
    
    
End Sub
Sub salva()
    Dim strSQL As String
    Dim iCount As Integer
  
    If Me.tag = "" Then
        strSQL = F_GeraNovoMovInt(Left(cmbxTipo.Text, 1), txtbxDesc.Text, lblValor.caption)
        Me.tag = strSQL
    Else
        strSQL = "update tab_movint set valor_mov=" & F_GravaDec(lblValor) & ","
        strSQL = strSQL & "tipo_mov='" & Left(cmbxTipo.Text, 1) & "',"
        strSQL = strSQL & "desc_mov='" & txtbxDesc.Text & "' "
        strSQL = strSQL & " where num_mov='" & Me.tag & "' "
        conn.Execute strSQL
        strSQL = Me.tag
    End If
    If strSQL <> "" Then
        P_InsereProd
    End If
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_erro
    strSQL = "select * from tab_movint where num_mov='" & Me.tag & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        txtbxDesc.Text = !Desc_mov
        .Close
        
    End With
    
    strSQL = "select tp.*,ti.* from tab_itmov as ti,tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and nuM_mov='" & Me.tag & "' "
    strSQL = strSQL & " order by seq_ite "
    rPesq.Open strSQL, conn_client
    With rPesq
        lblValor.caption = "0"
        Do While Not .EOF
            strSQL = msflxgrd.rows & Chr(9) & !nome_pro & Chr(9) & !qtd_ite & Chr(9)
            strSQL = strSQL & FormatVal(!custo_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!custo_ite * !qtd_ite) & Chr(9) & !cod_pro
            msflxgrd.AddItem strSQL
                    
            lblValor.caption = FormatVal(CCur(lblValor.caption) + !custo_ite * !qtd_ite)
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid msflxgrd, 1
    cmndbtnSalvar.Enabled = True
    cmndbtnImprimir.Enabled = True
    cmndbtnImprVenda.Enabled = True
    Exit Sub
    
trata_erro:
    MostraErro
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        editaItem
    End If
    If KeyCode = vbKeyF7 Then
        cmndbtnConcluir_Click
    End If
End Sub

Private Sub txtbxDesc_GotFocus()
    ganhaFoco txtbxDesc
End Sub

Private Sub txtbxDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtbxDesc_LostFocus()
    AjustaTexto txtbxDesc
End Sub

Private Sub Form_Load()
    bActive = False
    Me.Top = 10: Me.Left = 10
    With msflxgrd
        .FixedCols = 0
        .rows = 1
        .Cols = 6
        .col = 0: .Text = "..."
        .col = 1: .Text = "Descrição"
        .col = 2: .Text = "Qtd"
        .col = 3: .Text = "Unit.(R$)"
        .col = 4: .Text = "Total(R$)"
        .col = 5: .Text = "Código"
        .ColWidth(0) = 300
        .ColWidth(1) = 3500
        .ColWidth(2) = 600
        .ColWidth(3) = 700
        .ColWidth(4) = 800
        .ColWidth(5) = 2000
        
    End With
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyF4 Then
        exclui
    End If
    If KeyCode = vbKeyReturn Then
        With msflxgrd
            If .col = 2 Or .col = 3 Then
                txtbxGrid.Text = .Text
                txtbxGrid.Left = .CellLeft + msflxgrd.Left
                txtbxGrid.Top = .CellTop + msflxgrd.Top
                txtbxGrid.Height = .CellHeight
                txtbxGrid.Width = .CellWidth
                txtbxGrid.Visible = True
                txtbxGrid.SetFocus
                
            End If
        End With
    End If
End Sub
Sub exclui()
    With msflxgrd
        If .rows = 2 Then
            .col = 4
            lblValor.caption = FormatVal(CCur(lblValor.caption) - CCur(.Text))
            msflxgrd.rows = 1
        ElseIf .rows > 2 Then
            
            .col = 4
            lblValor.caption = FormatVal(CCur(lblValor.caption) - CCur(.Text))
            msflxgrd.RemoveItem .Row
        Else
            MsgBox "Não há o que excluir!", vbCritical
        End If
    End With
    If msflxgrd.rows > 1 Then
        cmndbtnImprimir.Enabled = True
    End If
    cmndbtnSalvar.Enabled = True
    ZebraGrid msflxgrd, 2, &HD3E7EB
    lblProd.caption = msflxgrd.rows
End Sub


Private Sub txtbxGrid_GotFocus()
    ganhaFoco txtbxGrid
End Sub

Private Sub txtbxGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        PerdeFoco txtbxGrid
        If CCur(txtbxGrid.Text) <= 0 Then
            MsgBox "Valor incorreto!", vbCritical
            txtbxGrid.SetFocus
            Exit Sub
        End If
        msflxgrd.Text = txtbxGrid.Text
        txtbxGrid.Visible = False
        SomaLinha
        SomaGrid
    End If
End Sub

Private Sub txtbxGrid_LostFocus()
    PerdeFoco txtbxGrid
    txtbxGrid.Visible = False
End Sub
Sub SomaGrid()
    Dim iCount As Integer
    Dim cTotal As Currency
    With msflxgrd
        For iCount = 1 To .rows - 1
            .Row = iCount
            .col = 4
            cTotal = cTotal + CCur(.Text)
            
        Next
        lblValor.caption = FormatVal(cTotal)
        lblProd.caption = .rows - 1
    End With
End Sub
Sub SomaLinha()
    Dim cQtd As Currency
    Dim cUnit As Currency
    msflxgrd.col = 2
    cQtd = CCur(msflxgrd.Text)
    msflxgrd.col = 3
    cUnit = CCur(msflxgrd.Text)
    msflxgrd.col = 4
    msflxgrd.Text = FormatVal(cUnit * cQtd)
    
End Sub
Private Sub txtbxProduto_Change()
    txtbxProduto.tag = ""
End Sub

Private Sub txtbxProduto_GotFocus()
    ganhaFoco txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmbxTipo.Text = "" Then
            MsgBox "Primeiro informe o tipo de Movimento Interno!", vbExclamation
            cmbxTipo.SetFocus
            Exit Sub
        End If
        If F_Vazio(txtbxProduto.Text) Then
            MsgBox "Informe o produto!", vbExclamation
            txtbxProduto.SetFocus
            Exit Sub
        End If
        insere
    End If
End Sub

Private Sub txtbxQtd_GotFocus()
    
    ganhaFoco txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo trata_erro
    If KeyCode = vbKeyReturn Then
        If CCur(txtbxQtd.Text) = 0 Then
            MsgBox "O Valor Informado não está Correto!", vbCritical
            txtbxQtd.SetFocus
            Exit Sub
        End If
        txtbxQtd.Enabled = False
        txtbxProduto.SetFocus
    End If
Exit Sub
trata_erro:
    MsgBox Err.Description, vbcre
End Sub

Private Sub txtbxQtd_LostFocus()
    On Error Resume Next
    If CCur(txtbxQtd.Text) > 999 Then
        If MsgBox("A quantidade digitada é muito grande." & _
            "Deseja Continuar?", vbYesNo + vbCritical + vbDefaultButton2, "Atenção!!!") = vbNo Then
            txtbxQtd.Text = "1"
            Exit Sub
        End If
    End If
    PerdeFoco txtbxQtd, "###,##0.000"
    txtbxQtd.Enabled = False
    'txtbxProduto.Text = ""
    txtbxProduto.Enabled = True
    txtbxProduto.SetFocus
    
    
    
End Sub
Sub editaItem()
    txtbxQtd.Enabled = True
    'txtbxQtd.Text = "1"
    txtbxQtd.SetFocus
End Sub

Private Sub txtbxUnit_GotFocus()
    ganhaFoco txtbxUnit
End Sub

Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        insereManual
        
    End If
End Sub
Sub insereManual()
    Dim strSQL As String
    If txtbxProduto.tag = "" Then
        MsgBox "Informe o produto!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxUnit.Text) <= 0 Then
        MsgBox "O valor unitário está zerado!", vbExclamation
        Exit Sub
    End If
    If CCur(txtbxQtd.Text) <= 0 Then
        MsgBox "Quantidade está zerada!", vbCritical
        Exit Sub
    End If
    
    strSQL = msflxgrd.rows & Chr(9) & txtbxProduto.Text & Chr(9) & txtbxQtd.Text & Chr(9)
    strSQL = strSQL & FormatVal(txtbxUnit.Text) & Chr(9)
    strSQL = strSQL & FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text)) & Chr(9) & txtbxProduto.tag
    msflxgrd.AddItem strSQL
                    
    lblValor.caption = FormatVal(CCur(lblValor.caption) + CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text))
    ZebraGrid msflxgrd, 2, &HD3E7EB
    cmndbtnImprimir.Enabled = True
    cmndbtnSalvar.Enabled = True
    txtbxProduto.Text = ""
    'txtbxQtd.Text = "1"
    txtbxProduto.SetFocus
    
End Sub
Private Sub txtbxUnit_LostFocus()
    PerdeFoco txtbxUnit
End Sub



