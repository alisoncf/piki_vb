VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frmConsPreco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Preços de Mercadorias"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsPreco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbxOrd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmConsPreco.frx":014A
      Left            =   3240
      List            =   "frmConsPreco.frx":0157
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox cmbxBusca 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmConsPreco.frx":017A
      Left            =   1680
      List            =   "frmConsPreco.frx":0190
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtbxExpr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      MaxLength       =   40
      TabIndex        =   5
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmndbtnBusca 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8910
      Picture         =   "frmConsPreco.frx":01D0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEF4E4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3105
      ScaleWidth      =   9225
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Label lblCategoria 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   2040
         TabIndex        =   16
         Tag             =   "!"
         Top             =   2520
         Width           =   1740
      End
      Begin VB.Label lblMarca 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   2040
         TabIndex        =   15
         Tag             =   "!"
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   120
         TabIndex        =   14
         Tag             =   "!"
         Top             =   2520
         Width           =   1875
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   120
         TabIndex        =   13
         Tag             =   "!"
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label lblPreço 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   825
         Left            =   120
         TabIndex        =   12
         Tag             =   "!"
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lblNomePro 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         TabIndex        =   11
         Tag             =   "!"
         Top             =   60
         Width           =   1650
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fechar: Esc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7920
         TabIndex        =   10
         Tag             =   "!"
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   9255
      Begin VB.Data dtPro 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmConsPreco.frx":031A
         Height          =   2295
         Left            =   120
         OleObjectBlob   =   "frmConsPreco.frx":032E
         TabIndex        =   7
         Top             =   240
         Width           =   9015
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "frmConsPreco.frx":1727
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Ordenar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E&xpressão:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Buscar Por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmConsPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBusca_Click()
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    If cmbxBusca.ListIndex = 0 Then
        sCampo = "tp.nome_pro"
    ElseIf cmbxBusca.ListIndex = 1 Then
        sCampo = "tp.codfab_pro"
    ElseIf cmbxBusca.ListIndex = 2 Then
        sCampo = "tb.codbarra_pro"
    ElseIf cmbxBusca.ListIndex = 3 Then
        sCampo = "tp.cod_pro"
    ElseIf cmbxBusca.ListIndex = 4 Then
        sCampo = "tg.nome_cat"
    Else
        sCampo = "te.nome_mar"
    End If
    If cmbxOrd.ListIndex = 1 Then
        sOrder = "tp.cod_pro,tp.nome_pro"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrder = "tp.codfab_pro,tp.nome_pro,tp.cod_pro"
    Else
        sOrder = "tp.nome_pro,tp.cod_pro "
    End If
    strSQL = "select distinct tg.nome_cat, te.nome_mar, tp.* "
    strSQL = strSQL & " from tab_produtos as tp,tab_categorias as tg, tab_codbarras as tb, "
    strSQL = strSQL & " tab_marcas as te "
    strSQL = strSQL & " where " & sCampo & " like '%"
    strSQL = strSQL & txtbxExpr.Text & "%' "
    strSQL = strSQL & " and tp.cod_pro=tb.cod_pro"
    strSQL = strSQL & " and tg.cod_cat=tp.cod_cat and te.cod_mar=tp.cod_mar "
    strSQL = strSQL & " order by " & sOrder
    dtPro.RecordSource = strSQL
    dtPro.Refresh
    If dtPro.Recordset.RecordCount > 0 Then
        Picture1.Visible = True
    Else
        Picture1.Visible = False
    End If
    dbgrd.caption = "[Listadas..: " & dtPro.Recordset.RecordCount & _
        "   Cadastradas: " & F_BuscaUmCampo("count(*)", "tab_produtos", "") & "]"
        txtbxExpr.SetFocus
        'txtbxExpr_GotFocus
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 1
End Sub

Private Sub dtPro_Reposition()
    With dtPro.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        lblNomePro.caption = !nome_pro
        lblPreço.caption = "R$ " & Format(!pV_pro, "##,###,###0.00")
        lblMarca.caption = F_BuscaUmCampo("nome_mar", "tab_marcas", "cod_mar ='" & !cod_mar & "'")
        lblCategoria.caption = F_BuscaUmCampo("nome_cat", "tab_categorias", "cod_cat ='" & !cod_cat & "'")
    End With
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    txtbxExpr.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        
    End If
End Sub

Private Sub Form_Load()
    dtPro.Connect = Banco.Connect
    P_CentraForm Me
    cmbxBusca.ListIndex = 2
    cmbxOrd.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    P_Fim
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        cmndbtnBusca_Click
        
        dbgrd.SetFocus
        txtbxExpr.SetFocus
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dbgrd.SetFocus
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
End Sub
