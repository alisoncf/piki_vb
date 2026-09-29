VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSelectProd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleção de Produtos"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7290
   ControlBox      =   0   'False
   Icon            =   "frmSelectProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Composição:"
      Height          =   1455
      Left            =   2520
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ComboBox cmbxCat 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Gênero:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblCat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1680
         TabIndex        =   32
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Produtos"
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7095
      Begin VB.Data dtPro 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmSelectProd.frx":08CA
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "frmSelectProd.frx":08DE
         TabIndex        =   8
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblLocal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
         ForeColor       =   &H8000000C&
         Height          =   195
         Left            =   6480
         TabIndex        =   30
         Top             =   4560
         Width           =   390
      End
      Begin VB.Label lblNome_Mar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         ForeColor       =   &H8000000C&
         Height          =   195
         Left            =   3000
         TabIndex        =   29
         Tag             =   "!"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Mc:"
         Height          =   195
         Left            =   2640
         TabIndex        =   28
         Top             =   4560
         Width           =   270
      End
      Begin VB.Label lblNome_Cat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         ForeColor       =   &H8000000C&
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Tag             =   "!"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cat:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   4560
         Width           =   285
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   7095
      Begin VB.ComboBox cmbxMar 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmSelectProd.frx":12C2
         Left            =   2040
         List            =   "frmSelectProd.frx":12D2
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   3600
         MaxLength       =   40
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblMar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1680
         TabIndex        =   2
         Top             =   120
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   3600
         TabIndex        =   5
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   2040
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   4800
      TabIndex        =   9
      Top             =   5400
      Width           =   2415
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmSelectProd.frx":12F6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   120
         Picture         =   "frmSelectProd.frx":1680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   4575
      Begin VB.Label lblCapPEI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PEI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   25
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblPEI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3960
         TabIndex        =   24
         Tag             =   "!"
         Top             =   480
         Width           =   465
      End
      Begin VB.Label lblCapTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   23
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3360
         TabIndex        =   22
         Tag             =   "!"
         Top             =   480
         Width           =   465
      End
      Begin VB.Label lblCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   2280
         TabIndex        =   20
         Tag             =   "!"
         Top             =   480
         Width           =   945
      End
      Begin VB.Label lblCapCusto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblEstoque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   1200
         TabIndex        =   18
         Tag             =   "!"
         Top             =   480
         Width           =   945
      End
      Begin VB.Label lblCapEst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Est.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblPreco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Tag             =   "!"
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7440
      TabIndex        =   21
      Top             =   -360
      Width           =   540
   End
End
Attribute VB_Name = "frmSelectProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxBusca_Click()
    Busca
End Sub

Private Sub cmbxDiv_Click()
    lblDiv.caption = F_BuscaUmCampo("Cod_div", "tab_divisoes", "nome_div='" & cmbxDiv.Text & "'")
    Busca
End Sub

Private Sub cmbxDiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxDiv.ListIndex = -1
    End If
End Sub

Private Sub cmbxEsp_Click()
    lblEsp.caption = F_BuscaUmCampo("Cod_Esp", "tab_Especies", "nome_esp='" & cmbxEsp.Text & "'")
    Busca
End Sub

Private Sub cmbxEsp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxEsp.ListIndex = -1
    End If
End Sub

Private Sub cmbxGenero_Click()
    lblGen.caption = F_BuscaUmCampo("Cod_Gen", "tab_Generos", "nome_gen='" & cmbxGenero.Text & "'")
    'P_PreencheDiv cmbxDiv, lblGen
    Busca
End Sub

Private Sub cmbxGenero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxGenero.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrd_Click()
    Busca
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxMar_Click()
    lblMar.caption = F_BuscaUmCampo("cod_mar", "tab_marcas", "nome_mar='" & cmbxMar.Text & "'")
End Sub

Private Sub cmbxMar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMar.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    If dtPro.Recordset.RecordCount = 0 Then
        MsgBox "Não há Produtos a serem selecionados!", vbExclamation
        Exit Sub
    End If
    If dtPro.Recordset!estoque <= 0 And Me.Tag = "V" Then
        'MsgBox "Esta Mercadoria não possui estoque!", vbExclamation
        frmVendaEstZer.lblNome_Pro.caption = dtPro.Recordset!nome_pro
        On Error Resume Next
        frmVendaEstZer.Show 1
        If frmVendaEstZer.Tag = "" Then
            Unload frmVendaEstZer
        Else
            Unload frmVendaEstZer
        End If
        
        
        Exit Sub
    End If
    Me.Tag = dtPro.Recordset!cod_pro
    Me.Hide
End Sub

Private Sub cmndbtnNew_Click()
    frmIncPro.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    Else
        If InStr(1, "NDPC", UCase(Chr(KeyCode))) > 0 Then
            
            If KeyCode = vbKeyC Then
                    frmBusca.caption = "Localizar Produtos por Código"
                Else
                    frmBusca.caption = "Localizar Produtos por Nome"
                End If
            On Error Resume Next
            frmBusca.Show 1
            If frmBusca.Tag <> "" Then
                If KeyCode = vbKeyC Then
                    cmbxBusca.ListIndex = 1
                Else
                    cmbxBusca.ListIndex = 0
                End If
                txtbxExpr.Text = frmBusca.txtbxExpr.Text
                
            End If
            Unload frmBusca
        End If
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtPro_Reposition()
    If dtPro.Recordset.RecordCount = 0 Then
    Else
        With dtPro.Recordset
            lblPreco.caption = FormatVal(dtPro.Recordset!preco_pro)
            lblEstoque.caption = FormatVal(dtPro.Recordset!estoque_pro - dtPro.Recordset!reservado_pro)
            lblCusto.caption = FormatVal(dtPro.Recordset!custo_pro)
            lblNome_Cat.caption = !nome_cat
            lblLocal.caption = !Endereco_pro
            lblNome_Mar.caption = !nome_mar
            lblPEI.caption = !pei_pro
            lblTipo.caption = !tipo_pro
        End With
    End If
End Sub

Private Sub Form_Activate()
    txtbxExpr.SetFocus
End Sub

Private Sub Form_Load()
    P_CentraForm Me, 0
    cmbxBusca.ListIndex = 0
    dtPro.Connect = banco.Connect
    
   P_PreencheMarca cmbxMar
   ' P_PreencheCateg cmbxCat
    'Busca "20"
    MudaAparencia Me
      'If Me.Tag = "C" Then
       ' dbgrd_KeyDown vbKeyC, 0
      'Else
'        dbgrd_KeyDown vbKeyN, 0
      'End If
End Sub

Private Sub txtbxExpr_Change()
    Busca
End Sub

Sub Busca(Optional limit As String = "")
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    If cmbxBusca.ListIndex = 0 Then
        sCampo = "tp.nome_pro"
    ElseIf cmbxBusca.ListIndex = 1 Then
        sCampo = "tp.cod_pro"
    ElseIf cmbxBusca.ListIndex = 2 Then
        sCampo = "nome_mar"
    ElseIf cmbxBusca.ListIndex = 1 Then
        sCampo = "nome_cat"
    End If
    
        sOrder = "tp.nome_pro,tp.cod_pro "
    
    strSQL = "select tc.nome_Cat, tm.nome_mar,"
    strSQL = strSQL & "tp.*,(tp.estoque_pro-tp.reservado_pro) as estoque"
    strSQL = strSQL & " from tab_produtos as tp,tab_categorias as tc, "
    strSQL = strSQL & " tab_marcas as tm  "
    strSQL = strSQL & " where " & sCampo & " like '"
    strSQL = strSQL & txtbxExpr.Text & "%' "
    strSQL = strSQL & " and tc.cod_cat=tp.cod_Cat and tm.cod_mar=tp.cod_mar "
    
    If cmbxCat.ListIndex <> -1 Then
        strSQL = strSQL & " and tp.cod_Cat='" & lblCat.caption & "'  "
    End If
    If cmbxMar.ListIndex <> -1 Then
        strSQL = strSQL & " and tp.cod_mar='" & lblMar.caption & "'  "
    End If
    
    strSQL = strSQL & " order by " & sOrder
    If limit <> "" Then
        strSQL = strSQL & " limit " & limit
    End If
    dtPro.RecordSource = strSQL
    dtPro.Refresh
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxExpr_Change
        dbgrd.SetFocus
    
    End If
    If KeyCode = vbKeyDown Then
        dbgrd.SetFocus
    End If
End Sub
