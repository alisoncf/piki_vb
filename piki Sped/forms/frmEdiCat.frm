VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCat 
   Caption         =   "Edição de Categorias"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
   Icon            =   "frmEdiCat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   11220
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frmEdiCat.frx":014A
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8281
      _Version        =   393216
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "cod_cat"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome_cat"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "qtdmin_cat"
         Caption         =   "Qtd. Atacado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "descmax_cat"
         Caption         =   "Desc.%"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4394,835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260,284
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   4935
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   11190
      TabIndex        =   1
      Top             =   5535
      Width           =   11220
      Begin VB.CommandButton cmndSelect 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmEdiCat.frx":015C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   9480
         Picture         =   "frmEdiCat.frx":04BD
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Transferência"
         Height          =   735
         Left            =   5160
         Picture         =   "frmEdiCat.frx":0628
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiCat.frx":0772
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmEdiCat.frx":0AFC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar por:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Menu mnopt 
      Caption         =   "&Opções"
      Visible         =   0   'False
      Begin VB.Menu mnEditar 
         Caption         =   "&Editar"
      End
      Begin VB.Menu mnSelecionar 
         Caption         =   "&Selecionar"
      End
      Begin VB.Menu mnVerProd 
         Caption         =   "Ver produtos"
      End
   End
End
Attribute VB_Name = "frmEdiCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadCat
    On Error Resume Next
    If ado.RecordCount > 0 Then
        fr.tag = ado!cod_cat
        fr.txtbxNome.Text = ado!nome_cat
        fr.txtDescMax.Text = FormatVal(ado!descmax_Cat)
        fr.txtQtdAcacado.Text = ado!qtdmin_cat
        fr.Show 1
        RefreshAdoPosition ado
    Else
        MsgBox "Não há o que ser Editado!", vbExclamation
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadCat
        
    frmCadCat.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub cmndSelect_Click()
    On Error Resume Next
    If Me.tag <> "" Then
        Me.tag = ado!cod_sub
        
        Me.Hide
    End If
    
End Sub

Private Sub grid_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmndSelect_Click
    End If
End Sub

Private Sub Form_Activate()
    busca
    If Me.tag <> "" Then
        cmndSelect.Enabled = True
    Else
        cmndSelect.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    Set ado = New ADODB.Recordset
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select * from tab_categorias "
    strSQL = strSQL & " where nome_cat like '%" & txtNome.Text & "%' "

    
    
    strSQL = strSQL & " order by nome_cat "
    abreObjetoAdo ado, strSQL, grid
    
    
    If Me.tag <> "" And ado.RecordCount = 1 Then
    End If
    
    
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grid_DblClick
    End If
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpt
    End If
End Sub

Private Sub mnSelecionar_Click()
    cmndSelect_Click
End Sub

Private Sub mnVerProd_Click()
    verProp
    
End Sub
Sub verProp()
    frmEdiPro.txtCat.Text = ado!nome_cat
    frmEdiPro.txtCat.tag = ado!cod_cat
    frmEdiPro.tag = "x"
    frmEdiPro.Show 1
End Sub
Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        grid.SetFocus
    End If
End Sub
