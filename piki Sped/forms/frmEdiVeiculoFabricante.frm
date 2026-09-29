VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiVeiculoFabricante 
   Caption         =   "Fabricantes"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775
   Icon            =   "frmEdiVeiculoFabricante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPesquisar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   4200
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
      ScaleWidth      =   8745
      TabIndex        =   0
      Top             =   4590
      Width           =   8775
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmEdiVeiculoFabricante.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiVeiculoFabricante.frx":10994
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Transferência"
         Height          =   735
         Left            =   5160
         Picture         =   "frmEdiVeiculoFabricante.frx":10D1E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   6840
         Picture         =   "frmEdiVeiculoFabricante.frx":10E68
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndSelect 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmEdiVeiculoFabricante.frx":10FD3
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frmEdiVeiculoFabricante.frx":114D5
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6588
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
      Caption         =   "Registros Encontrados"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "descricao"
         Caption         =   "Descrição"
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
            ColumnWidth     =   4470,236
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDescricao 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
End
Attribute VB_Name = "frmEdiVeiculoFabricante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_Erro
    frmCadVeiculoFabricante.tag = ado!id
    frmCadVeiculoFabricante.txtCodigo.Text = ado!id
    frmCadVeiculoFabricante.txtDescricao.Text = ado!descricao
    frmCadVeiculoFabricante.Show 1
    ado.Requery
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    frmCadVeiculoFabricante.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub cmndSelect_Click()
    On Error Resume Next
    Me.tag = ado!id
    Me.Hide
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    busca
    If Me.tag = "" Then
        cmndSelect.Enabled = False
    Else
        cmndSelect.Enabled = True
    End If
    bActive = True
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    bActive = False
End Sub

Private Sub grid_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmndSelect_Click
    End If
End Sub

Private Sub grid_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo trata_Erro
    
    
    Dim s As String
    s = grid.Columns(ColIndex).DataField
          ado.Sort = s
        
 
    
    
trata_Erro:
    
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grid_DblClick
    End If
End Sub

Private Sub txtPesquisar_Change()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select * from veiculo_fabricante "
    strSQL = strSQL & " where 1 > 0 "
    strSQL = strSQL & " and (descricao like '%" & txtPesquisar.Text & "%'"
    strSQL = strSQL & " or id='" & txtPesquisar.Text & "')"
    
    abreObjetoAdo ado, strSQL, grid
    If Me.tag <> "" Then
        If ado.RecordCount = 1 Then
            cmndSelect_Click
            
        End If
    End If
    
    
End Sub

Private Sub txtPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        grid.SetFocus
    
    End If
    
    
End Sub
