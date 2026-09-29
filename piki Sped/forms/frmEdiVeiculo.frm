VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiVeiculo 
   Caption         =   "Veículos"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11130
   Icon            =   "frmEdiVeiculo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFabricante 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   10
      Top             =   5280
      Width           =   3615
   End
   Begin VB.TextBox txtPesquisar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   5280
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
      ScaleWidth      =   11100
      TabIndex        =   1
      Top             =   5655
      Width           =   11130
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmEdiVeiculo.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiVeiculo.frx":10994
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Transferência"
         Height          =   735
         Left            =   5160
         Picture         =   "frmEdiVeiculo.frx":10D1E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   9360
         Picture         =   "frmEdiVeiculo.frx":10E68
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndSelect 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmEdiVeiculo.frx":10FD3
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frmEdiVeiculo.frx":114D5
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8493
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
      ColumnCount     =   3
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
      BeginProperty Column02 
         DataField       =   "fabricante"
         Caption         =   "Fabricante"
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
            ColumnWidth     =   5174,929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505,26
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fabricante"
      Height          =   195
      Left            =   5160
      TabIndex        =   9
      Top             =   5040
      Width           =   750
   End
   Begin VB.Label lblDescricao 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   690
   End
End
Attribute VB_Name = "frmEdiVeiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_Erro
    frmCadVeiculo.tag = ado!id
    frmCadVeiculo.txtCodigo.Text = ado!id
    frmCadVeiculo.txtDescricao.Text = ado!descricao
    frmCadVeiculo.txtFabricante.Text = ado!fabricante
    frmCadVeiculo.txtFabricante.tag = ado!id_fab
    frmCadVeiculo.caption = "Veículo: " & ado!id & " : " & ado!descricao
    frmCadVeiculo.Show 1
    ado.Requery
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    frmCadVeiculo.Show 1
    ado.Requery
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

Private Sub txtFabricante_Change()
    busca
End Sub

Private Sub txtFabricante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        grid.SetFocus
    End If
End Sub

Private Sub txtPesquisar_Change()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select v.*,f.descricao as fabricante from veiculo as v inner join veiculo_fabricante as f on v.id_fab=f.id"
    strSQL = strSQL & " where 1 > 0 "
    strSQL = strSQL & " and (v.descricao like '%" & txtPesquisar.Text & "%'"
    strSQL = strSQL & " or v.id='" & txtPesquisar.Text & "')"
    If txtFabricante.Text <> "" Then
        strSQL = strSQL & " and (f.descricao like '%" & txtFabricante.Text & "%'  "
        strSQL = strSQL & " or f.id='" & txtFabricante.Text & "')"
    End If
    strSQL = strSQL & " order by v.descricao "
    abreObjetoAdo ado, strSQL, grid
    
    If Me.tag <> "" Then
        If ado.RecordCount = 1 Then
            cmndSelect_Click
            
        Else
            
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
