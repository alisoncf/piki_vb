VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiMar 
   Caption         =   "Edição de Marcas:"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   8460
   Icon            =   "frmEdiMar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8460
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   8400
      TabIndex        =   1
      Top             =   5580
      Width           =   8460
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1680
         Picture         =   "frmEdiMar.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   0
         Picture         =   "frmEdiMar.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Transferência"
         Height          =   735
         Left            =   5040
         Picture         =   "frmEdiMar.frx":0D9E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   6840
         Picture         =   "frmEdiMar.frx":0EE8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndSelect 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3360
         Picture         =   "frmEdiMar.frx":1053
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   8175
   End
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frmEdiMar.frx":1555
      Height          =   4695
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
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
      Caption         =   "Marcas Encontradas"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cod_mar"
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
         DataField       =   "nome_mar"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4394,835
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   765
   End
End
Attribute VB_Name = "frmEdiMar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadMar
    If ado.RecordCount > 0 Then
        fr.tag = ado!cod_mar
        fr.txtbxNome.Text = ado!nome_mar
        fr.Show 1
    Else
        MsgBox "Não há o que ser Editado!", vbExclamation
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadMar
    fr.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub dtgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dtgrd_GotFocus()
    grid.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dtgrd_LostFocus()
    grid.MarqueeStyle = 0
End Sub

Private Sub cmndSelect_Click()
    If Me.tag <> "" Then
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

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grid_DblClick
    End If
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        busca
    End If
    bActive = True
End Sub

Private Sub Form_Load()
    Dim strSQL As String

    
    Set ado = New ADODB.Recordset
    bActive = False
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select * from " & TAB_MARCAS
    strSQL = strSQL & " where (nome_mar like '%" & txtNome.Text & "%') "
    strSQL = strSQL & " order by nome_mar "
    abreObjetoAdo ado, strSQL, grid
    If Me.tag <> "" And ado.RecordCount = 1 Then
        cmndSelect_Click
    End If
End Sub
Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        grid.SetFocus
    End If
    If KeyCode = vbKeyReturn Then
    
        busca
    End If
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        grid.SetFocus
    End If
End Sub
