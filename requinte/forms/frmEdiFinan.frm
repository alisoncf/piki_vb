VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiFinan 
   Caption         =   "Edição de Financeiras"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   8625
   Icon            =   "frmEdiFinan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtExpr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   4935
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8625
      TabIndex        =   3
      Top             =   3960
      Width           =   8625
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Selecionar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1800
         Picture         =   "frmEdiFinan.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiFinan.frx":0219
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   3480
         Picture         =   "frmEdiFinan.frx":05A3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   6840
         Picture         =   "frmEdiFinan.frx":06ED
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frmEdiFinan.frx":0858
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "cod_Fin"
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
         DataField       =   "nome_fin"
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
         DataField       =   "cnpj"
         Caption         =   "CNPJ"
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
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4215,118
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   735
   End
End
Attribute VB_Name = "frmEdiFinan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdSelect_Click()
    On Error Resume Next
    Me.tag = ado!cod_fin
    Me.Hide
End Sub

Private Sub cmndbtnEdit_Click()
    With ado
        If .RecordCount <= 0 Then
            MsgBox "Não há o que ser editado!", vbExclamation
            Exit Sub
        End If
        Dim fr As New frmCadFinan
        fr.tag = !cod_fin
        fr.txtbxNome.Text = !nome_fin
        fr.txtbxTel.Text = !tel_fin
        If Not IsNull(!CNPJ) Then
            fr.txtCNPJ.Text = !CNPJ
        End If
        P_LocalizaOperacao !tipo_fin, fr.cmbxBco
        P_LocalizaOperacao !tBand, fr.cmbxTBand
        
        fr.Show 1
            
        
    End With
End Sub

Private Sub cmndbtnNew_Click()
    frmCadFinan.Show 1
    
    
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        busca
        cmdSelect.Enabled = True
    Else
        cmdSelect.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Set ado = New ADODB.Recordset
    
End Sub
Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select * from tab_financeiras "
    If txtExpr.Text <> "" Then
        strSQL = strSQL & " where cod_fin='" & txtExpr.Text & "' "
        strSQL = strSQL & " or nome_fin like '" & txtExpr.Text & "%' "
        strSQL = strSQL & " or cnpj like '" & txtExpr.Text & "%' "
    End If
    
    
    
    strSQL = strSQL & " order by nome_fin "
    abreObjetoAdo ado, strSQL, grid
    
    If Me.tag <> "" Then
        If ado.RecordCount = 1 Then
            Exit Sub
        End If
    End If
    
    
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub grid_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmdSelect_Click
    
    End If
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grid_DblClick
    End If
End Sub

Private Sub txtExpr_Change()
    busca
End Sub

Private Sub txtExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        grid.SetFocus
    End If
End Sub
