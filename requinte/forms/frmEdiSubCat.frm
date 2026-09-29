VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiSubSec 
   Caption         =   "Subseções"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9750
   Icon            =   "frmEdiSubCat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9750
   Begin VB.TextBox txtSecao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   4
      Top             =   5040
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "cod_sub"
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
         DataField       =   "nome_sub"
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
         DataField       =   "nome_sec"
         Caption         =   "Seção"
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
         DataField       =   "descmax_sub"
         Caption         =   "Desc. Máx (%)"
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
      BeginProperty Column04 
         DataField       =   "comiss_sub"
         Caption         =   "Comissão(%)"
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
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2924,788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2849,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065,26
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9750
      TabIndex        =   9
      Top             =   5475
      Width           =   9750
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "Selecionar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmEdiSubCat.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiSubCat.frx":6BB3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmEdiSubCat.frx":6F3D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   7920
         Picture         =   "frmEdiSubCat.frx":7087
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.TextBox txbxExpr 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   4935
   End
   Begin VB.CommandButton cmdSecao 
      Height          =   315
      Left            =   9240
      Picture         =   "frmEdiSubCat.frx":71F2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Seção:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   3
      Top             =   4800
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisar por:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1005
   End
End
Attribute VB_Name = "frmEdiSubSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado   As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdSecao_Click()
    Dim s As secao.secao
    s = F_BuscaSec(txtSecao.Text)
    If s.cod_sec <> "" Then
        txtSecao.Text = s.Nome_Sec
        txtSecao.tag = s.cod_sec
        busca
    End If
    
End Sub

Private Sub cmdSelecionar_Click()
    On Error Resume Next
    Me.tag = ado!cod_sub
    Me.Hide
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_Erro
    Dim fr As New frmCadSubSec
    With ado
        fr.tag = !cod_sub
        fr.txtbxNome.Text = !nome_sub
        fr.cmbxSecao.Text = !Nome_Sec
        fr.txtbxComiss.Text = !comiss_sub
        fr.txtbxDescmax.Text = !descmax_sub
        fr.Show 1
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadSubSec
    fr.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        Me.caption = "Selecione uma Subseção"
        cmdSelecionar.Visible = True
    Else
        cmdSelecionar.Visible = False
    End If
    busca
    
End Sub

Private Sub Form_Load()
        
        
    P_CentraForm Me
    'P_PreencheSecao cmbxSecao
    
    
    Set ado = New ADODB.Recordset
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select ts.nome_sec,tss.* from tab_secoes as ts inner join tab_subsec as tss on ts.cod_sec=tss.cod_sec "
    
    
    
    strSQL = strSQL & " where nome_sub like '" & txbxExpr.Text & "%'"
    If txtSecao.tag = "" Then
          strSQL = strSQL & " and  nome_sec like '" & txtSecao.Text & "%' "
    Else
          strSQL = strSQL & " and  ts.cod_sec= '" & txtSecao.tag & "'"
    End If
    
    abreObjetoAdo ado, strSQL, grid
    
    If Me.tag <> "" And ado.RecordCount = 1 Then
        Me.tag = ado!cod_sub
        Me.Hide
    End If
    
    
    
End Sub

Private Sub grid_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmdSelecionar_Click
    End If
End Sub

Private Sub grid_GotFocus()
    grid.RecordSelectors = True
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grid_DblClick
    End If
    
End Sub

Private Sub grid_LostFocus()
    grid.RecordSelectors = False
End Sub

Private Sub txbxExpr_GotFocus()
    P_SelecTexto txbxExpr
End Sub

Private Sub txbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        grid.SetFocus
    End If
End Sub

Private Sub txtSecao_Change()
    txtSecao.tag = ""
    busca
End Sub
