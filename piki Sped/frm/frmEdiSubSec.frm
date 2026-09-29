VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiSubSec 
   Caption         =   "Edição de Subseções"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9750
   Icon            =   "frmEdiSubSec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9750
   Begin VB.TextBox txtSecao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8281
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
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
      TabIndex        =   7
      Top             =   5865
      Width           =   9750
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmEdiSubSec.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiSubSec.frx":6D54
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmEdiSubSec.frx":70DE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   7920
         Picture         =   "frmEdiSubSec.frx":7228
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.TextBox txbxExpr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   5040
      Width           =   6015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expressão:"
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   4800
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seção:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   510
   End
End
Attribute VB_Name = "frmEdiSubSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado   As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdSelect_Click()
    Me.Hide
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_Erro
    Dim fr As New frmCadSubSec
    With ado
        fr.tag = !cod_sub
        fr.txtbxNOme.Text = !nome_sub
        fr.txtSecao.Text = !nome_sec
        fr.txtSecao.tag = !cod_sec
        fr.txtbxComiss.Text = !comiss_sub
        fr.txtbxDescmax.Text = !descmax_sub
        fr.Show 1
    End With
    Exit Sub
trata_Erro:
    MostraErro
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
    busca
    If Me.tag <> "" Then
        cmdSelect.Visible = True
    Else
        cmdSelect.Visible = False
    End If
End Sub

Private Sub Form_Load()
        
        
    P_CentraForm Me
    'P_PreencheSecao  cmbxSecao
    
    
    Set ado = New ADODB.Recordset
    
End Sub
Sub busca()
    Dim sSec As String
    'ajustatexto txbxExpr
    sSec = "nome_sub like '" & txbxExpr.Text & "%'"
    Dim strSQL As String
    If txtSecao.tag <> "" Then
          sSec = sSec & " and  nome_sec ='" & cmbxSecao.Text & "' "
    End If
    FechaObjeto ado
     
    strSQL = "select tsu.*,ts.nome_sec from " & TAB_SUBSEC & " as tsu, tab_secoes as ts "
    strSQL = strSQL & " where ts.cod_sec=tsu.cod_sec and nome_sub like '" & txbxExpr.Text & "%'"
    
    If txtSecao.tag <> "" Then
        strSQL = strSQL & " and ts.cod_sec='" & txtSecao.Text & "' "
    End If
    
    
    
    abreObjetoAdo ado, strSQL, grid
    
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.tag <> "" Then
            cmdSelect_Click
        Else
        cmndbtnEdit_Click
        End If
    End If
End Sub

Private Sub txbxExpr_GotFocus()
        ganhaFoco txbxExpr
End Sub

Private Sub txbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtSecao_Change()
    busca
End Sub
