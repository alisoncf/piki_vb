VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiSubSec 
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9750
   Icon            =   "frmEdiSubSec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9750
   Begin MSDataGridLib.DataGrid grid 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7011
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
      TabIndex        =   11
      Top             =   5865
      Width           =   9750
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiSubSec.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmEdiSubSec.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   7920
         Picture         =   "frmEdiSubSec.frx":6D26
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.TextBox txbxExpr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   5040
      Width           =   6015
   End
   Begin VB.ComboBox cmbxSecao 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5040
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9855
      TabIndex        =   8
      Top             =   0
      Width           =   9855
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         Picture         =   "frmEdiSubSec.frx":6E91
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   9
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblIns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edição de Subseções"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   10
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expressão:"
      Height          =   195
      Left            =   3600
      TabIndex        =   3
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
Dim WithEvents ado   As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_Erro
    Dim fr As New frmCadSubSec
    With ado
        fr.tag = !cod_sub
        fr.txtbxNome.Text = !nome_sub
        fr.cmbxSecao.Text = !Nome_Sec
        fr.txtbxComiss.Text = !comiss_sub
        fr.txtbxDescmax.Text = !descmax_sub
        fr.Show 0, mdiFrm
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadSubSec
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
        
        
    P_CentraForm Me
    P_PreencheSecao cmbxSecao
    
    
    Set ado = New ADODB.Recordset
    
End Sub
Sub busca()
    Dim sSec As String
    P_AjustaTexto txbxExpr
    sSec = "nome_sub like '" & txbxExpr.Text & "%'"
    If cmbxSecao.Text <> "" Then
          sSec = sSec & " and  nome_sec ='" & cmbxSecao.Text & "' "
    End If
    FechaObjeto ado
    ado.Open subsecAO.F_SelectSubSec(sSec), Conn_Client
    
    Set grid.DataSource = ado
    
End Sub

Private Sub txbxExpr_GotFocus()
    P_SelecTexto txbxExpr
End Sub

Private Sub txbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
