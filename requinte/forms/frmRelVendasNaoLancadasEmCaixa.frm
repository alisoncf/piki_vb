VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelVendasNaoLancadasEmCaixa 
   Caption         =   "Vendas Efetivadas que Não Cairam no Caixa"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   9900
      TabIndex        =   11
      Top             =   7455
      Width           =   9960
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   2295
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   9930
      TabIndex        =   0
      Top             =   0
      Width           =   9960
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   4680
         Picture         =   "frmRelVendasNaoLancadasEmCaixa.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   6360
         Picture         =   "frmRelVendasNaoLancadasEmCaixa.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   8040
         Picture         =   "frmRelVendasNaoLancadasEmCaixa.frx":0797
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.ComboBox cmbxCai 
         Height          =   315
         ItemData        =   "frmRelVendasNaoLancadasEmCaixa.frx":088C
         Left            =   2520
         List            =   "frmRelVendasNaoLancadasEmCaixa.frx":0899
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   120
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmRelVendasNaoLancadasEmCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnBuscar_Click()
    If cmbxCai.text = "" Then
        MsgBox "Informe o caixa!", vbExclamation
        cmbxCai.SetFocus
        Exit Sub
    End If
    busca
End Sub
Sub busca()
    
    Dim strSQL As String
    Dim sIn As String
    sIn = F_NotIn("distinct num_ven", "tab_caixa", "data_cai>=" & F_GravaData(txtDe.text) & " And data_cai<=" & F_GravaData(txtAte.text) & " and cod_cai='" & Left(cmbxCai.text, 2) & "'")
    
    strSQL = "select tv.*,tvf.* from tab_vendas as  tv left join tab_venfec as tvf "
    strSQL = strSQL & " on tv.num_ven=tv.num_ven "
    strSQL = strSQL & " where tv.num_ven not in " & sIn
    
    If cmbxCai.text <> "" Then
        strSQL = strSQL & " and tv.cod_cai='" & Left(cmbxCai.text, 2) & "' "
    End If
    strSQL = strSQL & " and data_ven>=" & F_GravaData(txtDe.text)
    strSQL = strSQL & " and data_ven<=" & F_GravaData(txtAte.text)
    strSQL = strSQL & " and tipo_operacao=1"
    strSQL = strSQL & " order by data_ven,tv.cod_cai, tv.num_ven desc "
    
    abreObjetoAdo ado, strSQL, grid
    
    
    
    
    
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    P_PreencheCaixas cmbxCai
    txtAte.text = Date
    txtDe.text = Date
    Set ado = New ADODB.Recordset
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
End Sub


