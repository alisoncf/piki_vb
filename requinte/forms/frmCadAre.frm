VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCadAre 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Áreas"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7860
   Icon            =   "frmCadAre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtbxBairro 
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Top             =   5280
      Width           =   2655
   End
   Begin VB.ComboBox cmbxUF 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5280
      Width           =   615
   End
   Begin VB.ComboBox cmbxCidade 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton cmndbtnNovo 
      Caption         =   "&Novo"
      Height          =   615
      Left            =   120
      Picture         =   "frmCadAre.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSelect 
      Caption         =   "&Editar"
      Height          =   615
      Left            =   1560
      Picture         =   "frmCadAre.frx":02A6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "Sai&r"
      Height          =   615
      Left            =   5280
      Picture         =   "frmCadAre.frx":0375
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmCadAre.frx":046A
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "frmCadAre.frx":047B
      TabIndex        =   6
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro:"
      Height          =   195
      Left            =   3960
      TabIndex        =   9
      Top             =   5040
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade:"
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   5040
      Width           =   540
   End
End
Attribute VB_Name = "frmCadAre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dtgrd_Click()

End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    'dt.Connect = banco.Connect
    Adodc.ConnectionString = sNomeBanco
    strSQL = "select ta.*,tc.nome_cid,tc.uf_cid from tab_cidades as tc,"
    strSQL = strSQL & " tab_areas as ta where tc.cod_cid=ta.cod_cid order by tc.nome_cid "
    Adodc.RecordSource = strSQL
    Adodc.Refresh
    P_CentraForm Me
End Sub

Private Sub tlbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next '
    If Button.Index = 1 Then
        With Adodc.Recordset
            frmIncArea.tag = Adodc.Recordset!cod_are
            frmIncArea.cmbxUf = !uf_cid
            frmIncArea.caption = !cod_are & "-" & !nome_are
            frmIncArea.txtbxNome.Text = !nome_are
            P_PreencheCidades frmIncArea.cmbxCidade, !uf_cid
            frmIncArea.cmbxCidade.Text = !nome_cid
            frmIncArea.cmbxCidade.Enabled = False
            frmIncArea.cmbxUf.Enabled = False
            frmIncArea.Show 1
            
        End With
    ElseIf Button.Index = 2 Then
        On Error Resume Next
        frmIncArea.Show 1
    ElseIf Button.Index = 3 Then
    ElseIf Button.Index = 6 Then
        Unload Me
    End If
    
End Sub
