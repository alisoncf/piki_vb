VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCadCai 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Caixa"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcIp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2625
      ScaleWidth      =   6345
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Data dt 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnAdd 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   4320
         TabIndex        =   12
         Top             =   195
         Width           =   1215
      End
      Begin VB.TextBox txtbxIP 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmCadCai.frx":0000
         Height          =   1815
         Left            =   720
         OleObjectBlob   =   "frmCadCai.frx":0011
         TabIndex        =   11
         Top             =   645
         Width           =   4815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP:"
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   5040
      Picture         =   "frmCadCai.frx":086E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Salvar"
      Height          =   615
      Left            =   3480
      Picture         =   "frmCadCai.frx":09D9
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox txtbxConta 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtbxCod 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conta Vinc:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "frmCadCai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnAdd_Click()
    Dim strSQL As String
    If Me.tag = "" Then
        Exit Sub
    End If
    If txtbxIP.Text = "" Then
        MsgBox "Informe o IP", vbExclamation
        txtbxIP.SetFocus
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("cod_cai", "tab_caixasip", "ip_cai='" & txtbxIP.Text & "'")
    If strSQL <> "" Then
        MsgBox "Este ip já está conectado ao caixa " & strSQL, vbExclamation
        txtbxIP.SetFocus
        Exit Sub
    End If
    
    strSQL = "insert into tab_caixasip(cod_cai,ip_cai)values("
    strSQL = strSQL & "'" & txtbxCod.Text & "',"
    strSQL = strSQL & "'" & txtbxIP.Text & "')"
    Banco.Execute strSQL
    
    
    buscaIps
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If strFechames_usu <> "S" Then
        MsgBox "Não Tem Permissão para Este acesso!", vbExclamation
        Exit Sub
    End If
    If Not F_senha Then
        Exit Sub
    End If
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub grava()
    Dim strSQL As String
    strSQL = "insert into tab_caixas(cod_cai,nome_cai,cod_con)values("
    strSQL = strSQL & "'" & txtbxCod.Text & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "',"
    
    strSQL = strSQL & "'" & txtbxConta.Text & "')"
    Banco.Execute strSQL
    
    
    
    MsgBox "Dados Gravados!", vbInformation
    
    Me.tag = txtbxCod.Text
    txtbxCod.Enabled = False
    pcIp.Visible = True
End Sub
Sub altera()
    Dim strSQL As String
    strSQL = "update tab_caixas "
    strSQL = strSQL & "set cod_cai='" & txtbxCod.Text & "',"
    strSQL = strSQL & "nome_cai='" & txtbxNome.Text & "',"
    strSQL = strSQL & "ip_cai='" & txtbxIP.Text & "',"
    strSQL = strSQL & "cod_con='" & txtbxConta.Text & "' "
    strSQL = strSQL & " where cod_Cai='" & Me.tag & "' "
    Banco.Execute strSQL
    MsgBox "Dados Alterados!", vbInformation
End Sub
Sub buscaIps()
    Dim strSQL As String
    
    strSQL = "select * from tab_caixasip where cod_cai='" & txtbxCod.Text & "'"
    dt.RecordSource = strSQL
    dt.Refresh
    
End Sub

Private Sub Form_Activate()
    If Me.tag = "" Then
        pcIp.Visible = False
    Else
        pcIp.Visible = True
        buscaIps
    End If
End Sub

Private Sub Form_Load()
    
    dt.Connect = Banco.Connect
End Sub

