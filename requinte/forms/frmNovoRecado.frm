VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNovoRecado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Postar um Recado"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1920
      ScaleHeight     =   615
      ScaleWidth      =   4935
      TabIndex        =   8
      Top             =   320
      Width           =   4935
      Begin VB.CheckBox chkTodas 
         Caption         =   "Todas"
         Height          =   195
         Left            =   3960
         TabIndex        =   17
         Top             =   0
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnTerceiro 
         Height          =   315
         Left            =   4560
         Picture         =   "frmNovoRecado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxTerceiro 
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "Buscar"
      Height          =   615
      Left            =   4920
      Picture         =   "frmNovoRecado.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnMarcar 
      Caption         =   "Marcar"
      Height          =   615
      Left            =   2520
      Picture         =   "frmNovoRecado.frx":0467
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnExcl 
      Caption         =   "E&xcluir"
      Height          =   615
      Left            =   1320
      Picture         =   "frmNovoRecado.frx":05CB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnNovo 
      Caption         =   "&Novo"
      Height          =   615
      Left            =   120
      Picture         =   "frmNovoRecado.frx":0955
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   6240
      Picture         =   "frmNovoRecado.frx":0AB9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnEnviar 
      Caption         =   "&Enviar"
      Height          =   615
      Left            =   3720
      Picture         =   "frmNovoRecado.frx":0C24
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtbxMsg 
      Height          =   2175
      Left            =   120
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   7215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1720
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Novo Recado"
      TabPicture(0)   =   "frmNovoRecado.frx":1058
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmbxTipo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Verificar Recados"
      TabPicture(1)   =   "frmNovoRecado.frx":1074
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optnSeu"
      Tab(1).Control(1)=   "optnRep"
      Tab(1).ControlCount=   2
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmNovoRecado.frx":1090
         Left            =   120
         List            =   "frmNovoRecado.frx":109A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   540
         Width           =   1575
      End
      Begin VB.OptionButton optnSeu 
         Caption         =   "Seu"
         Height          =   195
         Left            =   -74760
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optnRep 
         Caption         =   "Represenante"
         Height          =   195
         Left            =   -74760
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   780
   End
End
Attribute VB_Name = "frmNovoRecado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Me.tag = "x"
    If optnRep.Value = True And txtbxTerceiro.tag = "" Then
        MsgBox "Escolha o terceiro", vbExclamation
        txtbxTerceiro.SetFocus
        Exit Sub
    End If
    
    If optnRep.Value Then
        strSQL = F_SelectMensagem("R", txtbxTerceiro.tag, chkTodas.Value = 1)
        
    Else
        strSQL = F_SelectMensagem("U", strCod_Usu, chkTodas.Value = 1)
        
    End If
    dt.RecordSource = strSQL
    dt.Refresh
    
        
End Sub

Private Sub cmndbtnEnviar_Click()
    If txtbxTerceiro.tag = "" Then
        MsgBox "Informe o Destino!", vbExclamation
        txtbxTerceiro.SetFocus
        Exit Sub
    End If
    If txtbxMsg.Text = "" Then
        MsgBox "Informe a mensagem!", vbExclamation
        txtbxMsg.SetFocus
    End If
    mensagem.PostaMensagem Left(cmbxTipo.Text, 1), txtbxTerceiro.tag, txtbxMsg.Text
    MsgBox "Mensagem Enviada!", vbInformation
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If MsgBox("Confirma exclusão do recado " & dt.Recordset!seq_rec & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "delete from tab_recados where seq_rec='" & dt.Recordset!seq_rec & "'"
    Banco.Execute strSQL
    
    
    dt.Refresh
    Exit Sub
trata_Erro:
End Sub

Private Sub cmndbtnMarcar_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    With dt.Recordset
        If !status_rec = "1" Then
            strSQL = "update tab_recados set status_rec='2' where seq_rec='" & dt.Recordset!seq_rec & "'"
            Banco.Execute strSQL
        Else
            strSQL = "update tab_recados set status_rec='1' where seq_rec='" & dt.Recordset!seq_rec & "'"
            Banco.Execute strSQL
        End If
    End With
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmndbtnNovo_Click()
    Me.tag = ""
    
    mudaControles
End Sub
Sub mudaControles()
    If Me.tag = "" Then
        
        cmndbtnExcl.Value = False
        cmndbtnMarcar.Visible = False
        
        dt.Visible = False
        txtbxMsg.Text = ""
        txtbxMsg.SetFocus
    Else
        txtbxMsg.Text = ""
        cmndbtnExcl.Value = False
        cmndbtnMarcar.Visible = False
        dt.Visible = True
    End If
    
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTerceiro_Click()
    If cmbxTipo.ListIndex = 0 Then
        buscaRep
    Else
        buscaFun
    End If
End Sub

Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.Text)
        If r.cod_Rep <> "" Then
            txtbxTerceiro.Text = r.nome_rep
            txtbxTerceiro.tag = r.cod_Rep
            txtbxMsg.SetFocus
        End If
    
End Sub

Sub buscaFun()
    Dim r As funcionario.funcionario
        r = funcionario.F_SelectFuncionario(txtbxTerceiro.Text)
        If r.cod_fun <> "" Then
            txtbxTerceiro.Text = r.nome_fun
            txtbxTerceiro.tag = r.cod_fun
            txtbxMsg.SetFocus
        End If
    
End Sub

Private Sub dt_Reposition()
    On Error GoTo trata_Erro
    With dt
    dt.caption = "Mensagem " & dt.Recordset.AbsolutePosition + 1 & "/" & dt.Recordset.RecordCount
        
        txtbxMsg.Text = dt.Recordset!mensagem_rec & Chr(13) + Chr(10) & "De: " & dt.Recordset!de & " em " & .Recordset!data_rec
        If .Recordset!status_rec = "1" Then
            txtbxMsg.FontBold = True
        Else
            txtbxMsg.FontBold = False
        End If
        
        cmndbtnMarcar_Click
        
    End With
    
    Exit Sub
trata_Erro:
    dt.caption = ""
    txtbxMsg.Text = ""
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    SSTab1.Tab = 1
    If Me.tag = "R" Then
        optnRep.Value = True
        cmndbtnBuscar_Click
    End If
    If Me.tag = "U" Then
        optnSeu.Value = True
        cmndbtnBuscar_Click
    End If
    
    
    bActive = True
End Sub

Private Sub Form_Load()
    cmbxTipo.ListIndex = 0
    dt.Connect = Banco.Connect
End Sub

Private Sub optnSeu_Click()
    txtbxTerceiro.Text = strNome_Usu
    txtbxTerceiro.Enabled = False
    cmndbtnTerceiro.Enabled = False
    txtbxTerceiro.tag = ""
End Sub

Private Sub optnRep_Click()
    txtbxTerceiro.Text = ""
    txtbxTerceiro.Enabled = True
    cmndbtnTerceiro.Enabled = True
    txtbxTerceiro.tag = ""
    txtbxTerceiro.SetFocus
End Sub


Private Sub SSTab1_DblClick()
    If SSTab1.Tab = 0 Then
        Me.tag = ""
        mudaControles
        
    Else
        Me.tag = "x"
        mudaControles
        
    End If
End Sub

Private Sub txtbxMsg_LostFocus()
    P_AjustaTexto txtbxMsg, 2
End Sub
