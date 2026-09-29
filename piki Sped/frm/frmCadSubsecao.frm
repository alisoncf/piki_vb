VERSION 5.00
Begin VB.Form frmCadsubsecao 
   Caption         =   "Inclusão de Seção"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6750
   Icon            =   "frmCadSubsecao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcSubSec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6000
      Picture         =   "frmCadSubsecao.frx":0ECA
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtSecao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6750
      TabIndex        =   7
      Top             =   2445
      Width           =   6750
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Salvar"
         Height          =   735
         Left            =   3360
         Picture         =   "frmCadSubsecao.frx":1275
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   4920
         Picture         =   "frmCadSubsecao.frx":16FC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seção"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmCadsubsecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Public campo1 As String
Public campo2 As String
Public campo3 As String
Public tabela As String
Public tabela2 As String

Private Sub cmdNovo_Click()
    If txtSecao.tag = "" Then
    
        MsgBox "escolha a seção!", vbExclamation
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
    strSQL = "insert into " & tabela & "(" & campo2 & ")values("
    
    strSQL = strSQL & "'" & campo2 & "')"
    conn.Execute strSQL
    MsgBox "Ok", vbInformation
    Unload Me
End Sub
Sub altera()
    Dim strSQL As String
    strSQL = "update " & tabela
    strSQL = strSQL & " set " & campo2 & "='" & txtDescricao & "' "
    strSQL = strSQL & " where " & campo1 & "='" & Me.tag & "'"
    conn.Execute strSQL
    MsgBox "Alterado", vbInformation
End Sub
Private Sub cmdSair_Click()
     Unload Me
End Sub
Public Sub Inicia(tabela As String, campo1 As String, campo2 As String, campo3 As String, captionForm As String)
    Me.campo1 = campo1
    Me.campo2 = campo2
    Me.campo3 = campo3
    Me.tabela = tabela
    Me.tabela2 = tabela2
    If captionForm <> "" Then
        Me.caption = captionForm
    Else
        Me.caption = "Gravar: (" & tabela & ") " & campo1 & "/" & campo2
    End If
End Sub


Private Sub cmdSelecionar_Click()
    Me.tag = ado!codigo
    Me.Hide
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    tabela = "secao"
    campo1 = "id"
    campo2 = "nome"
    campo3 = "id_secao"
    bActive = False
End Sub

Sub carrega()
    Dim strSQL As String
    
    strSQL = "select sec.nome as nome_secao, sub.* from secao as sec, subsecao as sub "
    strSQL = strSQL & " where sec.id=sub.id_sub "
    strSQL = strSQL & " and id=" & Me.tag
    abreObjetoAdo ado, strSQL
    
    
    With ado
        txtCodigo.Text = !id
        txtDescricao.Text = !nome
        txtSecao.Text = !nome_secao
        txtSecao.tag = !id_sec
        Me.caption = "Subseção: " & !id & "-" & !nome & "(Editar)"
    End With
    
End Sub

Private Sub pcSubSec_Click()
    buscaSecao
End Sub
Sub buscaSecao()
    Dim p As secao.secao
    p = SelecionaSecao(txtSecao.Text)
    If p.cod_sec <> "" Then
        txtSecao.Text = p.nome_sec
        txtSecao.tag = p.cod_sec
        cmdNovo.SetFocus
    End If
End Sub
Private Sub txtDescricao_GotFocus()
    ganhaFoco txtDescricao
End Sub

Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtDescricao_LostFocus()
    AjustaTexto txtDescricao
End Sub

Private Sub txtSecao_Change()
    txtSecao.tag = ""
End Sub

Private Sub txtSecao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaSecao
    End If
End Sub
