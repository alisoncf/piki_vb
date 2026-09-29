VERSION 5.00
Begin VB.Form frmCadSetor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inclusão de Setores"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7230
   Icon            =   "frmCadSetor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmCadSetor.frx":0442
         Left            =   1320
         List            =   "frmCadSetor.frx":0458
         TabIndex        =   1
         Text            =   "Vila"
         Top             =   120
         Width           =   5415
      End
      Begin VB.TextBox txtbxCobrador 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   2280
         Width           =   5175
      End
      Begin VB.CommandButton cmndbtnCobrador 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6480
         Picture         =   "frmCadSetor.frx":0490
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox txtbxCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   960
         Width           =   5175
      End
      Begin VB.TextBox txtbxNome 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   5415
      End
      Begin VB.CommandButton cmndbtnBuscaCid 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6480
         Picture         =   "frmCadSetor.frx":081A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Tipo:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Cobrador:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label lblUF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&UF:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Nome do Setor"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C&idade:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   5520
      Picture         =   "frmCadSetor.frx":0BA4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Height          =   735
      Left            =   3840
      Picture         =   "frmCadSetor.frx":0C99
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "frmCadSetor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxCidade_GotFocus()
    If cmbxUf.Text = "" Then
        MsgBox "Escolha primeiro o estado!", vbExclamation
        cmbxUf.SetFocus
        Exit Sub
    End If
End Sub

Private Sub cmbxCidade_LostFocus()
    lblCod_Cid.caption = F_BuscaUmCampo("cod_Cid", "tab_cidades", _
            "nome_cid='" & cmbxCidade.Text & "' and uf_cid='" & _
            cmbxUf.Text & "'")
End Sub

Private Sub cmbxUf_LostFocus()
    
End Sub

Private Sub cmbxTipo_GotFocus()
    P_SelecTexto cmbxTipo
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_LostFocus()
    cmbxTipo.Text = F_AjustaTexto(cmbxTipo.Text)
End Sub

Private Sub cmndbtnBuscaCid_Click()
    buscaCidade
End Sub
Sub buscaCidade()
    Dim c As Cidades.Cidade
    c = Cidades.F_BuscaCidades(txtbxCidade.Text)
    
    If c.cod_cid <> "" Then
        txtbxCidade.Text = c.nome_cid
        txtbxCidade.tag = c.cod_cid
        lblUF.caption = c.uf_cid
    End If
    
    
End Sub

Private Sub cmndbtnCobrador_Click()
    buscaCobrador
End Sub
Sub buscaCobrador()
   Dim c As ContaCob
    c = F_BuscaContaCob(txtbxCobrador.Text)
    If c.cod_cob <> "" Then
        
        txtbxCobrador.Text = c.nome_cob
        txtbxCobrador.tag = c.cod_cob
    End If
    
    
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    
    
    
    If Me.tag = "" Then
        If F_BuscaUmCampo("Nome_set", "tab_setores", _
        "nome_set='" & txtbxNome.Text & "'" & _
        "Cod_cid='" & txtbxCidade.tag & "'") <> "" Then
            MsgBox "Este setor já está cadastrado!", vbExclamation
            Exit Sub
        End If
        If txtbxNome.Text = "" Then
            MsgBox "Informe o nome da área!", vbExclamation
            Exit Sub
        End If
        If txtbxCidade.Text = "" Then
            MsgBox "Informe a cidade corretamente!", vbExclamation
            txtbxCidade.SetFocus
            Exit Sub
        End If
        
        grava
    Else
        altera
        
    End If
End Sub
Sub grava()
    Me.tag = Setores.F_gravaSetor(cmbxTipo.Text, txtbxNome.Text, txtbxCidade.tag, txtbxCobrador.tag)
    If Me.tag <> "" Then
        MsgBox "Setor cadastrado!", vbInformation
        
    End If
End Sub
Sub altera()
    Dim strSQL As String
    If Setores.F_EditaSetor(Me.tag, cmbxTipo.Text, txtbxNome.Text, txtbxCidade.tag, txtbxCobrador.tag) Then
        MsgBox "Setor alterado!", vbInformation
    Else
        MsgBox "Erro"
    End If
End Sub





Private Sub txtbxCidade_GotFocus()
    P_SelecTexto txtbxCidade
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaCid_Click
    End If
End Sub

Private Sub txtbxCidade_LostFocus()
    P_AjustaTexto txtbxCidade
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub
