VERSION 5.00
Begin VB.Form frmCadFreteiros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Freteiros"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4170
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtbxCidade 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   8
         Top             =   2160
         Width           =   3015
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         ItemData        =   "frmCadFreteiros.frx":0000
         Left            =   3240
         List            =   "frmCadFreteiros.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2145
         Width           =   735
      End
      Begin VB.TextBox txtbxBairro 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox txtbxEndereco 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   4
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtbxTel 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox txtbxNome 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&UF:"
         Height          =   195
         Left            =   3240
         TabIndex        =   9
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Bairro"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fone/Fax:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Top             =   3240
      Width           =   4095
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   2160
         Picture         =   "frmCadFreteiros.frx":004D
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   120
         Picture         =   "frmCadFreteiros.frx":03D7
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmCadFreteiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxNome.Text = "" Then
        MsgBox "Informe o nome do freteiro!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    If Me.Tag = "" Then
        If F_BuscaUmCampo("cod_fre", "tab_freteiros", "nome_fre='" & txtbxNome.Text & "'") <> "" Then
            MsgBox "Já exite este freteiro!", vbExclamation
            txtbxNome.SetFocus
            Exit Sub
        End If
        grava
    Else
        altera
        
    End If
    
End Sub
Sub grava()
    Dim strSQL As String
    Dim sCod_Fre As String
    sCod_Fre = ChaveSQL("tab_freteiros", "cod_fre", "00#")
    strSQL = "insert into tab_freteiros(cod_fre,nome_fre,endereco_fre,"
    strSQL = strSQL & "bairro_fre,cidade_fre,uf_fre,fone_fre)Values("
    strSQL = strSQL & "'" & sCod_Fre & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "',"
    strSQL = strSQL & "'" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "'" & txtbxBairro.Text & "',"
    strSQL = strSQL & "'" & txtbxCidade.Text & "',"
    strSQL = strSQL & "'" & cmbxUF.Text & "',"
    strSQL = strSQL & "'" & txtbxTel.Text & "')"
    banco.Execute strSQL
    MsgBox "Freteiro Gravado!", vbInformation
    txtbxNome.Text = ""
    txtbxBairro.Text = ""
    txtbxTel.Text = ""
    txtbxNome.SetFocus
    
    If IsLoad("frmEdiFreteiros") Then
        frmEdiFreteiros.dtServ.Refresh
    End If
End Sub
Sub altera()
    Dim strSQL As String
    
    
    strSQL = "update tab_freteiros set "
    strSQL = strSQL & "nome_fre='" & txtbxNome.Text & "',"
    strSQL = strSQL & "endereco_fre='" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "bairro_fre='" & txtbxBairro.Text & "',"
    strSQL = strSQL & "cidade_fre='" & txtbxCidade.Text & "',"
    strSQL = strSQL & "uf_fre='" & cmbxUF.Text & "',"
    strSQL = strSQL & "fone_fre='" & txtbxTel.Text & "'  "
    strSQL = strSQL & " where cod_Fre='" & Me.Tag & "'"
    banco.Execute strSQL
    MsgBox "Freteiro Alterado!", vbInformation
    
    
    If IsLoad("frmEdiFreteiros") Then
        frmEdiFreteiros.dtServ.Refresh
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    P_PreencheUF cmbxUF
End Sub

Private Sub txtbxBairro_GotFocus()
    P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxBairro_LostFocus()
    P_AjustaTexto txtbxBairro
End Sub

Private Sub txtbxCidade_GotFocus()
    P_SelecTexto txtbxCidade
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCidade_LostFocus()
    P_AjustaTexto txtbxCidade
End Sub

Private Sub txtbxEndereco_GotFocus()
    P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxEndereco_LostFocus()
    P_AjustaTexto txtbxEndereco
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub

Private Sub txtbxTel_GotFocus()
    P_SelecTexto txtbxTel
End Sub

Private Sub txtbxTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub
