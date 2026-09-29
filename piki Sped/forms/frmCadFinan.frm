VERSION 5.00
Begin VB.Form frmCadFinan 
   Caption         =   "Cadastro de Financeiras"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   7410
   Icon            =   "frmCadFinan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7410
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   7410
      TabIndex        =   15
      Top             =   4575
      Width           =   7410
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   5160
         Picture         =   "frmCadFinan.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   3120
         Picture         =   "frmCadFinan.frx":0554
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtSigla 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2520
         Width           =   4575
      End
      Begin VB.ComboBox cmbxOperacao 
         Height          =   315
         ItemData        =   "frmCadFinan.frx":069E
         Left            =   1920
         List            =   "frmCadFinan.frx":06A8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2160
         Width           =   4575
      End
      Begin VB.TextBox txtCNPJ 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1800
         Width           =   4575
      End
      Begin VB.ComboBox cmbxTBand 
         Height          =   315
         ItemData        =   "frmCadFinan.frx":06BD
         Left            =   1920
         List            =   "frmCadFinan.frx":06DF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txtbxNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtbxTel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   4
         Top             =   720
         Width           =   4575
      End
      Begin VB.ComboBox cmbxBco 
         Height          =   315
         ItemData        =   "frmCadFinan.frx":0775
         Left            =   1920
         List            =   "frmCadFinan.frx":0785
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sigla:"
         Height          =   195
         Left            =   1260
         TabIndex        =   13
         Top             =   2520
         Width           =   390
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação:"
         Height          =   195
         Left            =   930
         TabIndex        =   11
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   1260
         TabIndex        =   9
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bandeira:"
         Height          =   195
         Left            =   1035
         TabIndex        =   7
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome da Financeira:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fone/Fax:"
         Height          =   195
         Left            =   975
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   1350
         TabIndex        =   5
         Top             =   1080
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmCadFinan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxTBand_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxNome.Text = "" Then
        MsgBox "Informe o Nome da Financeira!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
    End If
    If cmbxBco.ListIndex = -1 Then
        MsgBox "Informe o Tipo da Financeira!", vbExclamation
        cmbxBco.SetFocus
        Exit Sub
    End If
    If Me.tag = "" Then
        P_Grava
    Else
        P_Altera
    
    End If
End Sub
Sub P_Grava()
    Dim strSQL As String
    Dim sCod As String
    sCod = MaxKeySQL("tab_financeiras", "cod_fin", "", "0#")
    strSQL = "insert into tab_financeiras(cod_fin,tel_fin,tipo_fin,nome_fin,sigla_fin,operacao_fin, tband, cnpj)Values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxTel.Text & "',"
    strSQL = strSQL & "'" & Left(cmbxBco.Text, 1) & "',"
    strSQL = strSQL & "'" & txtbxNome.Text & "',"
    strSQL = strSQL & "'" & txtSigla.Text & "',"
    strSQL = strSQL & "'" & Left(cmbxOperacao.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxTBand, 2) & "',"
    strSQL = strSQL & "'" & txtCNPJ.Text & "')"
    conn.Execute strSQL
    MsgBox "Financeira Gravada!", vbInformation
    
    txtbxNome.Text = ""
    txtbxTel.Text = ""
    cmbxBco.ListIndex = -1
    
    txtbxNome.SetFocus
End Sub
Sub P_Altera()
    strSQL = "update tab_financeiras set "
    
    strSQL = strSQL & " tel_Fin='" & txtbxTel.Text & "',"
    strSQL = strSQL & "tipo_fin='" & Left(cmbxBco.Text, 1) & "',"
    strSQL = strSQL & "nome_fin='" & txtbxNome.Text & "', "
    strSQL = strSQL & " tband='" & Left(cmbxTBand.Text, 2) & "',"
    strSQL = strSQL & " sigla_fin='" & txtSigla.Text & "',"
    strSQL = strSQL & " operacao_fin='" & Left(cmbxOperacao.Text, 1) & "',"
    strSQL = strSQL & " cnpj='" & txtCNPJ.Text & "' "
    strSQL = strSQL & " where cod_fin='" & Me.tag & "' "
    conn.Execute strSQL
    MsgBox "Financeira Alterada!", vbInformation
    
    Unload Me
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

Private Sub txtbxTel_GotFocus()
    P_SelecTexto txtbxTel
End Sub

Private Sub txtbxTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.Text = F_AjustaTelefone(txtbxTel.Text)
    
End Sub

Private Sub txtCNPJ_GotFocus()
    P_SelecTexto txtCNPJ
End Sub

Private Sub txtCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCNPJ_LostFocus()
    txtCNPJ.Text = FormataCNPJNfe(txtCNPJ)
End Sub
