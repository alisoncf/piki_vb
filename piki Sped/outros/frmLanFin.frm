VERSION 5.00
Begin VB.Form frmLanFin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lançamento Financeiro"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5145
   Icon            =   "frmLanFin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   5115
      TabIndex        =   14
      Top             =   0
      Width           =   5145
      Begin VB.TextBox txtbxData_Cai 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optnCerd 
         Caption         =   "Crédito"
         Height          =   615
         Left            =   120
         Picture         =   "frmLanFin.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optnDeb 
         Caption         =   "Debito"
         Height          =   615
         Left            =   1560
         Picture         =   "frmLanFin.frx":0481
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data:"
         Height          =   195
         Left            =   3720
         TabIndex        =   2
         Top             =   120
         Width           =   390
      End
   End
   Begin VB.ComboBox cmbxCentro 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3480
      Width           =   4815
   End
   Begin VB.ComboBox cmbxConta 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2760
      Width           =   4815
   End
   Begin VB.TextBox txtbxValor_Cai 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "0,00"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtbxHist 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   1920
      Picture         =   "frmLanFin.frx":0A33
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   3480
      Picture         =   "frmLanFin.frx":0EDB
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C&entro de Custo:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Conta Corrente:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Histórico:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   660
   End
End
Attribute VB_Name = "frmLanFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxCentro_Click()
    cmbxCentro.Tag = F_BuscaUmCampo("cod_pla", "tab_planocontas", "nome_pla='" & _
        cmbxCentro.Text & "' and tipo_pla='" & Left(lblTipo.caption, 1) & "'")
        
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.Tag = "" Then
        If txtbxValor_Cai.Text = "0,00" Then
            MsgBox "Informe um Valor válido!", vbExclamation
            txtbxValor_Cai.SetFocus
            Exit Sub
        End If
        If F_Vazio(txtbxHist.Text) Then
            MsgBox "Informe um histórico!", vbExclamation
            txtbxHist.SetFocus
            Exit Sub
        End If
        
        P_Grava
        
    Else
        P_Altera
    End If
    
End Sub
Sub P_Grava()
    Dim strSQL As String
    strSQL = F_GravaFinanceiro(Left(lblTipo.caption, 1), txtbxData_Cai.Text, _
                txtbxValor_Cai.Text, txtbxHist.Text, "DIN", _
                cmbxConta.Text, "", "", "", cmbxCentro.Tag)
    If strSQL = "" Then
        MsgBox "Erro!", vbCritical
        Exit Sub
    End If
    'P_Class CCur(txtbxValor_Cai.Text), Left(lblTipo.caption, 1), strSQL
    MsgBox "Lançamento Gravado!", vbInformation
    txtbxData_Cai.Text = Date
    txtbxHist.Text = ""
    txtbxValor_Cai.Text = "0,00"
    txtbxHist.SetFocus
End Sub

Sub P_Altera()
    Dim strSQL As String
    strSQL = "Update tab_FINANCEIRO set historico_FIN='" & txtbxHist.Text & "',"
    strSQL = strSQL & "valor_FIN=" & F_GravaDec(txtbxValor_Cai.Text)
    strSQL = strSQL & ", COD_cON='" & Left(cmbxConta.Text, 3) & "' "
    strSQL = strSQL & ", COD_pla='" & cmbxCentro.Tag & "' "
    
    strSQL = strSQL & " where seq_FIN='" & Me.Tag & "' "
    Banco.Execute strSQL
    
    MsgBox "Dados Alterados!", vbInformation
     'P_Class CCur(txtbxValor_Cai.Text), Left(lblTipo.caption, 1), Me.Tag
    Unload Me
    
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    
    P_CentraForm Me
    P_PreencheContas cmbxConta
    P_PreenchePlanoContas cmbxCentro, True
    cmbxConta.ListIndex = -1
    txtbxData_Cai.Text = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad("frmEdiCai") Then
        RefreshDataPosition frmEdiCai.dtServ
    End If
    Unload Me
End Sub

Private Sub optnCerd_Click()
    lblTipo.caption = "Crédito"
    P_PreenchePlanoContas cmbxCentro, True
End Sub

Private Sub optnCerd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnDeb_Click()
            lblTipo.caption = "Débito"
            P_PreenchePlanoContas cmbxCentro, False
End Sub

Private Sub optnDeb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxData_Cai_GotFocus()
    P_SelecTexto txtbxData_Cai
End Sub

Private Sub txtbxData_Cai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxData_Cai_LostFocus()
    P_AjustaData txtbxData_Cai
End Sub

Private Sub txtbxHist_GotFocus()
    P_SelecTexto txtbxHist
End Sub

Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxHist_LostFocus()
    P_AjustaTexto txtbxHist
End Sub

Private Sub txtbxValor_Cai_GotFocus()
    P_SelecTexto txtbxValor_Cai
End Sub

Private Sub txtbxValor_Cai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Cai_LostFocus()
    P_ValorLostFocus txtbxValor_Cai
End Sub
