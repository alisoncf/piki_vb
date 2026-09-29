VERSION 5.00
Begin VB.Form frmDadosChe 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dados do Cheque"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5085
   Icon            =   "frmDadosChe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C6DADD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4335
      ScaleWidth      =   4815
      TabIndex        =   18
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtbxCPF 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   1
         Top             =   120
         Width           =   3495
      End
      Begin VB.TextBox txtbxAgencia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtbxNome_Che 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   3
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtbxNum_Che 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtbxVencimento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   13
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnConfirmar 
         Caption         =   "Con&firmar"
         Height          =   615
         Left            =   1140
         Picture         =   "frmDadosChe.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3240
         Width           =   1455
      End
      Begin VB.ComboBox cmbxBanco 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1140
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtbxContaCorrente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmbxSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   3180
         Picture         =   "frmDadosChe.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtbxValor_Che 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   14
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblJam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         Height          =   195
         Left            =   3600
         TabIndex        =   22
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J.A.M%:"
         Height          =   195
         Left            =   2835
         TabIndex        =   21
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         Height          =   195
         Left            =   3600
         TabIndex        =   20
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Va&lor:"
         Height          =   195
         Left            =   3000
         TabIndex        =   19
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label lblCPF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CPF:"
         Height          =   195
         Left            =   660
         TabIndex        =   0
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Agência:"
         Height          =   195
         Left            =   375
         TabIndex        =   6
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cheque Nº:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Titular:"
         Height          =   195
         Left            =   525
         TabIndex        =   2
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Banco:"
         Height          =   195
         Left            =   495
         TabIndex        =   4
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&onta C.:"
         Height          =   195
         Left            =   345
         TabIndex        =   8
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Va&lor:"
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   2760
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmDadosChe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSair_Click()
Me.Hide
End Sub

Private Sub cmndbtnConfirmar_Click()
    If txtbxNum_Che.Text = "" Then
        MsgBox "informe o número do Cheque", vbExclamation
        txtbxNum_Che.SetFocus
        Exit Sub
    End If
    If txtbxValor_Che.Text = "0,00" Then
        MsgBox "Informe um valor válido!", vbExclamation
        txtbxValor_Che.SetFocus
        Exit Sub
    End If
    If CDate(txtbxVencimento.Text) < Date Then
 '       MsgBox "A data não pode ser retroativa!", vbExclamation
        
'        Exit Sub
    End If
    Me.Tag = txtbxNum_Che.Text
    Me.Hide
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    P_PreencheBanco cmbxBanco
    MudaAparencia Me
End Sub

Private Sub Picture1_GotFocus()
    txtbxNum_Che.SetFocus
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtbxAgencia_GotFocus()
    P_SelecTexto txtbxAgencia
End Sub

Private Sub txtbxAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAgencia_LostFocus()
    P_AjustaTexto txtbxAgencia
End Sub


Private Sub txtbxCNPJ_LostFocus()
    
End Sub

Private Sub txtbxContaCorrente_GotFocus()
    P_SelecTexto txtbxContaCorrente
End Sub

Private Sub txtbxContaCorrente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxContaCorrente_LostFocus()
    P_AjustaTexto txtbxContaCorrente
End Sub

Private Sub txtbxCpf_GotFocus()
    P_SelecTexto txtbxCPF
End Sub

Private Sub txtbxCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxCPF_LostFocus
        
    End If
End Sub
Sub buscaCH()
    Dim c As cheque
    If txtbxCPF.Text = "" Then Exit Sub
    c = F_HistCheCPF(txtbxCPF.Text)
    
    If c.cod_bco <> "" Then
        txtbxNome_Che.Text = c.nome_che
        txtbxAgencia.Text = c.age_che
        txtbxContaCorrente.Text = c.cco_che
        P_LocalizaOperacao c.cod_bco, cmbxBanco
        txtbxNum_Che.SetFocus
    Else
        SendKeys Chr(9)
    End If
    
    
    
End Sub
Private Sub txtbxCPF_LostFocus()
    
    
    
    txtbxCPF.Text = F_SoNumero(txtbxCPF.Text)
    If Len(txtbxCPF.Text) = 0 Then
        
    ElseIf Len(txtbxCPF.Text) >= 14 Then
        lblCPF.caption = "&CNPJ:"
        P_FormataCGC txtbxCPF
        If Not F_CGC(txtbxCPF.Text) Then
            MsgBox "CNPJ Inválido!", vbExclamation
            txtbxCPF.SetFocus
            Exit Sub
        End If
    Else
        lblCPF.caption = "&CPF:"
        P_FormataCPF txtbxCPF
        If Not F_CPF(txtbxCPF.Text) Then
            MsgBox "CPF Inválido!", vbExclamation
            txtbxCPF.SetFocus
            Exit Sub
        End If
    End If
    buscaCH
End Sub

Private Sub txtbxNome_Che_GotFocus()
    P_SelecTexto txtbxNome_Che
End Sub

Private Sub txtbxNome_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Che_LostFocus()
    P_AjustaTexto txtbxNome_Che
End Sub

Private Sub txtbxNum_Che_GotFocus()
    P_SelecTexto txtbxNum_Che
End Sub

Private Sub txtbxNum_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNum_Che_LostFocus()
    P_AjustaTexto txtbxNum_Che
End Sub

Private Sub txtbxValor_Che_GotFocus()
    P_SelecTexto txtbxValor_Che
End Sub

Private Sub txtbxValor_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Che_LostFocus()
    P_ValorLostFocus txtbxValor_Che
End Sub

Private Sub txtbxVencimento_GotFocus()
    P_SelecTexto txtbxVencimento
End Sub

Private Sub txtbxVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxVencimento_LostFocus()
    P_AjustaData txtbxVencimento
End Sub
