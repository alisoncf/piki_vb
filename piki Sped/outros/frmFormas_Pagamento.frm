VERSION 5.00
Begin VB.Form frmFormas_Pagamento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formas de Pagamento"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6105
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Não efetua forma de pagamentos, nem venda de nota fiscal."
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtbxDescricao 
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Informe alguma descrição complementar."
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtbxValor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   3240
      MaxLength       =   80
      TabIndex        =   3
      Text            =   "0,00"
      ToolTipText     =   "Informe um valor."
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cmbxFormaPagamento 
      BackColor       =   &H80000014&
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Selecione a forma de pagamento"
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Confirma a forma de pagamento."
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblSimbol 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   990
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblSimbol 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   870
      Index           =   1
      Left            =   255
      TabIndex        =   10
      Top             =   525
      Width           =   705
   End
   Begin VB.Label lblDescricao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Descrição:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label lblRestante 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0,00"
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "Indica o valor restante a pagar."
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblValorRestante 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V&alor Restante:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   1125
   End
   Begin VB.Label lblValor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Forma de Pagamento:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1560
   End
   Begin VB.Label lblSimbol 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   870
      Index           =   0
      Left            =   270
      TabIndex        =   9
      Top             =   555
      Width           =   705
   End
End
Attribute VB_Name = "frmFormas_Pagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Este formulário é utilizado no momento do fechamento do
'cupom fiscal.
'   O cupom fiscal tem que estar aberto


Private Sub cmbxFormaPagamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnConfirma_Click()
    'Confirma a forma lançada
    Dim intRet As Integer
    
    If F_TextToVal(txtbxValor.Text) > _
    F_TextToVal(lblRestante.caption) And _
    cmbxFormaPagamento.ListIndex Then
       'O valor digitado é maior que o valor restante!!!
         
        MsgBox "O valor está acima do limite!", vbExclamation
        txtbxValor.SetFocus
        Exit Sub
    End If
    If cmbxFormaPagamento.ListIndex < 0 Then
        'Não foi selecionada nenhuma forma de pagamento
        MsgBox "Você deve selecionar uma forma de pagamento!", vbExclamation
        cmbxFormaPagamento.SetFocus
        Exit Sub
    End If
    If F_TextToVal(txtbxValor.Text) = 0 Then
       MsgBox "Você tem que informar um valor válido!", vbExclamation
       txtbxValor.SetFocus
       Exit Sub
    End If
    
    lblRestante.caption = F_ValToText( _
    F_TextToVal(lblRestante.caption) - F_TextToVal(txtbxValor.Text))
    
    'intRet = F_FormaPagamento(Format(cmbxFormaPagamento.ListIndex + 1, _
    "0#"), F_TextToVal(txtbxValor.Text), Trim(txtbxDescricao.Text))
    
'    If Not intRet Then
'       MsgBox "Ocoreu algum erro com a impressora fiscal", vbCritical
'    End If
    
    
    'Preenchendo TAG conforme descrições passadas!
    Me.Tag = Me.Tag & Format(cmbxFormaPagamento.ListIndex + 1, "0#") & _
    "|" & Format(F_TextToVal(txtbxValor.Text) * 100, "0#############") _
    & IIf(F_Vazio(txtbxDescricao.Text), "", "|" & txtbxDescricao.Text) _
    & Chr(13)
       
    txtbxValor.Text = lblRestante.caption
    If F_TextToVal(lblRestante.caption) <= 0 Then
       'Terminou o lançamento
       Me.Hide 'Esconde o formulário passando os dados
    End If
End Sub

Private Sub cmndbtnFechar_Click()
    Me.Tag = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    'Verifica o valor do cupom fiscal
    'lblRestante.Caption = Format(F_ECFSubTotal, "###,##0.00")
    
    If F_TextToVal(lblRestante.caption) = 0 Then
       MsgBox "Ocorreu algum erro no cupom fiscal!", vbCritical
    End If
    P_PreencheFormaPagamento cmbxFormaPagamento
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Dim strMens As String
    If KeyCode = vbKeyF1 Then
        'P_MostraTeclaAtalho "Alt + C => Confirma a forma de pagamento."
    ElseIf KeyCode = vbKeyC And Shift = vbAltMask Then
'        If F_ControleAtivo(cmndbtnCOnfirma) Then
'            cmndbtnConfirma_Click
'        End If
    End If
End Sub

Private Sub txtbxDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_GotFocus()
    'Seleciona o valor
    P_ValorGotFocus txtbxValor
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    'Ajusta o valor
    P_ValorLostFocus txtbxValor
End Sub
