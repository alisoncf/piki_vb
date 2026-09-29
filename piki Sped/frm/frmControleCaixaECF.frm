VERSION 5.00
Begin VB.Form frmControleCaixaECF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle do Caixa do ECF"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Categoria"
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   2055
      Begin VB.OptionButton optnbtnSaida 
         Caption         =   "Saída"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         ToolTipText     =   "Indica que o valor é uma saída de dinheiro do caixa (SANGRIA)."
         Top             =   240
         Width           =   848
      End
      Begin VB.OptionButton optnbtnEntrada 
         Caption         =   "Entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Indica que o valor é uma entrada de dinheiro em caixa(SUPRIMENTO)."
         Top             =   240
         Value           =   -1  'True
         Width           =   848
      End
   End
   Begin VB.TextBox txtbxDescricao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   1320
      MaxLength       =   16
      TabIndex        =   3
      Text            =   "Caixa Inicial."
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Appearance      =   0  'Flat
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      Picture         =   "frmControleCaixaECF.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Confirma o Valor digitado."
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtbxValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "0,00"
      ToolTipText     =   "Informe o valor do caixa inicial."
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   2760
      Picture         =   "frmControleCaixaECF.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancela o valor digitado."
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblDescricao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Descrição:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   765
   End
   Begin VB.Label lblInforme 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Valor:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "frmControleCaixaECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmndbtnConfirma_Click()
   Dim strExpr As String

   P_Status "Informando dados à impressora fiscal"
   If F_ValToText(txtbxValor.Text) = 0 Then
      MsgBox "Deve-se informar um valor!", vbExclamation
      Exit Sub
   ElseIf F_Vazio(txtbxDescricao.Text) Then
      MsgBox "Deve-se informar uma descrição!", vbExclamation
      Exit Sub
   End If

   'Abrindo comunicação com a impressora
   If Not F_AbrePortaECF Then
      MsgBox "Não foi possível programar valor", vbExclamation
      P_Status
      Unload Me
   End If
   
   strExpr = conAbre_Protocolo_Fiscal
   strExpr = strExpr & "25|" & IIf(optnbtnEntrada.Value, "SU", "SA")
   strExpr = strExpr & "|" & Format(F_ValToText(txtbxValor.Text) * 100, "0#############")
   strExpr = strExpr & "|" & txtbxDescricao.Text
   strExpr = strExpr & conFecha_Protocolo_Fiscal
   
   'Passando comandos
   If Not F_ComandoECF(strExpr) Then
      P_Status
      FechaPorta
      Unload Me
   End If
   
   'Fechando porta de comunicação
   If Not F_FechaPortaECF Then
      P_Status
      FechaPorta
      Unload Me
   End If
   P_Status
   MsgBox "Valor confirmado!", vbInformation
   Unload Me
   
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strMens As String
    If KeyCode = vbKeyF1 Then
        strMens = "<ENTER> => Confirma os dados."
        ' P_MostraTeclaAtalho strMens
    End If
End Sub

Private Sub txtbxDescricao_GotFocus()
   'Seleciona o texto digitado
   P_SelecTexto txtbxDescricao
End Sub

Private Sub txtbxDescricao_LostFocus()
   'Ajusta o texto digitado
   P_AjustaTexto txtbxDescricao
End Sub

Private Sub txtbxValor_GotFocus()
   'Seleciona o valor digitado
   P_ValorGotFocus txtbxValor
End Sub

Private Sub txtbxValor_LostFocus()
   'Formata o valor digitado
   P_ValorLostFocus txtbxValor
End Sub
