VERSION 5.00
Begin VB.Form frmTotalizadorNF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Totalizador Não Fiscal"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nomear Novo Totalizador Não Fiscal:"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3855
      Begin VB.CommandButton cmndbtnNomeia 
         Caption         =   "Nomear Totalizador Não Fiscal"
         Height          =   615
         Left            =   960
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtbxDescricao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtbxIndice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Índice:"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.TextBox txtbxInfo 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   8535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totalizadores:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "frmTotalizadorNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnNomeia_Click()
    Dim i As Integer
    
    i = Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms(CInt(txtbxIndice.Text), txtbxDescricao.Text)
    If i = 0 Then
        MsgBox "Erro de comunicação!", vbExclamation
    ElseIf i = 1 Then
        MsgBox "Ok!", vbInformation
        cmndbtnRefresh_Click
    ElseIf i = -2 Then
        MsgBox "Parâmetro inválido na função.", vbExclamation
    ElseIf i = -4 Then
        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows.", vbExclamation
        
    ElseIf i = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação. ", vbExclamation

    ElseIf i = -27 Then
        MsgBox "Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2). ", vbExclamation

        
    End If
    
End Sub

Private Sub cmndbtnRefresh_Click()
    
    Dim i As Integer
    Dim sTot As String
    sTot = Space(179)
    i = Bematech_FI_VerificaTotalizadoresNaoFiscais(sTot)
    If i = 0 Then
        MsgBox "Erro de comunicação!", vbExclamation
    ElseIf i = 1 Then
     '   MsgBox "Ok!", vbInformation
    ElseIf i = -2 Then
        MsgBox "Parâmetro inválido na função.", vbExclamation
    ElseIf i = -4 Then
        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows.", vbExclamation
        
    ElseIf i = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação. ", vbExclamation

    ElseIf i = -27 Then
        MsgBox "Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2). ", vbExclamation

        
    End If
    txtbxInfo.Text = sTot
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
