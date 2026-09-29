VERSION 5.00
Begin VB.Form frmMensagem 
   Caption         =   "Mensagem"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtMsg 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   5895
   End
   Begin VB.TextBox txtPara 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Para:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviar_Click()
    If txtMsg.Text = "" Then
        MsgBox "Mensagem em branco.", vbExclamation
        Exit Sub
    End If
    If txtPara.tag = "" Then
        MsgBox "Informe destinatário", vbExclamation
        txtPara.Text = ""
        Exit Sub
    End If
    
    GravaMensagem txtPara.tag, txtMsg.Text
    MsgBox "Mensagem enviada", vbInformation
    txtPara.Text = ""
    
    txtPara.SetFocus
    
    
    
    
    
End Sub

Private Sub txtPara_Change()
    
    txtPara.tag = ""
End Sub

Private Sub txtPara_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscausu
    End If
    
End Sub
Sub buscausu()
    Dim f As funcionario.funcionario
    
    f = funcionario.F_SelectFuncionario(txtPara.Text)
    
    
    txtPara.Text = f.guerra_fun
    txtPara.tag = f.cod_fun
    
    txtMsg.SetFocus
    
End Sub
