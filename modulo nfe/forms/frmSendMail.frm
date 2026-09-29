VERSION 5.00
Begin VB.Form frmSendMail 
   Caption         =   "Enviar Email"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMSG 
      Height          =   1215
      Left            =   600
      TabIndex        =   1
      Text            =   "texto"
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   3000
      Width           =   1575
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
End Sub

Private Sub cmdEnviar_Click()
    Dim s() As String
    SendMailGmail "s10alisoncf@gmail.com", "Anapolis1", "alisoncf@gmail.com", "assunto", "msg", s
End Sub
