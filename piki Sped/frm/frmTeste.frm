VERSION 5.00
Begin VB.Form frmTeste 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "frmTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Single
'    Printer.PaperSize = 256
    Printer.Width = Printer.ScaleX(140, vbMillimeters, vbTwips)
    Printer.Height = Printer.ScaleX(210, vbMillimeters, vbTwips)
    For i = 0 To Printer.Height Step 100
        Printer.Print i
        
    Next
    Printer.EndDoc
    
    
End Sub

Private Sub Command2_Click()
    
    If Not F_OpenLpt Then
        Exit Sub
    End If
    Dim iCount As Integer
    Dim str As String
    Dim Margem As Integer
    For iCount = 1 To 80
        str = str & Left(iCount, 1)
    Next
    P_PrintLpt str, 0, False
    P_PrintLpt str, 0, False
    
    
    If Not F_CloseLpt Then
        Exit Sub
    End If
End Sub
