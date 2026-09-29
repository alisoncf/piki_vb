VERSION 5.00
Begin VB.Form frmODBC 
   Caption         =   "Escolha um Driver ODBC"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lst 
      Height          =   4545
      ItemData        =   "frmODBC.frx":0000
      Left            =   120
      List            =   "frmODBC.frx":0002
      TabIndex        =   1
      Top             =   240
      Width           =   8775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Avançar >>"
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "frmODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
        Dim c As New Collection
   Set c = ListODBCDrivers
   lst.Clear
    Dim i As Integer
    For i = 1 To c.Count
        lst.AddItem c(i)
    Next
End Sub

