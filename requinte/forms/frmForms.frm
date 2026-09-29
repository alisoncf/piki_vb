VERSION 5.00
Begin VB.Form frmForms 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Janelas Ativas"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2790
   ClipControls    =   0   'False
   Icon            =   "frmForms.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   1680
   End
   Begin VB.ListBox lstForms 
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    MudaAparencia Me
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Timer1_Timer
    Me.Width = lstForms.Width + 20
    Me.Height = lstForms.Height + 40
    Me.Left = 7500
    Me.Top = 1500
    If lstForms.ListCount = 0 Then Unload Me
    
End Sub

Private Sub lstForms_Click()
    Dim i As Integer
    On Error Resume Next
    
    For i = 1 To Forms.count - 1
        If Forms(i).caption = lstForms.Text Then
            Forms(i).Show 0
            Forms(i).SetFocus
        End If
    Next i
    
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    lstForms.Clear
        For i = 1 To Forms.count - 1
        If Forms(i).caption <> Me.caption Then
            lstForms.AddItem Forms(i).caption
        End If
    Next i
    If lstForms.ListCount = 0 Then
        Me.Hide
    End If
End Sub
