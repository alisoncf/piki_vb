VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFechaDia 
   Caption         =   "Fechamento"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5318
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   9225
      TabIndex        =   0
      Top             =   0
      Width           =   9225
      Begin VB.CommandButton cmdGerar 
         Caption         =   "Gerar"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dia:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmFechaDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGerar_Click()
    Dim strSQL As String
    
    
    
    
End Sub

Private Sub Form_Load()
    ajustaGrid
End Sub
Sub ajustaGrid()
    With grid
        .Cols = 7
    End With
End Sub
Private Sub txtData_GotFocus()
    ganhaFoco txtData
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
    If kecyode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtData_LostFocus()
    P_AjustaData txtData
    
End Sub
