VERSION 5.00
Begin VB.Form frmTesteDraw 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "testa impressao"
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   480
      ScaleHeight     =   2625
      ScaleWidth      =   7425
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "frmTesteDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.Cls
    
    P_Quadrado Picture1, 100, 100, 2000, 6000, vbSolid, 1
End Sub

Private Sub Command2_Click()
    With frmVisualizadorImpressao.pcImpressao
        Load frmVisualizadorImpressao
        frmVisualizadorImpressao.pcImpressao.Print "teste"
        frmVisualizadorImpressao.addNewPage
        frmVisualizadorImpressao.pcImpressao.FontBold = True
        frmVisualizadorImpressao.pcImpressao.fontsize = 14
        frmVisualizadorImpressao.pcImpressao.Print "teste2"
        frmVisualizadorImpressao.addNewPage
        
        
        frmVisualizadorImpressao.Show 1
        
    
    End With
End Sub
