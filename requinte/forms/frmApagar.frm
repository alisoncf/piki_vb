VERSION 5.00
Begin VB.Form frmApagar 
   Caption         =   "Apagar"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1305
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas (!)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmApagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select num_ven from tab_vendas "
    
    
End Sub
