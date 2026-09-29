VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmListaDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   Icon            =   "frmListaDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3315
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3300
      Width           =   405
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   3240
      Width           =   480
   End
End
Attribute VB_Name = "frmListaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub msflxgrd_Click()
    soma
End Sub

Sub soma()
    Dim iCount As Integer
    Dim cTotal As Currency
    For iCount = 0 To msflxgrd.rows - 1
        msflxgrd.col = 1
        msflxgrd.Row = iCount
        cTotal = cTotal + CCur(msflxgrd.Text)
    Next
    lblTotal.caption = FormatVal(cTotal) '
End Sub

