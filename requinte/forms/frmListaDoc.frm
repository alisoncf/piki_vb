VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmListaDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "frmListaDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   3900
      TabIndex        =   5
      Top             =   3975
      Width           =   3960
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
         Left            =   600
         TabIndex        =   7
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   60
         Width           =   405
      End
   End
   Begin VB.CommandButton cmndbtnBuscaCCo 
      Caption         =   "..."
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtbxConta 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   3735
   End
   Begin VB.TextBox txtbxGrid 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conta Creditada:"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   1185
   End
End
Attribute VB_Name = "frmListaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscaCCo_Click()
    Dim c As ContaCorrente
    c = F_BuscaContaCorrente(txtbxConta.Text)
    If c.cod_con <> "" Then
        txtbxConta.Text = c.nome_con
        txtbxConta.tag = c.cod_con
    End If
End Sub

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

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If msflxgrd.col = 1 Then
            txtbxGrid.Text = msflxgrd.Text
            txtbxGrid.Left = msflxgrd.CellLeft
            txtbxGrid.Top = msflxgrd.CellTop
            txtbxGrid.Visible = True
            txtbxGrid.SetFocus
        End If
    End If
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        msflxgrd.RemoveItem msflxgrd.Row
        soma
    End If
End Sub

Private Sub txtbxGrid_GotFocus()
    P_SelecTexto txtbxGrid
End Sub

Private Sub txtbxGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_ValorLostFocus txtbxGrid
        msflxgrd.col = 2
        If CCur(txtbxGrid.Text) > CCur(msflxgrd.Text) Then
            MsgBox "Valor Errado!", vbExclamation
            Exit Sub
        End If
        msflxgrd.col = 1
        msflxgrd.Text = txtbxGrid.Text
        soma
        SendKeys Chr(9)
        
    End If
End Sub

Private Sub txtbxGrid_LostFocus()
    
    txtbxGrid.Visible = False
End Sub
