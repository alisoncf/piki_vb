VERSION 5.00
Begin VB.Form frmSaidaDiaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saída Diária"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   525
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton cmnxdbtnConf 
         Caption         =   "&Confirmar"
         Height          =   525
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtbxData 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtbxNome_Pro 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Saída:"
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome do Produto:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSaidaDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmnxdbtnConf_Click()
    If Me.Tag = "" Then
        
        grava
    Else
    
    End If
End Sub
Sub grava()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "insert into tab_ajsai(cod_pro,data_sai,qtd_sai)values("
    strSQL = strSQL & "'" & txtbxNome_Pro.Tag & "',"
    strSQL = strSQL & F_GravaData(txtbxData.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxQtd.Text) & ") "
    banco.Execute strSQL
    txtbxNome_Pro.Text = ""
    txtbxQtd.Text = "0,00"
    MsgBox "Produto lançado!", vbInformation
    txtbxNome_Pro.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
End Sub

Private Sub txtbxNome_Pro_DblClick()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxNome_Pro.Text)
    txtbxNome_Pro.Text = p.nome_pro
    txtbxNome_Pro.Tag = p.cod_pro
    If txtbxNome_Pro.Tag <> "" Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_Pro_DblClick
    End If
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd
End Sub
