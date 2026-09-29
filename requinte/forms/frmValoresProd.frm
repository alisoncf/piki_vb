VERSION 5.00
Begin VB.Form frmValoresProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valores do Produto Selecionado"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   ControlBox      =   0   'False
   Icon            =   "frmValoresProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   7995
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtbxIPI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar:"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   750
      Width           =   1815
   End
   Begin VB.TextBox txtbxPreco 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Tag             =   "!"
      Text            =   "0,00"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtbxCusto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "0,00"
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtbxQtd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3720
      TabIndex        =   7
      Tag             =   "!"
      Text            =   "0,00"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblIPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Ipi:"
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Qtd:"
      Height          =   195
      Left            =   3720
      TabIndex        =   6
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço &Unit:"
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   795
   End
   Begin VB.Label lblCustoCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custo:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   1320
      TabIndex        =   12
      Top             =   0
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCod_Pro 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9999999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Tag             =   "!"
      Top             =   240
      Width           =   840
   End
   Begin VB.Label lblNome_Pro 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   1320
      TabIndex        =   9
      Tag             =   "!"
      Top             =   240
      Width           =   1665
   End
End
Attribute VB_Name = "frmValoresProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnConfirma_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub txtbxIPI_GotFocus()
    P_SelecTexto txtbxIPI
End Sub

Private Sub txtbxIPI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIPI_LostFocus()
    P_ValorLostFocus txtbxIPI
End Sub

Private Sub txtbxPreco_GotFocus()
    P_SelecTexto txtbxPreco
End Sub

Private Sub txtbxPreco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPreco_LostFocus()
    P_ValorLostFocus txtbxPreco
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
    Dim cEstoque As Currency
    P_ValorLostFocus txtbxQtd
    Dim sEst As String
    cEstoque = CCur(F_BuscaUmCampo("Estoque_pro-Reservado_pro", "tab_Produtos", "cod_pro='" & lblCod_Pro.caption & "'"))
    sEst = (F_BuscaUmCampo("qtd_ite", "tab_itvendas", "cod_pro='" & lblCod_Pro.caption & "' and num_ven='" & Me.Tag & "'"))
    If sEst <> "" Then
        cEstoque = cEstoque + CCur(sEst)
    End If
    If CCur(txtbxQtd.Text) > cEstoque Then
        If MsgBox("Este valor excede o estoque disponível!", vbExclamation + vbOKCancel) = vbCancel Then
            Unload Me
        Else
            txtbxQtd.Text = FormatVal(CStr(cEstoque))
            txtbxQtd.SetFocus
            Exit Sub
        End If
        
    End If
End Sub
