VERSION 5.00
Begin VB.Form frmValoresEntr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valores do Produto Selecionado para Entrada"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frmValoresEntr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   8130
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtbxDesp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      TabIndex        =   31
      Text            =   "0,0000"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   29
      Text            =   "0,0000"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtbxLucro 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   27
      Text            =   "0,0000"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtbxNovoCusto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   25
      Text            =   "0,00"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   23
      Text            =   "0,00"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtbxSujerido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   21
      Text            =   "0,00"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtbxIpi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtbxTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "0,00"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtbxQtd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Text            =   "0,00"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtbxCusto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "0,00"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtbxPreco 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Text            =   "0,0000"
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnConfirma 
      Caption         =   "&Confirmar:"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Despesas"
      Height          =   195
      Left            =   5160
      TabIndex        =   32
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lucro Liq(%)"
      Height          =   195
      Left            =   4200
      TabIndex        =   30
      Top             =   1800
      Width           =   870
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lucro(%)"
      Height          =   195
      Left            =   3240
      TabIndex        =   28
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custo Atal:"
      Height          =   195
      Left            =   1080
      TabIndex        =   26
      Top             =   1800
      Width           =   765
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7920
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço. &Futuro:"
      Height          =   195
      Left            =   4440
      TabIndex        =   24
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço. Suj:"
      Height          =   195
      Left            =   3360
      TabIndex        =   22
      Top             =   960
      Width           =   780
   End
   Begin VB.Label lblIPI_Rest 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ipi Rest: 0,00"
      Height          =   195
      Left            =   6240
      TabIndex        =   20
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label lblRestante 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr Rest: 0,00"
      Height          =   195
      Left            =   6240
      TabIndex        =   19
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lblMar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6000
      TabIndex        =   18
      Tag             =   "!"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      Height          =   195
      Left            =   5400
      TabIndex        =   17
      Tag             =   "!"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Categoria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   16
      Tag             =   "!"
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cat:"
      Height          =   195
      Left            =   2640
      TabIndex        =   15
      Tag             =   "!"
      Top             =   120
      Width           =   285
   End
   Begin VB.Label lblNome_Pro 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Tag             =   "0"
      Top             =   360
      Width           =   1485
   End
   Begin VB.Label lblCod_Pro 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Tag             =   "0"
      Top             =   360
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custo Ant.:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço &Unit:"
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Qtd:"
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   300
   End
   Begin VB.Label lblIPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Ipi:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmValoresEntr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnConfirma_Click()
    If Round(CCur(txtbxQtd.Text) * CCur(txtbxPreco.Text), 2) > Round(CCur(txtbxPreco.Tag), 2) Then
        MsgBox "O Valor está maior que o permitido!!", vbExclamation
        
        txtbxTotal.SetFocus
        Exit Sub
    End If
    Me.Tag = "X"
    Me.Hide
End Sub

Private Sub Form_Activate()
'    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Tag = ""
        Me.Hide
    End If
    
End Sub

Private Sub txtbxIpi_GotFocus()
    GanhaFoco txtbxIpi
End Sub

Private Sub txtbxIpi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIpi_LostFocus()
    PerdeFoco txtbxIpi
    If CCur(txtbxIpi.Text) > CCur(txtbxIpi.Tag) Then
        MsgBox "Não é possível este valor para IPI!", vbExclamation
        txtbxIpi.Text = txtbxIpi.Tag
        txtbxIpi.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtbxPreco_GotFocus()
    GanhaFoco txtbxPreco
End Sub

Private Sub txtbxPreco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPreco_LostFocus()
    PerdeFoco txtbxPreco, "###,###,##0.0000"
    
End Sub

Private Sub txtbxQtd_GotFocus()
    GanhaFoco txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    Dim cEstoque As Currency
    PerdeFoco txtbxQtd
    On Error Resume Next
    txtbxPreco.Text = FormatVal(CCur(txtbxTotal.Text) / CCur(txtbxQtd.Text), "###,###,##0.0000")
End Sub

Private Sub txtbxTotal_GotFocus()
    GanhaFoco txtbxTotal
End Sub

Private Sub txtbxTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxTotal_LostFocus()
    PerdeFoco txtbxTotal
    On Error Resume Next
    txtbxPreco.Text = FormatVal(CCur(txtbxTotal.Text) / CCur(txtbxQtd.Text), "###,###,##0.0000")
End Sub
