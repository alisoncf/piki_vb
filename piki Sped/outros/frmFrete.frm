VERSION 5.00
Begin VB.Form frmFrete 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Frete/Dados Adicionais"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnLimpar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   6690
      TabIndex        =   34
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmndbtnConfirmar 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   4770
      TabIndex        =   33
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Transportadora / Volumes Transportados"
      Height          =   2055
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.TextBox txtbxNomeFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtbxPlacaFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxUfVeiFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtbxCnpjFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtbxEnderecoFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtbxMuniFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   12
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtbxIEFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtbxUfFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6360
         TabIndex        =   14
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtbxQtdFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtbxEspFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtbxMarcaFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   22
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtbxNumFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   24
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtbxPBrutoFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   26
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtbxPLiqFre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Nome/Razão Social"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Placa Veíc:"
         Height          =   195
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF:"
         Height          =   195
         Left            =   6000
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Município"
         Height          =   195
         Left            =   3480
         TabIndex        =   11
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Insc. Est:"
         Height          =   195
         Left            =   6840
         TabIndex        =   15
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   6360
         TabIndex        =   13
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Espécie:"
         Height          =   195
         Left            =   1200
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Left            =   3000
         TabIndex        =   21
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   4440
         TabIndex        =   23
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "P.Bruto:"
         Height          =   195
         Left            =   5520
         TabIndex        =   25
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "P.Líq:"
         Height          =   195
         Left            =   6840
         TabIndex        =   27
         Top             =   1440
         Width           =   435
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frete por Conta:"
      Height          =   735
      Left            =   4770
      TabIndex        =   31
      Top             =   2280
      Width           =   3495
      Begin VB.TextBox txtbxContaFre 
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "1 - Emitente"
         Height          =   195
         Left            =   720
         TabIndex        =   36
         Top             =   285
         Width           =   840
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "2 - Destinatário"
         Height          =   195
         Left            =   1920
         TabIndex        =   35
         Top             =   285
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Adicionais:"
      Height          =   1455
      Left            =   90
      TabIndex        =   29
      Top             =   2280
      Width           =   4575
      Begin VB.TextBox txtbxDadosAdd 
         Height          =   1125
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmFrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnConfirmar_Click()
    Me.Tag = "x"
    Me.Hide
End Sub

Private Sub cmndbtnLimpar_Click()
    Dim iCount As Integer
    On Error Resume Next
    If MsgBox("Isto irá limpar todos os campos. Confirma?", vbYesNo + vbQuestion) = vbYes Then
        For iCount = 0 To Me.Controls.count - 1
            Me.Controls(iCount).Text = ""
            
        Next
    End If
    Unload Me
End Sub

Private Sub txtbxCnpjFre_GotFocus()
    P_SelecTexto txtbxCnpjFre
End Sub

Private Sub txtbxCnpjFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCnpjFre_LostFocus()
    P_AjustaTexto txtbxCnpjFre
End Sub

Private Sub txtbxEnderecoFre_GotFocus()
    P_SelecTexto txtbxEnderecoFre
End Sub

Private Sub txtbxEnderecoFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEnderecoFre_LostFocus()
    P_AjustaTexto txtbxEnderecoFre
End Sub

Private Sub txtbxEspFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIEFre_GotFocus()
    P_SelecTexto txtbxIEFre
End Sub

Private Sub txtbxIEFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIEFre_LostFocus()
    P_AjustaTexto txtbxIEFre
End Sub

Private Sub txtbxMarcaFre_Change()
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMuniFre_GotFocus()
    P_SelecTexto txtbxMuniFre
End Sub

Private Sub txtbxMuniFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMuniFre_LostFocus()
    P_AjustaTexto txtbxMuniFre
End Sub

Private Sub txtbxNomeFre_GotFocus()
    P_SelecTexto txtbxNomeFre
End Sub

Private Sub txtbxNomeFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNomeFre_LostFocus()
    P_AjustaTexto txtbxNomeFre
End Sub

Private Sub txtbxNumFre_Change()
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPBrutoFre_Change()
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPlacaFre_GotFocus()
    P_SelecTexto txtbxPlacaFre
End Sub

Private Sub txtbxPlacaFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPlacaFre_LostFocus()
    'P_AjustaPlaca txtbxPlacaFre
End Sub

Private Sub txtbxPLiqFre_Change()
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtdFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUfFre_GotFocus()
    P_SelecTexto txtbxUfFre
End Sub

Private Sub txtbxUfFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUfFre_LostFocus()
    P_AjustaTexto txtbxUfFre
End Sub

Private Sub txtbxUfVeiFre_GotFocus()
    P_SelecTexto txtbxUfFre
End Sub

Private Sub txtbxUfVeiFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUfVeiFre_LostFocus()
    P_AjustaTexto txtbxUfVeiFre
End Sub
