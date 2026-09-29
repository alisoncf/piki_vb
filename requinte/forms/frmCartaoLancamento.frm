VERSION 5.00
Begin VB.Form frmCartaoLancamento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lançamentos de Cartões"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   4320
      TabIndex        =   14
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   2400
      TabIndex        =   13
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtAut 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   4560
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtCoo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtParcelas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   240
      TabIndex        =   4
      Text            =   "0,00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtBandeira 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmndbtnOrgao 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3360
      Picture         =   "frmCartaoLancamento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autorização"
      Height          =   195
      Left            =   4560
      TabIndex        =   11
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOC"
      Height          =   195
      Left            =   3480
      TabIndex        =   9
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COO"
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bandeira:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   675
   End
End
Attribute VB_Name = "frmCartaoLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGravar_Click()
    If txtBandeira.tag = "" Then
        MsgBox "Informe a bandeira", vbInformation
        Exit Sub
    End If
    If CCur(txtValor.Text) <= 0 Then
        MsgBox "valor inválido", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub
Sub grava()
    
     strSQL = F_BuscaUmCampo("valor", "cartao_movimento", "financeira='" & txtBandeira.tag & "' and doc='" & txtDoc.Text & "' and doc <> ''")
    
    If strSQL <> "" Then
        If MsgBox("Já existe um lançamento para esta bandeira com este número de doc de R$ " & strSQL & " deseja continuar?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    
    
    Cartoes.gravaCartao_Movimento txtDia.Text, txtCoo.Text, txtDoc.Text, txtAut.Text, txtValor.Text, txtParcelas.Text, "N", txtBandeira.tag
    
    
    MsgBox "Movimento gravado!", vbExclamation
    txtValor.Text = "0"
    txtValor.SetFocus
    
End Sub
Sub altera()

End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnOrgao_Click()
    buscaFin
End Sub
Sub buscaFin()
    Dim f As financeira.financeira
    f = F_BuscaFinanceira(txtBandeira.Text)
    If f.cod_fin <> "" Then
        txtBandeira.Text = f.nome_fin
        txtBandeira.tag = f.cod_fin
        
        txtValor.SetFocus
        
        
        
    End If
End Sub

Private Sub Form_Load()
    txtDia.Text = Date
End Sub

Private Sub txtAut_GotFocus()
    P_SelecTexto txtAut
End Sub

Private Sub txtAut_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtAut_LostFocus()
    P_AjustaTexto txtAut
End Sub

Private Sub txtCoo_GotFocus()
    P_SelecTexto txtCoo
End Sub

Private Sub txtCoo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtCoo_LostFocus()
    P_AjustaTexto txtCoo
End Sub

Private Sub txtDia_GotFocus()
    P_SelecTexto txtDia
End Sub

Private Sub txtDia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtDia_LostFocus()
    P_AjustaData txtDia
End Sub

Private Sub txtDoc_GotFocus()
    P_SelecTexto txtDoc
End Sub

Private Sub txtDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtDoc_LostFocus()
    P_AjustaTexto txtDoc
End Sub

Private Sub txtParcelas_GotFocus()
    P_SelecTexto txtParcelas
End Sub

Private Sub txtParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtParcelas_LostFocus()
    P_ValorLostFocus txtParcelas, "#"
End Sub

Private Sub txtValor_GotFocus()
    P_SelecTexto txtValor
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtValor_LostFocus()
    P_ValorLostFocus txtValor
End Sub
