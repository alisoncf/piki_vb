VERSION 5.00
Begin VB.Form frmIncServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Serviços"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4560
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   855
      Left            =   3480
      Picture         =   "frmIncServ2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   855
      Left            =   2400
      Picture         =   "frmIncServ2.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtbxComissao_Ser 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtbxValor_Ser 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtbxNome_Ser 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Comissão(%):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   5
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Serviço:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmIncSErv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    sCod = F_BuscaUmCampo("Max(cod_Ser)", "tab_Servicos", "")
    If sCod = "" Then
        sCod = "001"
    Else
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "insert into tab_servicos(cod_Ser,nome_ser,valor_Ser,comissao_Ser)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxNome_Ser.Text & "',"
    strSQL = strSQL & f_GravaDec(txtbxValor_Ser.Text) & ","
    strSQL = strSQL & f_GravaDec(txtbxComissao_Ser.Text) & ")"
    banco.Execute strSQL
    MsgBox "Serviço cadastrado!", vbInformation
    txtbxNome_Ser.Text = ""
    txtbxNome_Ser.SetFocus
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.Tag = "" Then
        P_Salva
    Else
    End If
End Sub

Private Sub txtbxComissao_Ser_GotFocus()
    P_SelecTexto txtbxComissao_Ser
End Sub

Private Sub txtbxComissao_Ser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxComissao_Ser_LostFocus()
    P_ValorLostFocus txtbxComissao_Ser
    If CCur(txtbxComissao_Ser.Text) > 100 Then
        MsgBox "Ficou maluco?!", vbExclamation
        txtbxComissao_Ser.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtbxNome_Ser_GotFocus()
    P_SelecTexto txtbxNome_Ser
End Sub

Private Sub txtbxNome_Ser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Ser_LostFocus()
    P_AjustaTexto txtbxNome_Ser
End Sub

Private Sub txtbxValor_Ser_GotFocus()
    P_SelecTexto txtbxValor_Ser
End Sub

Private Sub txtbxValor_Ser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Ser_LostFocus()
    P_ValorLostFocus txtbxValor_Ser
End Sub
