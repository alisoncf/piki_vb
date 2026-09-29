VERSION 5.00
Begin VB.Form frmIncServic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Serviços"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frmIncServic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   Begin VB.Frame frControle 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4335
         Begin VB.TextBox txtbxNome_Ser 
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   6
            Top             =   480
            Width           =   3975
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
            TabIndex        =   5
            Text            =   "0,00"
            Top             =   1320
            Width           =   1455
         End
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
            TabIndex        =   4
            Text            =   "0,00"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Serviço:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   8
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   7
            Top             =   960
            Width           =   1470
         End
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   855
         Left            =   2400
         Picture         =   "frmIncServic.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   855
         Left            =   3480
         Picture         =   "frmIncServic.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2160
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmIncServic"
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
Sub P_Altera()
    
    strSQL = "Update tab_Servicos set "
    
    strSQL = strSQL & "Nome_Ser='" & txtbxNome_Ser.Text & "',"
    strSQL = strSQL & "valor_ser=" & f_GravaDec(txtbxValor_Ser.Text) & ","
    strSQL = strSQL & "comissao_Ser=" & f_GravaDec(txtbxComissao_Ser.Text) & " "
    strSQL = strSQL & " where cod_ser = '" & Me.Tag & "' "
    banco.Execute strSQL
    MsgBox "Serviço Alterado!", vbInformation
    Unload Me
    If IsLoad("frmEdiServ") Then
        RefreshDataPosition frmEdiServ.dtServ
    End If
End Sub
Private Sub cmndbtnSalvar_Click()
    If Me.Tag = "" Then
        P_Salva
    Else
        P_Altera
    End If
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    Me.Height = 3405
    Me.Width = 4680
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
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
