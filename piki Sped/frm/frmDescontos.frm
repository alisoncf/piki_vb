VERSION 5.00
Begin VB.Form frmDescontos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comissão Decrescente"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Cadastro de Descontos"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton cmndbtnFechar 
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   1440
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2535
         Begin VB.TextBox txtbxComiss_Des 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            Text            =   "0,0"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtbxValor_Des 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Text            =   "0,0"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1800
            TabIndex        =   9
            Top             =   240
            Width           =   150
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1800
            TabIndex        =   7
            Top             =   720
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&omissão:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Desconto:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton cmndbtnConf 
         Caption         =   "&Confirmar"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmDescontos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub CmndbtnConf_Click()
    If CCur(txtbxComiss_Des.Text) <= 0 Then
        MsgBox "Informe um valor válido para a comissão!", vbExclamation
        txtbxComiss_Des.SetFocus
        Exit Sub
    End If
    If CCur(txtbxValor_Des.Text) <= 0 Then
        MsgBox "Informe um valor válido para o desconto!", vbExclamation
        txtbxValor_Des.SetFocus
        Exit Sub
    End If
    If Me.Tag = "" Then
        If F_BuscaUmCampo("valor_des", "tab_descontos", "valor_ven=" & _
             f_GravaDec(txtbxValor_Des.Text)) <> "" Then
            MsgBox "Já existe uma comissão para este desconto!", vbExclamation
            Exit Sub
        End If
        grava
    Else
        altera
    End If
End Sub
Sub grava()
    Dim strSQL As String
    strSQL = "insert into tab_descontos(cod_des,valor_des,comiss_des)values("
    strSQL = strSQL & "'" & ChaveSQL("tab_descontos", "cod_des", "#") & "',"
    strSQL = strSQL & f_GravaDec(txtbxValor_Des.Text) & ","
    strSQL = strSQL & f_GravaDec(txtbxComiss_Des.Text) & ")"
    banco.Execute strSQL
    MsgBox "Desconto Cadastrado!", vbExclamation
    If IsLoad("frmEdiDes") Then
        frmEdiDes.dtServ.Refresh
    End If
    txtbxComiss_Des.Text = "0,0"
    txtbxValor_Des.Text = "0,0"
    txtbxValor_Des.SetFocus
End Sub
Sub altera()
    Dim strSQL As String
    strSQL = "update tab_descontos set "
    strSQL = strSQL & "valor_des=" & f_GravaDec(txtbxValor_Des.Text) & ","
    strSQL = strSQL & "comiss_Des=" & f_GravaDec(txtbxComiss_Des.Text) & ""
    strsqwl = strSQL & " cod_ven='" & Me.Tag & "' "
    banco.Execute strSQL
    MsgBox "Desconto Alterado!", vbExclamation
    If IsLoad("frmEdiDes") Then
        frmEdiDes.dtServ.Refresh
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    P_CentraForm Me
End Sub

Private Sub txtbxComiss_Des_GoTFocus()
    P_SelecTexto txtbxComiss_Des
End Sub

Private Sub txtbxComiss_Des_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxComiss_Des_LostFocus()
    P_ValorLostFocus txtbxComiss_Des
End Sub

Private Sub txtbxValor_Des_gotFocus()
    P_SelecTexto txtbxValor_Des
End Sub

Private Sub txtbxValor_Des_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxValor_Des_LostFocus()
    P_ValorLostFocus txtbxValor_Des
End Sub
