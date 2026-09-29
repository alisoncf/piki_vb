VERSION 5.00
Begin VB.Form frmDePara 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "De Para.:.Utilitário de Transferência de Produtos"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "frmDePara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6045
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   4440
      Picture         =   "frmDePara.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnConf 
      Caption         =   "&Confirmar"
      Height          =   615
      Left            =   3000
      Picture         =   "frmDePara.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnBuscarPara 
      Caption         =   "B&uscar"
      Height          =   405
      Left            =   4920
      TabIndex        =   5
      Top             =   1260
      Width           =   855
   End
   Begin VB.TextBox txtbxPara 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CommandButton cmndbtnBuscaDe 
      Caption         =   "&Buscar"
      Height          =   405
      Left            =   4920
      TabIndex        =   2
      Top             =   420
      Width           =   855
   End
   Begin VB.TextBox txtbxDe 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disp:"
      Height          =   195
      Left            =   3240
      TabIndex        =   17
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label lblEF_Para 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   3720
      TabIndex        =   16
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Est:"
      Height          =   195
      Left            =   1920
      TabIndex        =   15
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label lblER_Para 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   2400
      TabIndex        =   14
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disp:"
      Height          =   195
      Left            =   3240
      TabIndex        =   13
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblEF_de 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   3720
      TabIndex        =   12
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Est:"
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   120
      Width           =   270
   End
   Begin VB.Label lblEr_de 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para: O produto que receberá a movimentação..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DE: o produto que irá desaparecer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   2820
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Para:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmDePara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscaDe_Click()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxDe.Text, "B")
    txtbxDe.Text = p.nome_pro
    txtbxDe.tag = p.cod_pro
    lblEr_de.caption = p.est.er_pro
    lblEF_de.caption = p.est.ef_pro
    If txtbxDe.tag <> "" Then
        txtbxPara.SetFocus
    End If
End Sub

Private Sub cmndbtnBuscarPara_Click()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxPara.Text, "B")
    txtbxPara.Text = p.nome_pro
    txtbxPara.tag = p.cod_pro
    lblER_Para.caption = p.est.er_pro
    lblEF_Para.caption = p.est.ef_pro
End Sub

Private Sub CmndbtnConf_Click()
    Dim strSQL As String
    If txtbxDe.tag = "" Then
        MsgBox "Escolha o produto!", vbExclamation
        txtbxDe.SetFocus
        Exit Sub
    End If
    If txtbxPara.tag = "" Then
        MsgBox "Escolha o produto!", vbExclamation
        txtbxPara.SetFocus
        Exit Sub
    End If
    strSQL = "O Produto: " & txtbxDe.Text & Chr(13) & "Irá desaparecer "
    strSQL = strSQL & "e terá todo sua movimentação (entrada,saída,vendas) "
    strSQL = strSQL & "transferida para o Produto: " & Chr(13) & txtbxPara.Text & "" & Chr(13)
    strSQL = strSQL & "Confirma Operação?"
    If MsgBox(strSQL, vbYesNo + vbCritical, "Atenção!") = vbYes Then
        DePara
    End If
End Sub
Sub DePara()
    Dim strSQL As String
    strSQL = "update tab_itvendas set cod_pro='" & txtbxPara.tag & "' "
    strSQL = strSQL & " where cod_pro='" & txtbxDe.tag & "' "
    Banco.Execute strSQL
    strSQL = "update tab_itsai set cod_pro='" & txtbxPara.tag & "' "
    strSQL = strSQL & " where cod_pro='" & txtbxDe.tag & "' "
    Banco.Execute strSQL
    strSQL = "update tab_itcmp set cod_pro='" & txtbxPara.tag & "' "
    strSQL = strSQL & " where cod_pro='" & txtbxDe.tag & "' "
    Banco.Execute strSQL
    strSQL = "update tab_itcompras set cod_pro='" & txtbxPara.tag & "' "
    strSQL = strSQL & " where cod_pro='" & txtbxDe.tag & "' "
    Banco.Execute strSQL
    strSQL = "update tab_devolucoes set cod_pro='" & txtbxPara.tag & "' "
    strSQL = strSQL & " where cod_pro='" & txtbxDe.tag & "' "
    Banco.Execute strSQL
    'transf tab_itvendas, itcompras, itcmp,tab_itsai,tab_itcont,'tab_invent'
    'excluir produto
    strSQL = "delete from tab_produtos where cod_pro = '" & txtbxDe.tag & "'"
        Banco.Execute strSQL
        strSQL = "delete from tab_embalagens where cod_pro = '" & txtbxDe.tag & "'"
        Banco.Execute strSQL
        strSQL = "delete from tab_codbarras where cod_pro = '" & txtbxDe.tag & "'"
        Banco.Execute strSQL
    
    If CCur(lblEr_de.caption) > 0 Then
        If MsgBox("O produto: " & txtbxDe.Text & " Possui estoque." & _
        Chr(13) & "Deseja adicioná-lo a " & Chr(13) & txtbxPara.Text & _
        " Deixando-o com: " & CCur(lblEr_de.caption) + CCur(lblER_Para.caption) & _
        "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "update tab_produtos set er_pro=" & F_GravaDec(CCur(lblEr_de.caption) + CCur(lblER_Para.caption))
        strSQL = strSQL & " where cod_pro='" & txtbxPara.tag & "' "
        Banco.Execute strSQL
        End If
    End If
    If CCur(lblEF_de.caption) > 0 Then
        If MsgBox("O produto: " & txtbxDe.Text & " Possui estoque disp. " & _
        Chr(13) & "Deseja adicioná-lo a " & Chr(13) & txtbxPara.Text & _
        "Deixando-o com: " & CCur(lblEF_de.caption) + CCur(lblEF_Para.caption) & _
        "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "update tab_produtos set ef_pro=" & F_GravaDec(CCur(lblEF_de.caption) + CCur(lblEF_Para.caption))
        strSQL = strSQL & " where cod_pro='" & txtbxPara.tag & "' "
        Banco.Execute strSQL
        End If
    End If
    MsgBox "Operação efetuada com sucesso!", vbInformation
    txtbxDe.Text = ""
    txtbxDe.tag = ""
    txtbxPara.Text = ""
    txtbxPara.tag = ""
    txtbxDe.SetFocus
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub txtbxDe_Change()
    txtbxDe.tag = ""
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaDe_Click
    End If
End Sub

Private Sub txtbxPara_Change()
    txtbxPara.tag = ""
End Sub

Private Sub txtbxPara_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscarPara_Click
    End If
End Sub
