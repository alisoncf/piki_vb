VERSION 5.00
Begin VB.Form frmEstoqueIni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estoque Inicial"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmAjFornadas 
      Caption         =   "Dados do Produto:"
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   735
         Left            =   4680
         TabIndex        =   9
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtbxFornada 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtbxNome_Pro 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmnxdbtnConf 
         Caption         =   "&Confirmar"
         Height          =   525
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   525
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornada:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   615
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd"
         Height          =   195
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmEstoqueIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sRest As String
Private Sub cmnxdbtnConf_Click()
    If Me.Tag = "" Then
        
            grava
            
            
        
        
    Else
    
    End If
End Sub
Sub grava()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "insert into tab_fornadas(cod_pro,qtd_pro,fornada_pro)values("
    strSQL = strSQL & "'" & txtbxNome_Pro.Tag & "',"
    strSQL = strSQL & F_GravaDec(txtbxQtd.Text) & ", "
    strSQL = strSQL & "'" & txtbxFornada.Text & "') "
    banco.Execute strSQL
    txtbxFornada.Text = ""
    txtbxQtd.Text = "0,00"
    
    MsgBox "Produto lançado!", vbInformation
    
    'txtbxNome_Pro.Text = sRest
    txtbxNome_Pro.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub txtbxData_Change()

End Sub

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select sum(tf.qtd_pro) as soma, tp.* from tab_produtos as tp,"
    strSQL = strSQL & "tab_fornadas as tf where tp.cod_pro=tf.cod_pro group by tf.cod_pro"
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_produtos set estoque_pro=" & F_GravaDec(!soma)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
            banco.Execute strSQL
            strSQL = "update tab_estoque set estoque_pro=" & F_GravaDec(!soma)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
            banco.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    MsgBox "terminou tbm"
End Sub

Private Sub txtbxFornada_GotFocus()
    P_SelecTexto txtbxFornada
End Sub

Private Sub txtbxFornada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxFornada_LostFocus()
    P_AjustaTexto txtbxFornada
End Sub

Private Sub txtbxNome_Pro_Change()
    txtbxNome_Pro.Tag = ""
End Sub

Private Sub txtbxNome_Pro_DblClick()
    Dim sExp As String
    Dim p As Produto
    sRest = txtbxNome_Pro.Text
    p = F_BuscaProdSaida(txtbxNome_Pro.Text)
    txtbxNome_Pro.Text = p.nome_pro
    txtbxNome_Pro.Tag = p.cod_pro
    If p.cod_pro <> "" Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxNome_Pro_GotFocus()
    P_SelecTexto txtbxNome_Pro
End Sub

Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_Pro_DblClick
    End If
End Sub


Private Sub txtbxQtd_For_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
