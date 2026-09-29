VERSION 5.00
Begin VB.Form frmQueima 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Queima"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxFab 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      MaxLength       =   40
      TabIndex        =   13
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtbxSecao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtbxSubSec 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Apenas quem estiver fora da promoção"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.TextBox txtDesc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Text            =   "0,00"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Retirar da Promoção"
      Height          =   195
      Left            =   2400
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.OptionButton optnAdd 
      Caption         =   "Adicionar na Promoção"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
   Begin VB.CommandButton cmndbtnSecao 
      Height          =   285
      Left            =   2400
      Picture         =   "frmQueima.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmndbtnSubSec 
      Height          =   285
      Left            =   5040
      Picture         =   "frmQueima.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmndbtnFabr 
      Height          =   285
      Left            =   7320
      Picture         =   "frmQueima.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblSubCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Subseção:"
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label lblSecCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Seção:"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   1080
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricante/Marca:"
      Height          =   195
      Left            =   5400
      TabIndex        =   17
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Label lblAfet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6120
      TabIndex        =   8
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "Produtos que serão afetados:"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto  (%):"
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Palavra-chave:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmQueima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirmar_Click()
    If Not F_senha Then Exit Sub
    If strAltpreco_usu <> "S" Then
        MsgBox "Você naõ tem permissão para alterar preços!", vbExclamation
        Exit Sub
    End If
    If optnAdd Then
        If MsgBox("Deseja adicionar na promoção " & lblAfet.caption & " itens?", vbYesNo + vbQuestion) = vbYes Then
            adicionar
        End If
    Else
        If MsgBox("Deseja retirar da promoção " & lblAfet.caption & " itens?", vbYesNo + vbQuestion) = vbYes Then
            retira
        End If
    End If
    
End Sub
Sub retira()
    Dim strSQL As String
    Dim Where As String
    Dim rPesq As New ADODB.Recordset
    Dim sql_codigos As String
    strSQL = "update tab_produtos set pq_pro=0 "
    
    
    Where = " where nome_pro like '" & txtNome.text & "%' "
    strSQL = strSQL & Where
    Conn.Execute strSQL
    
    
    sql_codigos = "select cod_pro from tab_produtos " & Where
    abreObjetoAdo rPesq, sql_codigos
    With rPesq
        Do While Not .EOF
            insere_fila_produtos !cod_pro
            .MoveNext
        Loop
    End With
    executa_fila_produtos
    
    MsgBox "Produtos alterados!", vbInformation
    buscaAfet
End Sub
Sub adicionar()
    Dim strSQL As String
    Dim i As Integer
    Dim Where As String
    Dim rPesq As New ADODB.Recordset
    Dim sql_codigos As String
    
    
    
    
    strSQL = "update tab_produtos set pq_pro= pv_pro*" & F_GravaDec((100 - CCur(txtDesc.text)) / 100)
    Where = "  where nome_pro like '" & txtNome.text & "%' "
    If txtbxFab.tag <> "" Then
        Where = Where & " and cod_fab='" & txtbxFab.tag & "'"
    End If
    If Check1.Value = 1 Then
        Where = Where & " and pq_pro=0"
    End If
    strSQL = strSQL & Where
    Conn.Execute strSQL, i
    
    MsgBox i & " Produtos alterados!", vbInformation
    buscaAfet
    
    
    sql_codigos = "select cod_pro from " & TAB_PRODUTOS & " " & Where
    
    abreObjetoAdo rPesq, sql_codigos
    
    Do While Not rPesq.EOF
        insere_fila rPesq!cod_pro, 3
        
        
        rPesq.MoveNext
    Loop
    
    P_Status "enviando pra internet"
    executa_fila_produtos
    P_Status
    
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnFabr_Click()
    buscaFor
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_SelectForn(txtbxFab.text)
    If f.cod_For <> "" Then
        Label4.caption = f.cod_For
        txtbxFab.text = f.nome_for
        txtbxFab.tag = f.cod_For
        buscaAfet
    End If
    
End Sub
Private Sub cmndbtnSecao_Click()
    Dim s As secao.secao
    s = F_BuscaSec(txtbxSecao.text)
    If s.cod_sec <> "" Then
        txtbxSecao.text = s.nome_sec
        txtbxSecao.tag = s.cod_sec
        txtbxSubSec.text = ""
        txtbxSubSec.tag = ""
        buscaAfet
    End If
    
    
    
End Sub

Private Sub cmndbtnSubSec_Click()
    Dim s As subsec
    s = F_BuscaSubSec(txtbxSecao.tag, txtbxSubSec.text)
    If s.cod_sub <> "" Then
        txtbxSubSec.text = s.nome_sub
        txtbxSubSec.tag = s.cod_sub
        buscaAfet
        
    End If
End Sub


Private Sub Option1_Click()
    buscaAfet
End Sub

Private Sub optnAdd_Click()
    buscaAfet
End Sub
Sub buscaAfet()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    

    strSQL = "select count(*) as total from tab_produtos as tp  inner join tab_subsec as ts"
    strSQL = strSQL & " on tp.cod_sub=ts.cod_sub and nome_pro like '" & txtNome.text & "%' "
    If optnAdd.Value Then
        'strSQL = strSQL & " and pq_pro=0 "
        If Check1.Value = 1 Then
            strSQL = strSQL & " and pq_pro=0"
        End If
    Else
        strSQL = strSQL & " and pq_pro<>0 "
    End If
    
    If txtbxFab.text <> "" Then
        strSQL = strSQL & " and tp.cod_fab='" & txtbxFab.tag & "'    "
    End If
    If txtbxSecao.tag <> "" Then
        strSQL = strSQL & " and ts.cod_sec='" & txtbxSec.tag & "'    "
    End If
    If txtbxSubSec.tag <> "" Then
        strSQL = strSQL & " and tp.cod_sub='" & txtbxSubSec.tag & "'    "
    End If
    
    
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        lblAfet.caption = rPesq!total
        rPesq.Close
    End If
    
    
    
End Sub

Private Sub txtbxFab_Change()
    txtbxFab.tag = ""
End Sub

Private Sub txtbxSecao_Change()
    txtbxSecao.tag = ""
End Sub

Private Sub txtbxSubSec_Change()
    txtbxSubSec.tag = ""
End Sub

Private Sub txtNome_Change()
    buscaAfet
End Sub
