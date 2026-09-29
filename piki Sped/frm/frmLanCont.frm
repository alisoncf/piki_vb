VERSION 5.00
Begin VB.Form frmLanCont 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Contagem de Produto"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmLanCont.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optnBar 
      Caption         =   "&B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.OptionButton optnN 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmndbtnFEchar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   4320
      Picture         =   "frmLanCont.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   2040
      Picture         =   "frmLanCont.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtbxQtd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtbxCodBarra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "N"
      Top             =   600
      Width           =   4335
   End
   Begin VB.OptionButton optnC 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblCont 
      Caption         =   "Contagem:"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblMarca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2520
      TabIndex        =   12
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label lblCodBar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód Barra:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblCod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Qtd:"
      Height          =   195
      Left            =   4680
      TabIndex        =   2
      Tag             =   "B"
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblNome_pro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Width           =   645
   End
   Begin VB.Label lblExpr 
      AutoSize        =   -1  'True
      Caption         =   "Cod. Barra:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Tag             =   "B"
      Top             =   360
      Width           =   795
   End
End
Attribute VB_Name = "frmLanCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    If txtbxCodBarra.tag = "" Then
        MsgBox "Escolha primeiro o produto!", vbExclamation
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    If MsgBox("Incluir contagem de: " & lblNome_pro.caption & " com " & _
        txtbxQtd.Text & " ?", vbYesNo + vbQuestion) = vbYes Then
        
        strSQL = F_BuscaUmCampo("qtd_itc", "tab_itcont", "seq_cnt='" & Me.tag & _
            "' and cod_pro='" & txtbxCodBarra.tag & "'")
        If strSQL <> "" Then
            If MsgBox("Este item já se encontra lançado com " & strSQL & _
                 " Deseja alterar sua quantidade? ", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
                
        
            
        End If
        strSQL = " delete from tab_itcont where seq_cnt='" & Me.tag & "' "
        strSQL = strSQL & " and cod_pro='" & txtbxCodBarra.tag & "'"
        conn.Execute strSQL
        strSQL = " insert into tab_itcont(seq_cnt,cod_pro,qtd_itc,data_itc)values( "
        strSQL = strSQL & "'" & Me.tag & "', "
        strSQL = strSQL & "'" & txtbxCodBarra.tag & "', "
        strSQL = strSQL & F_GravaDec(txtbxQtd) & ","
        strSQL = strSQL & F_GravaDataTempo(Now) & ") "
        
        conn.Execute strSQL
        txtbxCodBarra.Text = ""
        txtbxCodBarra.tag = ""
        txtbxQtd.Text = "0,00"
        txtbxCodBarra.SetFocus
        
    End If
End Sub

Private Sub Form_Load()
    Me.tag = F_BuscaUmCampo("max(seq_cnt)", "tab_contagens", "")
    lblCont.caption = "Contagem: " & Me.tag & Chr(13) & _
    F_BuscaUmCampo("data_cnt", "tab_contagens", "seq_cnt='" & Me.tag & "'")
    optnN_Click
End Sub

Private Sub optnBar_Click()
    lblExpr.caption = "C&od. Barra:"
    lblExpr.tag = "B"
    txtbxCodBarra.SetFocus
End Sub

Private Sub optnC_Click()
        lblExpr.caption = "C&od. Prod.:"
        lblExpr.tag = "C"
        txtbxCodBarra.SetFocus
End Sub



Private Sub optnN_Click()
    lblExpr.caption = "N&ome Prod.:"
    lblExpr.tag = "N"
    
End Sub
Private Sub txtbxCodBarra_GotFocus()
    ganhaFoco txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_LostFocus()
    AjustaTexto txtbxCodBarra
    If optnC.Value = True Then
        txtbxCodBarra.Text = Format(txtbxCodBarra.Text, "0000000#")
        
    End If
    
End Sub
Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
        
    End If
End Sub
Sub busca()
    Dim p As Produto
    txtbxCodBarra_LostFocus
    p = F_BuscaProdSaida(txtbxCodBarra.Text)
    If p.cod_pro <> "" Then
        If optnN.Value Then
            txtbxCodBarra.Text = p.nome_pro
        ElseIf optnC.Value Then
            txtbxCodBarra.Text = p.cod_pro
        Else
                
        End If
        txtbxCodBarra.tag = p.cod_pro
        lblNome_pro.caption = "Prod: " & p.nome_pro
        lblCod.caption = "Cód: " & p.cod_pro
        lblCodBar.caption = "Cód.Barra: " & F_BuscaUmCampo("codbarra_pro", "tab_codbarras", "cod_pro='" & p.cod_pro & "'")
        lblMarca.caption = "Marca: " & p.nome_mar
        On Error Resume Next
        TrocaTabKey 13
    End If
End Sub

Private Sub txtbxQtd_GotFocus()
    ganhaFoco txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnOk_Click
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    PerdeFoco txtbxQtd
End Sub
