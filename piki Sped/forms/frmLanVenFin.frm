VERSION 5.00
Begin VB.Form frmLanVenFin 
   Caption         =   "Movimento de Financeira"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7020
   Icon            =   "frmLanVenFin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxFinanceira 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox txtbxAutorizacao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtbxParc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   4
      Text            =   "1"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtbxValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6990
      TabIndex        =   0
      Top             =   4785
      Width           =   7020
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   5400
         Picture         =   "frmLanVenFin.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Salvar"
         Height          =   735
         Left            =   3840
         Picture         =   "frmLanVenFin.frx":109B5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmndbtnBuscarControl 
      Height          =   360
      Left            =   5160
      Picture         =   "frmLanVenFin.frx":10E5D
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financeira:"
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Autorização:"
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Label lblParcelas 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas:"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblCNPJ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   720
      Width           =   405
   End
End
Attribute VB_Name = "frmLanVenFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmndbtnBuscarControl_Click()
    busca
    
End Sub
Sub busca()
    Dim f As Financeira.Financeira
    f = F_BuscaFinanceira(txtbxFinanceira.Text)
    If f.cod_fin <> "" Then
        txtbxFinanceira.Text = f.nome_fin
        txtbxFinanceira.tag = f.cod_fin
        lblOperacao.caption = f.operacao_fin
        lblTipo.caption = f.tipo_fin
        lblCNPJ.caption = f.CNPJ
        txtbxAutorizacao.SetFocus
    End If
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag <> "" Then
        Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If txtbxFinanceira.tag = "" Then
        MsgBox "Informe a financeira!", vbExclamation
        txtbxFinanceira.SetFocus
        Exit Sub
    End If
    If Me.tag <> "" Then
        strSQL = "Update " & TAB_VENFIN
        strSQL = strSQL & " set autorizacao_fin='" & txtbxAutorizacao.Text & "',"
        strSQL = strSQL & " cod_fin='" & txtbxFinanceira.tag & "'"
        strSQL = strSQL & " where num_ven='" & Me.tag & "'"
       ' strSQL = strSQL & " and parcela_fin='" & txtbxParc.Text & "' "
        conn.Execute strSQL
        
        MsgBox "dados gravados!", vbInformation
        
        
    End If
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select tf.nome_fin,tvf.* from tab_financeiras as tf inner join "
    strSQL = strSQL & TAB_VENFIN & "   as tvf on tvf.cod_fin=tf.cod_fin "
    strSQL = strSQL & " left join " & TAB_VENDAS & " as tv on tv.num_ven=tvf.num_ven "
    strSQL = strSQL & " where tvf.num_ven='" & Me.tag & "'"
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            txtbxAutorizacao.Text = !autorizacao_fin
            txtbxFinanceira.Text = !nome_fin
            txtbxFinanceira.tag = !cod_fin
            txtbxValor.Text = FormatVal(!valor_fin)
            txtbxParc.Text = !parcela_fin
            lblParcelas.caption = "Parc:" & !parcela_fin & "/" & !parcelas_fin & ""

            txtbxParc.Enabled = False
            txtbxValor.Enabled = False
        End If
        .Close
    End With
    
    
End Sub
Private Sub Form_Load()
    bActive = False
End Sub

Private Sub txtbxFinanceira_Change()
    txtbxFinanceira.tag = ""
End Sub

Private Sub txtbxFinanceira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
    End If
End Sub


