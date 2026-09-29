VERSION 5.00
Begin VB.Form frmAltProPCP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Incluir/Alterar Produto do Processo de Compra:"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnHist 
      Caption         =   "Histórico de Compras"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame frAltMer 
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Text            =   "0,00"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbxForn 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   330
         Left            =   3360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtbxQFA_Ipc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "0,00"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnSim 
         Caption         =   "Sim"
         Height          =   615
         Left            =   7200
         Picture         =   "frmAltProPCP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   390
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnNao 
         Caption         =   "Não"
         Height          =   615
         Left            =   8520
         Picture         =   "frmAltProPCP.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Unitário:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total.:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornecedor:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblNome_Mer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome da Mercadoria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   40
         Width           =   2040
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. Comprar:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº:"
      Height          =   195
      Left            =   2160
      TabIndex        =   24
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label lblNumNot 
      AutoSize        =   -1  'True
      Caption         =   "00/00/0000"
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
      Left            =   2400
      TabIndex        =   23
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label lblEmb 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
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
      Left            =   6480
      TabIndex        =   22
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Embalagem:"
      Height          =   195
      Left            =   5520
      TabIndex        =   21
      Top             =   1800
      Width           =   870
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
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
      Left            =   8520
      TabIndex        =   20
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço Unitário:"
      Height          =   195
      Left            =   7320
      TabIndex        =   19
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label lblQtd 
      AutoSize        =   -1  'True
      Caption         =   "0,00"
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
      Left            =   4800
      TabIndex        =   18
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd Comprada:"
      Height          =   195
      Left            =   3600
      TabIndex        =   17
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label lblNomeFor 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Fornecedor"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   1560
      Width           =   1785
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor:"
      Height          =   195
      Left            =   4560
      TabIndex        =   15
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblUltComp 
      AutoSize        =   -1  'True
      Caption         =   "00/00/0000"
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
      Left            =   3360
      TabIndex        =   14
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Última Compra:"
      Height          =   195
      Left            =   2160
      TabIndex        =   13
      Top             =   1560
      Width           =   1065
   End
End
Attribute VB_Name = "frmAltProPCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxForn_Click()
    cmbxForn.Tag = F_BuscaUmCampo("cod_For", "tab_fornecedores", "nome_for='" & cmbxForn.Text & "'")
End Sub

Private Sub cmbxForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnHist_Click()
    P_InfoEntradasPro lblNome_Mer.Tag
End Sub

Private Sub cmndbtnNao_Click()
    
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim iRow As Integer

        
        strSQL = "update tab_itenspcp set sts_ipc='N',"
        strSQL = strSQL & "QFA_ipc=" & Str(txtbxQFA_Ipc.Text)
        
        strSQL = strSQL & " where Cod_pro='" & lblNome_Mer.Tag & "' "
        strSQL = strSQL & " and seq_pcp='" & Me.Tag & "' "
        Banco.Execute strSQL
        
        Me.Hide
    


End Sub

Private Sub cmndbtnSim_Click()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim iRow As Integer
        If cmbxForn.Tag = "" Then
            MsgBox "Informe o fornecedor!", vbExclamation
            cmbxForn.SetFocus
            Exit Sub
        End If
        If Int(F_BuscaUmCampo("count(*)", "tab_itenspcp", "cod_pro='" & lblNome_Mer.Tag & "'" & _
            " and seq_pcp='" & Me.Tag & "'")) > 1 Then   'tem q ser mudado em caso de mais lugares
            MsgBox "Este item já está duplicado!", vbCritical
            Exit Sub
        End If
        
        
        strSQL = "update tab_itenspcp set sts_ipc='S',"
        strSQL = strSQL & "QFA_ipc=" & Str(txtbxQFA_Ipc.Text) & ","
        strSQL = strSQL & "total_ipc=" & Str(txtbxTotal.Text) & ","
        strSQL = strSQL & "cod_for='" & cmbxForn.Tag & "' "
        strSQL = strSQL & " where Cod_pro='" & lblNome_Mer.Tag & "' "
        strSQL = strSQL & " and seq_pcp='" & Me.Tag & "' "
        
        
        Banco.Execute strSQL
        
    
        Me.Hide

End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Activate()
    BuscaUltEntrada
End Sub
Sub VerificaProcesso()
    Dim strSQL As String
    
End Sub
Private Sub Form_Load()
    MudaAparencia Me
    P_PreencheFornecedores cmbxForn
End Sub
Sub BuscaUltEntrada()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_erro
    strSQL = "select num_not,entrada_not,tf.nome_for,ti.* "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti, "
    strSQL = strSQL & " tab_fornecedores as tf "
    strSQL = strSQL & " where tf.cod_for=tc.cod_for "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and ti.cod_pro='" & lblNome_Mer.Tag & "' "
    strSQL = strSQL & " order by entrada_not desc limit 1 "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
            lblUltComp.caption = ""
            lblNomeFor.caption = ""
            lblNumNot.caption = ""
            lblQtd.caption = ""
            lblUnit.caption = ""
            lblEmb.caption = ""
        Else
            lblUltComp.caption = !entrada_not
            lblNomeFor.caption = !nome_for
            lblNumNot.caption = !Num_Not
            lblQtd.caption = !qtd_ite
            lblUnit.caption = !unit_ite
            lblEmb.caption = !qtd_emb
        End If
        .Close
    End With
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Private Sub Text1_GotFocus()
    
End Sub

Private Sub txtbxQFA_Ipc_GotFocus()
    P_SelecTexto txtbxQFA_Ipc
End Sub

Private Sub txtbxQFA_Ipc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQFA_Ipc_LostFocus()
    P_ValorLostFocus txtbxQFA_Ipc
    TotalItem
End Sub

Sub TotalItem()
    If CCur(txtbxQFA_Ipc.Text) > 0 Then
        txtbxUnit.Text = FormatVal(CCur(txtbxTotal.Text) / CCur(txtbxQFA_Ipc.Text), "###,##0.0000")
    Else
        txtbxUnit.Text = "0"
    End If
End Sub
Private Sub txtbxTotal_GotFocus()
    P_SelecTexto txtbxTotal
End Sub

Private Sub txtbxTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxTotal_LostFocus()
    P_ValorLostFocus txtbxTotal
    TotalItem
End Sub
