VERSION 5.00
Begin VB.Form frmLanFin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lançamento Financeiro"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7350
   Icon            =   "frmLanFin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   6975
         Begin VB.ComboBox cmbxConta 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   360
            Width           =   5175
         End
         Begin VB.TextBox txtbxClass 
            Appearance      =   0  'Flat
            BackColor       =   &H0005C9EB&
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Tag             =   "000"
            Text            =   "NÃO CLASSIFICADO"
            Top             =   960
            Width           =   4935
         End
         Begin VB.CommandButton cmndbtnPlanoContas 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5040
            Picture         =   "frmLanFin.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Conta Corrente:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   1110
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Classificação:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   3960
         Picture         =   "frmLanFin.frx":0694
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   5640
         Picture         =   "frmLanFin.frx":0B3C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tipo do Movimento:"
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3015
         Begin VB.OptionButton optnDeb 
            Caption         =   "Débito"
            Height          =   615
            Left            =   1560
            Picture         =   "frmLanFin.frx":0CA7
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optnCerd 
            Caption         =   "Crédito"
            Height          =   615
            Left            =   120
            Picture         =   "frmLanFin.frx":1971
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtbxData_Cai 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame frch 
         Caption         =   "Cheque"
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   5280
         Visible         =   0   'False
         Width           =   2415
         Begin VB.Label lblNum_Che 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   720
            TabIndex        =   30
            Tag             =   "!"
            Top             =   480
            Width           =   45
         End
         Begin VB.Label lblCod_Con 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "001"
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   720
            TabIndex        =   29
            Tag             =   "!"
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conta:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   6975
         Begin VB.TextBox txtbxValor_Cai 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5520
            TabIndex        =   10
            Text            =   "0,00"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtbxHist 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   5295
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor:"
            Height          =   195
            Left            =   5520
            TabIndex        =   9
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Histórico:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Frame frPagamento 
         Caption         =   "Pagamento"
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   6975
         Begin VB.TextBox txtbxTipo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtbxFuncionario 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   4935
         End
         Begin VB.CommandButton cmndbtnFun 
            Height          =   315
            Left            =   5040
            Picture         =   "frmLanFin.frx":4D53
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnFormasPag 
            Height          =   315
            Left            =   2520
            Picture         =   "frmLanFin.frx":50DD
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Funcionário"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   825
         End
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   5760
         TabIndex        =   31
         Tag             =   "!"
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data:"
         Height          =   195
         Left            =   5640
         TabIndex        =   4
         Top             =   120
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmLanFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxConta_Click()
    
'    lblNome_Con.caption = F_BuscaUmCampo("NOme_con", "tab_contas", "cod_con='" & cmbxConta.Text & "'")
    If cmbxConta.Text <> constContaCaixa And optnDeb.Value = True Then
        Dim ch As chePag
        'ch = F_DadosChePag(txtbxData_Cai.Text, lblNum_Che.caption, _
                cmbxConta.Text, txtbxValor_Cai.Text, True)
        'If ch.num_che <> "" Then
        '    frch.Visible = True
        '    lblNum_Che.caption = ch.num_che
        '    lblCod_Con.caption = ch.cod_con
        'Else
        '    frch.Visible = False
        'End If
    End If
End Sub

Private Sub cmndbtnPlanoContas_Click()
    Dim p As PlanoContas
    p = F_BuscaPlanoContas(Left(lblTipo.caption, 1), "")
    If p.cod_pla <> "" Then
        txtbxClass.Text = p.nome_pla
        txtbxClass.tag = p.cod_pla
        
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.tag = "" Then
        If txtbxValor_Cai.Text = "0,00" Then
            MsgBox "Informe um Valor válido!", vbExclamation
            txtbxValor_Cai.SetFocus
            Exit Sub
        End If
        If F_Vazio(txtbxHist.Text) Then
            MsgBox "Informe um histórico!", vbExclamation
            txtbxHist.SetFocus
            Exit Sub
        End If
        
        P_Grava
        
    Else
        P_Altera
    End If
    
End Sub
Sub P_Grava()
    Dim strSQL As String
    Dim id As String
    id = F_GravaFinanceiro(Left(lblTipo.caption, 1), txtbxData_Cai.Text, _
                txtbxValor_Cai.Text, txtbxHist.Text, Left(cmbxConta.Text, 3), _
                 "N", txtbxClass.tag)
    If strSQL = "" Then
        MsgBox "Erro!", vbCritical
        Exit Sub
    End If
    
    If optnDeb.Value And txtbxTipo.tag <> "" And txtbxFuncionario.tag <> "" Then
        If Not F_GravaPagFun(strSQL, txtbxFuncionario.tag, txtbxTipo.tag) Then
        End If
    End If
    
    MsgBox "Lançamento Gravado!", vbInformation
    txtbxData_Cai.Text = Date
    txtbxHist.Text = ""
    txtbxValor_Cai.Text = "0,00"
    txtbxHist.SetFocus
End Sub

Sub P_Altera()
    Dim strSQL As String
    strSQL = "Update tab_FINANCEIRO set historico_FIN='" & txtbxHist.Text & "',"
    strSQL = strSQL & "valor_FIN=" & F_GravaDec(txtbxValor_Cai.Text)
    strSQL = strSQL & ", COD_cON='" & cmbxConta.Text & "' "
    strSQL = strSQL & " where seq_FIN='" & Me.tag & "' "
    Conn.Execute strSQL
    
    If optnDeb.Value And txtbxTipo.tag <> "" And txtbxFuncionario.tag <> "" Then
        If Not F_GravaPagFun(strSQL, txtbxFuncionario.tag, txtbxTipo.tag) Then
        End If
    End If
    
    
    MsgBox "Dados Alterados!", vbInformation
     'P_Class CCur(txtbxValor_Cai.Text), Left(lblTipo.caption, 1), Me.Tag
    
    
    
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    
    P_CentraForm Me
    P_PreencheContas cmbxConta
    'cmbxConta.ListIndex = 0
    txtbxData_Cai.Text = Date
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad("frmEdiCai") Then
        
    End If
    Unload Me
End Sub

Private Sub optnCerd_Click()
    lblTipo.caption = "Crédito"
    txtbxClass.Text = ""
    txtbxClass.tag = "000"
End Sub

Private Sub optnCerd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnDeb_Click()
    lblTipo.caption = "Débito"
    txtbxClass.Text = ""
    txtbxClass.tag = "000"
End Sub

Private Sub optnDeb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxData_Cai_GotFocus()
    P_SelecTexto txtbxData_Cai
End Sub

Private Sub txtbxData_Cai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxData_Cai_LostFocus()
    P_AjustaData txtbxData_Cai
End Sub

Private Sub txtbxHist_GotFocus()
    P_SelecTexto txtbxHist
End Sub

Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxHist_LostFocus()
    P_AjustaTexto txtbxHist
End Sub

Private Sub txtbxValor_Cai_GotFocus()
    P_SelecTexto txtbxValor_Cai
End Sub

Private Sub txtbxValor_Cai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Cai_LostFocus()
    P_ValorLostFocus txtbxValor_Cai
End Sub
