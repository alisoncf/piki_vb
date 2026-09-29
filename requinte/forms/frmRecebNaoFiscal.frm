VERSION 5.00
Begin VB.Form frmRecebNaoFiscal 
   Caption         =   "Recebimento Não Fiscal - TEF"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   8805
   Icon            =   "frmRecebNaoFiscal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8805
      TabIndex        =   29
      Top             =   4590
      Width           =   8805
      Begin VB.CommandButton cmndbtnOk 
         Caption         =   "Efetuar Recebimento"
         Height          =   735
         Left            =   4800
         Picture         =   "frmRecebNaoFiscal.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   6720
         Picture         =   "frmRecebNaoFiscal.frx":04B4
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.TextBox txtParcelas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   26
      Text            =   "1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtFinanceira 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   7095
   End
   Begin VB.PictureBox pcPre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1080
      ScaleHeight     =   1665
      ScaleWidth      =   7425
      TabIndex        =   6
      Top             =   2280
      Width           =   7455
      Begin VB.TextBox txtMoeda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   23
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cmbxTipoParcelamento 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRecebNaoFiscal.frx":061F
         Left            =   3720
         List            =   "frmRecebNaoFiscal.frx":0638
         TabIndex        =   21
         Top             =   480
         Width           =   1470
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         TabIndex        =   19
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTroco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   17
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtValorOriginal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCapacidades 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtTipoOp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtOp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtRede 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Moeda:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4560
         TabIndex        =   24
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Parcelamento"
         Height          =   195
         Left            =   3720
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         Height          =   195
         Left            =   3720
         TabIndex        =   20
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Troco:"
         Height          =   195
         Left            =   2520
         TabIndex        =   18
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Orignal"
         Height          =   195
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidades"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Op."
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação"
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rede:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.TextBox txtbxHist 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   7455
   End
   Begin VB.TextBox txtbxValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0073E1F4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Text            =   "0,00"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdPag 
      Height          =   390
      Left            =   8160
      Picture         =   "frmRecebNaoFiscal.frx":06AD
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gerencial"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas"
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Financeira:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Histórico:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   405
   End
End
Attribute VB_Name = "frmRecebNaoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lanca_caixa As Boolean
Private Sub cmdPag_Click()
    buscaFin
End Sub

Private Sub cmndbtnOk_Click()
    Dim num_doc As String
    Dim sseq_cai As String
    num_doc = "RNF" & strCod_Cai & F_SoNumero(Now)
    
    If txtFinanceira.tag = "" Then
        MsgBox "Informe a financeira!", vbExclamation
        txtFinanceira.SetFocus
        Exit Sub
    End If
    If TefNaoFiscalSemECF(txtbxValor.Text, txtRede.Text, txtOp.Text, txtTipoOp.Text, Left(cmbxTipoParcelamento.Text, 1), _
        txtParcelas.Text, txtMoeda.Text, txtCapacidades.Text, True) Then
        
        P_Status "Transação efetuada com sucesso!"
        
        
            
        
        
            
            sseq_cai = F_GravaCai("R", Date, txtbxValor.Text, txtbxHist.Text, "", "CRT", strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, "000", "S")
            
            If F_GravaVenFin(sseq_cai, txtFinanceira.tag, txtbxValor.Text, Date, TransacaoTEFAtual.codigo_autorizacao, txtParcelas.Text, "R", "", TransacaoTEFAtual.via) Then
                
            End If
            Me.tag = sseq_cai
        If Me.tag <> "" Then
            Me.Hide
        Else
            txtValorOriginal.Text = "0,00"
            txtbxHist.Text = ""
        End If
        
            
    Else
        MsgBox "Erro na transação TEF.", vbExclamation
    End If
    P_Status ""
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
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

Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    P_ValorLostFocus txtbxValor
End Sub

Private Sub txtFinanceira_Change()
    txtFinanceira.tag = ""
End Sub

Private Sub txtFinanceira_GotFocus()
    P_SelecTexto txtFinanceira
End Sub

Private Sub txtFinanceira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFin
    End If
End Sub
Sub buscaFin()
    Dim f As financeira.financeira
    f = F_BuscaFinanceira(txtFinanceira.Text)
    If f.cod_fin <> "" Then
        txtFinanceira.Text = f.nome_fin
        txtFinanceira.tag = f.cod_fin
        cmndbtnOk.SetFocus
    End If
    
End Sub

Private Sub txtParcelas_GotFocus()
    P_SelecTexto txtParcelas
End Sub

Private Sub txtParcelas_LostFocus()
    On Error Resume Next
    If CInt(txtParcelas.Text) > 1 Then
        cmbxTipoParcelamento.ListIndex = 4
    Else
        cmbxTipoParcelamento.ListIndex = 0
    End If
End Sub
