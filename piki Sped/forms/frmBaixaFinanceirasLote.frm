VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBaixaFinanceirasLote 
   Caption         =   "Baixa em Lote de Movimento de Financeiras"
   ClientHeight    =   5820
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11505
   Icon            =   "frmBaixaFinanceirasLote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5820
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   11445
      TabIndex        =   3
      Top             =   5445
      Width           =   11505
      Begin VB.Label lblFoco 
         AutoSize        =   -1  'True
         Caption         =   "Foco"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   360
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7858
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      Begin VB.TextBox txtConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   5520
         ScaleHeight     =   705
         ScaleWidth      =   5505
         TabIndex        =   4
         Top             =   120
         Width           =   5535
         Begin VB.CommandButton cmdProcessar 
            Caption         =   "Processar"
            Height          =   375
            Left            =   3360
            TabIndex        =   13
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox txtColVal 
            Height          =   285
            Left            =   2400
            TabIndex        =   9
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtColParc 
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtColAut 
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Colunas"
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
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   2400
            TabIndex        =   10
            Top             =   0
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parcela"
            Height          =   195
            Left            =   1680
            TabIndex        =   8
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Aut."
            Height          =   195
            Left            =   960
            TabIndex        =   6
            Top             =   0
            Width           =   630
         End
      End
      Begin VB.CommandButton cmdRede 
         Caption         =   "REDE"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBuscaCCo 
         Height          =   285
         Left            =   4080
         Picture         =   "frmBaixaFinanceirasLote.frx":1084A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
         Height          =   195
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmBaixaFinanceirasLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql_baixas As String

Private Sub cmdProcessar_Click()
        
    If grade.rows < 2 Then
        Exit Sub
    End If
    If txtConta.tag = "" Then
        MsgBox "Escolha a conta para baixa", vbExclamation
        Exit Sub
    End If
        
    processa
        
        
        
End Sub
Sub processa()
    Dim i As Integer
    Dim j As Integer
    Dim aut As String
    Dim valor As String
    Dim parcela As String
    Dim p As financeira.parcela_cartao
    Dim total As Currency
    Dim baixas As Integer
    Dim conta_corrente As String
    Dim Agencia As String
    Dim texto As String
    
    For i = 1 To grade.rows - 1
        grade.row = i
        grade.col = txtColAut.Text
        aut = grade.Text
        grade.col = txtColParc.Text
        parcela = grade.Text
        grade.col = txtColVal.Text
        valor = grade.Text
        
        
        p = financeira.BuscaParcelaPorAutorizacao(aut, parcela, valor)
        If p.autorizacao_fin <> "" Then
            If CCur(p.reman_fin) = CCur(valor) Then
                If Baixar_parcela_fin(p.num_ven, p.parcela_fin, valor, txtConta.tag) Then
                    If texto = "" Then
                        texto = "venda;parcela;vencimento;valor;usuario;data_baixa" & CrLf
                    End If
                    ZebraLinha grade, grade.row, vbGreen
                    texto = texto & p.num_ven & ";" & p.parcela_fin & ";" & p.vencimento_fin
                    texto = texto & ";" & p.valor_fin & ";" & strCod_Usu & ";" & Now
                    
                End If
            End If
        End If
        P_Status "processando linha " & grade.row & " de " & grade.rows - 1
        
    Next
    If texto <> "" Then
        GravaArquivoExterno "", texto, True, "csv"
    Else
        MsgBox "Não foram encontradas parcelas para baixar que coincidam com este arquivo", vbInformation
    End If
    P_Status ""
End Sub

Private Sub cmdRede_Click()
    
    AbrirCSV grade
    buscar_col
    
    
    
End Sub
Sub buscar_col()
    txtColAut.Text = getColByText(grade, "Autorização", 0, True)
    txtColParc.Text = getColByText(grade, "Parcela", 0, False)
    txtColVal.Text = getColByText(grade, "Valor Bruto", 0, True)
    
End Sub

Private Sub cmndbtnBuscaCCo_Click()
     Dim c As ContaCorrente
    c = F_BuscaContaCorrente(txtConta.Text)
    If c.cod_con <> "" Then
        txtConta.Text = c.nome_con
        txtConta.tag = c.cod_con
    End If
End Sub

Private Sub Form_DblClick()
    Form_Resize
End Sub

Private Sub Form_Load()
    resizeForm Me
End Sub

Private Sub Form_Resize()
    PosicionaGrade grade, Me, pcBot
    grade.width = grade.width - 100
End Sub

Private Sub grade_RowColChange()
    lblFoco.caption = "Coluna: " & grade.col & " Linha: " & grade.row
End Sub

Private Sub Label4_Click()
    buscar_col
End Sub
