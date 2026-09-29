VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVendaSup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Venda de Produtos"
   ClientHeight    =   9555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16605
   ControlBox      =   0   'False
   Icon            =   "frmVendaSup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   16605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmr 
      Interval        =   60000
      Left            =   5880
      Top             =   6000
   End
   Begin VB.PictureBox pcGrade 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   6480
      ScaleHeight     =   6855
      ScaleWidth      =   9015
      TabIndex        =   27
      Top             =   120
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid gradeItens 
         Height          =   6255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   11033
         _Version        =   393216
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   0
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox pcTotal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   120
      Picture         =   "frmVendaSup.frx":1CFA
      ScaleHeight     =   2460
      ScaleWidth      =   5760
      TabIndex        =   10
      Top             =   240
      Width           =   5760
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   2040
         TabIndex        =   11
         Top             =   600
         Width           =   1680
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   16605
      TabIndex        =   3
      Top             =   7245
      Width           =   16605
      Begin VB.PictureBox pcProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   6360
         Picture         =   "frmVendaSup.frx":2D104
         ScaleHeight     =   1575
         ScaleWidth      =   7935
         TabIndex        =   25
         Top             =   0
         Width           =   7935
         Begin VB.TextBox txtProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H00916024&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   720
            Left            =   360
            TabIndex        =   26
            Top             =   360
            Width           =   7215
         End
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   1  'Right Justify
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
         Left            =   5880
         TabIndex        =   24
         Text            =   "1,000"
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox pcLogo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   1680
         Left            =   0
         Picture         =   "frmVendaSup.frx":52B0E
         ScaleHeight     =   1680
         ScaleWidth      =   5520
         TabIndex        =   4
         Top             =   0
         Width           =   5520
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade (F3)"
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.PictureBox pcLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3825
      Left            =   0
      ScaleHeight     =   3825
      ScaleWidth      =   5175
      TabIndex        =   2
      Top             =   3015
      Width           =   5175
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   120
         Picture         =   "frmVendaSup.frx":5B52C
         ScaleHeight     =   585
         ScaleWidth      =   360
         TabIndex        =   22
         Top             =   3120
         Width           =   360
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   120
         Picture         =   "frmVendaSup.frx":5C066
         ScaleHeight     =   585
         ScaleWidth      =   360
         TabIndex        =   20
         Top             =   720
         Width           =   360
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   120
         Picture         =   "frmVendaSup.frx":5CBA0
         ScaleHeight     =   585
         ScaleWidth      =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   360
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   120
         Picture         =   "frmVendaSup.frx":5D6DA
         ScaleHeight     =   585
         ScaleWidth      =   360
         TabIndex        =   16
         Top             =   1920
         Width           =   360
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   120
         Picture         =   "frmVendaSup.frx":5E214
         ScaleHeight     =   585
         ScaleWidth      =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   360
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   120
         Picture         =   "frmVendaSup.frx":5ED4E
         ScaleHeight     =   585
         ScaleWidth      =   360
         TabIndex        =   12
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Sair (F12)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   23
         Top             =   3285
         Width           =   675
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cancelar Item (F4)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   21
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Concluir (F7)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   19
         Top             =   2685
         Width           =   885
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cancelar Venda (F6)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   17
         Top             =   2085
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "NFE/NFC-E (F5)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   15
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cliente (F2)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   16605
      TabIndex        =   1
      Top             =   0
      Width           =   16605
   End
   Begin VB.PictureBox pcBotControles 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   16605
      TabIndex        =   0
      Top             =   8940
      Width           =   16605
      Begin VB.PictureBox pcHora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   14520
         Picture         =   "frmVendaSup.frx":5F888
         ScaleHeight     =   615
         ScaleWidth      =   1815
         TabIndex        =   31
         Top             =   0
         Width           =   1815
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            Caption         =   "0"
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
            Left            =   720
            TabIndex        =   33
            Top             =   0
            Width           =   120
         End
         Begin VB.Label lblHora 
            AutoSize        =   -1  'True
            Caption         =   "0"
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
            Left            =   720
            TabIndex        =   32
            Top             =   240
            Width           =   120
         End
      End
      Begin VB.PictureBox pcAtendente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         Picture         =   "frmVendaSup.frx":5FDB0
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   5
         Top             =   60
         Width           =   615
      End
      Begin VB.Label lblTurno 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         TabIndex        =   37
         Top             =   0
         Width           =   120
      End
      Begin VB.Label lblCodCai 
         AutoSize        =   -1  'True
         Caption         =   "00"
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
         Left            =   1200
         TabIndex        =   36
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Turno:"
         Height          =   195
         Left            =   1800
         TabIndex        =   35
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Caixa:"
         Height          =   195
         Left            =   720
         TabIndex        =   34
         Top             =   0
         Width           =   435
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         Caption         =   "Venda a Consumidor"
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
         Left            =   8160
         TabIndex        =   30
         Tag             =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   6600
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblNumVen 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   4680
         TabIndex        =   8
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAtendente 
         AutoSize        =   -1  'True
         Caption         =   "Atendente"
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
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmVendaSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prodSelecionado As Produto
Dim pa As produtoAtacado
Dim adoItens As ADODB.Recordset
Dim bActive As Boolean
Dim tipo As tipo_venda
Private Sub Form_Activate()
    Form_Resize
    If bActive Then
        Exit Sub
    End If
    If Me.tag <> "" Then
        CarregarDados
        
    
    End If
    
    posicionaADireita pcHora, pcBot
End Sub
Sub CarregarDados()
    Carregar_Venda
    carregar_Itens
    AtualizaDadosVenda
    If tipo.codigo = "D" Then
        Me.BackColor = vbGreen
    End If
    
    On Error Resume Next
    txtProduto.SetFocus
End Sub
Sub BuscaPic()
    On Error Resume Next
    pcLogo.Picture = LoadPicture(App.Path & "\pic\logo.jpg")
    lblAtendente.caption = strNome_Usu
    lblCodCai.caption = strCod_Cai
    lblTurno.caption = strCod_Tur
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
    ElseIf KeyCode = vbKeyF2 Then
        InformaCli
    ElseIf KeyCode = vbKeyF3 Then
       txtProduto.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
        Excluir
    ElseIf KeyCode = vbKeyF5 Then
        frmEdiNFe.Show 1
    ElseIf KeyCode = vbKeyF6 Then
        CancelarVendaAtual
    ElseIf KeyCode = vbKeyF7 Then
        conclui
    ElseIf KeyCode = vbKeyF11 Then
        If gradeItens.rows > 1 Then
            gradeItens.TopRow = 1
        End If
    ElseIf KeyCode = vbKeyF12 Then
        sair
    End If
End Sub
Sub CancelarVendaAtual()
    If strCancelaVenda_Usu = "S" Then
        If Not F_senha Then
            Exit Sub
        End If
    Else
        usu = F_SenhaCodBar
        If F_BuscaUmCampo("cancelarvenda_usu", TAB_DEFUSU, "cod_usu='" & usu & "'") <> "S" Then
            MsgBox "Não autenticado!", vbExclamation
            Exit Sub
        End If
        
    End If
    
    If MsgBox("Confirma cancelamento da venda?", vbYesNo + vbCritical) = vbYes Then
    
    
        If CancelaItem(Me.tag, "todos") Then
            strSQL = "delete from " & TAB_VENDAS & " where num_ven='" & Me.tag & "'"
            conn.Execute strSQL
            Me.tag = F_GeraNovaVenda("000000", strFormaPagamentoPadrao, "")
            Form_Load
            Form_Activate
        End If
    End If
    
    
    
    
End Sub
Sub conclui()
    Dim t As String
    Dim cpf As String
    t = F_Troco(Me.tag)

    If t = "" Then
        
    Else
        P_DeletaVendaAberta Me.tag
        cpf = InformaCPFClienteVenda
    
    
        Me.tag = F_GeraNovaVenda("00000", "01", cpf)
        
    End If
End Sub
Sub AtualizaDadosVenda()
    
    P_AtualizaVendaCliente Me.tag, lblNomeCli.tag, lblNomeCli.caption, "", lblCPF.caption
    P_AtualizaValorTotalVenda Me.tag, lblTotal.caption, "0"
    
End Sub
Sub InformaCli()
    Dim cpf As String
    Dim nome As String
    Dim c As ClienteVenda
    cpf = lblCPF.caption
    nome = lblNomeCli.caption
    
    If cpf = "" Or cpf = "0" Then
    End If
    
    
    c = Vendas.BuscaClienteRapidoVenda(lblNomeCli.tag, cpf, nome)
    If c.cpfcnpj <> "" And c.cpfcnpj <> "0" Then
        lblCPF.caption = c.cpfcnpj
        
        lblCPF.Visible = True
        
        
        
    Else
        lblCPF.Visible = False
    End If
    If c.nomecli_ven <> "" Then
        
        lblNomeCli.caption = c.nomecli_ven
        lblNomeCli.tag = c.cod_cli
        lblNomeCli.Visible = True
    Else
        lblNomeCli.caption = "Venda a Consumidor"
        lblNomeCli.Visible = False
    End If
    If c.cod_cli <> "" Then
        lblNomeCli.tag = c.cod_cli
    End If
    
       
    
    AtualizaDadosVenda
End Sub
Sub Excluir(Optional item As String = "")
    Dim usu As String
    Dim prod As String
    Dim qtd As String
    Dim seq As String
    
    If gradeItens.row < 1 Then
        Exit Sub
    End If
    
    If item = "" Then
        item = gradeItens.row
    End If
    gradeItens.col = 0
    seq = gradeItens.Text
    gradeItens.col = 3
    qtd = gradeItens.Text
    gradeItens.col = 6
    prod = gradeItens.Text
    'gradeItens.Row
    
    If MsgBox("Deseja excluir o item " & prod & " com a quantidade " & qtd & "?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    
    
    If strCancelarItem_Usu = "S" Then
        If Not F_senha Then
            Exit Sub
        End If
    Else
        usu = F_SenhaCodBar
        If F_BuscaUmCampo("cancelaritem_usu", TAB_DEFUSU, "cod_usu='" & usu & "'") <> "S" Then
            MsgBox "Não autenticado!", vbExclamation
            Exit Sub
        End If
        
    End If
    
    Vendas.CancelaItem Me.tag, seq
    carregar_Itens
    
End Sub
Sub sair()
    Dim strSQL As String
    If CCur(lblTotal.caption) > 0 Then
        If MsgBox("Tem certeza que deseja abandonar a venda?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
    Else
        If gradeItens.rows <= 1 Then
            strSQL = "delete from " & TAB_VENDAS & " where num_ven='" & Me.tag & "'"
            conn.Execute strSQL
        End If
        
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    ajustaGrid
    BuscaPic
    Set adoItens = New ADODB.Recordset
    tmr_Timer
End Sub
Sub Carregar_Venda()
    Dim strSQL As String
    Dim ado As New ADODB.Recordset
    On Error GoTo trata_erro
    strSQL = "select tv.*,tc.cpf_cli,tc.cnpj_cli from " & TAB_VENDAS & " as tv left join " & TAB_CLIENTES & " as tc"
    strSQL = strSQL & " on tv.cod_cli=tc.cod_cli "
    
    strSQL = strSQL & " where num_ven='" & Me.tag & "' "
    
    
    abreObjetoAdo ado, strSQL
    
    lblNumVen.caption = Me.tag
    lblCPF.caption = Format(ado!cpfcli_ven)
    If lblCPF.caption <> "" Then
        lblCPF.Visible = True
    End If
    
    lblNomeCli.caption = Format(ado!nomecli_ven)
    lblNomeCli.tag = ado!cod_cli
    
    tipo.codigo = ado!tipo_ven
    
    P_GravaVendaAberta Me.tag
    Exit Sub
trata_erro:
    P_MostraErro
    
End Sub
Sub carregar_Itens()
    Dim cTotal As Currency
    
    Dim cDesconto As Currency
    
    Dim strSQL As String
    gradeItens.Visible = False
    
    strSQL = "select ti.*,tp.nome_pro,tp.ean_pro from " & TAB_ITVENDAS & " as ti inner join " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " on ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " where num_ven='" & Me.tag & "'"
    strSQL = strSQL & " order by seq_ite "
    
    
    abreObjetoAdo adoItens, strSQL
    
    
    'ZebraGrid grid, 1
    gradeItens.rows = 1
    gradeItens.GridLines = flexGridNone
'    ZebraGrid gradeItens, 1, vbGrayed
    
    With adoItens
        Do While Not .EOF
            strSQL = !seq_ite & Chr(9)
            strSQL = strSQL & .AbsolutePosition & Chr(9) & !ean_ite & Chr(9)
            strSQL = strSQL & FormatVal(!qtd_ite, "###,##0.000") & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite, "###,##0.00") & Chr(9)
            strSQL = strSQL & FormatVal(!final_ite, "###,##0.00") & Chr(9)
            strSQL = strSQL & !desc_ite & Chr(9) & !atacado_ite
            gradeItens.AddItem strSQL
            
            If !atacado_ite = "S" Then
                ZebraLinha gradeItens, gradeItens.rows - 1, 16700080, 0, True
            
            
            
            Else
                If .AbsolutePosition Mod 2 = 0 Then
                    ZebraLinha gradeItens, gradeItens.rows - 1, &H8000000F
                End If
            End If
            cTotal = cTotal + !final_ite
            .MoveNext
        Loop
    End With
    lblTotal.caption = FormatVal(cTotal)
    
    gradeItens.row = gradeItens.rows - 1
    If gradeItens.rows > 1 Then
        gradeItens.TopRow = gradeItens.rows - 1
    End If
    gradeItens.GridLines = flexGridNone
    gradeItens.GridLinesFixed = flexGridNone
    gradeItens.GridColor = vbWhite
    gradeItens.Visible = True
    gradeItens.ColWidth(6) = gradeItens.width - gradeItens.ColPos(6)
    
    
End Sub
Sub ajustaGrid()
    With gradeItens
        .FixedCols = 1
        .FixedRows = 1
        .Cols = 8
        .rows = 1
        .GridLines = flexGridNone
        .FillStyle = flexFillSingle
        
        '.Row = 1
        .col = 0: .CellFontBold = True:    .Text = "seq_ite": .CellAlignment = 3
        .col = 1: .CellFontBold = True: .Text = "*": .CellAlignment = 3
        .col = 2: .CellFontBold = True: .Text = "EAN": .CellAlignment = 3
        .col = 3: .CellFontBold = True: .Text = "Qtd": .CellAlignment = 3
        .col = 4: .CellFontBold = True: .Text = "Unitário": .CellAlignment = 3
        .col = 5: .CellFontBold = True: .Text = "Total": .CellAlignment = 3
        .col = 6: .CellFontBold = True: .Text = "Descrição": .CellAlignment = 3
        .col = 7: .CellFontBold = True: .Text = "A": .CellAlignment = 3
        .ColWidth(0) = 0
        .ColWidth(1) = 500
        .ColWidth(2) = 2000
        .ColWidth(6) = 5000
        .ColWidth(7) = 500
        
        .ColAlignment(2) = 3
        .ColAlignment(1) = 3
        .ColAlignment(3) = 3
        .ColAlignment(4) = 3
        .ColAlignment(5) = 3
        .ColAlignment(6) = 2
        .ScrollBars = flexScrollBarVertical
    End With

End Sub
Sub ajustaControles()
    On Error Resume Next
    pcGrade.width = Me.width - pcGrade.Left - 300
    pcGrade.Height = pcBot.Top - pcGrade.Top
    gradeItens.Height = pcGrade.Height - gradeItens.Top - 200
    gradeItens.width = pcGrade.width - gradeItens.Left - 200
    'pcGrade.ColWidth(pcGrade.Cols - 1) = Me.width - pcGrade.width - (pcGrade.ColPos(pcGrade.Cols - 1) - pcGrade.Left)
     
'    posicionaADireita pcGrade, Me

End Sub

Private Sub Form_Resize()
    ajustaControles
End Sub



Private Sub lblCodCai_Click()
    If strFechacai_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    frmFechaCaixa.Show 1
End Sub

Private Sub tmr_Timer()
    lblData.caption = Date
    lblHora.caption = Format(Time, "HH:MM")
    
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        gradeItens.SetFocus
    End If
End Sub
Sub InsereProduto()
    If CCur(txtQtd.Text) <= "0" Then
        MsgBox "Quantidade inválida!", vbExclamation
        Exit Sub
    End If
    If prodSelecionado.pv_pro <= 0 Then
        MsgBox "Produto com preço zerado!", vbExclamation
        Exit Sub
    End If
    If prodSelecionado.balanca_pro <> "S" Then
        If CCur(txtQtd.Text) <> CInt(txtQtd.Text) Then
            MsgBox "Produto não autorizado à venda com quantidade não inteira!", vbExclamation
            Exit Sub
        End If
    End If
    
    If prodSelecionado.cod_pro = "" Then
        Exit Sub
    End If
    
    If Vendas.F_AdicionaItem(Me.tag, prodSelecionado, txtQtd.Text, "0") Then
        txtProduto.Text = "" 'txtQtd.Text & "*" & prodSelecionado.ean_pro
        

    End If
    
    
    
    
    carregar_Itens
    txtProduto.SetFocus
    P_SelecTexto txtProduto
    
End Sub
Sub buscaProd()
    Dim c As New Collection
    Dim i As Integer
    Dim tag As String
    Dim qtd As String
    If txtQtd.Text = "" Then
        txtQtd.Text = "1"
    End If
    
    On Error GoTo trata_erro
    tag = txtProduto.Text
    i = InStr(1, txtProduto.Text, "*")
    
    If i > 0 Then
        qtd = Trim(Left(txtProduto.Text, i - 1))
        
        If CCur(qtd) <= 0 Then
            Exit Sub
        End If
        If txtQtd.Text > 1000 Then
            If MsgBox("Tem certeza que deseja adicionar " & qtd & " itens do produto?", vbYesNo + vbCritical) = vbNo Then
                Exit Sub
            End If
        End If
        
        
        txtQtd.Text = CCur(qtd)
        
        tag = Trim(Mid(txtProduto, i + 1))
    Else
        txtQtd.Text = "1"
    End If
    
    
    
    
    
    
    
    prodSelecionado = Produtos.F_BuscaProdSaida(tag, txtQtd.Text)
    
    If prodSelecionado.cod_pro = "" Then
        Exit Sub
    End If
    
    If prodSelecionado.qtd <> "" Then
        txtQtd.Text = prodSelecionado.qtd
    Else
        prodSelecionado.qtd = "0"
        txtQtd.Text = "1"
    End If
    
    
    
    If prodSelecionado.qtd > 1 Then
        
        pa = buscaProduto_atacado(prodSelecionado.qtd, prodSelecionado.cod_pro)
                
        If prodSelecionado.qtd = pa.qtd Then
            prodSelecionado.Atacado = "S"
            prodSelecionado.pv_pro = pa.Preco
            
        Else
            prodSelecionado.Atacado = "N"
        End If
    
    End If
    
    
    InsereProduto
    
    
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
