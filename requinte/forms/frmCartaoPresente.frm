VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCartaoPresente 
   Caption         =   "Cartão Presente"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
   Icon            =   "frmCartaoPresente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2775
      ScaleWidth      =   11775
      TabIndex        =   30
      Top             =   4440
      Width           =   11775
      Begin VB.TextBox txtParcelas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1680
         TabIndex        =   39
         Text            =   "1"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtDataBase 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   37
         Top             =   600
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade 
         Height          =   1455
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   2566
         _Version        =   393216
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdDistr 
         Caption         =   "Calcular"
         Height          =   465
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtValorAp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   31
         Text            =   "0,00"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas:"
         Height          =   195
         Left            =   1680
         TabIndex        =   40
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2520
         TabIndex        =   38
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagamento a prazo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   2790
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.PictureBox pcPgto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   11775
      TabIndex        =   23
      Top             =   3240
      Width           =   11775
      Begin VB.TextBox txtValorAv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   28
         Text            =   "0,00"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdPag 
         Height          =   465
         Left            =   6480
         Picture         =   "frmCartaoPresente.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtFormasPag 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1680
         TabIndex        =   24
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma"
         Height          =   195
         Left            =   1680
         TabIndex        =   35
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lblPgto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagamento à vista:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIN"
         Height          =   195
         Left            =   3000
         TabIndex        =   26
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1680
      Width           =   11775
   End
   Begin VB.PictureBox pcCartao 
      Appearance      =   0  'Flat
      BackColor       =   &H0073E1F4&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   11745
      TabIndex        =   11
      Top             =   2280
      Width           =   11775
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1680
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtVenc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3120
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lblId 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         TabIndex        =   22
         Top             =   240
         Width           =   165
      End
      Begin VB.Label lblTransacao 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transação"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7080
         TabIndex        =   21
         Top             =   308
         Width           =   765
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Expira em:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   16
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número do Cartão"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   1290
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   12030
      TabIndex        =   0
      Top             =   0
      Width           =   12030
      Begin VB.TextBox txtContato 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         TabIndex        =   6
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtCpf 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton cmdCliente 
         Caption         =   "..."
         Height          =   285
         Left            =   4440
         TabIndex        =   41
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contato:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   5
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   3
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   555
      End
   End
   Begin VB.TextBox txtPresenteado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   4575
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   12030
      TabIndex        =   18
      Top             =   7305
      Width           =   12030
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   855
         Left            =   9600
         Picture         =   "frmCartaoPresente.frx":04AE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdConf 
         Caption         =   "Confirmar"
         Height          =   855
         Left            =   7320
         Picture         =   "frmCartaoPresente.frx":0A3C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outras Informações:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presenteada (o):"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1170
   End
End
Attribute VB_Name = "frmCartaoPresente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCliente_Click()
    Dim c As cliente
    c = F_BuscaCliente(txtNome.Text)
    If c.cod_cli Then
        txtNome.Text = c.nome_cli
        txtNome.tag = c.cod_cli
        txtCPF.Text = c.cpf_cli
    End If
End Sub

Private Sub cmdConf_Click()
    If txtFormasPag.tag = "" And CCur(txtValorAv.Text) > 0 Then
        MsgBox "Informe a forma de pagamento!", vbExclamation
        Exit Sub
    End If
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso: LanCaixa", vbExclamation
        Exit Sub
    End If
    If Not F_senha Then
        Exit Sub
    End If
    If CCur(txtValor.Text) <= 0 Then
        MsgBox "Informe um valor válido!", vbExclamation
        Exit Sub
    End If
        
    If CCur(txtValorAp.Text) > 0 And txtNome.tag = "" Then
        MsgBox "Informe o cliente", vbExclamation
        Exit Sub
    End If
        
    If CCur(txtValor.Text) >= 10000 Then
        MsgBox "O valor está muito alto!", vbExclamation
        Exit Sub
    
    End If
    
    If CCur(txtValor.Text) <> CCur(txtValorAp.Text) + CCur(txtValorAv.Text) Then
        MsgBox "Os valores não batem!", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtNome.Text) = "" Then
        MsgBox "Informe o nome!", vbExclamation
        txtNome.SetFocus
        Exit Sub
    End If
    If txtContato.Text = "" Then
        MsgBox "Informe o contato!", vbExclamation
        txtContato.SetFocus
        Exit Sub
    End If
    If CDate(txtVenc.Text) <= DateAdd("D", 3, CDate(Date)) Then
        MsgBox "A data deve ser igual ou superior a três dias!", vbExclamation
        txtVenc.SetFocus
        Exit Sub
    End If
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
    
    
End Sub
Sub altera()
    
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para essa operação!", vbExclamation
        Exit Sub
    End If
    If F_BuscaUmCampo("status", "cartao_presente", "id='" & Me.tag & "'") <> "1" Then
        MsgBox "Este cartão não pode ser alterado!", vbExclamation
        Exit Sub
    End If
    
    If CDate(txtVenc.Text) <= DateAdd("d", 1, Date) Then
        MsgBox "Informe uma data válida!", vbExclamation
        Exit Sub
    End If
    
    AlteraVencCartaoPresente Me.tag, txtVenc.Text
    MsgBox "Dados alterados", vbInformation
End Sub
Sub grava()
    Dim id As String
    Dim seq As String
    Dim strSQL As String
    Dim seq_doc As String
    Dim id_doc As String
    Dim venc As String
    Dim valor As String
    If Not Cartao_Presente.CartaoPresenteEhValido(txtNumero.Text) Then
        Exit Sub
    End If
    If F_IsCaixa = False Then
        MsgBox "Este computador não é um caixa!", vbExclamation
        Exit Sub
    End If
    
    If txtFormasPag.tag = "" And CCur(txtValorAv.Text) > 0 Then
        MsgBox "Informe a forma de pagamento!", vbExclamation
        Exit Sub
        
    End If
    
    If CCur(txtValorAv.Text) > 0 Then
        If lblTipo.caption = "CRD" Or lblTipo.caption = "CRT" Then
            seq = Cartoes.F_PagamentoCartaoAvulso("Cartão-presente " & txtNumero.Text & "-" & txtNome.Text, txtValorAv.Text)
            If seq = "" Then
                MsgBox "pagamento não efetuado!", vbExclamation
                Exit Sub
            End If
        End If
    End If
    id = Cartao_Presente.GravaCartao_Presente(txtNome.Text, txtCPF.Text, txtContato.Text, txtPresenteado.Text, _
        txtObs.Text, Now, txtVenc.Text, txtNumero.Text, txtValor.Text, txtFormasPag.tag)
    If id <> "" Then
        
        If CCur(txtValorAv.Text) > 0 And lblTipo.caption = "CRD" And lblTipo.caption = "CRT" Then
            seq = F_GravaCai("C", Date, txtValorAv.Text, "CARTÃO-PRESENTE: " & txtNumero.Text, id, lblPgto.tag, strCod_Cai, txtValorAv.Text, "0", strCod_Cai, "1", "000", "S", "0")
        End If
        
        If CCur(txtValorAp.Text) > 0 Then
            seq = F_GravaCai("C", Date, txtValorAp.Text, "CARTÃO-PRESENTE: " & txtNumero.Text, id, "DUP", strCod_Cai, txtValorAp.Text, txtValorAp.Text, strCod_Cai, "1", "000", "S", "0")
        End If
        
        
        strSQL = "Update " & TAB_CAIXA & " set num_ven='" & id & "' where seq_cai='" & seq & "'"
        Conn.Execute strSQL
        
        Dim i As Integer
        If CCur(txtValorAp.Text) > 0 Then
            id_doc = doc.F_GravaDoc("CARTÃO-PRESENTE: " & txtNumero.Text, "CP" & id, txtValorAp.Text, Date, txtNome.tag, "R", id, grade.rows - 1, "V", "C", "")
        
            If id_doc <> "" Then
                
                For i = 1 To grade.rows - 1
                    grade.row = i
                    grade.col = 1
                    venc = grade.Text
                    grade.col = 2
                    valor = grade.Text
                    If doc.F_GravaItDoc(id_doc, valor, venc, venc, CStr(i), a_receber_avulso) = "" Then
                        MsgBox "Erro na gravação da parcela!", vbExclamation
                    End If
                Next
            End If
            APIInserirDoc id_doc, True
        End If
        
        
        Cartao_Presente.ImprimeCartao_Presente id
        MsgBox "Transação realizada!", vbInformation
        
        
        
        
        
        
        
        
        Unload Me
        
        
        
        
    End If
    
    
    
    
    
End Sub

Private Sub cmdDistr_Click()
    On Error GoTo trata_Erro:
    distribui
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub distribui()
    Dim i As Integer
    Dim s As String
    grade.rows = 1
    For i = 1 To Int(txtParcelas.Text)
        grade.AddItem i & Chr(9) & DateAdd("M", i, Date) & Chr(9) & FormatVal(CCur(txtValorAp.Text) / CCur(txtParcelas.Text))
    Next
    ZebraGrid grade, 1
End Sub
Private Sub cmdPag_Click()
    buscaPag
End Sub
Sub buscaPag()
    Dim p As FormasPag.FormasPag
    p = FormasPag.F_BuscaFormasPag("", txtFormasPag.Text, "N", True)
    If p.cod_For <> "" Then
        txtFormasPag.Text = p.nome_for
        txtFormasPag.tag = p.cod_For
        lblPgto.tag = p.tipo_cai
        lblTipo.caption = p.tipo_cai
    End If
    
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub Command1_Click()
    Dim f As FormasPag.FormasPag
    f = FormasPag.F_BuscaFormasPag
    If f.avista_For = "S" Then
        Exit Sub
    End If
    
    txtFormaAp.Text = f.nome_for
    
    
End Sub

Private Sub Form_Activate()
    txtVenc.Text = Format(DateAdd("D", 30, Date), "DD/MM/YYYY")
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        
        cmdConf.Enabled = False
    End If
    
    
    ajustaGrid
    If Me.tag <> "" Then
        carrega
    End If
    
End Sub
Sub ajustaGrid()
    On Error Resume Next
    AjustaGridData grade, 3, 1, 0, "Item, Vencimento, Valor", "", "1000,2000,2000", Nothing, "", True
End Sub
Function Restante() As Currency
    Restante = CCur(txtValor.Text) - txtval
End Function
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from cartao_presente as cp left join tab_formaspag as tf on cp.cod_For=tf.cod_For "
    strSQL = strSQL & " where cp.id='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
           MsgBox "Não encontrado!", vbExclamation
           .Close
        End If
        txtContato.Text = TiraNull(!contato)
        txtCPF.Text = TiraNull(!cpf)
        txtNome.Text = TiraNull(!nome)
        txtValor.Text = FormatVal(!valor)
        txtNumero.Text = !numero
        txtVenc.Text = TiraNull(!vencimento)
        txtObs.Text = TiraNull(!obs)
        txtPresenteado.Text = TiraNull(!presenteado)
        lblTransacao.caption = !id
        txtFormasPag.Text = !nome_for
        txtFormasPag.tag = !cod_For
        
            cmdConf.Enabled = False
            txtNome.Locked = True
            txtContato.Locked = True
            txtValor.Enabled = False
            txtObs.Enabled = False
            txtCPF.Enabled = False
            txtPresenteado.Enabled = False
            txtVenc.Enabled = False
            txtFormasPag.Enabled = False
        
        
        If !status = "1" Then
            txtVenc.Enabled = True
            cmdConf.Enabled = True
        End If
        
        
         .Close
    End With
End Sub

Private Sub Form_Load()
    txtDataBase.Text = Date
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtFormasPag_Click()
    txtFormasPag_GotFocus
End Sub

Private Sub txtFormasPag_GotFocus()
    P_SelecTexto txtFormasPag
End Sub

Private Sub txtFormasPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdPag_Click
    End If
End Sub

Private Sub txtNome_Change()
    txtNome.tag = ""
End Sub

Private Sub txtNome_GotFocus()
    P_SelecTexto txtNome
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNome_LostFocus()
    P_AjustaTexto txtNome
End Sub


Private Sub txtContato_GotFocus()
    P_SelecTexto txtContato
End Sub

Private Sub txtContato_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtContato_LostFocus()
    P_AjustaTexto txtContato, 2
End Sub
Private Sub txtCPF_GotFocus()
    P_SelecTexto txtCPF
End Sub

Private Sub txtCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCPF_LostFocus()
    P_AjustaTexto txtCPF
End Sub


Private Sub txtParcelas_GotFocus()
    P_SelecTexto txtParcelas
End Sub

Private Sub txtParcelas_LostFocus()
    P_ValorLostFocus txtParcelas, "#"
    cmdDistr_Click
End Sub

Private Sub txtPresenteado_GotFocus()
    P_SelecTexto txtPresenteado
End Sub

Private Sub txtPresenteado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtPresenteado_LostFocus()
    P_AjustaTexto txtPresenteado
End Sub

Private Sub txtNumero_GotFocus()
    P_SelecTexto txtNumero
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNumero_LostFocus()
    P_AjustaTexto txtNumero
End Sub

Private Sub txtValor_GotFocus()
    P_SelecTexto txtValor
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtValor_LostFocus()
    P_ValorLostFocus txtValor
End Sub

Private Sub txtValorAp_GotFocus()
    txtValorAp.Text = CCur(txtValor.Text) - CCur(txtValorAv.Text)
    
    P_SelecTexto txtValorAp
End Sub

Private Sub txtValorAp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtValorAp_LostFocus()
    P_ValorLostFocus txtValorAp
    cmdDistr_Click
End Sub
Private Sub txtValorAv_GotFocus()
    txtValorAv.Text = CCur(txtValor.Text) - CCur(txtValorAp.Text)
    
    P_SelecTexto txtValorAv
End Sub

Private Sub txtValorAv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtValorAv_LostFocus()
    P_ValorLostFocus txtValorAv
End Sub



Private Sub txtVenc_GotFocus()
    P_SelecTexto txtVenc
End Sub

Private Sub txtVenc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtVenc_LostFocus()
    P_AjustaData txtVenc
End Sub
Private Sub txtObs_GotFocus()
    P_SelecTexto txtObs
End Sub

Private Sub txtObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtObs_LostFocus()
    P_AjustaTexto txtObs
End Sub

