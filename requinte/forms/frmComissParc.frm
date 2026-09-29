VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmComissParc 
   Caption         =   "Relatório de Comissão Permanente:"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   705
   ClientWidth     =   14835
   Icon            =   "frmComissParc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   14835
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14805
      TabIndex        =   13
      Top             =   0
      Width           =   14835
      Begin VB.CommandButton cmndbtnDoc 
         Caption         =   "&Visualiza Doc"
         Height          =   615
         Left            =   10440
         Picture         =   "frmComissParc.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnVisComissPrev 
         Caption         =   "&Visualiza Acerto"
         Height          =   615
         Left            =   9000
         Picture         =   "frmComissParc.frx":0243
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   7560
         Picture         =   "frmComissParc.frx":033C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   11880
         Picture         =   "frmComissParc.frx":06C6
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   13320
         Picture         =   "frmComissParc.frx":0AD3
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   285
         Left            =   6840
         Picture         =   "frmComissParc.frx":0C3E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14805
      TabIndex        =   0
      Top             =   8325
      Width           =   14835
      Begin VB.CommandButton cmdUltimo 
         Caption         =   "Último Compr."
         Height          =   615
         Left            =   12240
         Picture         =   "frmComissParc.frx":0FC8
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   0
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Já Incluso na Comissão"
         ForeColor       =   &H80000008&
         Height          =   715
         Left            =   8160
         TabIndex        =   28
         Top             =   -20
         Width           =   2775
         Begin VB.Label lblDifTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1440
            TabIndex        =   32
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferença Total:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label lblDifParc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1440
            TabIndex        =   30
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferenças Parc:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.CommandButton cmndbtnCompr 
         Caption         =   "Comprovante"
         Height          =   615
         Left            =   11040
         Picture         =   "frmComissParc.frx":13D5
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnPagarCom 
         Caption         =   "Pagar Comissão"
         Height          =   615
         Left            =   13440
         Picture         =   "frmComissParc.frx":17E2
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnTodos 
         Caption         =   "Todos"
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkAberta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Comissões Não Pagas"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   0
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Selecionado:"
         Height          =   195
         Left            =   5400
         TabIndex        =   24
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblSelec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6960
         TabIndex        =   23
         Top             =   480
         Width           =   390
      End
      Begin VB.Label lblAVencer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   6960
         TabIndex        =   19
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comissões A Vencer:"
         Height          =   195
         Left            =   5400
         TabIndex        =   18
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label lblVencidas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3960
         TabIndex        =   17
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comissões Vencidas:"
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comissão Total:"
         Height          =   195
         Left            =   2280
         TabIndex        =   12
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3960
         TabIndex        =   11
         Top             =   120
         Width           =   390
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   6735
      Left            =   0
      TabIndex        =   21
      Top             =   1200
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   11880
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comissões Encontradas - (Duplo Click para Selecionar)"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      TabIndex        =   22
      Top             =   960
      Width           =   14775
   End
   Begin VB.Menu mnParcela 
      Caption         =   "&Parcela"
      Begin VB.Menu mnMarcar 
         Caption         =   "&Marcar/ Desmarcar"
      End
      Begin VB.Menu mnVisAcerto 
         Caption         =   "&Visualizar Acerto"
      End
      Begin VB.Menu mnVisDoc 
         Caption         =   "&Visualizar Documento"
      End
   End
   Begin VB.Menu mnAção 
      Caption         =   "&Ação"
      Begin VB.Menu mnPagarCom 
         Caption         =   "Pagar Comissões Selecionadas"
      End
      Begin VB.Menu mnImpr 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "frmComissParc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Dim cDifP As Currency
Dim cDifT As Currency

Private Sub cmdUltimo_Click()
    P_ImprUltimoComprPagComiss txtbxNomeCli.tag
End Sub

Private Sub cmndbtnBuscar_Click()
    P_Status "Buscando dados..."
    busca
    P_Status
End Sub
Function select_(sTab As String) As String
    Dim strSQL As String
    
    If sTab = "tab_itdoc" Then
        strSQL = "select concat(parcela_itd,'/',parcelas_doc) as parcela,tc.*,diferenca_ace/parcelas_doc as difpar,"
        strSQL = strSQL & "if(status_com='2',valor_com,0) as comiss_pg,diferenca_ace,"
        strSQL = strSQL & "ti.reman_itd,"
        strSQL = strSQL & "tcli.nome_cli, ta.data_ace,tc.seq_ace,ti.valor_itd,td.num_doc,td.parcelas_doc,td.valor_doc "
        strSQL = strSQL & " from tab_doc as td, tab_comissparc as tc, tab_acertos as ta, tab_clientes as tcli,"
        strSQL = strSQL & sTab & " as ti  "
        strSQL = strSQL & " where ti.seq_doc=td.seq_doc "
        strSQL = strSQL & " and tc.seq_itd=ti.seq_itd "
        strSQL = strSQL & " and ta.cod_cli=tcli.cod_cli"
    Else
        strSQL = "select concat(parc_che,'/',parcelas_doc) as parcela,diferenca_ace,tc.*,comissao_ace,diferenca_ace/parcelas_doc as difpar,"
        strSQL = strSQL & "if(status_com='2',valor_com,0) as comiss_pg,tcli.nome_cli,"
        strSQL = strSQL & "if(dep_che='0000-00-00',valor_che,0) as reman_itd,"
        strSQL = strSQL & " ta.data_ace,tc.seq_ace,ti.valor_che as valor_itd,td.num_doc,td.parcelas_doc,td.valor_doc "
        strSQL = strSQL & " from tab_doc as td, tab_comissparc as tc, tab_acertos as ta,tab_clientes as tcli, "
        strSQL = strSQL & sTab & " as ti  "
        strSQL = strSQL & " where ti.seq_doc=td.seq_doc "
        strSQL = strSQL & " and tc.seq_itd=ti.seq_che "
        strSQL = strSQL & " and ta.cod_cli=tcli.cod_cli"
    End If
    
    
    
    
    strSQL = strSQL & " and tipo_com=" & IIf(sTab = "tab_itdoc", "'1'", "'2'")
    
    
    
    strSQL = strSQL & " and ta.seq_ace=tc.seq_ace "
    If txtbxDataFin.text <> "" Then
        strSQL = strSQL & " and venc_com<=" & F_GravaData(txtbxDataFin.text)
    End If
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and venc_com>=" & F_GravaData(txtbxDataIni.text)
    End If
    If chkAberta.Value Then
        strSQL = strSQL & " and status_com='1' "
    End If
    If txtbxNomeCli.tag <> "" Then
        strSQL = strSQL & " and tc.cod_rep='" & txtbxNomeCli.tag & "' "
    End If
    strSQL = strSQL & " and status_ace=2 "
    strSQL = strSQL & " order by venc_com "
    select_ = strSQL
End Function
Sub busca()
    
    Dim strSQL As String
    Dim sCampo As String
    Dim sTab As String
    Dim rPesq As New ADODB.Recordset
    Dim iCount As Integer
    cDifP = 0
    cDifT = 0
    If txtbxNomeCli.tag = "" Then
        MsgBox "Escolha o Representante!", vbExclamation
        txtbxNomeCli.SetFocus
        Exit Sub
    End If
    msflxgrd.rows = 1
    
    
    
    
    For iCount = 1 To 2
        If iCount = 1 Then
            sTab = "tab_itdoc"
        Else
            sTab = "tab_cheques"
        End If
        strSQL = select_(sTab)
        abreObjetoAdo rPesq, strSQL
        With rPesq
            Do While Not .EOF
                strSQL = .AbsolutePosition + 1 & Chr(9) & !seq_itd & Chr(9)
                strSQL = strSQL & !seq_ace & Chr(9)
                strSQL = strSQL & !data_ace & Chr(9)
                strSQL = strSQL & !venc_com & Chr(9)
                strSQL = strSQL & FormatVal(!valor_com) & Chr(9)
                strSQL = strSQL & FormatVal(!valor_itd) & Chr(9)
                strSQL = strSQL & FormatVal(!reman_itd) & Chr(9)
                
                If !tipo_com = "1" Then
                    strSQL = strSQL & "1. Carnê"
                Else
                    strSQL = strSQL & "2. Cheque"
                End If
                strSQL = strSQL & Chr(9)
                If !status_com = "1" Then
                    strSQL = strSQL & "1. Em Aberto"
                Else
                    strSQL = strSQL & "2. Paga"
                End If
                strSQL = strSQL & Chr(9)
                strSQL = strSQL & !nome_Cli & Chr(9) & !seq_cai & Chr(9)
                strSQL = strSQL & FormatVal(!difpar) & Chr(9)
                strSQL = strSQL & FormatVal(!diferenca_ace)
                cDifT = cDifT + !diferenca_ace
                cDifP = cDifP + !difpar
                msflxgrd.AddItem strSQL
                
                .MoveNext
            Loop
            .Close
        End With
    Next
    lblDifParc.caption = FormatVal(cDifP)
    lblDifTotal.caption = FormatVal(cDifT)
    
    soma
    'dt.RecordSource = strSQL
    'dt.Refresh
    'ajustaMsf
End Sub
Private Sub cmndbtnCliente_Click()
    Dim r As representante.representante
    
    r = F_BuscaRep(txtbxNomeCli.text)
    If r.cod_Rep <> "" Then
        txtbxNomeCli.tag = r.cod_Rep
        txtbxNomeCli.text = r.nome_rep
        
        cmndbtnBuscar_Click
    End If
    
    
End Sub

Private Sub cmndbtnCompr_Click()
    On Error GoTo trata_Erro
    With msflxgrd
        .col = 11
        If .text <> "" Then
             P_ImprimePagamentoComissRepr .text
        End If
    End With
    Exit Sub
trata_Erro:
End Sub

Private Sub cmndbtnDoc_Click()
    msflxgrd.col = 1
    imprimeCarne msflxgrd.text
End Sub

Private Sub cmndbtnImprimir_Click()
    P_MsflxToHtml "Relatório de Comissão Parcelada", msflxgrd, "Representante:" & txtbxNomeCli.text, _
         "Comissão Total " & lblTotal.caption & " -  Vencidas " & lblVencidas.caption & " - A Vencer: " & lblAVencer.caption
    
    
End Sub

Private Sub cmndbtnPagarCom_Click()
    Dim strSQL As String
    Dim iCount As Integer
    Dim sSeq As String
    If Not F_IsCaixa Then
        MsgBox "Este computador não pode efetuar pagamento por que este computador não é um caixa!", vbExclamation
        Exit Sub
    End If
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para fazer lançamento de caixa!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma Pagamento das Comissões?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    sSeq = F_GravaCai("D", Date, lblSelec.caption, "PAG. COM. " & txtbxNomeCli.text, "", "DIN", _
        strCod_Cai, 0, lblSelec.caption, strCod_Cai, strCod_Tur, "000", "N")
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            If .CellBackColor = vbYellow Then
                .col = 1
                strSQL = "update tab_comissparc set status_com=2,"
                strSQL = strSQL & "seq_cai='" & sSeq & "' "
                strSQL = strSQL & " where seq_itd='" & .text & "'"
                Conn.Execute strSQL
            End If
        Next
        busca
    End With
    P_ImprimePagamentoComissRepr sSeq
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub cmndbtnTodos_Click()
    Dim iCount As Integer
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0
            ZebraLinha msflxgrd, .row, vbWhite
        Next
    End With
    If cmndbtnTodos.tag = "" Then
        cmndbtnTodos.tag = "x"
        cmndbtnTodos.caption = "&Nenhum"
        With msflxgrd
            For iCount = 1 To .rows - 1
                .row = iCount
                .col = 0
                ZebraLinha msflxgrd, .row, vbYellow
            Next
        End With
    Else
        cmndbtnTodos.tag = ""
        cmndbtnTodos.caption = "&Todos"
    End If
    soma
End Sub

Private Sub cmndbtnVisComissPrev_Click()
    msflxgrd.col = 2
    P_VisualizaAcerto msflxgrd.text
End Sub

Private Sub Form_Load()
 
    Set ado = New ADODB.Recordset
 
    resizeForm Me
    AjustaGrid
    Me.top = 900
End Sub
Sub AjustaGrid()
    With msflxgrd
        .Cols = 14
        .FixedCols = 0
        .rows = 1
        .col = 0: .text = "..."
        .col = 1: .text = "Seq.Itd."
        .col = 2: .text = "Seq.Ace."
        .col = 3: .text = "Dt.Acerto"
        .col = 4: .text = "Vencimento"
        .col = 5: .text = "Comissão"
        .col = 6: .text = "Vlr Parc."
        .col = 7: .text = "Reman."
        .col = 8: .text = "Tipo"
        .col = 9: .text = "Status"
        .col = 10: .text = "Cliente"
        .col = 11: .text = "Seq.Cai"
        .col = 12: .text = "Dif.Parcela"
        .col = 13: .text = "Dif.Acerto"
        .ColAlignment(8) = 1
        .ColAlignment(9) = 1
        .ColWidth(0) = 500
        .ColWidth(10) = 3500
    End With
End Sub
Sub ajustaMsf()
    Dim iCount As Integer
    msflxgrd.FixedCols = 0
    msflxgrd.Cols = 1
    msflxgrd.rows = 0
    On Error Resume Next

    End Sub
Private Sub lblAp_Click()

End Sub
Sub soma()
    Dim cTotal As Currency
    Dim cVenc As Currency
    Dim iCount As Integer
    Dim cSel As Currency
    Dim iRow As Integer
    
    
    On Error GoTo trata_Erro
    
    With msflxgrd
        
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 9
            
            
            If Left(.text, 1) = "1" Then
                .col = 5
                If CDate(.text) < Date Then
                    .col = 5
                    cVenc = cVenc + CCur(.text)
                End If
                
                
                
            End If
            .col = 5
            cTotal = cTotal + CCur(.text)
                
            .col = 5
            If .CellBackColor = vbYellow Then
                cSel = cSel + CCur(.text)
            End If
            
            
        Next
        
        lblTotal.caption = FormatVal(cTotal)
        lblSelec.caption = FormatVal(cSel)
        lblVencidas.caption = FormatVal(cVenc)
        lblAVencer.caption = FormatVal(cTotal - cVenc)
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGradeWidth msflxgrd, Me
    PosicionaGradeWidth lblCaption, Me
End Sub

Private Sub mnMarcar_Click()
    msflxgrd_DblClick
End Sub

Private Sub msflxgrd_DblClick()
    With msflxgrd
        .col = 9
        If Left(.text, 1) = "1" Then
            If .CellBackColor = vbYellow Then
                ZebraLinha msflxgrd, .row, vbWhite
            Else
                ZebraLinha msflxgrd, .row, vbYellow
            End If
        
            soma
        End If
    End With
End Sub

Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnParcela
    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnAção
    End If
End Sub


Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub
