VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMigraMovInt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inserir Movimento Externo"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtbxDest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton cmndbtnBuscaArq 
         Height          =   315
         Left            =   3360
         Picture         =   "frmMigraMovInt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir Arquivo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   945
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   11520
      Picture         =   "frmMigraMovInt.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnTransf 
      Caption         =   "&Transferir"
      Height          =   615
      Left            =   9720
      Picture         =   "frmMigraMovInt.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   11456
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label lblProtocolo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   17
      Top             =   97
      Width           =   360
   End
   Begin VB.Label lblDataAcao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   16
      Top             =   840
      Width           =   360
   End
   Begin VB.Label lblDataMov 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   15
      Top             =   600
      Width           =   360
   End
   Begin VB.Label lblFunc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   14
      Top             =   360
      Width           =   360
   End
   Begin VB.Label lblCnpj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   13
      Top             =   75
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Funcionário:"
      Height          =   195
      Left            =   4320
      TabIndex        =   12
      Top             =   360
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data Ação:"
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data Movimento:"
      Height          =   195
      Left            =   4320
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Protocolo:"
      Height          =   195
      Left            =   7680
      TabIndex        =   9
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ:"
      Height          =   195
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmMigraMovInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sArquivo As String
Private Sub cmndbtnBuscaArq_Click()
    
    mdiFrm.cmndlg.ShowOpen
    
    sArquivo = mdiFrm.cmndlg.FileName
    txtbxDest.Text = sArquivo
    If sArquivo <> "" Then
        buscaArquivo
    End If
End Sub

Sub buscaArquivo()
    On Error GoTo trata_Erro
    Dim iCount As Integer
    msflxgrd.rows = 1
    Dim sBuff As String
    Dim sCodPro As String
    Dim sSeqIte As String
    Dim sNomePro As String
    Dim sQtdIte As String
    Dim sCodSub As String
    Dim sNomeSub As String
    Dim sCodSec As String
    Dim sNomeSec As String
    Dim sLucroPro As String
    Dim sStPro As String
    Dim sIcmsPro As String
    Dim sUnidadePro As String
    Dim sArq As New Collection
    Dim sCrPro As String
    Dim sPvPro As String
    Dim sCbPro As String
    Dim sCfPro As String
    Dim sCodFor As String
    Dim sNomeFor As String
    
    Open sArquivo For Input As #1
    Do While Not EOF(1)
        Line Input #1, sBuff
        sArq.Add sBuff
    Loop
    
    Close #1
    ajusta_Grid
    
    For iCount = 1 To sArq.Count
        'colocar o esquema
        P_Status "carregando " & iCount & "/" & sArq.Count
        If Left(sArq(iCount), 4) = "1000" Then
            lblcnpj.caption = Mid(sArq(iCount), 5)
        End If
        If Left(sArq(iCount), 4) = "1001" Then
            lblProtocolo.caption = Mid(sArq(iCount), 5)
        End If
        If Left(sArq(iCount), 4) = "1002" Then
            lblDataMov.caption = Mid(sArq(iCount), 5)
        End If
        If Left(sArq(iCount), 4) = "1003" Then
            lblDataAcao.caption = Mid(sArq(iCount), 5)
        End If
        If Left(sArq(iCount), 4) = "1004" Then
            lblFunc.tag = Mid(sArq(iCount), 5)
        End If
        If Left(sArq(iCount), 4) = "1005" Then
            lblFunc.caption = Mid(sArq(iCount), 5)
        End If
        If Left(sArq(iCount), 1) = "2" Then
            sCodPro = Trim(Mid(sArq(iCount), 5, 8))
            sSeqIte = Trim(Mid(sArq(iCount), 13, 3))
            sQtdIte = Trim(Mid(sArq(iCount), 16, 8))
            sNomePro = Trim(Mid(sArq(iCount), 24, 50))
            sCodSub = Trim(Mid(sArq(iCount), 74, 3))
            sNomeSub = Trim(Mid(sArq(iCount), 77, 20))
            sCodSec = Trim(Mid(sArq(iCount), 97, 3))
            sNomeSec = Trim(Mid(sArq(iCount), 100, 20))
            sCrPro = Trim(Mid(sArq(iCount), 120, 8))
            sCfPro = Trim(Mid(sArq(iCount), 128, 8))
            sCbPro = Trim(Mid(sArq(iCount), 136, 8))
            sPvPro = Trim(Mid(sArq(iCount), 144, 8))
            sLucroPro = Trim(Mid(sArq(iCount), 152, 8))
            sSt = Trim(Mid(sArq(iCount), 160, 1))
            sIcmsPro = Trim(Mid(sArq(iCount), 161, 2))
            sUnidadePro = Trim(Mid(sArq(iCount), 163, 3))
            sCodFor = Trim(Mid(sArq(iCount), 166, 5))
            sNomeFor = Trim(Mid(sArq(iCount), 171, 50))
            
            strSQL = sCodPro & Chr(9) & sSeqIte & Chr(9)
            strSQL = strSQL & sQtdIte & Chr(9)
            strSQL = strSQL & sNomePro & Chr(9)
            strSQL = strSQL & sCodSub & Chr(9)
            strSQL = strSQL & sNomeSub & Chr(9)
            strSQL = strSQL & sCodSec & Chr(9)
            strSQL = strSQL & sNomeSec & Chr(9)
            strSQL = strSQL & FormatVal(sCrPro) & Chr(9)
            strSQL = strSQL & FormatVal(sCfPro) & Chr(9)
            strSQL = strSQL & FormatVal(sCbPro) & Chr(9)
            strSQL = strSQL & FormatVal(sPvPro) & Chr(9)
            strSQL = strSQL & FormatVal(sLucroPro) & Chr(9)
            strSQL = strSQL & sSt & Chr(9)
            strSQL = strSQL & sIcmsPro & Chr(9)
            strSQL = strSQL & sUnidadePro & Chr(9)
            strSQL = strSQL & sCodFor & Chr(9)
            strSQL = strSQL & sNomeFor
            If sCodPro <> "" Then
            
                msflxgrd.AddItem strSQL
            End If
        End If
        
    Next
    P_Status ""
    
            Exit Sub
trata_Erro:
P_Status ""
            P_MostraErro
            On Error Resume Next
            Close #1
            
End Sub
Sub ajusta_Grid()
    With msflxgrd
        .Cols = 18
        .rows = 1
        .col = 0: .Text = "Código"
        .col = 1: .Text = "Séq."
        .col = 2: .Text = "Qtd"
        .col = 3: .Text = "Nome"
        .col = 4: .Text = "Cód.Sub."
        .col = 5: .Text = "nome Subseção"
        .col = 6: .Text = "Cód.Seç."
        .col = 7: .Text = "nome Seção"
        .col = 8: .Text = "CR"
        .col = 9: .Text = "CF"
        .col = 10: .Text = "CB"
        .col = 11: .Text = "PV"
        .col = 12: .Text = "Lucro"
        .col = 13: .Text = "St"
        .col = 14: .Text = "ICMS"
        .col = 15: .Text = "Und"
        .col = 16: .Text = "Cód.For"
        .col = 17: .Text = "Fabricante/Fornecedor"
        .ColWidth(3) = 3000
        .ColWidth(5) = 2000
        .ColWidth(4) = 2000
        .ColWidth(1) = 500
        .ColWidth(13) = 800
        .ColWidth(14) = 800
        .ColWidth(2) = 800
        .ColWidth(4) = 800
        
    End With
End Sub
'Sub s()
'     Print #1, "0000"
'        Print #1, "1000" & F_SoNumero(strCNPJ_Emp)
'        Print #1, "1001" & protocolo_mov
'        Print #1, "1002" & data_mov
'        Print #1, "1003" & Now
'        Print #1, "1004" & strCod_Usu
'        Print #1, "1005" & strNome_Usu
'        'Do While Not .EOF
'            Print #1, "200" & .AbsolutePosition & F_EsticaTexto(!cod_pro, 8);
'            Print #1, F_EsticaTexto(!seq_ite, 3);
'            Print #1, F_EsticaTexto(!qtd_ite, 8);
'            Print #1, F_EsticaTexto(!nome_pro, 50);
'            Print #1, F_EsticaTexto(!cod_sub, 3);
'            Print #1, F_EsticaTexto(!nome_sub, 20);
'            Print #1, F_EsticaTexto(!cod_sec, 3);
'            Print #1, F_EsticaTexto(!nome_sec, 20);
'            Print #1, F_EsticaTexto(!cr_pro, 8);
'            Print #1, F_EsticaTexto(!cf_pro, 8);
'            Print #1, F_EsticaTexto(!cb_pro, 8); 'custo bruto
'            Print #1, F_EsticaTexto(!pv_pro, 8); 'preço  venda
'            Print #1, F_EsticaTexto(!lucro_pro, 8);
'            Print #1, F_EsticaTexto(!st_pro, 1);
'            Print #1, F_EsticaTexto(!icms_pro, 2);
'            Print #1, F_EsticaTexto(!unidade_pro, 3)
'
'
'            .MoveNext
'        'Loop
'        Print #1, "9999"
'End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
Function Verifica_Produto_Cadastrado() As Boolean
    Dim i As Integer
    With msflxgrd
        For i = 1 To .rows - 1
            .row = i
            .col = 0
            
            If .Text <> "" Then
            
                If F_BuscaUmCampo("Cod_Pro", "tab_produtos", "cod_pro='" & .Text & "'") = "" Then
                    
                    MsgBox "O produto de código " & .Text & " não se encontra cadastrado. Não é possível prosseguir sem migrar este produto", vbExclamation
                    
                    
                    Verifica_Produto_Cadastrado = False
                    Exit Function
                End If
            End If
        Next
    End With
    Verifica_Produto_Cadastrado = True
End Function
Private Sub cmndbtnTransf_Click()
    Dim strSQL As String
    Dim sSeqMov As String
    Dim iCount As Integer
    Dim cod_pro As String
    Dim qtd_ite As String
    Dim cr_pro As String
    Dim cf_pro As String
    Dim cb_pro As String
    Dim pv_pro As String
    Dim cod_sec As String
    Dim Nome_Sec As String
    Dim nome_pro As String
    Dim cod_sub As String
    Dim nome_sub As String
    Dim lucro_pro As String
    Dim st_pro As String
    Dim unidade_pro As String
    Dim icms_pro As String
    Dim cod_For As String
    Dim nome_for As String
    
    If Not Verifica_Produto_Cadastrado Then
        
        Exit Sub
    End If
    
    
    
    
    If msflxgrd.rows = 1 Then
        MsgBox "Não há produtos para serem importados!", vbExclamation
        Exit Sub
    End If
    
    strSQL = F_BuscaUmCampo("num_mov", "tab_movext", "protocolo_mov='" & lblProtocolo.caption _
        & "' and cnpj_mov='" & lblcnpj.caption & "'")
    If strSQL <> "" Then
        MsgBox "Este movimento já foi Importado!", vbExclamation
        Exit Sub
    End If

    sSeqMov = MovInt.F_GeraNovoMovExt(lblcnpj.caption, lblProtocolo.caption)
    
    If sSeqMov = "" Then
        MsgBox "Não gravado!", vbExclamation
        Exit Sub
        
    End If
    
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0
            cod_pro = .Text
            .col = 2
            qtd_ite = .Text
            .col = 3
            nome_pro = .Text
            .col = 4
            cod_sub = .Text
            .col = 5
            nome_sub = .Text
            .col = 6
            cod_sec = .Text
            .col = 7
            Nome_Sec = .Text
            .col = 8
            cr_pro = .Text
            .col = 9
            cf_pro = .Text
            .col = 10
            cb_pro = .Text
            .col = 11
            pv_pro = .Text
            .col = 13
            st_pro = .Text
            .col = 14
            icms_pro = .Text
            .col = 15
            unidade_pro = .Text
            P_InsereItMovExt sSeqMov, cod_pro, qtd_ite, cr_pro, cf_pro, cb_pro, pv_pro
            
            If F_BuscaUmCampo("cod_sec", "tab_secoes", "cod_sec='" & cod_sec & "'") = "" Then
                secao.P_GravaSecComCodigo cod_sec, Nome_Sec
            End If
            If F_BuscaUmCampo("cod_sub", "tab_subsec", "cod_sub='" & cod_sub & "'") = "" Then
                subsecAO.P_GravaSubSecComCodigo cod_sec, cod_sub, Nome_Sec
            End If
            If cod_For <> "" Then
                If F_BuscaUmCampo("cod_For", "tab_fornecedores", "cod_for='" & cod_For & "'") = "" Then
                    fornecedor.F_GravaFornCurto cod_For, nome_for, nome_for
                End If
            End If
            If F_BuscaUmCampo("cod_pro", "tab_produtos", "cod_pro='" & cod_pro & "'") = "" Then
                Produtos.P_GravaProdutoComCodigo cod_pro, nome_pro, cod_sub, cod_pro, cf_pro, cr_pro, _
                    pv_pro, pv_pro, unidade_pro, cb_pro, cod_For, "", icms_pro, lucro_pro, cb_pro, "0", st_pro
            End If
            strSQL = "update tab_produtos set er_pro=er_pro+" & F_GravaDec(qtd_ite)
            strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
            Conn.Execute strSQL
        Next
        
    End With
    MsgBox "Produtos Migrados!", vbInformation
    
End Sub

