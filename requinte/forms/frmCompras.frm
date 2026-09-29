VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCompras 
   Caption         =   "Compras de Produtos"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   Icon            =   "frmCompras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frControle 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.Frame frCab 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8175
         Begin VB.TextBox txtbxDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4080
            MaxLength       =   12
            TabIndex        =   57
            Tag             =   "!"
            Text            =   "0,00"
            ToolTipText     =   "Total da Entrada Fiscal - Total dos Prod. Fiscal + Total Produtos Real"
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton cmndbtnLocaliza 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4080
            Picture         =   "frmCompras.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtbxEntrada_Not 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            MaxLength       =   10
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxEmissao_Not 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   6000
            MaxLength       =   10
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxValor_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   3000
            MaxLength       =   12
            TabIndex        =   21
            Tag             =   "!"
            Text            =   "0,00"
            ToolTipText     =   "Total da Entrada Fiscal - Total dos Prod. Fiscal + Total Produtos Real"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxBC_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   15
            Text            =   "0,00"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtbxICMS_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   12
            TabIndex        =   17
            Text            =   "0,00"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtbxIPI_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            MaxLength       =   12
            TabIndex        =   19
            Text            =   "0,00"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtbxNOme_Ter 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   4
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox txtbxSerie_Not 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   5520
            MaxLength       =   4
            TabIndex        =   9
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtbxNum_Not 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxDepesas_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   6240
            MaxLength       =   12
            TabIndex        =   23
            Text            =   "0,00"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtbxFrete 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   7200
            MaxLength       =   12
            TabIndex        =   25
            Text            =   "0,00"
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblValorFinal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5160
            TabIndex        =   60
            Tag             =   "!"
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Valor Final:"
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
            Left            =   5160
            TabIndex        =   59
            Tag             =   "!"
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Des&conto:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   4080
            TabIndex        =   58
            Tag             =   "!"
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblUF_Ter 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Left            =   4005
            TabIndex        =   3
            Top             =   120
            Width           =   75
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "E&ntrada:"
            Height          =   195
            Left            =   7080
            TabIndex        =   12
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&Emissão: "
            Height          =   195
            Left            =   6000
            TabIndex        =   10
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "&ValorTotal:"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   3000
            TabIndex        =   20
            Tag             =   "!"
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "&B.Cálc.:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label7 
            Caption         =   "&ICMS:"
            Height          =   255
            Left            =   1080
            TabIndex        =   16
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "I&PI:"
            Height          =   255
            Left            =   2040
            TabIndex        =   18
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "&Fornecedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Série:"
            Height          =   195
            Left            =   5520
            TabIndex        =   8
            Top             =   120
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº: &Doc.:"
            Height          =   195
            Left            =   4440
            TabIndex        =   6
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&Frete:"
            Height          =   195
            Left            =   7200
            TabIndex        =   24
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&Despesas:"
            Height          =   195
            Left            =   6240
            TabIndex        =   22
            Top             =   720
            Width           =   750
         End
      End
      Begin VB.Frame frProd 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   10095
         Begin MSDBGrid.DBGrid dbgrdProd 
            Bindings        =   "frmCompras.frx":0454
            Height          =   3495
            Left            =   1920
            OleObjectBlob   =   "frmCompras.frx":0469
            TabIndex        =   36
            Top             =   120
            Width           =   8175
         End
         Begin VB.Frame frOrd 
            BorderStyle     =   0  'None
            Caption         =   "Ordem:"
            Height          =   320
            Left            =   1920
            TabIndex        =   37
            Top             =   3585
            Width           =   4215
            Begin VB.OptionButton optnSeq 
               Caption         =   "Seqüência"
               Height          =   195
               Left            =   3120
               TabIndex        =   41
               Top             =   120
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optnCod 
               Caption         =   "Có&digo"
               Height          =   195
               Left            =   840
               TabIndex        =   39
               Top             =   120
               Width           =   855
            End
            Begin VB.OptionButton optnDesc 
               Caption         =   "Descrição"
               Height          =   195
               Left            =   1920
               TabIndex        =   40
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ordem:"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   120
               Width           =   510
            End
         End
         Begin VB.Frame frOpt 
            Height          =   3855
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   1815
            Begin VB.CommandButton cmndbtnAddProd 
               Caption         =   "&Adicionar"
               Height          =   495
               Left            =   120
               TabIndex        =   33
               Top             =   1320
               Width           =   1575
            End
            Begin VB.CommandButton cmndbtnExc_Prod 
               Caption         =   "E&xcluir"
               Enabled         =   0   'False
               Height          =   495
               Left            =   120
               TabIndex        =   34
               Top             =   1920
               Width           =   1575
            End
            Begin VB.TextBox txtbxCodBarra 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   50
               TabIndex        =   32
               Top             =   960
               Width           =   1575
            End
            Begin VB.ComboBox cmbxBusca 
               Height          =   315
               ItemData        =   "frmCompras.frx":1882
               Left            =   120
               List            =   "frmCompras.frx":1892
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   360
               Width           =   1575
            End
            Begin VB.CommandButton cmndbtnEditProd 
               Caption         =   "&Editar"
               Enabled         =   0   'False
               Height          =   495
               Left            =   120
               TabIndex        =   35
               Top             =   2520
               Width           =   1575
            End
            Begin VB.Label lblMarca 
               AutoSize        =   -1  'True
               Height          =   195
               Left            =   120
               TabIndex        =   61
               Top             =   2400
               Width           =   45
            End
            Begin VB.Label lblExpr 
               AutoSize        =   -1  'True
               Caption         =   "Cod. Barra:"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Tag             =   "B"
               Top             =   720
               Width           =   795
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "&Buscar Por:"
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   120
               Width           =   825
            End
         End
         Begin VB.Data dtProd 
            Caption         =   "Data2"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   1  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   4200
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1800
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmndbtnEdiCab 
         Caption         =   "Editar Ca&beçalho"
         Height          =   600
         Left            =   8400
         Picture         =   "frmCompras.frx":18BF
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnAbreVenda 
         Caption         =   "C&onfirmar Cabeçalho"
         Height          =   600
         Left            =   8400
         Picture         =   "frmCompras.frx":1A09
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   840
         Width           =   1695
      End
      Begin VB.Frame frResu 
         Caption         =   "Resumo:"
         Height          =   975
         Left            =   120
         TabIndex        =   45
         Top             =   5400
         Width           =   5175
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Gastos:"
            Height          =   195
            Left            =   3600
            TabIndex        =   55
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblGastos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   3600
            TabIndex        =   54
            Tag             =   "!"
            Top             =   600
            Width           =   315
         End
         Begin VB.Label lblIpi_Lan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   2520
            TabIndex        =   53
            Tag             =   "!"
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "IPI Lançado:"
            Height          =   195
            Left            =   2520
            TabIndex        =   52
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lblValoLan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Tag             =   "!"
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total Lançado:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblItens 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   4320
            TabIndex        =   49
            Tag             =   "!"
            Top             =   600
            Width           =   90
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Itens:"
            Height          =   195
            Left            =   4320
            TabIndex        =   48
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Total Produtos:"
            Height          =   195
            Left            =   1320
            TabIndex        =   47
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblTotalProd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1320
            TabIndex        =   46
            Tag             =   "!"
            Top             =   600
            Width           =   315
         End
      End
      Begin VB.Frame frBotoes 
         Height          =   975
         Left            =   5400
         TabIndex        =   42
         Top             =   5400
         Width           =   4815
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   735
            Left            =   2520
            Picture         =   "frmCompras.frx":1B53
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   180
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnSalvar 
            Caption         =   "&Concluir"
            Height          =   735
            Left            =   120
            Picture         =   "frmCompras.frx":1EDD
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   180
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean 'para saber se o frm já foi aberto
Dim iTam As Integer 'tam do form


Sub P_Grava()
    Dim sTipo As String
    Dim strSQL As String
    On Error GoTo trata_erro
    sTipo = UCase(Left(lblTipo.caption, 1))
    If sTipo = "S" Then
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and tipo_not= 'S' and cod_Cli = '" & strCOD_CLI & "'") <> "" Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    Else
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and cod_ter ='" & txtbxNOme_Ter.Tag & "' " _
                & " and tipo_not= 'E' and cod_Cli = '" & strCOD_CLI & "'") <> "" Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "insert into tab_notasfiscais(Cod_Cli,Cod_Ter,Tipo_Not,Num_Not,Serie_Not, "
    strSQL = strSQL & "valor_not,emissao_not,bc_not,bcs_not,icms_Not,ipi_not,despesas_not,"
    strSQL = strSQL & "operacao_not,entrada_not,tipodoc_not,codval_not)Values("
    strSQL = strSQL & "'" & strCOD_CLI & "', "
    strSQL = strSQL & "'" & txtbxNOme_Ter.Tag & "',"
    strSQL = strSQL & "'" & sTipo & "',"
    strSQL = strSQL & "'" & txtbxNum_Not & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxValor_Not.Text) & ","
    strSQL = strSQL & F_GravaData(txtbxEmissao_Not.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxBC_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxBCS_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxICMS_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIPI_Not.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxDepesas_Not.Text) & ","

    strSQL = strSQL & "'" & Left(cmbxTipo_Op.Text, 5) & "', "
    strSQL = strSQL & F_GravaData(txtbxEntrada_Not.Text) & ","
    strSQL = strSQL & "'" & Left(cmbxtipo_Not.Text, 1) & "', "
    strSQL = strSQL & "'" & Left(cmbxCodVal_Not.Text, 1) & "') "
    banco.Execute strSQL
    MsgBox "Nota incluída com sucesso!", vbInformation
    P_Clear
    txtbxNOme_Ter.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Clear()
    txtbxNOme_Ter.Text = ""
    txtbxNOme_Ter.Tag = ""
    txtbxNum_Not.Text = ""
    txtbxSerie_Not.Text = ""
    txtbxEmissao_Not.Text = ""
    cmbxTipo_Op.ListIndex = -1
    txtbxBC_Not.Text = "0,00"
    txtbxBCS_Not.Text = "0,00"
    txtbxICMS_Not.Text = "0,00"
    txtbxIPI_Not.Text = "0,00"
    txtbxDepesas_Not.Text = "0,00"
    txtbxEntrada_Not.Text = ""
    cmbxCodVal_Not.ListIndex = -1
    cmbxtipo_Not.ListIndex = -1
    txtbxValor_Not.Text = "0,00"
End Sub


Sub abreNota()
       Dim strSQL As String
    If txtbxNOme_Ter.Tag = "" Then
        MsgBox "Selecione o Terceiro primeiramente!", vbExclamation
        txtbxNOme_Ter.SetFocus
        Exit Sub
    End If
    If CCur(txtbxBC_Not.Text) > 0 Then
        If CCur(txtbxValor_Not.Text) <> CCur(txtbxBC_Not.Text) + CCur(txtbxDepesas_Not.Text) + CCur(txtbxIPI_Not.Text) + CCur(txtbxFrete.Text) Then
            MsgBox "Os valores não estão corretos!", vbExclamation
            Exit Sub
        End If
    End If
    If txtbxValor_Not.Text = "0,00" Then
        MsgBox "Informe o Valor da Nota Fiscal!", vbExclamation
        txtbxValor_Not.SetFocus
        Exit Sub
    End If
    If F_BuscaUmCampo("num_not", "tab_compras", "cod_for='" & txtbxNOme_Ter.Tag & "' " & _
        " and num_not='" & txtbxNum_Not.Text & "' and serie_not ='" & _
        txtbxSerie_Not.Text & "'") <> "" Then
        MsgBox "Já existe uma nota com este número e esta série para este fornecedor!", vbCritical
        Exit Sub
    End If
    
    
    If MsgBox("Confirma Abertura da Nota?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    frBotoes.Visible = True
    strSQL = "insert into tab_Cmp(Cod_For,Num_not,Serie_not,Valor_not,Bc_Not,"
    strSQL = strSQL & "frete_not,icms_not,ipi_not,despesas_not,operacao_not,emissao_not,"
    strSQL = strSQL & "entrada_not,cod_usu,status_not)Values("
    strSQL = strSQL & "'" & txtbxNOme_Ter.Tag & "', "
    strSQL = strSQL & "'" & txtbxNum_Not.Text & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxValor_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxBC_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxFrete.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxICMS_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIPI_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxDepesas_Not.Text) & ","
    strSQL = strSQL & "'',"
    strSQL = strSQL & F_GravaData(txtbxEmissao_Not.Text) & ","
    strSQL = strSQL & F_GravaData(txtbxEntrada_Not.Text) & ","
    strSQL = strSQL & "'" & strCod_Usu & "','1') " 'em aberto
    banco.Execute strSQL
    Me.Tag = txtbxNum_Not.Text
    lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.Text) + _
                                         CCur(txtbxFrete.Text))
    frProd.Visible = True
    bActive = False
    Form_Activate
    'txtbxCodBarra.SetFocus
End Sub

Private Sub cmbxBusca_Click()
    lblExpr.caption = "&" & cmbxBusca.Text
    
    If cmbxBusca.ListIndex = 0 Then
        lblExpr.Tag = "B"
    Else
        lblExpr.Tag = Left(cmbxBusca.Text, 1)
    End If
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnAbreVenda_Click()
    
    
    Dim strSQL As String
    If Me.Tag <> "" Then
        strSQL = "update tab_cmp set valor_not = " & F_GravaDec(txtbxValor_Not.Text) & ","
        strSQL = strSQL & "frete_not=" & F_GravaDec(txtbxFrete.Text) & ","
        strSQL = strSQL & "despesas_not=" & F_GravaDec(txtbxDepesas_Not.Text)
        strSQL = strSQL & " where num_not='" & txtbxNum_Not.Text & "' "
        strSQL = strSQL & " and cod_for = '" & txtbxNOme_Ter.Tag & "' "
        strSQL = strSQL & " and serie_not ='" & txtbxSerie_Not.Text & "' "
        banco.Execute strSQL
    Else
    
         abreNota
         Me.Tag = txtbxNum_Not.Text
    End If
    
    
    'If IsLoad("frmNotaEntrada") Then
     '   RefreshDataPosition frmNotaEntrada.dtNot
    'End If
    Soma
    
    frProd.Enabled = True
    frCab.Enabled = False
    txtbxValor_Not.Enabled = False
    txtbxIPI_Not.Enabled = False
    txtbxDepesas_Not.Enabled = False
    cmndbtnAbreVenda.Enabled = False
    cmndbtnEdiCab.Enabled = True
    txtbxCodBarra.SetFocus
    
End Sub

'Private Sub cmndbtnAddProd_Click()
'    Dim p As Produto
'    'p = F_BuscaProd_Ent(CCur(txtbxIPI_Not.Text) - CCur(lblIpi_Lan.caption), _
'        CCur(txtbxValor_Not.Text) - CCur(lblValoLan.caption) - CCur(lblGastos.caption))
'
'    'p = F_BuscaProd_Ent(CCur(txtbxIPI_Not.Text) - CCur(lblIpi_Lan.caption), _
'        (CCur(txtbxValor_Not.Text) - CCur(txtbxIPI_Not.Text) - CCur(lblGastos.caption) - CCur(lblTotalProd.caption)))
'
'
'
'    Dim cFator As Currency 'para saber qual a relacao do produto pelo total
'    If p.ipi > (CCur(txtbxIPI_Not.Text) - CCur(lblIpi_Lan.caption)) Then
'        MsgBox "Esse valor para IPI não está correto!", vbExclamation
'        Exit Sub
'    ElseIf Round((p.preco * p.qtd_pro) + CCur(lblValoLan.caption), 2) > CCur(txtbxValor_Not.Text) Then
'        MsgBox "O valor não está correto! tente novamente", vbExclamation
'        Exit Sub
'    ElseIf ((p.preco * p.qtd_pro)) = 0 Then
'        MsgBox "O valor não está correto! tente novamente", vbExclamation
'        Exit Sub
'    End If
'
'
'    cFator = (p.preco * p.qtd_pro) / (CCur(txtbxValor_Not.Text) - CCur(lblGastos.caption))
'
'    P_InsereProd p.cod_pro, p.preco, p.qtd_pro, p.ipi, cFator
'    P_buscaItens
'    soma
'
'End Sub

Sub P_InsereProd(cod_pro As String, unit As Currency, qtd As Currency, _
                ipi As Currency, cFator As Currency)
    Dim strSQL As String
    Dim sSeq As String
    Dim cCusto As Currency
    Dim cGatos As Currency

    sSeq = F_BuscaUmCampo("Max(Seq_ite)", "tab_itcmp", "Num_Not='" & txtbxNum_Not.Text & "'" & _
        " and cod_for='" & txtbxNOme_Ter.Tag & "'")
    If sSeq = "" Then
        sSeq = "1"
    Else
        sSeq = Int(sSeq) + 1
    End If
    cCusto = unit + (CCur(lblGastos.caption) / qtd) + (ipi / qtd)
    cGatos = CCur(lblGastos.caption) * cFator
    
    strSQL = "insert into tab_itcmp(num_not,cod_for,serie_not,cod_pro,"
    strSQL = strSQL & "unit_ite,qtd_ite,ipi_ite,custo_ite,gastos_ite,seq_ite)values("
    strSQL = strSQL & "'" & txtbxNum_Not.Text & "',"
    strSQL = strSQL & "'" & txtbxNOme_Ter.Tag & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(unit) & ","
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(ipi) & ","
    strSQL = strSQL & F_GravaDec(cCusto) & ","
    strSQL = strSQL & F_GravaDec(cGatos) & ","
    
    
    strSQL = strSQL & sSeq & ") "
    banco.Execute strSQL
    
End Sub
Sub P_AlteraProd(cod_pro As String, unit As Currency, qtd As Currency, _
                ipi As Currency, cFator As Currency)
    Dim strSQL As String
    Dim cCusto As Currency
    Dim cGastos As Currency
    Dim iPos As Integer
    cGastos = CCur(lblGastos.caption) * cFator
    cCusto = unit + (cGastos / qtd) + Round((ipi / qtd), 4)
    
    strSQL = "update tab_itcmp set "
    
    
    strSQL = strSQL & "unit_ite=" & F_GravaDec(unit) & ","
    strSQL = strSQL & "qtd_ite=" & F_GravaDec(qtd) & ","
    strSQL = strSQL & "ipi_ite=" & F_GravaDec(ipi) & ","
    strSQL = strSQL & "gastos_ite=" & F_GravaDec(cGastos) & ","
    strSQL = strSQL & "custo_ite =" & F_GravaDec(cCusto) & " "
    strSQL = strSQL & " where  "
    strSQL = strSQL & "num_not='" & txtbxNum_Not.Text & "' "
    strSQL = strSQL & "and cod_for='" & txtbxNOme_Ter.Tag & "' "
    strSQL = strSQL & "and serie_not='" & txtbxSerie_Not.Text & "' "
    strSQL = strSQL & "and cod_pro= '" & cod_pro & "' "
    
    
    banco.Execute strSQL
    iPos = dtProd.Recordset.AbsolutePosition
    P_buscaitens
    dtProd.Recordset.Move iPos
End Sub

Private Sub cmndbtnConfirmar_Click()
    
    If Me.Tag = "" Then
        If txtbxNOme_Ter.Tag = "" Then
            MsgBox "Não há um cliente selecionado!", vbExclamation
            Exit Sub
        End If
        P_Grava
                
    Else
        P_Altera
    End If
End Sub
Sub P_Altera()
    Dim iPos As Integer
    Dim sTipo As String
    Dim strSQL As String
    On Error GoTo trata_erro
    sTipo = UCase(Left(lblTipo.caption, 1))
    If sTipo = "S" Then
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and tipo_not= 'S' and cod_Cli = '" & _
                strCOD_CLI) <> "" And _
                txtbxNum_Not.Text <> txtbxNum_Not.Tag Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    Else
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and cod_ter ='" & txtbxNOme_Ter.Tag & "' " _
                & " and tipo_not= 'E' and cod_Cli = '" & strCOD_CLI & "'") <> "" _
                And txtbxNum_Not.Text <> txtbxNum_Not.Tag Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "update tab_NotasFiscais set "
    
    strSQL = strSQL & "num_not='" & txtbxNum_Not & "',"
    strSQL = strSQL & "serie_not='" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & "valor_Not=" & F_GravaDec(txtbxValor_Not.Text) & ","
    strSQL = strSQL & "emissao_not=" & F_GravaData(txtbxEmissao_Not.Text) & ", "
    strSQL = strSQL & "bc_not=" & F_GravaDec(txtbxBC_Not.Text) & ","
    strSQL = strSQL & "bcs_not=" & F_GravaDec(txtbxBCS_Not.Text) & ","
    strSQL = strSQL & "icms_not=" & F_GravaDec(txtbxICMS_Not.Text) & ","
    strSQL = strSQL & "ipi_Not=" & F_GravaDec(txtbxIPI_Not.Text) & ", "
    strSQL = strSQL & "despesas_not=" & F_GravaDec(txtbxDepesas_Not.Text) & ","

    strSQL = strSQL & "operacao_not='" & Left(cmbxTipo_Op.Text, 5) & "', "
    strSQL = strSQL & "entrada_not=" & F_GravaData(txtbxEntrada_Not.Text) & ","
    strSQL = strSQL & "tipodoc_not='" & Left(cmbxtipo_Not.Text, 1) & "', "
    strSQL = strSQL & "codval_not='" & Left(cmbxCodVal_Not.Text, 1) & "' "
    strSQL = strSQL & " where cod_cli='" & strCOD_CLI & "' "
    strSQL = strSQL & " and num_not ='" & txtbxNum_Not.Tag & "' "
    strSQL = strSQL & " and tipo_not='" & sTipo & "' "
    strSQL = strSQL & " and cod_ter='" & Me.Tag & "' "
    strSQL = strSQL & " and serie_not='" & txtbxSerie_Not.Tag & "'"
    banco.Execute strSQL
    MsgBox "Nota Alterada com sucesso!", vbInformation
    If frmEditaMov.Visible = True Then
        iPos = frmEditaMov.dtNot.Recordset.AbsolutePosition
        frmEditaMov.dtNot.Refresh
        frmEditaMov.dtNot.Recordset.Move iPos
    End If
    Unload Me
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

 Sub Cancelar()
    If MsgBox("Deseja cancelar?", vbYesNo + vbCritical) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmndbtnAddProd_Click()
    Dim p As Produto
    If CCur(lblValoLan.caption) >= CCur(txtbxValor_Not.Text) Then
       
            MsgBox "Os valores estão completos!", vbExclamation
        
            Exit Sub
       
    End If
    
    
    p = F_buscacodbarra_ent(txtbxCodBarra.Text, CCur(txtbxIPI_Not.Text) - CCur(lblIpi_Lan.caption), _
        (CCur(txtbxValor_Not.Text) - CCur(txtbxIPI_Not.Text) - _
        CCur(lblGastos.caption) - CCur(lblTotalProd.caption)), lblExpr.Tag)
    
    If F_BuscaUmCampo("cod_pro", "tab_itcmp", "num_not='" & txtbxNum_Not.Text & _
        "' and serie_not='" & txtbxSerie_Not.Text & "' and cod_for='" & _
            txtbxNOme_Ter.Tag & "' and cod_pro='" & p.cod_pro & "'") <> "" Then
        MsgBox "Este produto já se encontra lançado!", vbExclamation
        Exit Sub
    End If
    
    If p.cod_pro = "" Then
        MsgBox "Produto não encontado!", vbExclamation
        Exit Sub
    End If
    
    
    Dim cFator As Currency 'para saber qual a relacao do produto pelo total
    If p.ipi > (CCur(txtbxIPI_Not.Text) - CCur(lblIpi_Lan.caption)) Then
        MsgBox "Esse valor para IPI não está correto!", vbExclamation
        Exit Sub
    ElseIf Round((p.qtd_pro * p.preco) + CCur(lblValoLan.caption), 2) > CCur(txtbxValor_Not.Text) Then
        MsgBox "O valor não está correto! tente novamente", vbExclamation
        Exit Sub
    ElseIf ((p.qtd_pro * p.preco)) = 0 Then
        MsgBox "O valor não está correto! tente novamente", vbExclamation
        Exit Sub
    End If
    
    
    cFator = (p.qtd_pro * p.preco) / (CCur(txtbxValor_Not.Text) - CCur(lblGastos.caption) - CCur(txtbxIPI_Not.Text))
    
    P_InsereProd p.cod_pro, p.preco, p.qtd_pro, p.ipi, cFator
    P_buscaitens
    Soma
    lblMarca.caption = p.nome_mar
End Sub


Private Sub cmndbtnEdiCab_Click()
    frCab.Enabled = True
    txtbxValor_Not.Enabled = True
    txtbxDepesas_Not.Enabled = True
    txtbxFrete.Enabled = True
    'txtbxNOme_Ter.Enabled = True
    'cmndbtnLocaliza.Enabled = True
    'txtbxNum_Not.Enabled = True
    'txtbxSerie_Not.Enabled = True
    frProd.Enabled = False
    cmndbtnAbreVenda.Visible = True
    cmndbtnAbreVenda.Enabled = True
    cmndbtnEdiCab.Enabled = False
    
    txtbxValor_Not.SetFocus
End Sub

Private Sub cmndbtnEditProd_Click()
        Dim cValor_Max As Currency
        Dim cIpi_Max As Currency
        Dim strSQL As String
        Dim cFator As Currency
        Dim preco As Currency
        Dim qtd_pro As Currency
        With dtProd.Recordset
            cIpi_Max = CCur(txtbxIPI_Not.Text) - CCur(lblIpi_Lan.caption)


    
            'cValor_Max = CCur(txtbxValor_Not.Text) - CCur(lblValoLan.caption) - CCur(lblGastos.caption)
            cValor_Max = CCur(txtbxValor_Not.Text) - CCur(txtbxIPI_Not.Text) - CCur(lblGastos.caption) - CCur(lblTotalProd.caption)
            
            cValor_Max = cValor_Max + (!qtd_ite * !unit_ite)
            cIpi_Max = cIpi_Max + !ipi_ite
            frmValoresEntr.lblCod_Pro.caption = !cod_pro
            frmValoresEntr.lblNome_pro.caption = !nome_pro
            frmValoresEntr.txtbxCusto.Text = FormatVal(!c2_pro)
            frmValoresEntr.txtbxPreco.Text = FormatVal(!p2_pro)
            frmValoresEntr.lblCat.caption = !nome_cat
            
            frmValoresEntr.lblMar.caption = !nome_mar
            frmValoresEntr.txtbxIpi.Visible = True
            frmValoresEntr.lblIPI.Visible = True
            frmValoresEntr.txtbxPreco.Text = "0,00"
            frmValoresEntr.txtbxPreco.Enabled = False
            frmValoresEntr.txtbxPreco.Tag = cValor_Max
            frmValoresEntr.txtbxIpi.Tag = cIpi_Max
            frmValoresEntr.lblRestante.caption = "Restante: " & FormatVal(cValor_Max)
            frmValoresEntr.txtbxIpi.Text = FormatVal(!ipi_ite)
            frmValoresEntr.txtbxIpi.Enabled = False
            frmValoresEntr.txtbxCusto.Text = FormatVal(!custo_ite)
            frmValoresEntr.txtbxQtd.Text = FormatVal(!qtd_ite)
            frmValoresEntr.txtbxQtd.Enabled = True
            frmValoresEntr.txtbxTotal.Text = FormatVal(!qtd_ite * !unit_ite)
            frmValoresEntr.txtbxPreco.Text = FormatVal(CCur(frmValoresEntr.txtbxTotal.Text) / CCur(!qtd_ite), "###,###,##0.0000")
            On Error Resume Next
            frmValoresEntr.Show 1
            If frmValoresEntr.Tag <> "" Then
                If MsgBox("Confirma Alteração do Produto?", vbYesNo + vbQuestion) = vbYes Then
                    preco = CCur(frmValoresEntr.txtbxPreco.Text)
                    qtd_pro = CCur(frmValoresEntr.txtbxQtd.Text)
                    cFator = (preco * qtd_pro) / (CCur(txtbxValor_Not.Text) - CCur(lblGastos.caption) - CCur(txtbxIPI_Not.Text))
                    
                    P_AlteraProd !cod_pro, CCur(frmValoresEntr.txtbxPreco.Text), _
                        CCur(frmValoresEntr.txtbxQtd.Text), _
                        CCur(frmValoresEntr.txtbxIpi.Text), cFator
                        
                End If
                Unload frmValoresEntr
            Else
                Unload frmValoresEntr
            End If
        End With
         
         
         Soma
End Sub

Private Sub cmndbtnExc_Prod_Click()
    Dim strSQL As String
    If MsgBox("Confirma exlusão do item " & dtProd.Recordset!nome_pro & "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from tab_itcmp where num_not='" & txtbxNum_Not.Text & "' "
        strSQL = strSQL & " and cod_for='" & txtbxNOme_Ter.Tag & "' "
        strSQL = strSQL & " and serie_not='" & txtbxSerie_Not.Text & "' "
        strSQL = strSQL & " and cod_pro='" & dtProd.Recordset!cod_pro & "' "
        banco.Execute strSQL
        
        
        
        
        If banco.RecordsAffected > 0 Then
            MsgBox "Item excuído! Fornadas Excluídas", vbInformation
        Else
            MsgBox "Item excuído!", vbInformation
        End If
        
        
    End If
    P_buscaitens
    Soma
End Sub


Private Sub cmndbtnLocaliza_Click()
    frmSelectCli.txtbxNome.Tag = "F"
    frmSelectCli.txtbxNome.Text = txtbxNOme_Ter.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxNOme_Ter.Text = frmSelectCli.dtBusca.Recordset!nome
    txtbxNOme_Ter.Tag = frmSelectCli.dtBusca.Recordset!COD
    Unload frmSelectCli
    txtbxNum_Not.SetFocus
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
    If IsLoad(frmEdiEntr.Name) Then
        frmEdiEntr.dtNot.Refresh
    End If
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL  As String
    Dim cNovoPreco As Currency
    Dim cLucro As Currency
    
    sDif = CCur(lblValoLan.caption) - CCur(txtbxValor_Not.Text)
    If sDif < 0 Then
        sDif = sDif * -1
    End If
    If sDif > 0.4 Then
        MsgBox "Os valores não estão completamente lançados!", vbExclamation
        Exit Sub
    ElseIf lblIpi_Lan.caption <> txtbxIPI_Not.Text Then
        MsgBox "O valor do IPI não está completamente lançado!", vbExclamation
        Exit Sub
    ElseIf dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não é possível concluir uma nota sem itens!", vbExclamation
        Exit Sub
    End If
    
    
    If MsgBox("Confirma Conclusão da Nota 'Fiscal'? ", vbYesNo + vbQuestion) = vbYes Then
        With dtProd.Recordset
            .MoveFirst
            strSQL = "update tab_cmp set status_not = '2' "
            strSQL = strSQL & " where cod_for='" & txtbxNOme_Ter.Tag & "' "
            strSQL = strSQL & " and num_not='" & txtbxNum_Not.Text & "' "
            strSQL = strSQL & " and serie_not ='" & txtbxSerie_Not.Text & "' "
            banco.Execute strSQL
'            Do While Not .EOF
'                strSQL = "update tab_produtos set estoque_pro=estoque_pro+" & f_GravaDec(!qtd_ite)
'
'                If !custo_ite > CCur(F_BuscaUmCampo("Custo_Pro", "tab_produtos", " cod_pro='" & !cod_pro & "'")) Then
'                    strSQL = strSQL & ", custo_pro=" & f_GravaDec(!custo_ite)
'                End If
'                cLucro = CCur(F_BuscaUmCampo("lucro_pro", "tab_produtos", "cod_pro='" & !cod_pro & "'"))
'
'                cNovoPreco = !custo_ite / (100 - (cLucro + crcTxFederativa_Par + crcTxAdm + crcTxMargemSeg_Par)) * 100
'                If cNovoPreco > CCur(F_BuscaUmCampo("Preco_Pro", "tab_produtos", " cod_pro='" & !cod_pro & "'")) Then
'                    strSQL = strSQL & ", preco_pro=" & f_GravaDec(cNovoPreco)
'                End If
'
'                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
'                banco.Execute strSQL
'                .MoveNext
'            Loop
            
        End With
        MsgBox "Nota Fiscal concluída!", vbInformation
        If IsLoad(frmNotaEntrada.Name) Then
            frmNotaEntrada.dtNot.Refresh
        End If
        Unload Me
    End If
    
    
    
End Sub

Private Sub dbgrdProd_DblClick()
    cmndbtnEditProd_Click
End Sub

Private Sub dbgrdProd_GotFocus()
    dbgrdProd.MarqueeStyle = 3
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEditProd_Click
    End If
    If KeyCode = vbKeyI And Shift = 1 Then
        P_InfPro dtProd.Recordset!cod_pro
    End If
End Sub

Private Sub dbgrdProd_LostFocus()
    dbgrdProd.MarqueeStyle = 0
    
End Sub

Private Sub dtProd_Reposition()
    With dtProd.Recordset
        If .RecordCount > 0 Then
            cmndbtnEditProd.Enabled = True
            cmndbtnExc_Prod.Enabled = True
            
        Else
            cmndbtnEditProd.Enabled = False
            cmndbtnExc_Prod.Enabled = False
            
        End If
    End With
End Sub

Sub Soma()
    Dim iPos As Integer
    Dim iCount As Integer
    Dim cValor As Currency
    Dim cIPI As Currency
    Dim cGastos As Currency
    With dtProd.Recordset
    On Error GoTo trata_erro
        If .RecordCount = 0 Then
            lblIpi_Lan.caption = "0,00"
            lblValoLan.caption = "0,00"
            lblTotalProd.caption = "0,00"
            lblItens.caption = "0"
            Exit Sub
        End If
        iPos = .AbsolutePosition
        
        .MoveFirst
        Do While Not .EOF
            cValor = cValor + !Total
            cIPI = cIPI + !ipi_ite
            iCount = iCount + !qtd_ite
            cGastos = cGastos + CCur(!gastos_ite)
            .MoveNext
        Loop
            lblIpi_Lan.caption = FormatVal(cIPI)
            lblValoLan.caption = FormatVal(cValor + cIPI + cGastos)
            lblTotalProd.caption = FormatVal(cValor)
            lblItens.caption = iCount
            'lblItens.caption = .RecordCount
            .MoveFirst
            .Move iPos
    End With
trata_erro:
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    dtProd.Connect = banco.Connect
 
 
    If Me.Tag <> "" Then
                    
        dbgrdProd.Visible = True
        frBotoes.Visible = True
        frOpt.Visible = True
        frOrd.Visible = True
        frResu.Visible = True
      
        dtProd.Connect = banco.Connect
        cmndbtnAbreVenda.Visible = True
        frCab.Enabled = False
        P_buscaitens
        lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.Text) + _
                             CCur(txtbxFrete.Text))
    Else
        dbgrdProd.Visible = False
        frBotoes.Visible = False
        frOpt.Visible = False
        frOrd.Visible = False
        frResu.Visible = False
      
        dtProd.Connect = banco.Connect
        cmndbtnAbreVenda.Visible = True
        frCab.Enabled = True
        frCab.Enabled = True
        txtbxValor_Not.Enabled = True
        txtbxDepesas_Not.Enabled = True
        txtbxFrete.Enabled = True
        txtbxNOme_Ter.Enabled = True
        cmndbtnLocaliza.Enabled = True
        txtbxNum_Not.Enabled = True
        txtbxSerie_Not.Enabled = True
        txtbxEmissao_Not.Enabled = True
        txtbxEntrada_Not.Enabled = True
        frProd.Enabled = False
        cmndbtnAbreVenda.Visible = True
        cmndbtnAbreVenda.Enabled = True
        cmndbtnEdiCab.Enabled = False
        txtbxNOme_Ter.SetFocus
    End If
            
    
    
    Soma
    bActive = True
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If Shift = 2 And KeyCode = vbKeyC Then
        optnC.Value = True
        txtbxCodBarra.SetFocus
    End If
    If Shift = 2 And KeyCode = vbKeyN Then
        optnN.Value = True
        txtbxCodBarra.SetFocus
    End If
    If Shift = 2 And KeyCode = vbKeyB Then
        optnBar.Value = True
        txtbxCodBarra.SetFocus
    End If
End Sub

Sub P_buscaitens()
    Dim strSQL As String
    
    
    strSQL = "select distinct ti.*,(ti.unit_ite*ti.qtd_ite) as total,ti.gastos_ite,"
    
    strSQL = strSQL & "tm.nome_mar,tc.nome_cat, te.c2_pro, te.p2_pro, "
    strSQL = strSQL & " tp.nome_pro from tab_itcmp as ti, tab_produtos as tp, "
    strSQL = strSQL & "tab_marcas as tm, tab_categorias as tc,tab_composicao as te, tab_codbarras as tb "
    strSQL = strSQL & " where ti.num_not='" & txtbxNum_Not.Text & "' "
    strSQL = strSQL & " and te.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar and tp.cod_cat=tc.cod_cat "
    strSQL = strSQL & " and tp.cod_pro=tb.cod_pro "
    strSQL = strSQL & " and ti.cod_for='" & txtbxNOme_Ter.Tag & "' and ti.serie_not='"
    
    strSQL = strSQL & txtbxSerie_Not.Text & "' "
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
    If optnCod.Value = True Then
        strSQL = strSQL & " order by ti.cod_pro"
    
    End If
    If optnDesc.Value = True Then
        strSQL = strSQL & " order by tp.nome_pro"
    
    End If
    If optnSeq.Value = True Then
        strSQL = strSQL & " order by ti.seq_ite"
    
    End If
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dtProd.Connect = banco.Connect
    'P_PreencheOperacoes cmbxTipo_Op, "E"
    cmbxBusca.ListIndex = 0
    bActive = False
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub Option2_Click()
    
End Sub

Private Sub optnBar_Click()
    lblExpr.caption = "C&od. Barra:"
    lblExpr.Tag = "B"
End Sub

Private Sub optnC_Click()
        lblExpr.caption = "C&od. Prod.:"
        lblExpr.Tag = "C"
End Sub



Private Sub optnN_Click()
    lblExpr.caption = "N&ome Prod.:"
    lblExpr.Tag = "N"
End Sub

Private Sub optnCod_Click()
    P_buscaitens
End Sub

Private Sub optnDesc_Click()
    P_buscaitens
End Sub
Private Sub optnSeq_Click()
    P_buscaitens
End Sub

Private Sub txtbxBC_Not_GotFocus()
    P_SelecTexto txtbxBC_Not
End Sub

Private Sub txtbxBC_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxBC_Not_LostFocus()
    P_ValorLostFocus txtbxBC_Not
End Sub


Private Sub txtbxBCS_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub



Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxCodBarra_LostFocus
        cmndbtnAddProd_Click
    End If
End Sub

Private Sub txtbxCodBarra_LostFocus()
    P_AjustaTexto txtbxCodBarra
    If lblExpr.Tag = "C" Then
        txtbxCodBarra.Text = Format(txtbxCodBarra.Text, "0000000#")
    End If
End Sub

Private Sub txtbxDepesas_Not_GotFocus()
    P_SelecTexto txtbxDepesas_Not
End Sub

Private Sub txtbxDepesas_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDepesas_Not_LostFocus()
    P_ValorLostFocus txtbxDepesas_Not
End Sub

Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    P_ValorLostFocus txtbxDesconto
End Sub

Private Sub txtbxEmissao_Not_GotFocus()
    P_SelecTexto txtbxEmissao_Not
End Sub

Private Sub txtbxEmissao_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEmissao_Not_LostFocus()
    P_AjustaData txtbxEmissao_Not
End Sub




Private Sub txtbxEntrada_Not_GotFocus()
    P_SelecTexto txtbxEntrada_Not
End Sub

Private Sub txtbxEntrada_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEntrada_Not_LostFocus()
    On Error GoTo trata_erro
    P_AjustaData txtbxEntrada_Not
    If txtbxEntrada_Not.Text = "" Then Exit Sub
    If CDate(txtbxEntrada_Not.Text) < CDate(txtbxEmissao_Not.Text) Then
        MsgBox "A data de Entrada não pode ser anterior a data de Emissão", vbExclamation
        txtbxEntrada_Not.SetFocus
    End If
trata_erro:
    
End Sub

Private Sub txtbxFrete_GotFocus()
    P_SelecTexto txtbxFrete
End Sub

Private Sub txtbxFrete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxFrete_LostFocus()
    P_ValorLostFocus txtbxFrete
End Sub

Private Sub txtbxICMS_Not_GotFocus()
    P_SelecTexto txtbxICMS_Not
End Sub

Private Sub txtbxICMS_Not_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxICMS_Not_LostFocus()
    P_ValorLostFocus txtbxICMS_Not
End Sub

Private Sub txtbxIPI_Not_GotFocus()
    P_SelecTexto txtbxIPI_Not
End Sub

Private Sub txtbxIPI_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIPI_Not_LostFocus()
    P_ValorLostFocus txtbxIPI_Not
End Sub

Private Sub txtbxNOme_Ter_GotFocus()
    P_SelecTexto txtbxNOme_Ter
End Sub

Private Sub txtbxNOme_Ter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxNum_Not_GotFocus()
    P_SelecTexto txtbxNum_Not
End Sub

Private Sub txtbxNum_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNum_Not_LostFocus()
     txtbxNum_Not.Text = Format(F_SoNumero(txtbxNum_Not.Text), "0#######")
End Sub


Private Sub txtbxSerie_Not_GotFocus()
    P_SelecTexto txtbxSerie_Not
End Sub

Private Sub txtbxSerie_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSerie_Not_LostFocus()
    P_AjustaTexto txtbxSerie_Not
End Sub

Private Sub txtbxValor_Not_GotFocus()
    P_SelecTexto txtbxValor_Not
End Sub

Private Sub txtbxValor_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Not_LostFocus()
    P_ValorLostFocus txtbxValor_Not
End Sub


