VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPerdas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Perdas:"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmPerdas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9060
   Begin VB.Frame frControle 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Frame frCab 
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4455
         Begin VB.TextBox txtbxNome_ter 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   3
            Top             =   360
            Width           =   3975
         End
         Begin VB.CommandButton cmndbtnLocaliza 
            Height          =   285
            Left            =   4080
            Picture         =   "frmPerdas.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblCnpj 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   420
         End
         Begin VB.Label Label12 
            Caption         =   "&Fornecedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   855
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
            TabIndex        =   18
            Top             =   120
            Width           =   75
         End
      End
      Begin VB.Frame frOpt 
         Caption         =   "Produtos:"
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   8775
         Begin VB.TextBox txtbxCusto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            MaxLength       =   50
            TabIndex        =   15
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3240
            MaxLength       =   50
            TabIndex        =   13
            Text            =   "0,0"
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmndbtnAddProd 
            Caption         =   "&Adicionar"
            Height          =   495
            Left            =   6120
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmndbtnExc_Prod 
            Caption         =   "E&xcluir"
            Height          =   495
            Left            =   7440
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtbxCodBarra 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   8
            Top             =   480
            Width           =   3015
         End
         Begin VB.OptionButton optnC 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   255
         End
         Begin VB.OptionButton optnN 
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optnBar 
            Caption         =   "&B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Custo:"
            Height          =   195
            Left            =   4200
            TabIndex        =   14
            Tag             =   "B"
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Quantidade:"
            Height          =   195
            Left            =   3240
            TabIndex        =   12
            Tag             =   "B"
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lblMarca 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   2400
            Width           =   45
         End
         Begin VB.Label lblExpr 
            AutoSize        =   -1  'True
            Caption         =   "&C&ód. Barra:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Tag             =   "B"
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame frBotoes 
         Height          =   975
         Left            =   4800
         TabIndex        =   19
         Top             =   120
         Width           =   4095
         Begin VB.CommandButton cmndbtnCancelar 
            Caption         =   "Cancelar Perda"
            Height          =   735
            Left            =   1440
            Picture         =   "frmPerdas.frx":058C
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   180
            Width           =   1215
         End
         Begin VB.CommandButton cmndbtnSalvar 
            Caption         =   "&Concluir"
            Height          =   735
            Left            =   120
            Picture         =   "frmPerdas.frx":06D6
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   180
            Width           =   1215
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   735
            Left            =   2760
            Picture         =   "frmPerdas.frx":0820
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.Frame frOrd 
         Caption         =   "Ordem:"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   5760
         Width           =   8775
         Begin VB.OptionButton optnSeq 
            Caption         =   "Seqüência"
            Height          =   195
            Left            =   960
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optnCod 
            Caption         =   "Códi&go"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optnDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   2160
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6240
            TabIndex        =   29
            Top             =   120
            Width           =   645
         End
         Begin VB.Label lblTotal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   7080
            TabIndex        =   28
            Top             =   120
            Width           =   465
         End
      End
      Begin MSAdodcLib.Adodc adc 
         Height          =   375
         Left            =   1320
         Top             =   8040
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dtgrd 
         Bindings        =   "frmPerdas.frx":096A
         Height          =   3375
         Left            =   120
         TabIndex        =   30
         Top             =   2280
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPerdas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Sub soma()
    Dim cTotal As Currency
    With dtProd.Recordset
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            cTotal = cTotal + !qtd_itp * !unit_itp
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    lblTotal.caption = FormatVal(cTotal)
End Sub
Private Sub cmndbtnAddProd_Click()
    Dim strSQL As String
    Dim sItem As String
    
    If Me.tag = "" Then
        MsgBox "Não existe uma perda gerada!", vbExclamation
        txtbxNome_ter.SetFocus
        Exit Sub
    End If
    If txtbxCodBarra.tag = "" Then
        MsgBox "Escolha o produto!", vbExclamation
        Exit Sub
    End If
    If F_BuscaUmCampo("cod_pro", "tab_itperdas", "seq_per='" & Me.tag & "' and cod_pro='" & _
        txtbxCodBarra.tag & "'") <> "" Then
        MsgBox "Este item já se encontra lançado!", vbExclamation
        Exit Sub
    End If
    
    strSQL = "insert into tab_itperdas(seq_per,cod_pro,qtd_itp,item_itp,unit_itp)values("
    strSQL = strSQL & "'" & Me.tag & "',"
    strSQL = strSQL & "'" & txtbxCodBarra.tag & "',"
    strSQL = strSQL & F_GravaDec(txtbxQtd.Text) & ","
    sItem = F_BuscaUmCampo("max(item_itp)", "tab_itperdas", "seq_per='" & Me.tag & "'")
    If sItem = "" Then
        sItem = "1"
    Else
        sItem = Int(sItem + 1)
    End If
    strSQL = strSQL & sItem & ","
    strSQL = strSQL & F_GravaDec(txtbxCusto.Text) & ") "
    
    conn.Execute strSQL
    busca
End Sub

Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    On Error Resume Next
    If Me.tag = "" Then
        Unload Me
        Exit Sub
    End If
    If MsgBox("Tem certeza que deseja cancelar a perda " & Me.tag & "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from tab_perdas where seq_per='" & Me.tag & "' "
        conn.Execute strSQL
        strSQL = "delete from tab_itperdas where seq_per='" & Me.tag & "' "
        conn.Execute strSQL
    End If
    MsgBox "Perda excluída!", vbExclamation
    If IsLoad("frmEdiPerdas") Then
        frmEdiPerdas.adc.Refresh
    End If
    Unload Me
End Sub

Private Sub cmndbtnEditProd_Click()

End Sub

Private Sub cmndbtnExc_Prod_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    With dtProd.Recordset
    If .RecordCount = 0 Then
        MsgBox "Não há o que ser excluído!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma exclusão do lançamento?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from tab_itperdas where cod_pro='" & !cod_pro & "' "
        strSQL = strSQL & " and seq_per='" & Me.tag & "' "
        conn.Execute strSQL
        busca
    End If
        
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxNome_ter.tag = "" Then
        MsgBox "Informe o fornecedor!", vbExclamation
        Exit Sub
    End If
    If Me.tag = "" Then
        
        If MsgBox("Confirma abertura da Perda?", vbYesNo + vbQuestion) = vbYes Then
        
            'txtbxNome_ter.Text = frmSelectCli.dtBusca.Recordset!nome
            'txtbxNome_ter.Tag = frmSelectCli.dtBusca.Recordset!cod
            If txtbxNome_ter.tag = "" Then
                Exit Sub
            End If
        
            Me.tag = F_GeraNovaPerda(txtbxNome_ter.tag, strCod_Usu)
            If Me.tag = "" Then
                MsgBox "Perda inválida!", vbExclamation
            End If
            Me.caption = "Perda nº " & Me.tag & "-" & txtbxNome_ter.Text
            cmndbtnSalvar.caption = "&Concluir Lançamento"
            
        End If
    Else
        If MsgBox("Confirma conclusão dos lançamentos?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        If txtbxNome_ter.tag = "" Then
            MsgBox "Escolha o Fornecedor!", vbExclamation
            Exit Sub
        End If
        If dtProd.Recordset.RecordCount = 0 Then
            MsgBox "Não há itens para conclusão!", vbExclamation
            Exit Sub
        End If
        Baixa
        MsgBox "Lançamento de Perdas concluído!", vbInformation
        imprimePerda Me.tag
        Unload Me
    End If
End Sub
Sub Baixa()
    Dim strSQL As String
    Dim cEst As Currency
    strSQL = "update tab_perdas set situacao_per='2' where seq_per='" & Me.tag & "' "
    conn.Execute strSQL

    With dtProd.Recordset
        Do While Not .EOF
            strSQL = F_BuscaUmCampo("eR_pro", "tab_produtos", "cod_pro='" & !cod_pro & "'")
            If !qtd_itp > CCur(strSQL) Then
                cEst = 0
            Else
                cEst = CCur(strSQL) - !qtd_itp
            End If
            strSQL = "update tab_produtos set er_pro=" & F_GravaDec(cEst)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            
            conn.Execute strSQL
            
            .MoveNext
        Loop
        
    End With
End Sub
Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    If Me.tag = "" Then
        cmndbtnSalvar.caption = "&Abrir Lançamento"
    Else
        cmndbtnSalvar.caption = "&Concluir Lançamento"
    End If
    busca
    bActive = True
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select ti.*,ti.unit_itp*ti.qtd_itp as total,tp.nome_pro from tab_produtos as tp, "
    strSQL = strSQL & "tab_itperdas ti, tab_perdas as ta where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & "and ta.seq_per=ti.seq_per "
    strSQL = strSQL & " and ti.seq_per='" & Me.tag & "' "
    If optnCod.Value = True Then
        strSQL = strSQL & "  order by ti.codfab_pro"
    ElseIf optnSeq.Value Then
        strSQL = strSQL & "  order by ti.item_itp"
    Else
        strSQL = strSQL & "  order by tp.nome_pro"
    End If
    
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    soma
End Sub
Private Sub Form_Load()
    P_CentraForm Me
    dtProd.ConnectionString = strConnect
End Sub

Private Sub txtbxCodBarra_GotFocus()
    ganhaFoco txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim p As Produto
    If KeyCode = vbKeyReturn Then
        txtbxCodBarra_LostFocus
        If optnBar.Value Then
            p = F_BuscaProdSaida(txtbxCodBarra.Text)
        ElseIf optnC.Value Then
            p = F_BuscaProdSaida(txtbxCodBarra.Text)
        Else
            p = F_BuscaProdSaida(txtbxCodBarra.Text)
        End If
    End If
    If p.cod_pro <> "" Then
        txtbxCodBarra.Text = p.nome_pro
        txtbxCodBarra.tag = p.cod_pro
        txtbxCusto.Text = FormatVal(p.cr_pro)
        txtbxQtd.SetFocus
    
    End If
End Sub

Private Sub txtbxCodBarra_LostFocus()
    AjustaTexto txtbxCodBarra
End Sub

Private Sub txtbxNOme_Ter_GotFocus()
    ganhaFoco txtbxNome_ter
End Sub

Private Sub txtbxNOme_Ter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub
Private Sub cmndbtnLocaliza_Click()
    bActive = True
    frmSelectCli.txtbxNome.tag = "F"
    frmSelectCli.txtbxNome.Text = txtbxNome_ter.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxNome_ter.Text = frmSelectCli.adc.Recordset!nome
    txtbxNome_ter.tag = frmSelectCli.adc.Recordset!cod
        
    Unload frmSelectCli
    bActive = False
        
    
    
End Sub

Private Sub txtbxQtd_GotFocus()
    ganhaFoco txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtbxQtd_LostFocus()
    PerdeFoco txtbxQtd, "###,##0.0"
End Sub
