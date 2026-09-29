VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmControleChequesPagar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controle de Cheques"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   15120
   Icon            =   "frmcontrolechequespagar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmcontrolechequespagar.frx":014A
      Height          =   3975
      Left            =   0
      TabIndex        =   35
      Top             =   840
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7011
      _Version        =   393216
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "status"
         Caption         =   "Status"
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
         DataField       =   "nome"
         Caption         =   "Nome"
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
      BeginProperty Column02 
         DataField       =   "num_che"
         Caption         =   "Número"
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
      BeginProperty Column03 
         DataField       =   "valor_che"
         Caption         =   "Valor"
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
      BeginProperty Column04 
         DataField       =   "venc_che"
         Caption         =   "Vencimento"
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
      BeginProperty Column05 
         DataField       =   "nome_con"
         Caption         =   "Conta"
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
      BeginProperty Column06 
         DataField       =   "num_doc"
         Caption         =   "Documento"
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
         BeginProperty Column00 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2819,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1874,835
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1214,929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   1200
      Top             =   8760
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
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
   Begin VB.PictureBox pcBusca 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   15090
      TabIndex        =   28
      Top             =   0
      Width           =   15120
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   645
         Left            =   14280
         Picture         =   "frmcontrolechequespagar.frx":015C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmcontrolechequespagar.frx":04E6
         Left            =   120
         List            =   "frmcontrolechequespagar.frx":04F3
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnConta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         Picture         =   "frmcontrolechequespagar.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxNumChe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10080
         MaxLength       =   50
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmcontrolechequespagar.frx":08B2
         Left            =   12600
         List            =   "frmcontrolechequespagar.frx":08C2
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbxSituacao 
         Height          =   315
         ItemData        =   "frmcontrolechequespagar.frx":08EE
         Left            =   2160
         List            =   "frmcontrolechequespagar.frx":08FE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12240
         Picture         =   "frmcontrolechequespagar.frx":092F
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "Conta:"
         Height          =   195
         Left            =   5760
         TabIndex        =   8
         Tag             =   "!"
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "Num Doc:"
         Height          =   195
         Left            =   4680
         TabIndex        =   6
         Tag             =   "!"
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblLocal 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "Num Che:"
         Height          =   195
         Left            =   3600
         TabIndex        =   4
         Tag             =   "!"
         Top             =   0
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   7920
         TabIndex        =   11
         Tag             =   "!"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   9000
         TabIndex        =   13
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor:"
         Height          =   195
         Left            =   10080
         TabIndex        =   15
         Tag             =   "!"
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar Por:"
         Height          =   195
         Left            =   12600
         TabIndex        =   18
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   2160
         TabIndex        =   2
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox pcBotes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   15090
      TabIndex        =   21
      Top             =   8325
      Width           =   15120
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmcontrolechequespagar.frx":0CB9
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmcontrolechequespagar.frx":0E03
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmcontrolechequespagar.frx":118D
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   9240
         Picture         =   "frmcontrolechequespagar.frx":12D7
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDev 
         Caption         =   "&Cheque Devolvido"
         Height          =   735
         Left            =   5160
         Picture         =   "frmcontrolechequespagar.frx":1825
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   30
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   6840
         Picture         =   "frmcontrolechequespagar.frx":196F
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   30
         Width           =   1575
      End
   End
   Begin VB.PictureBox pCdados 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   10905
      TabIndex        =   29
      Top             =   7800
      Width           =   10935
      Begin VB.Label lblDev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label10"
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
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblQtd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9240
         TabIndex        =   33
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados:"
         Height          =   195
         Left            =   8520
         TabIndex        =   32
         Top             =   60
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   4560
         TabIndex        =   31
         Top             =   60
         Width           =   405
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   30
         Top             =   0
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmControleChequesPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxBuscar_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxTipo_Click()
    If cmbxTipo.ListIndex = 0 Then
        dtgrd.Visible = True
        dtgrdPag.Visible = False
    Else
        dtgrd.Visible = False
        dtgrdPag.Visible = True
    End If
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnBaixar_Click()
    Dim sSeq As String
    Dim iCount As Integer
    Dim c As cheque
    Dim sChe As String
    Dim sEsp As String
    Dim cValor As Currency
    Dim strSQL As String
    Dim sCod_Con  As String
    Dim sData As String
    With adc.Recordset
    
        If .RecordCount = 0 Then
            MsgBox "Não há o que baixar!", vbExclamation
            Exit Sub
        End If
        
      
            sCod_Con = !cod_con
            sData = Date
      
        If sCod_Con = "" Then
            Exit Sub
        End If
        If MsgBox("Confirma baixa do cheque " & !num_che & " ?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        If sCod_Con = "" Then
            Exit Sub
        End If
        
        
        sChe = F_GravaFinanceiro("D", Date, CStr(!Valor_che), _
                 "Baixa Cheque " & !num_che & "-" & !nome, sCod_Con, "", "000")
                
        If sChe <> "" Then
            'P_Class !Valor_che, Left(cmbxTipo.Text, 1), sChe
            strSQL = "update tab_cheques set dep_che = " & F_GravaData(Date) & ","
            strSQL = strSQL & " status_Che=2 "
            strSQL = strSQL & " where seq_che = '" & !seq_che & "' "
            conn.Execute strSQL
            
            strSQL = "insert into tab_chequedep(seq_che,data_dep,cod_con)values("
            strSQL = strSQL & "'" & !seq_che & "',"
            strSQL = strSQL & F_GravaData(sData) & ","
            strSQL = strSQL & "'" & sCod_Con & "')"
            conn.Execute strSQL
            
            MsgBox "Cheque baixado!", vbInformation
            adc.Refresh
            
        End If
        
    End With
            
End Sub

Private Sub cmndbtnBuscar_Click()
    
    Dim strSQL As String
    Dim sOrderBy As String
    On Error GoTo trata_Erro
    
    
    If cmbxOrder.ListIndex = 0 Then
        sOrderBy = "venc_che"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "Nome"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "num_doc"
    Else
        sOrderBy = "Valor_Che"
    End If
    strSQL = Cheques.F_selectChPag(Left(cmbxTipoTer.Text, 1), txtbxExpr.tag, txtbxDataIni.Text, txtbxdataFin.Text, _
        txtbxNumChe.Text, "", Left(cmbxSituacao.Text, 1), txtbxConta.tag, "", "", txtbxNumDoc.Text, sOrderBy)
    adc.RecordSource = strSQL
    adc.Refresh
    
    soma
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnLocaliza_Click()
    frmBusca_Ter.txtbxCli.Text = UCase(txtbxExpr.Text)
    On Error Resume Next
    frmBusca_Ter.Show 1
    txtbxExpr.Text = frmBusca_Ter.dtCli.Recordset!Nome_Ter
    txtbxExpr.tag = frmBusca_Ter.dtCli.Recordset!cod_ter
    Unload frmBusca_Ter
End Sub

Private Sub cmndbtnConta_Click()
    buscaConta
End Sub
Sub buscaConta()
    Dim c As ContaCorrente
    c = F_BuscaContaCorrente(txtbxConta.Text)
    If c.cod_con <> "" Then
        txtbxConta.Text = c.nome_con
        txtbxConta.tag = c.cod_con
    End If
End Sub
Private Sub cmndbtnDev_Click()
    On Error GoTo trata_Erro
    With adc.Recordset
        'If F_DevolveChe(!seq_che) Then
        '    adc.Refresh
        'End If
    End With
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub buscaTer()
    
        buscaCLi
    
End Sub
Sub buscaCLi()
    Dim c As fornecedor.fornecedor
    c = fornecedor.F_BuscaFornecedor(txtbxExpr.Text)
    
    If c.cod_for <> "" Then
        txtbxExpr.Text = c.nome_for
        txtbxExpr.tag = c.cod_for
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    Dim c As cheque
    Dim cp As chePag
    Dim strSQL As String
    With adc.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser editado!", vbExclamation
            Exit Sub
        End If
        
            cp = F_DadosChePag(!Venc_che, !num_che, _
                 !cod_con, CStr(!Valor_che))
            If cp.num_che = "" Then
                MsgBox "Não foi informado um número para o cheque!", vbExclamation
                Exit Sub
            ElseIf cp.Valor_che = 0 Then
                MsgBox "Não foi informado um valor válido para o cheque!  ", vbExclamation
                Exit Sub
            End If
            If MsgBox("Confirma alteração do cheque?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "update tab_cheques set num_che = '" & cp.num_che & "', "
                
                
                strSQL = strSQL & "cod_con='" & cp.cod_con & "',"
                
                strSQL = strSQL & "venc_che=" & F_GravaData(cp.Venc_che) & ","
                strSQL = strSQL & "valor_che=" & F_GravaDec(cp.Valor_che)
                strSQL = strSQL & " where seq_che = '" & !seq_che & "' "
                conn.Execute strSQL
                MsgBox "cheque alterado!", vbInformation
                RefreshDataPosition adc
            End If
        
    
    End With
End Sub

Private Sub cmndbtnImpr_Click()
    If adc.Recordset.RecordCount <= 0 Then
        Exit Sub
    End If
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    
    Dim strSQL As String
    Printer.Orientation = vbPRORLandscape
    If cmbxTipo.ListIndex = 0 Then
        strSQL = "A Receber"
    End If
    If cmbxSituacao.ListIndex = 2 Then
        strSQL = strSQL & " Devolvidos "
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " de " & txtbxDataIni.Text
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " até " & txtbxDataIni.Text
    End If
    P_CabRelTimbr "Relatório de Cheques " & strSQL, obj
    
    With adc.Recordset
        obj.fontsize = 8
        obj.FontBold = True
        P_line1 obj
        P_ImprX obj, "Cliente", 1
        P_ImprX obj, "Titular", 60
        P_ImprX obj, "CPF", 120
        P_ImprX obj, "Número", 150
        P_ImprX obj, "Conta", 170
        P_ImprX obj, "Ag.", 185
        P_ImprX obj, "Valor", 210
        P_ImprX obj, "Bco", 220

        P_ImprX obj, "Data", 235
        P_ImprX obj, "Venc.", 250
        P_ImprX obj, "Status", 265, False
        P_line1 obj
        obj.FontBold = False
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 25 Then
                P_RodapeVis obj, "Continua..."
                P_AddPicVis obj
                obj.fontsize = 8
                obj.FontBold = True
                
                obj.Print
                P_line1 obj
                P_ImprX obj, "Cliente", 1
                P_ImprX obj, "Titular", 60
                P_ImprX obj, "CPF", 120
                P_ImprX obj, "Número", 150
                P_ImprX obj, "Conta", 170
                P_ImprX obj, "Ag.", 185
                P_ImprX obj, "Bco", 200
                P_ImprX obj, "Data", 215
                P_ImprX obj, "Venc.", 240
                P_ImprX obj, "Status", 255, False
                P_line1 obj
                obj.FontBold = False
            End If
            obj.fontsize = 8
            P_ImprX obj, !nome, 1
            P_ImprX obj, !nome_che, 60
            P_ImprX obj, !cpf_che, 120
            P_ImprX obj, !num_che, 150
            P_ImprX obj, !cco_che, 170
            P_ImprX obj, !age_che, 185
            P_ImprX obj, FormatVal(!Valor_che), 217, True, True
            P_ImprX obj, !cod_bco, 220
            P_ImprX obj, !data_doc, 235
            
            P_ImprX obj, !Venc_che, 250
            P_ImprX obj, !situacao, 265, False
            P_line1 obj, vbDot
            .MoveNext
        Loop
        P_line1 obj
        obj.FontBold = True
        P_ImprX obj, "Cheques listados: " & lblQtd.caption, 10
        P_ImprX obj, "Valor Total: " & lblTotal.caption, obj.ScaleWidth - 10, False, True
        obj.FontBold = False
        
        
        P_AddPicVis obj
    End With
    
    
    
End Sub
Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 0
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.BorderStyle = 0
End Sub

Private Sub cmndbtnTerc_Click()
    buscaTer
End Sub

Private Sub dtNot_Reposition()
    Dim strSQL As String
    With dtNot.Recordset
        If .RecordCount = 0 Then Exit Sub
        strSQL = "select * from tab_cheques where seq_doc='" & !seq_doc & "' "
        adc.RecordSource = strSQL
        adc.Refresh
        If adc.Recordset.RecordCount > 0 Then
            dtgrd.caption = "Cheques referentes ao documento " & !num_doc
            dtgrd.Visible = True
        Else
            dtgrd.Visible = False
        End If
        If .RecordCount > 0 Then
            If !reman_itd = 0 Then
                cmndbtnEdit.Enabled = False
                cmndbtnBaixar.Enabled = False
            Else
                cmndbtnEdit.Enabled = True
                cmndbtnBaixar.Enabled = True
            End If
            Exit Sub
        
        End If
        
    End With
End Sub


Private Sub dtgrd_DblClick()
    If cmndbtnEdit.Enabled Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub dtgrdPag_GotFocus()
    'dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrdPag_LostFocus()
'    dtgrd.MarqueeStyle = 0
End Sub

Sub soma()
    Dim cTotal As Currency
    
    On Error GoTo trata_Erro
    
    With adc.Recordset
        .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !Valor_che
            .MoveNext
        Loop
    End With
    lblTotal.caption = FormatVal(cTotal)
    Exit Sub
trata_Erro:
    lblTotal.caption = "0,00"
End Sub

Private Sub adc_Reposition()
    On Error Resume Next
    If adc.Recordset!status_che = 2 Then
        cmndbtnEdit.Enabled = False
        cmndbtnBaixar.Enabled = False
        
    
    Else
        cmndbtnEdit.Enabled = True
        cmndbtnBaixar.Enabled = True
    End If
    If adc.Recordset!status_che = "3" Then
        devChe
    Else
        lblDev.Visible = False
    End If
    
End Sub
Sub devChe()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    lblDev.caption = ""
    With adc.Recordset
        strSQL = "select * from tab_chequedev where seq_che='" & !seq_che & "' "
        rPesq.Open strSQL, conn_client
        
    End With
    With rPesq
        If .RecordCount > 0 Then
            lblDev.caption = !motivo_Dev & " em " & !data_dev
        End If
        .Close
    End With
    lblDev.Visible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    Dim strSQL As String
  '  On Error GoTo trata_Erro
    resizeForm Me
    txtbxDataIni.Text = "01" & Mid(Format(Date, "dd/mm/yyyy"), 3, 8)
    strSQL = DateAdd("m", 1, txtbxDataIni.Text)
    txtbxdataFin.Text = DateAdd("D", -1, strSQL)
   
    
    adc.ConnectionString = strConnect
    
    

    
    cmbxOrder.ListIndex = 0
    cmbxTipoTer.ListIndex = 0
    'cmndbtnBuscar_Click
    Exit Sub
trata_Erro:
    
End Sub

Private Sub optnCli_Click()
    frExpr.Enabled = True
End Sub

Private Sub optnTodas_Click()
    frExpr.Enabled = False
 
    
End Sub

Private Sub Form_Resize()
    posicionaControles
End Sub
Sub posicionaControles()
    On Error Resume Next
    dtgrd.width = Me.width - 100
    dtgrd.Height = pcBotes.Top - dtgrd.Top - 100
    
    cmndbtnSair.Left = Me.width - cmndbtnSair.width - 100
End Sub

Private Sub txbtxLocalche_Change()

End Sub

Private Sub txtbtxLocalche_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaLoc
    End If
End Sub
Sub buscaLoc()
    Dim l As localche
    l = f_BuscaLocalChe(txtbxLocal.Text)
    If l.cod_loc <> "" Then
        txtbxLocal.Text = l.nome_loc
        txtbxLocal.tag = l.cod_loc
        cmndbtnBuscar_Click
    End If
End Sub
Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_Change()
    cmndbtnBuscar_Click
End Sub

Private Sub txtbxExpr_GotFocus()
    ganhaFoco txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxExpr_LostFocus
        cmndbtnTerc_Click
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    If cmbxBuscar.ListIndex = 0 Then
        txtbxExpr.Text = F_SoNumero(txtbxExpr.Text)
    Else
        AjustaTexto txtbxExpr
    End If
End Sub


Private Sub txtbxLocal_GotFocus()
    'GanhaFoco txtbxLocal
End Sub

Private Sub txtbxLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaLoc
    End If
End Sub
