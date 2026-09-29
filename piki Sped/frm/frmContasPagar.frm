VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmContasPagar 
   Caption         =   "Contas A Pagar"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12645
   Icon            =   "frmContasPagar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   12645
   Begin TabDlg.SSTab sstb 
      Height          =   6375
      Left            =   0
      TabIndex        =   25
      Top             =   840
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Duplicatas"
      TabPicture(0)   =   "frmContasPagar.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dbgrd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "adc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Cheques"
      TabPicture(1)   =   "frmContasPagar.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "adcch"
      Tab(1).Control(1)=   "gridche"
      Tab(1).ControlCount=   2
      Begin MSAdodcLib.Adodc adc 
         Height          =   330
         Left            =   1440
         Top             =   5280
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSDataGridLib.DataGrid dbgrd 
         Bindings        =   "frmContasPagar.frx":688A
         Height          =   5655
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9975
         _Version        =   393216
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
      Begin MSAdodcLib.Adodc adcch 
         Height          =   330
         Left            =   -73200
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSDataGridLib.DataGrid gridche 
         Bindings        =   "frmContasPagar.frx":689C
         Height          =   5655
         Left            =   -74880
         TabIndex        =   34
         Top             =   360
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   9975
         _Version        =   393216
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
   Begin VB.PictureBox pcBusca 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12615
      TabIndex        =   7
      Top             =   0
      Width           =   12645
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   525
         Left            =   11640
         Picture         =   "frmContasPagar.frx":68B0
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox txtbxHistoricoDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtbxSeqDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxFornecedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmContasPagar.frx":6C3A
         Left            =   9960
         List            =   "frmContasPagar.frx":6C4D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbxSituacao 
         Height          =   315
         ItemData        =   "frmContasPagar.frx":6C91
         Left            =   8640
         List            =   "frmContasPagar.frx":6C9B
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         Picture         =   "frmContasPagar.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Histórico:"
         Height          =   195
         Left            =   4440
         TabIndex        =   31
         Tag             =   "!"
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Id. Doc:"
         Height          =   195
         Left            =   1200
         TabIndex        =   29
         Tag             =   "!"
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº Doc:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Tag             =   "!"
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   2280
         TabIndex        =   12
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
         Left            =   3360
         TabIndex        =   14
         Tag             =   "!"
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0B0B0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornecedor:"
         Height          =   195
         Left            =   6000
         TabIndex        =   16
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
         Left            =   9960
         TabIndex        =   8
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   8640
         TabIndex        =   10
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox pcBotes 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12615
      TabIndex        =   0
      Top             =   7485
      Width           =   12645
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   735
         Left            =   3480
         Picture         =   "frmContasPagar.frx":7044
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmContasPagar.frx":718E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1800
         Picture         =   "frmContasPagar.frx":7518
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   10800
         Picture         =   "frmContasPagar.frx":7662
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDev 
         Caption         =   "&Cheque Devolvido"
         Height          =   735
         Left            =   5160
         Picture         =   "frmContasPagar.frx":7BB0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   6840
         Picture         =   "frmContasPagar.frx":7CFA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
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
      TabIndex        =   19
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados:"
         Height          =   195
         Left            =   8520
         TabIndex        =   22
         Top             =   60
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   4560
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   0
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmContasPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbxBuscar_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub


Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub baixaCheque()
    Dim sSeq As String
    Dim iCount As Integer
    Dim c As cheque
    Dim sChe As String
    Dim sEsp As String
    Dim cValor As Currency
    Dim strSQL As String
    Dim sCod_Con  As String
    Dim sData As String
    With adcch.Recordset
    
        If .RecordCount = 0 Then
            MsgBox "Não há o que baixar!", vbExclamation
            Exit Sub
        End If
        
        If !tipo_che = "R" Then
            frmContaCorrente.Show 1
            frmContaCorrente.caption = "Baixa de Cheque"
            sCod_Con = Left(frmContaCorrente.tag, 3)
            sData = frmContaCorrente.txtbxData.Text
            
            Unload frmContaCorrente
        Else
            sCod_Con = !cod_con
            sData = Date
        End If
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
                "Baixa Cheque " & !num_che & "-" & !nome, "DIN", sCod_Con, "", _
                    adcch.Recordset!seq_che, "", !cod_pla)
                
        If sChe <> "" Then
            
            
            'P_Class !Valor_che, Left(cmbxTipo.Text, 1), sChe
            
            strSQL = "update tab_cheques1 set dep_che = " & F_GravaData(Date) & ","
            strSQL = strSQL & " status_Che=2 "
            strSQL = strSQL & " where seq_che = '" & !seq_che & "' "
            conn.Execute strSQL
            
            strSQL = "insert into tab_chequedep(seq_che,data_dep,cod_con)values("
            strSQL = strSQL & "'" & !seq_che & "',"
            strSQL = strSQL & F_GravaData(sData) & ","
            strSQL = strSQL & "'" & sCod_Con & "')"
            conn.Execute strSQL
            
            MsgBox "Cheque baixado!", vbInformation
            adcch.Refresh
            
        End If
        
    End With
            
End Sub

Private Sub cmndbtnBaixar_Click()
    If sstb.Tab = 0 Then
        baixaDup
    Else
        baixaCheque
    End If
End Sub
Sub baixaDup()
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma baixa da parcela?", vbYesNo + vbQuestion) = vbYes Then
        If BaixaItDocPag(adc.Recordset!seq_itd, adc.Recordset!Valor_Itd, "0") Then
            MsgBox "Documento baixado!", vbInformation
            adc.Refresh
        Else
            MsgBox "Erro na baixa", vbExclamation
            
        End If
    End If
    
End Sub
Private Sub cmndbtnBuscar_Click()
    On Error GoTo trata_erro
    Dim strSQL As String
    strSQL = Financeiro.F_SelectContasPagar(txtbxDataIni.Text, txtbxDAtaFIn.Text, _
        txtbxNumDoc.Text, txtbxSeqDoc.Text, "", txtbxFornecedor.tag, "", txtbxHistoricoDoc.Text, _
        IIf(Left(cmbxSituacao.Text, 1) = "2", True, False))
    adc.RecordSource = strSQL
    adc.Refresh
    
    
    strSQL = Cheques.F_selectChPag("F", txtbxFornecedor.tag, txtbxDataIni.Text, txtbxDAtaFIn.Text)
    adcch.RecordSource = strSQL
    adcch.Refresh
    
    
  '  lblQtd.caption = adcch.Recordset.RecordCount
  '  soma
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub


Private Sub cmndbtnDev_Click()
    On Error GoTo trata_erro
    With adcch.Recordset
        If F_DevolveChe(!seq_che) Then
            adcch.Refresh
        End If
    End With
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    Dim c As cheque
    Dim cp As chePag
    Dim strSQL As String
    Dim sTipo As String
    
    sTipo = "P"
    With adcch.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser editado!", vbExclamation
            Exit Sub
        End If
        If sTipo = "R" Then
            c = F_DadosCheque(True, !Venc_che, !nome_che, _
                !num_che, !cco_che, !age_che, !cod_bco, CStr(!Valor_che), Format(!cpf_che))
            If c.num_che = "" Then
'                MsgBox "Não foi informado um número para o cheque!", vbExclamation
                Exit Sub
            ElseIf c.Valor_che = 0 Then
                MsgBox "Não foi informado um valor válido para o cheque!  ", vbExclamation
                Exit Sub
            End If
            If MsgBox("Confirma alteração do cheque?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "update tab_cheques1 set num_che = '" & c.num_che & "', "
                strSQL = strSQL & "age_che='" & c.age_che & "',"
                strSQL = strSQL & "cco_che='" & c.cco_che & "',"
                strSQL = strSQL & "cod_bco='" & c.cod_bco & "',"
                strSQL = strSQL & "nome_che='" & c.nome_che & "',"
                strSQL = strSQL & "venc_che=" & F_GravaData(c.Venc_che) & ","
                strSQL = strSQL & "valor_che=" & F_GravaDec(c.Valor_che) & ","
                strSQL = strSQL & "cpf_che='" & c.cpf_che & "'"
                
                strSQL = strSQL & " where seq_che = '" & !seq_che & "' "
                conn.Execute strSQL
                MsgBox "cheque alterado!", vbInformation
                adcch.Refresh
            End If
        Else
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
                strSQL = "update tab_cheques1 set num_che = '" & cp.num_che & "', "
                
                
                strSQL = strSQL & "cod_con='" & cp.cod_con & "',"
                
                strSQL = strSQL & "venc_che=" & F_GravaData(cp.Venc_che) & ","
                strSQL = strSQL & "valor_che=" & F_GravaDec(cp.Valor_che)
                strSQL = strSQL & " where seq_che = '" & !seq_che & "' "
                conn.Execute strSQL
                MsgBox "cheque alterado!", vbInformation
                adcch.Refresh
            End If
        End If
    
    End With
End Sub

Private Sub cmndbtnImpr_Click()
    If adcch.Recordset.RecordCount <= 0 Then
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
    Dim sTipo As String
    sTipo = "P"
    Dim strSQL As String
    Printer.Orientation = vbPRORLandscape
    If sTipo = "R" Then
        strSQL = "A Receber"
    End If
    If cmbxSituacao.ListIndex = 2 Then
        strSQL = strSQL & " Devolvidos "
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " de " & txtbxDataIni.Text
    End If
    If txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & " até " & txtbxDataIni.Text
    End If
    P_CabRelTimbr "Relatório de Cheques " & strSQL, obj
    
    With adcch.Recordset
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
            P_ImprX obj, !Data_Doc, 235
            
            P_ImprX obj, !Venc_che, 250
            If !tipo_che = "R" Then
                P_ImprX obj, !situacao, 265, False
            Else
                P_ImprX obj, "", 265, False
            End If
            P_line1 obj, vbDot
            .MoveNext
        Loop
        P_line1 obj
        obj.FontBold = True
        P_ImprX obj, "Cheques listados: " & lblQtd.caption, 10
        P_ImprX obj, "Valor Total: " & lblTOtal.caption, obj.ScaleWidth - 10, False, True
        obj.FontBold = False
        
        
        P_AddPicVis obj
    End With
    
    
    
End Sub
Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 0
    
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    Unload Me
    adcch.Recordset.Close
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub



Private Sub cmndbtnTerc_Click()
   On Error GoTo trata_erro
    
    
    buscaFor
    
    
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub buscaCli()
    Dim c As tCliente
    c = F_BuscaCliente("C", txtbxFornecedor.Text)
    If c.cod_cli <> "" Then
        txtbxFornecedor.Text = c.nome_Cli
        txtbxFornecedor.tag = c.cod_cli
    End If
End Sub
Sub buscaFor()
    Dim c As tCliente
    c = F_BuscaCliente("F", txtbxFornecedor.Text)
    If c.cod_cli <> "" Then
        txtbxFornecedor.Text = c.nome_Cli
        txtbxFornecedor.tag = c.cod_cli
        cmndbtnBuscar_Click
    End If
End Sub
Private Sub dtNot_Reposition()
    Dim strSQL As String
    With adc.Recordset
        If .RecordCount = 0 Then Exit Sub
        strSQL = "select * from tab_cheques1 where seq_doc='" & !seq_doc & "' "
        adcch.RecordSource = strSQL
        adcch.Refresh
        If adcch.Recordset.RecordCount > 0 Then
            gridche.caption = "Cheques referentes ao documento " & !num_doc
            gridche.Visible = True
        Else
            gridche.Visible = False
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

Private Sub gridche_DblClick()
    If cmndbtnEdit.Enabled Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub gridche_GotFocus()
    gridche.MarqueeStyle = 3
End Sub

Private Sub gridche_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        gridche_DblClick
    End If
End Sub

Private Sub gridche_LostFocus()
    gridche.MarqueeStyle = 0
End Sub


Private Sub gridchePag_GotFocus()
    gridche.MarqueeStyle = 3
End Sub

Private Sub gridchePag_LostFocus()
    gridche.MarqueeStyle = 0
End Sub

Sub soma()
    Dim cTotal As Currency
    
    On Error GoTo trata_erro
    
    With adcch.Recordset
        .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !Valor_che
            .MoveNext
        Loop
    End With
    lblTOtal.caption = FormatVal(cTotal)
    Exit Sub
trata_erro:
    lblTOtal.caption = "0,00"
End Sub

Private Sub adcch_Reposition()
    On Error Resume Next
    If adcch.Recordset!status_che = 2 Then
        cmndbtnEdit.Enabled = False
        cmndbtnBaixar.Enabled = False
        
    
    Else
        cmndbtnEdit.Enabled = True
        cmndbtnBaixar.Enabled = True
    End If
    If adcch.Recordset!status_che = "3" Then
        devChe
    Else
        lblDev.Visible = False
    End If
    
End Sub
Sub devChe()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    lblDev.caption = ""
    With adcch.Recordset
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
    'On Error GoTo trata_erro
    resizeForm Me
    txtbxDataIni.Text = F_PrimDiaMes(Date)
    txtbxDAtaFIn.Text = F_UltDiaMes(Date)
    
    
    cmbxOrder.ListIndex = 0
    
    adc.ConnectionString = strConnect
    adcch.ConnectionString = strConnect
    
    
'    strSQL = "select tn.*,tt.* from tab_compras as tn, tab_fornecedores as tt "
'    strSQL = strSQL & " where tt.cod_for=tn.cod_for "
'
'    strSQL = strSQL & " order by num_not "
'    dtNot.RecordSource = strSQL
'    dtNot.Refresh
    
    cmbxOrder.ListIndex = 0
sstb.tag = 0
    'cmndbtnBuscar_Click
    Exit Sub
trata_erro:
    
End Sub


Private Sub Form_Resize()
    posicionaControles
End Sub
Sub posicionaControles()
    
    
End Sub

Private Sub txtbxdataFin_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxdataFin_LostFocus()
    P_AjustaData txtbxDAtaFIn
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



Private Sub txtbxFornecedor_Change()
    txtbxFornecedor.tag = ""
End Sub

Private Sub txtbxFornecedor_GotFocus()
    ganhaFoco txtbxFornecedor
End Sub

Private Sub txtbxFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxNumDoc_GotFocus()
    ganhaFoco txtbxNumDoc
End Sub

Private Sub txtbxNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNumDoc_LostFocus()
    AjustaTexto txtbxNumDoc
End Sub

Private Sub txtbxSeqDoc_GotFocus()
    ganhaFoco txtbxSeqDoc
End Sub

Private Sub txtbxSeSdoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxSeqDoc_LostFocus()
    AjustaTexto txtbxSeqDoc
End Sub
Private Sub txtbxHistoricoDoc_GotFocus()
    ganhaFoco txtbxHistoricoDoc
End Sub

Private Sub txtbxHistoricoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxHistoricoDoc_LostFocus()
    AjustaTexto txtbxHistoricoDoc
End Sub
