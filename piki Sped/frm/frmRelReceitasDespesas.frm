VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelReceitasDespesas 
   Caption         =   "Análise de Receitas e Despesas"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13395
   Icon            =   "frmRelReceitasDespesas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   13395
   Begin TabDlg.SSTab sstb 
      Height          =   6495
      Left            =   2640
      TabIndex        =   35
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Movimento Financeiro"
      TabPicture(0)   =   "frmRelReceitasDespesas.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Documentos (Previsão)"
      TabPicture(1)   =   "frmRelReceitasDespesas.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "gridDoc"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid grid 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6376
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridDoc 
         Height          =   3615
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6376
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7125
      Left            =   0
      ScaleHeight     =   7095
      ScaleWidth      =   2505
      TabIndex        =   22
      Top             =   0
      Width           =   2535
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmRelReceitasDespesas.frx":047A
         Left            =   120
         List            =   "frmRelReceitasDespesas.frx":0487
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Análise:"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2295
         Begin VB.OptionButton optnDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "De&talhada"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optnSint 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "S&intética"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1200
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   685
         Left            =   120
         Picture         =   "frmRelReceitasDespesas.frx":04A6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   685
         Left            =   120
         Picture         =   "frmRelReceitasDespesas.frx":07B0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnFechar 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   685
         Left            =   120
         Picture         =   "frmRelReceitasDespesas.frx":0ABA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4920
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico"
         Height          =   685
         Left            =   120
         Picture         =   "frmRelReceitasDespesas.frx":1384
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox cmbxContas 
         Height          =   315
         ItemData        =   "frmRelReceitasDespesas.frx":17C6
         Left            =   120
         List            =   "frmRelReceitasDespesas.frx":17D6
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D&e:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   1920
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Conta:"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   465
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13365
      TabIndex        =   15
      Top             =   7125
      Width           =   13395
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "=>"
         Height          =   195
         Left            =   4320
         TabIndex        =   34
         Top             =   450
         Width           =   180
      End
      Begin VB.Label lblFinal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   4800
         TabIndex        =   33
         Top             =   427
         Width           =   375
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "=>"
         Height          =   195
         Left            =   4320
         TabIndex        =   32
         Top             =   240
         Width           =   180
      End
      Begin VB.Label lblTotPag 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   4800
         TabIndex        =   31
         Top             =   217
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "=>"
         Height          =   195
         Left            =   4320
         TabIndex        =   30
         Top             =   0
         Width           =   180
      End
      Begin VB.Label lblTotRec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   4800
         TabIndex        =   29
         Top             =   -23
         Width           =   375
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   2280
         TabIndex        =   28
         Top             =   450
         Width           =   360
      End
      Begin VB.Label lblTotDoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   3240
         TabIndex        =   27
         Top             =   427
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Pagar"
         Height          =   195
         Left            =   2280
         TabIndex        =   26
         Top             =   240
         Width           =   570
      End
      Begin VB.Label lblPag 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3240
         TabIndex        =   25
         Top             =   217
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Receber"
         Height          =   195
         Left            =   2280
         TabIndex        =   24
         Top             =   0
         Width           =   765
      End
      Begin VB.Label lblRec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   3240
         TabIndex        =   23
         Top             =   -23
         Width           =   375
      End
      Begin VB.Label lblTotalD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   21
         Top             =   217
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Débito:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblTotalC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   840
         TabIndex        =   19
         Top             =   -23
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblTOtal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   840
         TabIndex        =   17
         Top             =   427
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   450
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmRelReceitasDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim WithEvents adoDoc As ADODB.Recordset
Attribute adoDoc.VB_VarHelpID = -1
Private Sub cmbxContas_Click()
    cmbxContas.tag = F_BuscaUmCampo("Cod_Pla", TAB_PLANOCONTAS, "nome_pla='" & _
        cmbxContas.Text & "'")
    If cmbxContas.tag <> "" Then
        cmndbtnGraf.Enabled = False
    Else
        If optnDet.Value = False Then
            cmndbtnGraf.Enabled = True
        End If
    End If
    
End Sub

Private Sub cmbxContas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDelete Then
        cmbxContas.ListIndex = -1
    End If
End Sub

Sub BuscaSint()
    Dim strSQL As String
    strSQL = "select sum(valor_fin) as total,tp.nome_pla,tipo_pla "
    strSQL = strSQL & " from " & TAB_PLANOCONTAS & " as tp,"
    strSQL = strSQL & "  " & TAB_FINANCEIRO & "  as tf "
    strSQL = strSQL & " where tp.cod_pla=tf.cod_pla "

    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If cmbxTipo.ListIndex = 2 Then
        strSQL = strSQL & " and tp.tipo_pla='D'"
    ElseIf cmbxTipo.ListIndex = 1 Then
        strSQL = strSQL & " and tp.tipo_pla='C'"
    End If
    If cmbxContas.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pla='" & cmbxContas.tag & "'"
    End If
    strSQL = strSQL & " group by tp.cod_pla "
    strSQL = strSQL & " order by tp.nome_pla"
    abreObjetoAdo ado, strSQL, grid
    
    strSQL = "select sum(reman_itd) as total,tp.nome_pla,tipo_pla "
    strSQL = strSQL & " from " & TAB_PLANOCONTAS & " as tp,"
    strSQL = strSQL & "  " & TAB_DOC & " 1 as td, " & TAB_ITDOC & " 1 as ti "
    strSQL = strSQL & " where tp.cod_pla=td.cod_pla "
    strSQL = strSQL & " and td.seq_doc=ti.seq_doc "
    

    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and Venc_itd<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and venc_itd>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If cmbxTipo.ListIndex = 2 Then
        strSQL = strSQL & " and tipo_doc='P'"
    ElseIf cmbxTipo.ListIndex = 1 Then
        strSQL = strSQL & " and tipo_doc='R'"
    End If
    If cmbxContas.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pla='" & cmbxContas.tag & "'"
    End If
    strSQL = strSQL & " group by tp.cod_pla "
    strSQL = strSQL & " order by tp.nome_pla"
    abreObjetoAdo adoDoc, strSQL, gridDoc
    
    
    
    soma
End Sub
Sub BuscaDet()
    Dim strSQL As String
    strSQL = "select valor_fin as total,tf.cod_con,tf.data_fin,tf.Historico_fin,tf.cod_usu, tp.nome_pla,tipo_pla "
    strSQL = strSQL & " from " & TAB_PLANOCONTAS & " as tp,"
    strSQL = strSQL & "  " & TAB_FINANCEIRO & "  as tf "
    strSQL = strSQL & " where tf.cod_pla=tp.cod_pla "

    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tf.data_fin>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If cmbxTipo.ListIndex = 2 Then
        strSQL = strSQL & " and tp.tipo_pla='D'"
    ElseIf cmbxTipo.ListIndex = 1 Then
        strSQL = strSQL & " and tp.tipo_pla='C'"
    End If
    If cmbxContas.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pla='" & cmbxContas.tag & "'"
    End If
    
    strSQL = strSQL & " order by tp.nome_pla,tf.data_fin"
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    
    
    
    strSQL = "select reman_itd as total,data_doc,Historico_doc,td.cod_usu, tp.nome_pla,tipo_pla "
    strSQL = strSQL & " from " & TAB_PLANOCONTAS & " as tp,"
    strSQL = strSQL & "  " & TAB_DOC & " 1 as td,  " & TAB_ITDOC & " 1 as ti  "
    strSQL = strSQL & " where td.cod_pla=tp.cod_pla "
    strSQL = strSQL & " and td.seQ_DOC=ti.seq_doc "

    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and venc_itd<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and venc_itd>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If cmbxTipo.ListIndex = 2 Then
        strSQL = strSQL & " and tipo_doc='P'"
    ElseIf cmbxTipo.ListIndex = 1 Then
        strSQL = strSQL & " and tipo_doc='R'"
    End If
    If cmbxContas.tag <> "" Then
        strSQL = strSQL & " and tp.cod_pla='" & cmbxContas.tag & "'"
    End If
    
    strSQL = strSQL & " order by tp.nome_pla,Venc_itd"
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    
    soma
End Sub

Sub soma()
    Dim cTotal As Currency
    Dim cTotalC As Currency
    Dim cTotalD As Currency
    With ado
            If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If !tipo_pla = "C" Then
                
                cTotalC = cTotalC + !total
            Else
                
                cTotalD = cTotalD + !total
                
            End If
            
            
            
            .MoveNext
        Loop
    End With
    cTotal = cTotalC - cTotalD
    
    lblTotal.caption = FormatVal(cTotal)
    
    lblTotalC.caption = FormatVal(cTotalC)
    lblTotalD.caption = FormatVal(cTotalD)
    somaDoc
    
    
End Sub
Sub somaDoc()
    Dim cTotal As Currency
    Dim cTotalC As Currency
    Dim cTotalD As Currency
    With adoDoc
    
            If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If !tipo_pla = "C" Then
                
                cTotalC = cTotalC + !total
            Else
                
                cTotalD = cTotalD + !total
                
            End If
            
            
            
            .MoveNext
        Loop
    End With
    cTotal = cTotalC - cTotalD
    
    lblTotDoc.caption = FormatVal(cTotal)
    
    lblRec.caption = FormatVal(cTotalC)
    lblPag.caption = FormatVal(cTotalD)
    
    lblTotRec.caption = FormatVal(CCur(lblTotalC.caption) + CCur(lblRec.caption))
    lblTotPag.caption = FormatVal(CCur(lblTotalD.caption) + CCur(lblPag.caption))
    lblFinal.caption = FormatVal(CCur(lblTotal.caption) + CCur(lblTotDoc.caption))
    
End Sub
Private Sub Command1_Click()
    
End Sub

Private Sub cmbxTipo_Click()
    P_PreenchePlanoContas cmbxContas, cmbxTipo.ListIndex = 1
    If cmbxTipo.ListIndex = 0 Then
        cmbxContas.Enabled = False
    Else
        cmbxContas.Enabled = True
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    
    If txtbxDataIni.Text = "" Then
        MsgBox "Informe a data Inicial!", vbExclamation
        txtbxDataIni.SetFocus
        Exit Sub
    End If
    If txtbxdataFin.Text = "" Then
        MsgBox "Informe a data Final!", vbExclamation
        txtbxdataFin.SetFocus
        Exit Sub
    End If
    
    P_Status "Aguarde..."
    If optnSint.Value Then
        BuscaSint
    Else
        BuscaDet
    End If
    P_Status ""
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmndbtnGraf_Click()
    Dim b As BlChart
    Dim iCount As Integer
    On Error GoTo trata_erro
    
    If ado.RecordCount = 0 Then
        MsgBox "Não há o que ser gerado!", vbExclamation
        Exit Sub
    End If
    With ado
        .MoveFirst
        b.count = .RecordCount
        For iCount = 0 To .RecordCount - 1
            
            b.val(iCount).caption = !nome_pla
            b.val(iCount).Value = !total
            .MoveNext
        Next
    End With
    'P_Graph  b, 1, True, "Gráfico de Análise de " & _
        IIf(optnRec.Value, "Receitas", "Despesas") & " De: " & txtbxDataIni.Text & _
        " Até: " & txtbxDAtaFIn.Text, "Fonte: " & constSigla_sis & " " & constVersao_Sis
    
    Exit Sub
trata_erro:
    MostraErro
End Sub
Sub imprimeSint(obj As Object)
    Dim max As Integer
    max = 40
    P_CabRelTimbr "Análise Sintética de " & _
    IIf(optnDesp.Value, "Despesas", "Receitas"), obj
    P_ImprLinha obj, "-"
    Dim iCount As Integer
    If cmbxContas.Text = "" Then
        P_ImprX obj, "Conta: (Todas) ", 0, True
        
    Else
        P_ImprX obj, "Conta: " & cmbxContas.Text, 0, True
    End If
    P_ImprX obj, "De: " & txtbxDataIni.Text, 50
    P_ImprX obj, "Até: " & txtbxdataFin.Text, 80, False
    P_ImprLinha obj, "-"
    P_ImprX obj, "Conta", 0
    P_ImprX obj, "T", 85
    P_ImprX obj, "Valor", 120, False, True
    
    P_ImprLinha obj, "-"
    With ado
        .MoveFirst
        Do While Not .EOF
            
            If iCount = max Then
                iCount = 0
                P_RodapeVis obj, strNomeSis & constVersao_Sis, strFantasia_Emp
                P_ImprX obj, "Conta", 0
                P_ImprX obj, "T", 85
                P_ImprX obj, "Valor", 120, False, True
                P_ImprLinha obj, "-"
            End If
            
            P_ImprX obj, !nome_pla, 0
            P_ImprX obj, !tipo_pla, 85
            P_ImprX obj, FormatVal(!total), 120, False, True
            
            
            iCount = iCount + 1
            .MoveNext
        Loop
        P_ImprLinha obj
        obj.FontBold = True
        P_ImprX obj, "Total...: " & lblTotal.caption, 120, False, True
        obj.FontBold = False
    End With
    
    P_RodapeVis obj, strNomeSis & constVersao_Sis, strFantasia_Emp
    
End Sub
Sub imprimeDet(obj As Object)
    Dim max As Integer
    max = 40
    P_CabRelTimbr "Análise Detalhada de " & _
    IIf(optnDesp.Value, "Despesas", "Receitas"), obj
    P_ImprLinha obj, "-"
    Dim iCount As Integer
    If cmbxContas.Text = "" Then
        P_ImprX obj, "Conta: (Todas) ", 0, True
        
    Else
        P_ImprX obj, "Conta: " & cmbxContas.Text, 0, True
    End If
    P_ImprX obj, "De: " & txtbxDataIni.Text, 50
    P_ImprX obj, "Até: " & txtbxdataFin.Text, 80, False
    P_ImprLinha obj, "-"
    P_ImprX obj, "Conta", 0
    P_ImprX obj, "T", 35
    P_ImprX obj, "Valor", 50, True, True
    P_ImprX obj, "Data", 55
    P_ImprX obj, "Histórico", 75
    P_ImprX obj, "C/C.", 165
    P_ImprX obj, "Usu.", 190, False
    P_ImprLinha obj, "-"
    With ado
        .MoveFirst
        Do While Not .EOF
            
            If iCount = max Then
                iCount = 0
                P_RodapeVis obj, strNomeSis & constVersao_Sis, strFantasia_Emp
                P_ImprX obj, "Conta", 0
                P_ImprX obj, "T", 35
                P_ImprX obj, "Valor", 50, True, True
                P_ImprX obj, "Data", 55
                P_ImprX obj, "Histórico", 75
                P_ImprX obj, "C/C.", 165
                P_ImprX obj, "Usu.", 190, False
                P_ImprLinha obj, "-"
            End If
            
            P_ImprX obj, !nome_pla, 0
            P_ImprX obj, !tipo_pla, 35
            P_ImprX obj, FormatVal(!total), 50, True, True
            P_ImprX obj, !data_fin, 55
            P_ImprX obj, Left(!historico_fin, 40), 75
            P_ImprX obj, F_BuscaUmCampo("nome_con", "tab_Contas", _
                "cod_con='" & !cod_con & "'"), 165
            P_ImprX obj, !cod_usu, 190, False
            
            iCount = iCount + 1
            .MoveNext
        Loop
        P_ImprLinha obj
        obj.FontBold = True
        P_ImprX obj, "Total...: " & lblTotal.caption, 50, False, True
        obj.FontBold = False
    End With
    
    P_RodapeVis obj, strNomeSis & constVersao_Sis, strFantasia_Emp
    
End Sub

Private Sub cmndbtnImprimir_Click()
    On Error GoTo trata_erro
    With ado
        If .RecordCount = 0 Then
            MsgBox "Não há o que ser impresso!", vbExclamation
            Exit Sub
        End If
        If optnDet.Value Then
            imprimeDet frmVisualiza.Pic
            frmVisualiza.Show 1
            If frmVisualiza.tag <> "" Then
                imprimeDet Printer
                Printer.EndDoc
            End If
            Unload frmVisualiza
        End If
        If optnSint.Value Then
            imprimeSint frmVisualiza.Pic
            frmVisualiza.Show 1
            If frmVisualiza.tag <> "" Then
                imprimeSint Printer
                Printer.EndDoc
            End If
            Unload frmVisualiza
        End If
    End With
Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub Form_Load()
    
    optnRec_Click
    optnSint_Click
    Set ado = New ADODB.Recordset
    Set adoDoc = New ADODB.Recordset
    resizeForm Me
End Sub

Private Sub optnDesp_Click()
    P_PreenchePlanoContas cmbxContas, False
End Sub

Private Sub optnDesp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub Form_Resize()
    
    PosicionaGrade sstb, Me, pcTop
    
    PosicionaGrade grid, sstb, sstb
    PosicionaGrade gridDoc, sstb, sstb
    
    
End Sub

Private Sub optnDet_Click()
    Dim iCount As Integer
    For iCount = 3 To 6
        dbgrd.Columns(iCount).Visible = True
    Next
    cmndbtnGraf.Enabled = False
End Sub

Private Sub optnDet_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnRec_Click()
    P_PreenchePlanoContas cmbxContas, True
End Sub

Private Sub optnRec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub optnSint_Click()
    Dim iCount As Integer
    On Error Resume Next
    For iCount = 3 To 6
        dbgrd.Columns(iCount).Visible = False
        dbgrdDoc.Columns(iCount).Visible = False
    Next
    cmndbtnGraf.Enabled = True
End Sub

Private Sub optnSint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
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
