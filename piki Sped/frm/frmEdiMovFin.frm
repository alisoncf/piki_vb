VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiMovFin 
   Caption         =   "Movimento de Financeiras"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   Icon            =   "frmEdiMovFin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   10440
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiMovFin.frx":0442
      Height          =   7215
      Left            =   2640
      TabIndex        =   18
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   12726
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
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   3840
      Top             =   7320
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
   Begin VB.Frame frControle 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CheckBox chkAberto 
         Caption         =   "Apenas Parc. em Aberto"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   5640
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Frame Frame5 
         Caption         =   "&Ordenar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
         Begin VB.ComboBox cmbxOrder 
            Height          =   315
            ItemData        =   "frmEdiMovFin.frx":0454
            Left            =   120
            List            =   "frmEdiMovFin.frx":0464
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "&Financeira:"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2175
         Begin VB.ComboBox cmbxFin 
            Height          =   315
            ItemData        =   "frmEdiMovFin.frx":048A
            Left            =   120
            List            =   "frmEdiMovFin.frx":0494
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2175
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxdataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1080
            TabIndex        =   6
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3495
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   2175
         Begin VB.CommandButton cmndbtnBuscar 
            Caption         =   "&Pesquisar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":04C8
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   160
            Width           =   1980
         End
         Begin VB.CommandButton cmndbtnBaixar 
            Caption         =   "&Baixar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":0852
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   795
            Width           =   1980
         End
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "&Imprimir"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":0BDC
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1440
            Width           =   1980
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":0F66
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2085
            Width           =   1980
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiMovFin.frx":10B0
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2745
            Width           =   1980
         End
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total: 0,00"
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
         Left            =   360
         TabIndex        =   17
         Top             =   6000
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmEdiMovFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAberto_Click()
    If chkAberto.Value = 1 Then
        cmndbtnBaixar.Enabled = True
    Else
        cmndbtnBaixar.Enabled = False
    End If
End Sub

Private Sub cmbxFin_Click()
    cmbxFin.tag = F_BuscaUmCampo("cod_fin", "tab_financeiras", "nome_fin='" & _
    cmbxFin.Text & "'")
End Sub

Private Sub cmbxFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxFin.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnBaixar_Click()
    Dim iCount As Integer
    Dim strSQL As String
    If IsLoad("frmListaDoc") Then
        If frmListaDoc.tag = Me.Name Then
            If MsgBox("Confirma baixa de todas as parcelas listadas no valor de " _
                & frmListaDoc.lblTOtal.caption & "?", vbYesNo + vbQuestion) = vbYes Then
                With frmListaDoc.msflxgrd
                    For iCount = 0 To .rows - 1
                        .Row = iCount
                        .col = 0
                        strSQL = "update tab_venfin set baixa_ven=" & F_GravaData(Date)
                        strSQL = strSQL & ",status_ven=2, "
                        strSQL = strSQL & "cod_usu='" & strCod_Usu & "'  "
                        strSQL = strSQL & " where num_ven='" & .Text & "' "
                        conn.Execute strSQL
                        
                    Next
                Unload frmListaDoc
                adc.Refresh
                End With
            End If
        End If
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = "select tf.nome_fin,tvf.*,tc.nome_cli from tab_financeiras as tf, "
    strSQL = strSQL & " tab_venfin as tvf," & TAB_CLIENTES & " as tc where tvf.cod_fin=tf.cod_fin "
    strSQL = strSQL & " and tc.cod_cli=tvf.cod_cli "
    strSQL = strSQL & " and tvf.data_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tvf.data_ven<=" & F_GravaData(txtbxDAtaFIn.Text)
    If chkAberto.Value = 1 Then
        strSQL = strSQL & " and tvf.status_ven=1 "
    End If
    If cmbxFin.tag <> "" Then
        strSQL = strSQL & " and tvf.cod_fin='" & cmbxFin.tag & "' "
    End If
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by nome_fin, data_ven "
    ElseIf cmbxOrder.ListIndex = 1 Then
        strSQL = strSQL & " order by  data_ven,nome_fin,num_ven "
    ElseIf cmbxOrder.ListIndex = 2 Then
        strSQL = strSQL & " order by  nome_cli,data_ven,nome_fin,num_ven "
    Else
        strSQL = strSQL & " order by  num_ven "
    End If
    adc.RecordSource = strSQL
    adc.Refresh
    If adc.Recordset.RecordCount > 0 Then
        cmndbtnBaixar.Enabled = True
        cmndbtnEdit.Enabled = True
        cmndbtnImpr.Enabled = True
    Else
        cmndbtnBaixar.Enabled = False
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
    End If
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    With adc.Recordset
        Do While Not .EOF
            cTotal = cTotal + !valor_ven
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    lblTOtal.caption = "Total " & FormatVal(cTotal)
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtgrd_DblClick()
    With adc.Recordset
        If !status_ven <> 1 Then
            Exit Sub
        End If
        If IsLoad("frmListaDoc") Then
            If frmListaDoc.tag <> Me.Name Then
                MsgBox "A lista está sendo usada por outro formulário!", vbExclamation
                Exit Sub
            End If
            If F_InFlxgrd(frmListaDoc.msflxgrd, !num_ven, 0) = -1 Then
                frmListaDoc.msflxgrd.AddItem !num_ven & Chr(9) & !valor_ven
            End If
            frmListaDoc.soma
        Else
            frmListaDoc.tag = Me.Name
            frmListaDoc.caption = "Documentos de Financeiras"
            frmListaDoc.msflxgrd.Cols = 2
            frmListaDoc.msflxgrd.rows = 0
            frmListaDoc.msflxgrd.AddItem !num_ven & Chr(9) & !valor_ven
            frmListaDoc.Show 0, mdiFrm
            frmListaDoc.soma
        End If
    
    End With
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Load()
    adc.ConnectionString = strConnect
    
    P_PreencheFinanceiras cmbxFin
    cmbxOrder.ListIndex = 0
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
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
