VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSaidas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caixa"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Totalização:"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   5040
      Width           =   9975
      Begin VB.Label lblFinan 
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
         Left            =   8400
         TabIndex        =   22
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finan:"
         Height          =   195
         Left            =   7920
         TabIndex        =   21
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblAprazo 
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
         Left            =   6600
         TabIndex        =   20
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Prazo:"
         Height          =   195
         Left            =   6000
         TabIndex        =   19
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAvista 
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
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Vista:"
         Height          =   195
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   540
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
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Efetivo:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblTotalCanc 
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
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelados:"
         Height          =   195
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtbxDataFin 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cmbxCaixa 
         Height          =   315
         ItemData        =   "frmSaidas.frx":0000
         Left            =   120
         List            =   "frmSaidas.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lbldataFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblDataIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Caixa:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         Picture         =   "frmSaidas.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1755
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmSaidas.frx":0182
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   160
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         Picture         =   "frmSaidas.frx":050C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   120
         Picture         =   "frmSaidas.frx":0656
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2520
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   3240
      Top             =   4920
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
      Bindings        =   "frmSaidas.frx":09E0
      Height          =   4575
      Left            =   2280
      TabIndex        =   23
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
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
Attribute VB_Name = "frmSaidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxCaixa_Click()
    busca
End Sub

Private Sub cmbxCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxCaixa.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanCaixa
    fr.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        frmEdiCai.Show 0
    ElseIf KeyCode = vbKeyF7 Then
           If strFechacai_usu = "S" Then
        frmFechaCaixa.Show 1
    Else
        MsgBox "Você não tem permissão para fechar o caixa!", vbExclamation
    End If
    ElseIf KeyCode = vbKeyF5 Then
        frmRelVendas.Show 0
    ElseIf KeyCode = vbKeyF4 Then
        frmEfetivaVen.Show 0
    ElseIf KeyCode = vbKeyF11 Then
        frmFechaFinan.Show 0
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dtServ.ConnectionString = strConnect
    P_PreencheCaixas cmbxCaixa
    txtbxdataFin.Text = Date
    txtbxDataIni.Text = Date
    
End Sub
Sub busca()
    Dim strSQL As String
    P_Status "Aguarde..."
    If txtbxdataFin.Visible = True Then
        If F_DataMaior(txtbxDataIni.Text, txtbxdataFin.Text) Then
            MsgBox "A data final deve ser maior que a inicial!", vbExclamation
            txtbxdataFin.SetFocus
            Exit Sub
        End If
    End If
    
    strSQL = "select tv.*,ts.*, tf.* from " & TAB_VENDAS & " as tv, cupons as ts,  " & TAB_VENFEC & "  as tf "
    
    strSQL = strSQL & " where "
    strSQL = strSQL & "tv.num_ven=ts.num_ven   "
    strSQL = strSQL & " and tv.num_ven=tf.num_ven "
    strSQL = strSQL & " and ts.data_cup >= " & F_GravaDataTempo(txtbxDataIni.Text)
    strSQL = strSQL & " and ts.data_cup <= " & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:59")
    If cmbxCaixa.ListIndex > -1 Then
        strSQL = strSQL & " and ts.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    End If
    strSQL = strSQL & " order by ts.cod_cai,emissao_sai "
    
    dtServ.RecordSource = strSQL
    dtServ.Refresh
    P_Status "Somando..."
    soma
    P_Status ""
End Sub
Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim cCanc As Currency
    Dim cAv As Currency
    Dim cAp As Currency
    Dim cFin As Currency
    On Error Resume Next
    With dtServ.Recordset
        Do While Not .EOF
            If !serie_sai = "CUPM" Then
                cTotal = cTotal + !valor_sai
                cAv = cAv + !avista_ven
                cAp = cAp + !aprazo_ven
                cFin = cFin + !finan_ven
            Else
                cCanc = cCanc + !valor_sai
            End If
            .MoveNext
        Loop
        lblTotal.caption = FormatVal(cTotal)
        lblTotalCanc.caption = FormatVal(cCanc)
        lblAPrazo.caption = FormatVal(cAp)
        lblAvIsta.caption = FormatVal(cAv)
        lblFinan.caption = FormatVal(cFin)
    End With
End Sub

Private Sub Label7_Click()

End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
    busca
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
