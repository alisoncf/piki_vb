VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDemonstrativoInvent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Demonstrativo de Estoque"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Sal&var"
      Height          =   735
      Left            =   6480
      Picture         =   "frmDemonstrativoInvent.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Width           =   1335
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   9105
      TabIndex        =   8
      Top             =   5520
      Width           =   9135
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Itens:"
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   390
      End
      Begin VB.Label lblItens 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   480
         TabIndex        =   11
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soma:"
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   0
         Width           =   450
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   2400
         TabIndex        =   9
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordenar Por:"
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmDemonstrativoInvent.frx":015C
         Left            =   120
         List            =   "frmDemonstrativoInvent.frx":0169
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Por:"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   1695
      Begin VB.OptionButton Option1 
         Caption         =   "Preço de Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optnCusto 
         Caption         =   "Custo"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   3600
      Picture         =   "frmDemonstrativoInvent.frx":0187
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   735
      Left            =   5040
      Picture         =   "frmDemonstrativoInvent.frx":0280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   7920
      Picture         =   "frmDemonstrativoInvent.frx":068D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   1200
      Top             =   4680
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
      Bindings        =   "frmDemonstrativoInvent.frx":0782
      Height          =   4815
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8493
      _Version        =   393216
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
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4394,835
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDemonstrativoInvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Busca
    
End Sub
Sub Busca()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim sField As String
    Dim sCampo As String
    If pc.BackColor = vbGreen Then
        sField = "ef_pro"
    Else
        sField = "er_pro"
    End If
    
    
    If optnCusto.Value Then
        If pc.BackColor = vbGreen Then
            sCampo = "cf_pro"
        Else
            sCampo = "cr_pro"
        End If
    Else
        sCampo = "pv_pro"
    End If
    strSQL = "select tp.*," & sCampo & " * " & sField & " as total "
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " where " & sField & " >0 "
    
    If cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by st_pro,nome_pro "
    ElseIf cmbxOrder.ListIndex = 0 Then
        strSQL = strSQL & " order by st_pro,cod_Pro "
    Else
        strSQL = strSQL & " order by st_pro,total"
    End If
    
    
    dt.RecordSource = strSQL
    dt.Refresh
    
    With dt.Recordset
        
        If .RecordCount > 0 Then
            Do While Not .EOF
                cTotal = cTotal + !total
                .MoveNext
            Loop
            .MoveFirst
        
            
        End If
        
        lblItens.caption = .RecordCount
        lblTotal.caption = FormatVal(cTotal)
        
        
        
        
    End With
    
    
    
End Sub

Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    Dim sPage As String
    Dim cTotal As Currency
    P_CabRelTimbr "DEMONSTRATIVO DE INVENTÁRIO EM " & Now, obj
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth, obj.CurrentY
    obj.FontName = "arial"
    obj.fontsize = 10
    P_ImprX obj, "UN.", 5
    P_ImprX obj, "QTD.", 15
    P_ImprX obj, "DISCRIMINAÇÃO", 35
    P_ImprX obj, "UNITÁRIO", 155, True, True
    P_ImprX obj, "TOTAL", 175, True, True
    P_ImprX obj, "T", 180, False
    P_Line obj, 0, obj.CurrentY, obj.ScaleWidth, obj.CurrentY
    
    With dt.Recordset
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 25 Then
                If TypeName(obj) = "Printer" Then
                    sPage = obj.Page
                Else
                    sPage = ""
                End If
                
                
                P_RodapeVis obj, "Continua...", "Pág: " & sPage
                P_ImprX obj, "UN.", 5
                P_ImprX obj, "QTD.", 15
                P_ImprX obj, "DISCRIMINAÇÃO", 35
                P_ImprX obj, "UNITÁRIO", 155, True, True
                P_ImprX obj, "TOTAL", 175, True, True
                P_ImprX obj, "T", 180, False
                P_Line obj, 0, obj.CurrentY, obj.ScaleWidth, obj.CurrentY
            End If
            
            
            
            P_ImprX obj, !unidade_pro, 5
            P_ImprX obj, FormatVal(dbgrd.Columns(3).Text), 15
            P_ImprX obj, !nome_pro, 35
            P_ImprX obj, FormatVal(dbgrd.Columns(4).Text), 155, True, True
            P_ImprX obj, FormatVal(!total), 175, True, True
            P_ImprX obj, FormatVal(!st_pro), 180, False
            .MoveNext
        Loop
        P_Line obj, 0, obj.CurrentY, obj.ScaleWidth, obj.CurrentY
        P_ImprX obj, "Itens.......: " & lblItens.caption, 5
        P_ImprX obj, "Total.......: " & lblTotal.caption, 175, False, True
        
    End With
    
End Sub
Private Sub cmndbtnSair_Click()
    dt.Recordset.Close
    Unload Me
    
End Sub


Private Sub cmndbtnSalvar_Click()
    Dim s As String
    On Error GoTo trata_erro
    Dim cSt As Currency
    Dim sSt As String
    mdiFrm.cmndlg.Filter = "Text (*.txt)|*.txt"
    mdiFrm.cmndlg.ShowSave
    s = mdiFrm.cmndlg.FileName
    MsgBox s
    If s = "" Then
        MsgBox "Não foi informado um nome correto para o arquivo!", vbInformation
        Exit Sub
    End If
    Open s For Output As #1
    Print #1, strNome_Emp
    Print #1, String(80, "-")
    Print #1, "DEMONSTRATIVO DE INVENTÁRIO EM " & Now
    Print #1, String(80, "=")
    With dt.Recordset
        .MoveFirst
        
        
        strSQL = F_EsticaTexto("UN.", 4, 1)
        strSQL = strSQL & F_EsticaTexto("QTD.", 13, 1)
        strSQL = strSQL & F_EsticaTexto("DISCRIMINAÇÃO", 40, 1)
        strSQL = strSQL & F_EsticaTexto("UNITÁRIO", 16, 1)
        strSQL = strSQL & F_EsticaTexto("TOTAL", 8, 1)
        strSQL = strSQL & F_EsticaTexto("T", 1, 1)
        Print #1, strSQL
        Print #1, String(80, "=")
        Do While Not .EOF
            strSQL = F_EsticaTexto(!unidade_pro, 4, 0)
            strSQL = strSQL & F_EsticaTexto(dbgrd.Columns(3).Text, 9, 2) & "  "
            strSQL = strSQL & F_EsticaTexto(!nome_pro, 40, 0)
            strSQL = strSQL & F_EsticaTexto(FormatVal(dbgrd.Columns(4).Text), 10, 2)
            strSQL = strSQL & F_EsticaTexto(FormatVal(!total), 14, 2) & " "
            strSQL = strSQL & F_EsticaTexto(FormatVal(!st_pro), 2, 1)
            Print #1, strSQL
            .MoveNext
        Loop
        Print #1, String(80, "=")
        Print #1, "ITENS: " & lblItens.caption & Space(50) & "TOTAL: " & lblTotal.caption
        Close #1
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
    On Error Resume Next
    Close #1
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        pc.BackColor = vbGreen
        dbgrd.Columns(3).DataField = "ef_pro"
        Busca
    ElseIf KeyCode = vbKeyF9 Then
        pc.BackColor = vbYellow
        dbgrd.Columns(3).DataField = "er_pro"
        Busca
    End If
End Sub

Private Sub Form_Load()
    pc.BackColor = vbGreen
    adc.ConnectionString = strConnect
    cmbxOrder.ListIndex = 0
    P_CentraForm Me
    Busca
End Sub

Private Sub Option1_Click()
    dbgrd.Columns(4).caption = "Preço:"
    dbgrd.Columns(4).DataField = "pv_pro"
End Sub

Private Sub optnCusto_Click()
    dbgrd.Columns(4).caption = "Custo"
    dbgrd.Columns(4).DataField = "cf_pro"
End Sub
