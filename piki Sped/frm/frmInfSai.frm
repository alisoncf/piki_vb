VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInfSai 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informações de Saída"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbxAte 
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtbxDe 
      Height          =   285
      Left            =   240
      MaxLength       =   40
      TabIndex        =   1
      Top             =   5280
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   2160
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   7920
      Picture         =   "frmInfSai.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   5760
      Picture         =   "frmInfSai.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1935
   End
   Begin VB.OptionButton optnFiscal 
      Caption         =   "Informações &Gerais"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton optnGer 
      Caption         =   "Usar &Filtro"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   4800
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmInfSai.frx":0294
      Height          =   4215
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7435
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Até:"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   5040
      Width           =   285
   End
   Begin VB.Label lblDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&De:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   195
      Left            =   4800
      TabIndex        =   12
      Top             =   5160
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr Gerado: :"
      Height          =   195
      Left            =   3840
      TabIndex        =   11
      Top             =   5160
      Width           =   885
   End
   Begin VB.Label lblQtd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   195
      Left            =   4800
      TabIndex        =   10
      Top             =   4920
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd Total:"
      Height          =   195
      Left            =   3840
      TabIndex        =   9
      Top             =   4920
      Width           =   705
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9840
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Produto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   2460
   End
End
Attribute VB_Name = "frmInfSai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Busca
End Sub
Sub Busca()
    Dim strSQL As String
    Dim sTabCab As String
    Dim sTabIte As String
    If optnFiscal.Value Then
        sTabCab = "tab_Compras"
        sTabIte = "tab_itcompras"
    Else
        sTabCab = "tab_Cmp"
        sTabIte = "tab_itcmp"
    End If
    If optnFiscal.Value = False Then
        dtgrd.Columns(0).caption = "NF.:"
        dtgrd.Columns(0).DataField = "nota_sai"
        strSQL = "select tv.data_ven,ts.nota_sai,tv.cod_usu,tv.nomecli_ven,"
        strSQL = strSQL & "ti.unit_ite as unit,ti.qtd_ite as qtd,ti.unit_ite*ti.qtd_ite as total"
        strSQL = strSQL & " from " & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti "
        strSQL = strSQL & " where ti.num_ven=tv.num_ven "
        strSQL = strSQL & " and ti.cod_pro='" & Me.tag & "' "
        If txtbxDe.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDe.Text)
        End If
        If txtbxAte.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxAte.Text)
        End If
        strSQL = strSQL & " order by data_ven "
    Else
        strSQL = "select tv.data_ven,tv.num_ven,tv.cod_usu,tv.nomecli_ven,"
        strSQL = strSQL & "ti.unit_ite as unit,ti.qtd_ite as qtd,ti.unit_ite*ti.qtd_ite as total,tp.nome_pro "
        strSQL = strSQL & " from " & TAB_VENDAS & " as tv, " & TAB_ITVENDAS & " as ti," & TAB_PRODUTOS & " as tp  "
        
        strSQL = strSQL & " where ti.num_ven=tv.num_ven "
        strSQL = strSQL & " and ti.cod_pro='" & Me.tag & "' "
        strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
        strSQL = strSQL & " and situacao_ven=4 "
        If txtbxDe.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDe.Text)
        End If
        If txtbxAte.Text <> "" Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxAte.Text)
        End If
        
        strSQL = strSQL & " order by data_ven "
    End If
    

    adc.RecordSource = strSQL
    adc.Refresh
    With adc.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        Me.caption = Me.tag & " - " & !nome_pro & " [Informações de Entradas]  "
        lblNome.caption = !nome_pro
        
    End With
        soma
    
End Sub
Sub soma()
    Dim cTotal As Currency
    Dim cQtd As Currency
    With adc.Recordset
        Do While Not .EOF
            cTotal = cTotal + !total
            cQtd = cQtd + !qtd
            .MoveNext
        Loop
    End With
    lblQtd.caption = cQtd
    lblTotal.caption = FormatVal(cTotal)
End Sub
Private Sub Form_Load()
    P_CentraForm Me
    txtbxAte.Text = F_UltDiaMes(Date)
    txtbxDe.Text = F_PrimDiaMes(Date)
    adc.ConnectionString = strConnect
    
End Sub

Private Sub optnFiscal_Click()
    Busca
End Sub

Private Sub optnGer_Click()
    Busca
End Sub

Private Sub txtbxAte_GotFocus()
    ganhaFoco txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
    Busca
End Sub

Private Sub txtbxDe_GotFocus()
    ganhaFoco txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub
