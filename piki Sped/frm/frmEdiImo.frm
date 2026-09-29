VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiImo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Imobilizados"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   360
      Top             =   3840
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
      Bindings        =   "frmEdiImo.frx":0000
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7223
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
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   98
      TabIndex        =   0
      Top             =   4215
      Width           =   7215
      Begin VB.CommandButton cmndbtnExcluir 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   3240
         Picture         =   "frmEdiImo.frx":0012
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiImo.frx":015C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiImo.frx":04E6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   5640
         Picture         =   "frmEdiImo.frx":0630
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
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
      Left            =   600
      TabIndex        =   6
      Top             =   5280
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   405
   End
End
Attribute VB_Name = "frmEdiImo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadImo
    With adc.Recordset
        fr.Tag = !cod_imo
        fr.txtbxNome.Text = !desc_IMO
        fr.txtbxQtd.Text = FormatVal(!qtd_imo)
        fr.txtbxUnit.Text = FormatVal(!unit_imo)
        fr.txtbxUni.Text = !unidade_imo
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnExcluir_Click()
    If MsgBox("Confirma exclusão do imobilizado?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from tab_imobilizado where cod_imo='" & adc.Recordset!cod_imo & "' "
        conn.Execute strSQL
        busca
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadImo
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        strSQL = "select * from tab_imob order by cod_imo"
    Else
        strSQL = "select * from tab_imob order by desc_imo"
    End If
    
    
    adc.RecordSource = strSQL
    adc.Refresh
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    busca
    
End Sub
Sub busca()
    strSQL = "select ti.*,ti.unit_imo*ti.qtd_imo as total from tab_imob as ti"
    adc.ConnectionString = strConnect
    adc.RecordSource = strSQL
    adc.Refresh
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    With adc.Recordset
        Do While Not .EOF
            cTotal = cTotal + !total
            .MoveNext
        Loop
        
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        lblTotal.caption = FormatVal(cTotal)
    End With
    
End Sub
Private Sub Form_Load()
    'MudaAparencia Me
    'Me.Height = 5205
    'Me.Width = 5145
    P_CentraForm Me
    
End Sub

