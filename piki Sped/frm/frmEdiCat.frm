VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCat 
   Caption         =   "Edição de Categorias"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   Icon            =   "frmEdiCat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   3135
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   8460
      TabIndex        =   1
      Top             =   5475
      Width           =   8520
      Begin VB.CommandButton cmndSelect 
         Caption         =   "&Selecionar"
         Height          =   735
         Left            =   3360
         Picture         =   "frmEdiCat.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   6840
         Picture         =   "frmEdiCat.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Transferência"
         Height          =   735
         Left            =   5040
         Picture         =   "frmEdiCat.frx":07B7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   0
         Picture         =   "frmEdiCat.frx":0901
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1680
         Picture         =   "frmEdiCat.frx":0C8B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   1320
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
      Bindings        =   "frmEdiCat.frx":0DD5
      Height          =   4815
      Left            =   120
      TabIndex        =   0
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
         DataField       =   "cod_cat"
         Caption         =   "Código"
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
         DataField       =   "nome_cat"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   -1  'True
         AllowSizing     =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4394,835
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   765
   End
End
Attribute VB_Name = "frmEdiCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadCat
    If adc.Recordset.RecordCount > 0 Then
        fr.Tag = adc.Recordset!cod_cat
        fr.txtbxNome.Text = adc.Recordset!nome_cat
        fr.Show 0
    Else
        MsgBox "Não há o que ser Editado!", vbExclamation
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadCat
    frmCadCat.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    If Me.Tag = "" Then
        Unload Me
    Else
        Me.Tag = ""
        Me.Hide
    End If
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub cmndSelect_Click()
    If Me.Tag <> "" Then
        Me.Hide
    End If
End Sub

Private Sub dbgrd_DblClick()
    If Me.Tag = "" Then
        cmndbtnEdit_Click
    Else
        cmndSelect_Click
    End If
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    adc.ConnectionString = strConnect
    
End Sub
Sub busca()
    strSQL = "select * from tab_categorias "
    strSQL = strSQL & " where (nome_cat like '%" & txtNome.Text & "%') "
    strSQL = strSQL & " order by nome_cat "
    adc.RecordSource = strSQL
    adc.Refresh
End Sub
Private Sub txtNome_Change()
    busca
End Sub
