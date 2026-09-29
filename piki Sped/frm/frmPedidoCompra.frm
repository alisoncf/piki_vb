VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPedidoCompra 
   Caption         =   "Pedido de Compra"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmPedidoCompra.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   15240
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5580
      Left            =   0
      ScaleHeight     =   5550
      ScaleWidth      =   2865
      TabIndex        =   30
      Top             =   1335
      Width           =   2895
      Begin VB.PictureBox pcInfoPro 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3825
         ScaleWidth      =   5985
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   6015
         Begin VB.Label lblNota 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00000000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   45
            Top             =   1920
            Width           =   840
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
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
            TabIndex        =   44
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblUltCusto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Último Valor:"
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
            TabIndex        =   42
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblDataNota 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00/00/0000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   41
            Top             =   1320
            Width           =   960
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota:"
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
            TabIndex        =   40
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lblFornecedor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "000000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   630
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor:"
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
            TabIndex        =   38
            Top             =   480
            Width           =   885
         End
         Begin VB.Label lblUltQtd 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1440
            TabIndex        =   37
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd"
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
            Left            =   1560
            TabIndex        =   36
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   15210
      TabIndex        =   14
      Top             =   6915
      Width           =   15240
      Begin VB.CommandButton Command1 
         Caption         =   "Confirmar Entrega"
         Height          =   615
         Left            =   8280
         Picture         =   "frmPedidoCompra.frx":0A22
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   14400
         Picture         =   "frmPedidoCompra.frx":16EC
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnConcluir 
         Caption         =   "Concluir"
         Height          =   615
         Left            =   6720
         Picture         =   "frmPedidoCompra.frx":1857
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   5160
         Picture         =   "frmPedidoCompra.frx":1BE1
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   3600
         Picture         =   "frmPedidoCompra.frx":3FA8
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.PictureBox pcPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H0071CACE&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   945
         ScaleWidth      =   3225
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   3255
         Begin VB.Label lblUsu 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "000000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   960
            TabIndex        =   25
            Top             =   720
            Width           =   630
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuário:"
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
            TabIndex        =   24
            Top             =   720
            Width           =   600
         End
         Begin VB.Label lblFin 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2520
            TabIndex        =   23
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Finalizado:"
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
            Left            =   1680
            TabIndex        =   22
            Top             =   480
            Width           =   765
         End
         Begin VB.Label lblIni 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   960
            TabIndex        =   21
            Top             =   480
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Iniciado:"
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
            TabIndex        =   20
            Top             =   480
            Width           =   585
         End
         Begin VB.Label lblData 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "00/00/0000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
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
            TabIndex        =   18
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblPedido 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "000000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   960
            TabIndex        =   17
            Top             =   0
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pedido Nº:"
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
            TabIndex        =   16
            Top             =   0
            Width           =   750
         End
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   15210
      TabIndex        =   13
      Top             =   0
      Width           =   15240
      Begin VB.PictureBox pcOpcao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4680
         ScaleHeight     =   585
         ScaleWidth      =   4425
         TabIndex        =   32
         Top             =   0
         Width           =   4455
         Begin VB.OptionButton optnTodos 
            Caption         =   "Mostrar Todos os Itens"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Value           =   -1  'True
            Width           =   3135
         End
         Begin VB.OptionButton optnForn 
            Caption         =   "Mostrar Somente Itens do Fornecedor"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   3135
         End
      End
      Begin VB.CommandButton cmndbtnAdd 
         Caption         =   "Adicionar"
         Height          =   285
         Left            =   7800
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   9
         Text            =   "0,0000"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   7
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtbxFornecedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmndbtnForn 
         Height          =   285
         Left            =   4320
         Picture         =   "frmPedidoCompra.frx":43B5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnProduto 
         Height          =   285
         Left            =   4320
         Picture         =   "frmPedidoCompra.frx":473F
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   6720
         TabIndex        =   10
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário:"
         Height          =   195
         Left            =   5640
         TabIndex        =   8
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd:"
         Height          =   195
         Left            =   4680
         TabIndex        =   6
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   4320
      Top             =   7200
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
      Bindings        =   "frmPedidoCompra.frx":4AC9
      Height          =   5295
      Left            =   3120
      TabIndex        =   46
      Top             =   1440
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9340
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
Attribute VB_Name = "frmPedidoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmbxForn_Change()

End Sub

Private Sub cmbxForn_Click()
    carrega
End Sub

Private Sub cmbxForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxForn.ListIndex = -1
    End If
End Sub

Private Sub cmndbtnAdd_Click()
    If txtbxProduto.tag = "" Then
        MsgBox "Informe o produto!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxTotal.Text) <= 0 Then
        MsgBox "Informe o valor do produto!", vbExclamation
        txtbxQtd.SetFocus
        Exit Sub
    End If
    If txtbxFornecedor.tag = "" Then
        MsgBox "Informe o fornecedor!", vbExclamation
        txtbxFornecedor.SetFocus
        Exit Sub
    End If
    If Me.tag = "" Then
        Me.tag = pedidoCompra.F_GravaPedidoCompra()
        If Me.tag = "" Then
            MsgBox "Erro na abertura do pedido de compra!", vbExclamation
            Exit Sub
        End If
        
    End If
    pedidoCompra.P_GravaItPed Me.tag, txtbxProduto.tag, txtbxQtd.Text, txtbxUnit.Text, txtbxFornecedor.tag, CStr(Date)
    carrega
    txtbxProduto.SetFocus
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If Me.tag <> "" Then
        strSQL = "select * from tab_pedidocompra where seq_ped=" & Me.tag
        rPesq.Open strSQL, conn_client
        With rPesq
            If .RecordCount > 0 Then
                lblData.caption = !data_ped
                lblFin.caption = !horafec_ped
                lblIni.caption = !horaini_ped
                lblPedido.caption = !seq_ped
                lblUsu.caption = F_BuscaUmCampo("guerra_fun", TAB_FUNCIONARIOS, "cod_fun='" & !cod_usu & "'")
                pcPedido.Visible = True
                If !status_ped <> "1" Then
                    cmndbtnCancel.Enabled = False
                    cmndbtnConcluir.Enabled = False
                    cmndbtnAdd.Enabled = False
                    
                End If
            End If
            .Close
        End With
    
        strSQL = F_SelectItPed(Me.tag, IIf(optnForn.Value = True, txtbxFornecedor.tag, ""))
        dt.RecordSource = strSQL
        dt.Refresh
    
    End If
    
    
End Sub

Private Sub cmndbtnCancel_Click()
    If Me.tag = "" Then
        Unload Me
    Else
    
        If MsgBox("Deseja realmente cancelar o pedido de compra " & Me.tag & "?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        strSQL = "delete from tab_pedidocompra where seq_ped=" & Me.tag
        conn.Execute strSQL
        strSQL = "delete from tab_itped where seq_ped=" & Me.tag
        conn.Execute strSQL
        MsgBox "Pedido Excluído!", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmndbtnConcluir_Click()
    Dim strSQL As String
    If Me.tag = "" Then
        MsgBox "Você não pode concluir um processo de compra não iniciado!", vbExclamation
        Exit Sub
    End If
    pedidoCompra.P_AlteraPedidoCompra Me.tag, Format(Time, "hh:mm"), "0"
    ImprimePedidoCompra Me.tag
    Unload Me
    
End Sub

Private Sub cmndbtnForn_Click()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_BuscaFornecedor(txtbxFornecedor.Text)
    If f.cod_for <> "" Then
        txtbxFornecedor.Text = f.nome_for
        txtbxFornecedor.tag = f.cod_for
        txtbxProduto.SetFocus
    End If
End Sub

Private Sub cmndbtnImprimir_Click()
    If Me.tag = "" Then
        MsgBox "Não há o que imprimir!", vbExclamation
        Exit Sub
    End If
    ImprimePedidoCompra Me.tag
    
End Sub

Private Sub cmndbtnProduto_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produtos.Produto
    Dim i As Produtos.InfoProd
    p = Produtos.F_BuscaProdSaida(txtbxProduto.Text)
    If p.cod_pro <> "" Then
        txtbxProduto.Text = p.nome_pro
        txtbxProduto.tag = p.cod_pro
        dtProd.RecordSource = Produtos.F_SelectInfoProd(p.cod_pro)
        dtProd.Refresh
        txtbxQtd.SetFocus
        
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtProd_Reposition()
    With dtProd.Recordset
        If .RecordCount > 0 Then
            lblDataNota.caption = !entrada_not
            lblUltCusto.caption = FormatVal(!unit_ite)
            lblNota.caption = !Num_not
            lblFornecedor.caption = !nome_for
            lblUltQtd.caption = !qtd_ite * !qtd_emb
            pcInfoPro.Visible = True
            dtProd.caption = .AbsolutePosition + 1 & "/" & .RecordCount
        Else
            pcInfoPro.Visible = False
        End If
        
    End With
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    
    If Me.tag <> "" Then
        carrega
    End If
End Sub

Private Sub Form_Load()
    resizeForm Me
    dt.ConnectionString = strConnect
    dtProd.ConnectionString = strConnect
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.Height = pcBot.Top - dbgrd.Top - 200
    dbgrd.width = Me.width - dbgrd.Left - 200
End Sub

Private Sub optnForn_Click()
    carrega
End Sub

Private Sub optnTodos_Click()
    carrega
End Sub

Private Sub txtbxFornecedor_GotFocus()
    ganhaFoco txtbxFornecedor
End Sub

Private Sub txtbxFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnForn_Click
    End If
End Sub

Private Sub txtbxProduto_GotFocus()
    ganhaFoco txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnProduto_Click
    End If
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
    PerdeFoco txtbxQtd
    P_Total
End Sub

Private Sub txtbxTotal_GotFocus()
    ganhaFoco txtbxTotal
End Sub

Private Sub txtbxTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxTotal_LostFocus()
    PerdeFoco txtbxTotal
    P_Unit
End Sub
Sub P_Total()
    txtbxTotal.Text = FormatVal(CCur(txtbxQtd.Text) * CCur(txtbxUnit.Text))
End Sub
Sub P_Unit()
    On Error Resume Next
    If CCur(txtbxQtd.Text) > 0 Then
        txtbxUnit.Text = FormatVal(CCur(txtbxTotal.Text) / CCur(txtbxQtd.Text), "###,##0.0000")
    Else
        txtbxUnit.Text = "0"
    End If
End Sub
Private Sub txtbxUnit_GotFocus()
    ganhaFoco txtbxUnit
End Sub

Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxUnit_LostFocus()
    PerdeFoco txtbxUnit, "###,##0.0000"
    P_Total
End Sub
