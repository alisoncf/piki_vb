VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmParcelasTer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Posição de Débito de Cliente"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12855
      TabIndex        =   11
      Top             =   0
      Width           =   12885
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Escolha um Cliente"
         Top             =   360
         Width           =   4575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   285
         Left            =   4680
         Picture         =   "frmParcelasTer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDataIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   5040
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lbldataFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
         Height          =   195
         Left            =   6120
         TabIndex        =   6
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   525
      End
      Begin VB.Label lblFones 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   720
         TabIndex        =   23
         Top             =   720
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fones:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblCidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   8640
         TabIndex        =   21
         Top             =   600
         Width           =   105
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         Height          =   195
         Left            =   8040
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   8640
         TabIndex        =   19
         Top             =   120
         Width           =   105
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   8640
         TabIndex        =   18
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblCpf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   3840
         TabIndex        =   17
         Top             =   720
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   8040
         TabIndex        =   16
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End:"
         Height          =   195
         Left            =   8040
         TabIndex        =   15
         Top             =   120
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblCodCli 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   0
         Width           =   675
      End
   End
   Begin TabDlg.SSTab sstb 
      Height          =   6495
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Parcelas (docs)"
      TabPicture(0)   =   "frmParcelasTer.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "gridparc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "adcparc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Cheques"
      TabPicture(1)   =   "frmParcelasTer.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridch"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "adcch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Picture3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Vendas"
      TabPicture(2)   =   "frmParcelasTer.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "gridvenda"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "adcvenda"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Picture4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Produtos Comprados"
      TabPicture(3)   =   "frmParcelasTer.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "adcprod"
      Tab(3).Control(1)=   "gridprod"
      Tab(3).ControlCount=   2
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74880
         ScaleHeight     =   465
         ScaleWidth      =   12345
         TabIndex        =   44
         Top             =   5760
         Width           =   12375
         Begin VB.CommandButton Command1 
            Caption         =   "&Imprimir"
            Height          =   495
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   10680
            TabIndex        =   46
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimentação Total:"
            Height          =   195
            Left            =   9000
            TabIndex        =   45
            Top             =   120
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74880
         ScaleHeight     =   465
         ScaleWidth      =   12345
         TabIndex        =   35
         Top             =   5880
         Width           =   12375
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Aberto:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   780
         End
         Begin VB.Label lblCheAberto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   960
            TabIndex        =   42
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A Vencer:"
            Height          =   195
            Left            =   3120
            TabIndex        =   41
            Top             =   120
            Width           =   705
         End
         Begin VB.Label lblCheDevolv 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   3960
            TabIndex        =   40
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quitados:"
            Height          =   195
            Left            =   6360
            TabIndex        =   39
            Top             =   120
            Width           =   675
         End
         Begin VB.Label lblCheQuit 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0005C9EB&
            Height          =   315
            Left            =   7200
            TabIndex        =   38
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimentação Total:"
            Height          =   195
            Left            =   9000
            TabIndex        =   37
            Top             =   120
            Width           =   1500
         End
         Begin VB.Label lblCheTotal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   10680
            TabIndex        =   36
            Top             =   60
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   12345
         TabIndex        =   24
         Top             =   5760
         Width           =   12375
         Begin VB.Label lblDocsTotal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   11160
            TabIndex        =   34
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimentação Total:"
            Height          =   195
            Left            =   9600
            TabIndex        =   33
            Top             =   120
            Width           =   1500
         End
         Begin VB.Label lblDocsQuitados 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0005C9EB&
            Height          =   315
            Left            =   8040
            TabIndex        =   32
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quitados:"
            Height          =   195
            Left            =   7200
            TabIndex        =   31
            Top             =   120
            Width           =   675
         End
         Begin VB.Label lblDocsAVencer 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   5400
            TabIndex        =   30
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A Vencer:"
            Height          =   195
            Left            =   4560
            TabIndex        =   29
            Top             =   120
            Width           =   705
         End
         Begin VB.Label lblDocsVencidos 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   3000
            TabIndex        =   28
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vencidos:"
            Height          =   195
            Left            =   2160
            TabIndex        =   27
            Top             =   120
            Width           =   705
         End
         Begin VB.Label lblDocAbertos 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   960
            TabIndex        =   26
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Aberto:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   780
         End
      End
      Begin MSAdodcLib.Adodc adcparc 
         Height          =   330
         Left            =   1680
         Top             =   5160
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
      Begin MSDataGridLib.DataGrid gridparc 
         Bindings        =   "frmParcelasTer.frx":03FA
         Height          =   5055
         Left            =   240
         TabIndex        =   48
         Top             =   480
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   8916
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "seq_doc"
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
         BeginProperty Column01 
            DataField       =   "venc_itd"
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
         BeginProperty Column02 
            DataField       =   "valor_itd"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "reman_itd"
            Caption         =   "Restam"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.00"
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
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adcch 
         Height          =   330
         Left            =   -73560
         Top             =   5400
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
      Begin MSDataGridLib.DataGrid gridch 
         Bindings        =   "frmParcelasTer.frx":0410
         Height          =   5295
         Left            =   -74760
         TabIndex        =   49
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9340
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "num_Che"
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
         BeginProperty Column01 
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "cco_Che"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adcvenda 
         Height          =   330
         Left            =   -73680
         Top             =   5040
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
      Begin MSDataGridLib.DataGrid gridvenda 
         Bindings        =   "frmParcelasTer.frx":0424
         Height          =   5295
         Left            =   -74880
         TabIndex        =   50
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   9340
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "num_ven"
            Caption         =   "Venda"
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
            DataField       =   "dataefet_ven"
            Caption         =   "Data"
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
            DataField       =   "valor_ven"
            Caption         =   "Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "guerra_fun"
            Caption         =   "Vendedor"
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
            DataField       =   "comprador_ven"
            Caption         =   "Comprador"
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
            BeginProperty Column02 
               Alignment       =   1
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adcprod 
         Height          =   330
         Left            =   -73680
         Top             =   5160
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
      Begin MSDataGridLib.DataGrid gridprod 
         Bindings        =   "frmParcelasTer.frx":043B
         Height          =   5775
         Left            =   -74880
         TabIndex        =   51
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   10186
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cod_pro"
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
            DataField       =   "codfab_pro"
            Caption         =   "Cód.Fab."
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
            DataField       =   "desc_ite"
            Caption         =   "Produto"
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
            DataField       =   "num_ven"
            Caption         =   "Venda"
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
            DataField       =   "data_ven"
            Caption         =   "Data"
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
            BeginProperty Column02 
               ColumnWidth     =   4034,835
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pcTotal 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12855
      TabIndex        =   0
      Top             =   7665
      Width           =   12885
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   735
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnVer 
         Caption         =   "&Ver Compra"
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmParcelasTer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub cmndbtnEdiCli_Click()
    Dim fr As New frmCadCli
    If txtbxNomeCli.tag = "" Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    strHora = Time
    fr.tag = txtbxNomeCli.tag
    
    fr.Show 0
End Sub

Private Sub cmndbtnBuscar_Click()
    buscaCLi
    
    busca
End Sub
Sub busca()
    Dim strSQL As String
    If Me.tag <> "" Then
        buscaCLi
    End If
    strSQL = financeiro.F_SelectContasReceber(txtDe.Text, txtAte.Text, "", "", "", txtbxNomeCli.tag, "", "", False)
    adcparc.RecordSource = strSQL
    adcparc.Refresh
    
    strSQL = Cheques.F_SelectChRec("C", txtbxNomeCli.tag, txtDe.Text, txtAte.Text)
    adcch.RecordSource = strSQL
    adcch.Refresh
    
    
    strSQL = Vendas.F_selectVendas(txtbxNomeCli.tag, txtDe.Text, txtAte.Text, "situacao_ven=4")
    adcvenda.RecordSource = strSQL
    adcvenda.Refresh
    
    strSQL = F_SelectItensCompradosVendas(txtbxNomeCli.tag, txtDe.Text, txtAte.Text)
    adcprod.RecordSource = strSQL
    adcprod.Refresh
    
        
    
    soma
        
        
        
    
End Sub
Sub soma()
    gridparc.Visible = False
    gridprod.Visible = False
    gridvenda.Visible = False
    
    somaDoc
    somaChe
    
    gridparc.Visible = True
    gridprod.Visible = True
    gridvenda.Visible = True
End Sub

Sub somaDoc()
    Dim cTotal As Currency
    Dim cVencidos As Currency
    Dim cAVencer As Currency
    Dim cQuitados As Currency
    With adcparc.Recordset
        Do While Not .EOF
            cTotal = cTotal + !valor_itd
            If !venc_itd < Date Then
                cVencidos = cVencidos + !reman_itd
            Else
                cAVencer = cAVencer + !reman_itd
            End If
            If !reman_itd <= 0 Then
                cQuitados = cQuitados + !valor_itd
            End If
            .MoveNext
        Loop
    End With
    lblDocAbertos.caption = FormatVal(cVencidos + cAVencer)
    lblDocsVencidos.caption = FormatVal(cVencidos)
    lblDocsAVencer.caption = FormatVal(cAVencer)
    lblDocsQuitados.caption = FormatVal(cQuitados)
    lblDocsTotal.caption = FormatVal(cTotal)
    
End Sub
Sub somaChe()
    Dim cTotal As Currency
    Dim cDev As Currency
    Dim cAVencer As Currency
    Dim cQuitados As Currency
    With adcch.Recordset
        Do While Not .EOF
            cTotal = cTotal + !Valor_che
            If !status_che = "1" Then
                cAVencer = cAVencer + !Valor_che
            ElseIf !status_che = "2" Then
                cQuitados = cQuitados + !Valor_che
            Else
                cDev = cDev = !Valor_che
            End If
            
            .MoveNext
        Loop
    End With
    lblCheAberto.caption = FormatVal(cAVencer)
    lblCheDevolv.caption = FormatVal(cDev)
    lblCheQuit.caption = FormatVal(cQuitados)
    
    lblDocsTotal.caption = FormatVal(cTotal)
    
End Sub




Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTodos_Click()
    
End Sub

Private Sub Command2_Click()
    
End Sub


Private Sub Form_Activate()
    If Me.tag <> "" Then
        busca
    End If
End Sub

Private Sub Form_Load()
    'resizeForm Me
    sstb.Tab = 0
    
    txtDe.Text = "01/01/" & Format(Date, "YYYY")
    txtAte.Text = Date
    
    adcparc.ConnectionString = strConnect
    adcch.ConnectionString = strConnect
    
    adcprod.ConnectionString = strConnect
    adcvenda.ConnectionString = strConnect
    
End Sub

Private Sub Form_Resize()
    'msflxgrd.Width = Me.Width - msflxgrd.Left - 100
    'msflxgrd.Height = pcTotal.Top - msflxgrd.Top - 100
End Sub


Private Sub txtbxNomeCli_GotFocus()
    ganhaFoco txtbxNomeCli
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCLi
        
        TrocaTabKey KeyCode
    End If
End Sub
Sub buscaCLi()
    Dim t As cliente
    If Me.tag <> "" Then
        t = F_BuscaCliente(txtbxNomeCli.tag)
        Me.tag = "" 'pra chamar o cliente certeiro da primeira vez q o formulário abrir
    Else
        t = F_BuscaCliente(txtbxNomeCli.Text)
    End If
    If t.cod_cli <> "" Then
        txtbxNomeCli.Text = t.nome_Cli
        txtbxNomeCli.tag = t.cod_cli
        lblCodCli.caption = t.cod_cli
        lblCpf.caption = t.cnpj_Cli
        lblEnd.caption = t.endereco_cli
        lblBairro.caption = t.bairro_cli
        lblCidade.caption = t.cidade_cli
        lblFones.caption = t.telefone_cli
    End If
End Sub


Private Sub txtbxNomeCli_LostFocus()
    AjustaTexto txtbxNomeCli
End Sub

Private Sub txtDe_GotFocus()
    ganhaFoco txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
Private Sub txtAte_GotFocus()
    ganhaFoco txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
    busca
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

