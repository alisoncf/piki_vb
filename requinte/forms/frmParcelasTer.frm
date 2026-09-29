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
      TabIndex        =   5
      Top             =   0
      Width           =   12885
      Begin VB.OptionButton optnRep 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Representante"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton optnCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cliente"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "Escolha um Cliente"
         Top             =   360
         Width           =   4575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   285
         Left            =   4680
         Picture         =   "frmParcelasTer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   255
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
         TabIndex        =   20
         Top             =   720
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fones:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
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
         Left            =   5400
         TabIndex        =   18
         Top             =   720
         Width           =   105
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         Height          =   195
         Left            =   5400
         TabIndex        =   17
         Top             =   480
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
         Left            =   7800
         TabIndex        =   16
         Top             =   240
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
         Left            =   7800
         TabIndex        =   15
         Top             =   720
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
         Left            =   5400
         TabIndex        =   14
         Top             =   240
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   7800
         TabIndex        =   13
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End:"
         Height          =   195
         Left            =   7800
         TabIndex        =   12
         Top             =   0
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Left            =   5400
         TabIndex        =   11
         Top             =   0
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
         TabIndex        =   10
         Top             =   0
         Width           =   675
      End
   End
   Begin TabDlg.SSTab sstb 
      Height          =   6495
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Carnês"
      TabPicture(0)   =   "frmParcelasTer.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtgrdDoc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "adcDoc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Cheques"
      TabPicture(1)   =   "frmParcelasTer.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dtgrdCheque"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Vendas"
      TabPicture(2)   =   "frmParcelasTer.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "dtgrdVendas"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Produtos Comprados"
      TabPicture(3)   =   "frmParcelasTer.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dtgProd"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Produtos Comprados de Representantes"
      TabPicture(4)   =   "frmParcelasTer.frx":03FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "dtgCons"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Acertos de Representantes"
      TabPicture(5)   =   "frmParcelasTer.frx":0416
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "dtgAcertos"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74880
         ScaleHeight     =   465
         ScaleWidth      =   12345
         TabIndex        =   41
         Top             =   5880
         Width           =   12375
         Begin VB.CommandButton Command1 
            Caption         =   "&Imprimir"
            Height          =   495
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimentação Total:"
            Height          =   195
            Left            =   9000
            TabIndex        =   42
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
         TabIndex        =   32
         Top             =   5880
         Width           =   12375
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Aberto:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A Vencer:"
            Height          =   195
            Left            =   3120
            TabIndex        =   38
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
            TabIndex        =   37
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quitados:"
            Height          =   195
            Left            =   6360
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimentação Total:"
            Height          =   195
            Left            =   9000
            TabIndex        =   34
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
            TabIndex        =   33
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
         TabIndex        =   21
         Top             =   5880
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
            TabIndex        =   31
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimentação Total:"
            Height          =   195
            Left            =   9600
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quitados:"
            Height          =   195
            Left            =   7200
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A Vencer:"
            Height          =   195
            Left            =   4560
            TabIndex        =   26
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
            TabIndex        =   25
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vencidos:"
            Height          =   195
            Left            =   2160
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Aberto:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   780
         End
      End
      Begin MSAdodcLib.Adodc adcDoc 
         Height          =   375
         Left            =   120
         Top             =   6000
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
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
      Begin MSDataGridLib.DataGrid dtgrdDoc 
         Bindings        =   "frmParcelasTer.frx":0432
         Height          =   5175
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   9128
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
      Begin MSDataGridLib.DataGrid dtgrdCheque 
         Bindings        =   "frmParcelasTer.frx":0447
         Height          =   5175
         Left            =   -74880
         TabIndex        =   46
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9128
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
      Begin MSDataGridLib.DataGrid dtgrdVendas 
         Bindings        =   "frmParcelasTer.frx":045F
         Height          =   5415
         Left            =   -74880
         TabIndex        =   47
         Top             =   360
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9551
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin MSDataGridLib.DataGrid dtgProd 
         Bindings        =   "frmParcelasTer.frx":0477
         Height          =   5055
         Left            =   -74880
         TabIndex        =   48
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin MSDataGridLib.DataGrid dtgCons 
         Bindings        =   "frmParcelasTer.frx":048D
         Height          =   4575
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin MSDataGridLib.DataGrid dtgAcertos 
         Bindings        =   "frmParcelasTer.frx":04A5
         Height          =   5775
         Left            =   -74880
         TabIndex        =   50
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   10186
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         TabIndex        =   3
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnVer 
         Caption         =   "&Ver Compra"
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
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

Dim WithEvents adodoc As ADODB.Recordset
Attribute adodoc.VB_VarHelpID = -1
Dim WithEvents adoprod As ADODB.Recordset
Attribute adoprod.VB_VarHelpID = -1
Dim WithEvents adoconsig As ADODB.Recordset
Attribute adoconsig.VB_VarHelpID = -1
Dim WithEvents adoAcertos As ADODB.Recordset
Attribute adoAcertos.VB_VarHelpID = -1
Dim WithEvents adoCheque As ADODB.Recordset
Attribute adoCheque.VB_VarHelpID = -1
Dim WithEvents AdoVendas As ADODB.Recordset
Attribute AdoVendas.VB_VarHelpID = -1


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
    If optnCli.Value Then
        buscaCli
    Else
        buscaRep
    End If
    busca
End Sub
Sub busca()
    Dim strSQL As String
    If Me.tag <> "" Then
        buscaCli
    End If
    strSQL = doc.F_SelectDocsReceber(IIf(optnCli.Value, "C", "R"), , , , , , , , , , , , txtbxNomeCli.tag, -1)
    
    abreObjetoAdo adodoc, strSQL, dtgrdDoc
    
    strSQL = Cheques.F_SelectChRec(IIf(optnCli.Value, "C", "R"), _
        txtbxNomeCli.tag)
    abreObjetoAdo adoCheque, strSQL, dtgrdCheque
    
    
    strSQL = Vendas.F_selectVendas(txtbxNomeCli.tag, "", "", "situacao_ven=4")
    abreObjetoAdo AdoVendas, strSQL, dtgrdVendas
    
    strSQL = F_SelectItensCompradosVendas(txtbxNomeCli.tag)
    abreObjetoAdo adoprod, strSQL, dtgProd
    
    strSQL = Acerto.F_SelectAcertosCli(txtbxNomeCli.tag, IIf(optnCli.Value, "C", "R"))
    abreObjetoAdo adoAcertos, strSQL, dtgAcertos
    
    strSQL = Acerto.F_SelectItAcertosCli(txtbxNomeCli.tag, IIf(optnCli.Value, "C", "R"))
    abreObjetoAdo adoconsig, strSQL, dtgCons
    
    
    
    soma
        
        
        
    
End Sub
Sub soma()
    somaDoc
    somaChe
End Sub

Sub somaDoc()
    Dim cTotal As Currency
    Dim cVencidos As Currency
    Dim cAVencer As Currency
    Dim cQuitados As Currency
    With adodoc
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
    With adoCheque
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


Sub buscaRep()
    Dim t As representante.representante
    
    
    t = representante.F_BuscaRep(txtbxNomeCli.Text)
    If t.cod_Rep <> "" Then
        txtbxNomeCli.Text = t.nome_rep
        txtbxNomeCli.tag = t.cod_Rep
        lblCodCli.caption = t.cod_Rep
        lblCPF.caption = t.cpf_rep
        lblEnd.caption = t.endereco_rep
        lblBairro.caption = t.set.nome
        lblCidade.caption = t.set.Cidade
        lblFones.caption = t.telefone_rep
    End If

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
        P_Status "buscando dados..."
        busca
        P_Status ""
    End If
End Sub

Private Sub Form_Load()
    'resizeForm Me
    sstb.Tab = 0
    Set adoAcertos = New ADODB.Recordset
    Set adoCheque = New ADODB.Recordset
    Set adoconsig = New ADODB.Recordset
    Set adodoc = New ADODB.Recordset
    Set adoprod = New ADODB.Recordset
    Set AdoVendas = New ADODB.Recordset
End Sub

Private Sub Form_Resize()
    'msflxgrd.Width = Me.Width - msflxgrd.Left - 100
    'msflxgrd.Height = pcTotal.Top - msflxgrd.Top - 100
End Sub


Private Sub txtbxNomeCli_GotFocus()
    P_SelecTexto txtbxNomeCli
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCli
        
        Sendkeys Chr(9)
    End If
End Sub
Sub buscaCli()
    Dim t As cliente
    If Me.tag <> "" Then
        t = Clientes.F_BuscaCliente(txtbxNomeCli.tag)
        Me.tag = "" 'pra chamar o cliente certeiro da primeira vez q o formulário abrir
    Else
        t = Clientes.F_BuscaCliente(txtbxNomeCli.Text)
    End If
    If t.cod_cli <> "" Then
        txtbxNomeCli.Text = t.nome_Cli
        txtbxNomeCli.tag = t.cod_cli
        lblCodCli.caption = t.cod_cli
        lblCPF.caption = t.cnpj_Cli
        lblEnd.caption = t.endereco_cli
        lblBairro.caption = t.set.nome
        lblCidade.caption = t.set.Cidade
        lblFones.caption = t.telefone_cli
    End If
End Sub


Private Sub txtbxNomeCli_LostFocus()
    P_AjustaTexto txtbxNomeCli
End Sub
