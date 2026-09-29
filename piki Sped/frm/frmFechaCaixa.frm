VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFechaCaixa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8940
   Icon            =   "frmFechaCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frControle 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin TabDlg.SSTab SSTab1 
         Height          =   5295
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   9340
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Movimento à Vista"
         TabPicture(0)   =   "frmFechaCaixa.frx":0442
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "adc"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "dbgrd"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Histórico das Vendas"
         TabPicture(1)   =   "frmFechaCaixa.frx":045E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "grid_vendas"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "adcvendas"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         Begin MSAdodcLib.Adodc adc 
            Height          =   330
            Left            =   -73680
            Top             =   5280
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
            Bindings        =   "frmFechaCaixa.frx":047A
            Height          =   4815
            Left            =   -74880
            TabIndex        =   44
            Top             =   600
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
         Begin MSAdodcLib.Adodc adcvendas 
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
         Begin MSDataGridLib.DataGrid grid_vendas 
            Bindings        =   "frmFechaCaixa.frx":048C
            Height          =   4815
            Left            =   0
            TabIndex        =   45
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
      Begin VB.Frame Frame3 
         Caption         =   "Vendas:"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   5520
         Width           =   6855
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Finan."
            Height          =   195
            Left            =   3240
            TabIndex        =   43
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Left            =   5160
            TabIndex        =   27
            Top             =   240
            Width           =   450
         End
         Begin VB.Label lblTotal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   5760
            TabIndex        =   26
            Tag             =   "!"
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblAVista 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   25
            Tag             =   "!"
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblFinan 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3720
            TabIndex        =   24
            Tag             =   "!"
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblAprazo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2280
            TabIndex        =   23
            Tag             =   "!"
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A Prazo:"
            Height          =   195
            Left            =   1560
            TabIndex        =   22
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A Vista:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Computador:"
         Height          =   495
         Left            =   7080
         TabIndex        =   28
         Top             =   5520
         Width           =   1695
         Begin VB.Label lblComputador 
            AutoSize        =   -1  'True
            Caption         =   "Label4"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "&Data"
         Height          =   2415
         Left            =   7080
         TabIndex        =   1
         Top             =   120
         Width           =   1695
         Begin VB.ComboBox cmbxTurno 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtbxFundo_Cai 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Text            =   "0,00"
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ComboBox cmbxCaixa 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtbxData 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "&Turno"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&Fundo de Caixa:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   1800
            Width           =   1155
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cai&xa"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   390
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   7080
         TabIndex        =   7
         Top             =   2475
         Width           =   1695
         Begin VB.CommandButton cmndbtnImprHist 
            Caption         =   "&Hist. das Vendas"
            Height          =   615
            Left            =   120
            Picture         =   "frmFechaCaixa.frx":049E
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1520
            Width           =   1455
         End
         Begin VB.CommandButton cmndbtnFechar 
            Caption         =   "&Confirmar "
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            Picture         =   "frmFechaCaixa.frx":05E8
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   160
            Width           =   1455
         End
         Begin VB.CommandButton cmndbtnImprREs 
            Caption         =   "&Imprimir Resumo"
            Height          =   615
            Left            =   120
            Picture         =   "frmFechaCaixa.frx":0732
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   120
            Picture         =   "frmFechaCaixa.frx":087C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2235
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Entradas"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   6000
         Width           =   1695
         Begin VB.Label lblEntr 
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
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   600
            TabIndex        =   14
            Tag             =   "!"
            Top             =   285
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Saídas:"
         Height          =   615
         Left            =   1920
         TabIndex        =   15
         Top             =   6000
         Width           =   1695
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblSai 
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   600
            TabIndex        =   16
            Tag             =   "!"
            Top             =   315
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Saldo da Conta:"
         Height          =   615
         Left            =   3720
         TabIndex        =   18
         Top             =   6000
         Width           =   1695
         Begin VB.Label lblSaldo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            TabIndex        =   19
            Tag             =   "!"
            Top             =   315
            Width           =   525
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Totais:"
         Height          =   855
         Left            =   5520
         TabIndex        =   30
         Top             =   6000
         Width           =   3255
         Begin VB.Label lblTotalSaldo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2040
            TabIndex        =   37
            Tag             =   "!"
            Top             =   480
            Width           =   525
         End
         Begin VB.Line Line1 
            X1              =   1800
            X2              =   2280
            Y1              =   720
            Y2              =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saídas"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   510
         End
         Begin VB.Label lblTotalSai 
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   33
            Tag             =   "!"
            Top             =   525
            Width           =   480
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entradas"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lblTotalEnt 
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   840
            TabIndex        =   31
            Tag             =   "!"
            Top             =   165
            Width           =   480
         End
      End
      Begin VB.Label lblDev 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   39
         Tag             =   "!"
         Top             =   6720
         Width           =   390
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Devoluções:"
         Height          =   195
         Left            =   2400
         TabIndex        =   38
         Top             =   6720
         Width           =   900
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   6720
         Width           =   615
      End
      Begin VB.Label lblVendas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1320
         TabIndex        =   35
         Tag             =   "!"
         Top             =   6720
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmFechaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Sub buscaAprazo()
    strSQL = "select tv .num_ven,tv.dataefet_ven, tv.nomecli_ven,tf.avista_ven,"
    strSQL = strSQL & "tf.aprazo_ven, tf.finan_ven from " & TAB_VENDAS & " as tv, "
    strSQL = strSQL & " " & TAB_VENFEC & "  as tf where tv.num_ven=tf.num_ven "
    strSQL = strSQL & " and tv.dataefet_ven=" & F_PesqData(txtbxData.Text)
    strSQL = strSQL & " and tf.cod_Cai='" & Left(cmbxCaixa.Text, 2) & "'"
    strSQL = strSQL & " and tf.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    strSQL = strSQL & " order by tv.num_ven "
    dtAprazo.RecordSource = strSQL
     dtAprazo.Refresh
End Sub
Sub busca(bTipoBusca As Integer)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim cVen As Currency
    Dim cCred As Currency
    Dim cDeb As Currency
    
    If bTipoBusca = 0 Then
        strSQL = "select * from tab_caixa where data_cai=" & F_PesqData(txtbxData.Text)
        strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
        strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
        strSQL = strSQL & " order by tipo_cai,seq_cai "
    Else
        strSQL = "select * from tab_caixa where data_cai=" & F_PesqData(txtbxData.Text)
        strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
        strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
        strSQL = strSQL & " order by tipo_cai,seq_cai "

    End If
    
    dtServ.RecordSource = strSQL
    dtServ.Refresh
    With dtServ.Recordset
        Do While Not .EOF
            If !tipo_cai = "C" Then
                cCred = cCred + !valor_Cai
            Else
                cDeb = cDeb + !valor_Cai
            End If
            .MoveNext
        Loop
        lblEntr.caption = FormatVal(cCred)
        lblSai.caption = FormatVal(cDeb)
        lblSaldo.caption = FormatVal(cCred - cDeb)
        
        lblVendas.caption = FormatVal(F_BuscaUmCampo("sum(valor_cai)", "tab_caixa", _
            "num_ven<>'' and data_cai=" & F_PesqData(txtbxData.Text) & _
            " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' and tipo_Cai='C'  and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "))
            
        lblDev.caption = FormatVal(F_BuscaUmCampo("sum(valor_cai)", "tab_caixa", _
            "num_ven<>'' and data_cai=" & F_PesqData(txtbxData.Text) & _
            " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' and tipo_Cai='D' AND Cod_tur='" & Left(cmbxTurno.Text, 1) & "'"))
        
        If lblDev.caption = "" Then
            lblDev.caption = "0,00"
        End If
        
        If lblTotalEnt.caption = "" Then lblTotalEnt.caption = "0,00"
        
        lblTotalEnt.caption = FormatVal(F_BuscaUmCampo("sum(valor_cai)", "tab_caixa", _
            " tipo_cai='C' and data_cai=" & F_PesqData(txtbxData.Text) & _
            " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' and tipo_Cai='D' AND Cod_tur='" & Left(cmbxTurno.Text, 1) & "'"))
        
        lblTotalSai.caption = FormatVal(F_BuscaUmCampo("sum(valor_cai)", "tab_caixa", _
            " tipo_cai='D' and data_cai=" & F_PesqData(txtbxData.Text) & _
            " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' and tipo_Cai='D' AND Cod_tur='" & Left(cmbxTurno.Text, 1) & "'"))
        If lblTotalSai.caption = "" Then lblTotalSai.caption = "0,00"
        If lblTotalEnt.caption = "" Then
            lblTotalEnt.caption = "0,00"
        End If
        
        lblTotalSaldo.caption = FormatVal(CCur(lblTotalEnt.caption) - CCur(lblTotalSai.caption))
        
        strSQL = "select sum(tf.avista_ven) as avista ,"
        strSQL = strSQL & "sum(tf.aprazo_ven) as aprazo ,"
        strSQL = strSQL & "sum(tf.finan_ven) as finan "
        strSQL = strSQL & " from  " & TAB_VENFEC & "  as tf, " & TAB_VENDAS & " as tv "
        strSQL = strSQL & " where tv.situacao_ven>3"
        strSQL = strSQL & " and tv.num_ven=tf.num_ven"
        strSQL = strSQL & " and tv.dataefet_ven=" & F_PesqData(txtbxData.Text)
        strSQL = strSQL & " and tf.cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' "
        strSQL = strSQL & " and tf.cod_Tur='" & Left(cmbxTurno.Text, 1) & "' "
        rPesq.Open strSQL, conn_client
        
        If rPesq.RecordCount > 0 Then
            lblAprazo.caption = FormatVal(rPesq!aprazo)
             lblAvista.caption = FormatVal(rPesq!avista)
            lblFinan.caption = FormatVal(rPesq!finan)
            lblTotal.caption = FormatVal(rPesq!aprazo + rPesq!avista + rPesq!finan)
        Else
            lblAprazo.caption = "0,00"
            lblAvista.caption = "0,00"
            lblFinan.caption = "0,00"
            lblTotal.caption = "0,00"
        End If
        rPesq.Close
    End With
    



Exit Sub




trata_erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmbxCaixa_Click()
    buscaTodos
End Sub

Private Sub cmbxCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxTurno_Click()
    buscaTodos
End Sub

Private Sub cmndbtnFechar_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = F_BuscaUmCampo("count(*)", "" & TAB_VENDAS & "", "situacao_ven=2")
    If Int(strSQL) > 0 Then
        MsgBox "Não é possível fechar o caixa pois existem vendas a efetivar!", vbExclamation
        Exit Sub
    End If
    
    If F_BuscaUmCampo("data_cai", "tab_fechacaixa", "data_cai=" & _
        F_PesqData(txtbxData.Text) & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "'") <> "" Then
        MsgBox "Este caixa já foi fechado neste dia!", vbExclamation
        Exit Sub
    End If
    If VendasFechadas = True Then
        MsgBox "Existem vendas condicionais a serem efetivadas!", vbExclamation
        On Error Resume Next
        frmEfetivaVen.Show 1
        Exit Sub
    End If
        If F_BuscaUmCampo("status_tur", "tab_turncai", "cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' and cod_tur='" & Left(cmbxTurno.Text, 1) & "' and data_tur=" & F_GravaData(txtbxData.Text)) = "F" Then
            MsgBox "Este turno já se encontra Fechado!", vbCritical
            Exit Sub
         End If
    If MsgBox("Deseja realmente fechar o Turno do dia " & txtbxData.Text _
        & " Com " & txtbxFundo_Cai.Text & " de Fundo de caixa?", vbYesNo + vbCritical) = vbYes Then
        
        If CCur(txtbxFundo_Cai.Text) = 0 Then
            If MsgBox("Este fechamento não contém fundo de caixa!" & Chr(13) & _
                "Deseja realmente continuar?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
             End If
        End If
        strSQL = "update tab_turncai set logout_usu='" & strCod_Usu & "',"
        strSQL = strSQL & "Fechamento_tur=" & F_GravaHora(Now) & ","
        strSQL = strSQL & "valorfechamento_tur=" & F_GravaDec(CCur(lblSaldo.caption) - CCur(txtbxFundo_Cai.Text)) & ","
        strSQL = strSQL & "status_tur='F' "
        strSQL = strSQL & " where cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' "
        strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
        
        conn.Execute strSQL
        strSQL = F_GravaCai("D", txtbxData.Text, CCur(lblSaldo.caption) - CCur(txtbxFundo_Cai.Text), _
            "TRANSF. CAIXA: " & Left(cmbxCaixa.Text, 2) & " TURNO:" & Left(cmbxTurno.Text, 1), "", "TRA", "01", 0, 0, Left(cmbxCaixa.Text, 2), Left(cmbxTurno.Text, 1))
        
        
        strSQL = F_GravaFinanceiro("C", txtbxData.Text, CCur(lblSaldo.caption) - CCur(txtbxFundo_Cai.Text), _
             "TRANSF. CAIXA: " & Left(cmbxCaixa.Text, 2) & " TURNO:" & Left(cmbxTurno.Text, 1), "TRA", constContaCaixa, "", "", "", "")
        MsgBox "Caixa " & cmbxCaixa.Text & "  Fechado com sucesso!", vbExclamation
        cmndbtnImprREs_Click
        If MsgBox("Deseja Fazer Redução 'Z'", vbYesNo + vbQuestion) = vbYes Then
            If Not F_ReducaoZ Then
            
            End If
        End If
    End If
        
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Function VendasFechadas() As Boolean
    If F_BuscaUmCampo("num_ven", "" & TAB_VENDAS & "", "situacao_ven=2") <> "" Then
        VendasFechadas = True
    Else
        VendasFechadas = False
    End If
End Function

Private Sub cmndbtnImprHist_Click()
    P_ImprimeHist frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprimeHist Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub P_ImprimeHist(obj As Object)
    On Error GoTo trata_erro
    Dim rPesq As New adodb.Recordset
    Dim rLan As New adodb.Recordset
    Dim sMsg As String
    P_CabRelTimbr "Histórico das Vendas em: " & txtbxData.Text, obj
    P_ImprX obj, "Venda", 5
    P_ImprX obj, "Cliente", 25
    P_ImprX obj, "A Vista", 135, True, True
    P_ImprX obj, "A prazo", 155, True, True
    P_ImprX obj, "Finan", 175, False, True
    P_ImprLinha obj
    
    With dtAprazo.Recordset
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, constNome_sis
                
            End If
            P_ImprX obj, !num_ven, 5
            P_ImprX obj, !nomecli_ven, 25
            P_ImprX obj, FormatVal(!avista_ven), 135, True, True
            P_ImprX obj, FormatVal(!aprazo_ven), 155, True, True
            P_ImprX obj, FormatVal(!finan_ven), 175, False, True
            strSQL = "select * from tab_doc1 where num_ven='" & !num_ven & "' "
            rPesq.Open strSQL, conn_client
            With rPesq
                If .RecordCount > 0 Then
                    P_ImprX obj, "Documento: " & !num_doc, 10
                    P_ImprX obj, "Valor: " & FormatVal(!valor_doc), 50
                    P_ImprX obj, "Parcelas: " & !Parc_Doc, 105, False
                    
                    strSQL = "select * from tab_itdoc1 where seq_doc='" & !seq_doc & "' "
                    rLan.Open strSQL
                    With rLan
                        Do While Not .EOF
                            P_ImprX obj, "Lançamento:" & !seq_itd, 10
                            P_ImprX obj, "Valor:" & FormatVal(!reman_itd), 50
                            P_ImprX obj, "Vencimento:" & !venc_itd, 105, False
                            .MoveNext
                        Loop
                        
                    End With
                    strSQL = "select * from tab_cheques1 where seq_doc='" & !seq_doc & "' "
                    rLan.Open strSQL
                    With rLan
                        Do While Not .EOF
                            P_ImprX obj, "Ch:" & !num_che, 10
                            P_ImprX obj, "Banco:" & !cod_bco, 28
                            P_ImprX obj, "Ag:" & !age_che, 45
                            P_ImprX obj, "C/C:" & !cco_che, 60
                            P_ImprX obj, "Valor:" & FormatVal(!Valor_che), 80
                            P_ImprX obj, "Vencimento:" & !Venc_che, 105
                            P_ImprX obj, "Titular:" & !nome_che, 145, False
                            
                            .MoveNext
                        Loop
                        
                    End With
                    P_ImprLinha obj, "-"
                End If
                .Close
            End With
            .MoveNext
        Loop
        P_ImprLinha obj, "-"
        P_ImprX obj, "Totais......:", 0
        P_ImprX obj, FormatVal(lblAvista.caption), 135, True, True
        P_ImprX obj, FormatVal(lblAprazo.caption), 155, True, True
        P_ImprX obj, FormatVal(lblFinan.caption), 175, True, True
        obj.FontBold = True
            P_ImprX obj, FormatVal(lblTotal.caption), 200, False, True
        
        obj.FontBold = False
        P_RodapeVis obj, constNome_sis
    End With
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical1
End Sub

Private Sub cmndbtnImprMovCco_Click()
    If dtCco.Recordset.RecordCount = 0 Then
        Exit Sub
    End If
    P_imprMovCco frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_imprMovCco Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub P_imprMovCco(obj As Object)
    P_CabRelTimbr dbgrdCco.caption, obj
    P_ImprX obj, "Seq", 5
    P_ImprX obj, "Histórico", 20
    P_ImprX obj, "Esp", 120
    P_ImprX obj, "T", 130
    P_ImprX obj, "Valor", 150, False, True
    P_ImprLinha obj
    With dtCco.Recordset
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj, constNome_sis
            End If
            P_ImprX obj, !seq_fin, 5
            P_ImprX obj, !historico_fin, 20
            P_ImprX obj, !Esp_fin, 120
            P_ImprX obj, !tipo_fin, 130
            P_ImprX obj, FormatVal(!valor_fin), 150, False, True
            .MoveNext
        Loop
        P_ImprLinha obj, "_"
        
        P_ImprX obj, "Crédito: " & lblCredCco.caption, 60
        P_ImprX obj, "Débito: " & lblDebCco.caption, 100
        P_ImprX obj, "Total: " & lblTotCco.caption, 140, False
        P_RodapeVis obj, constNome_sis
    End With
End Sub
Private Sub cmndbtnImprREs_Click()
    If dtServ.Recordset.RecordCount = 0 Then
        Exit Sub
    End If
    P_imprRes frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_imprRes Printer
    End If
    Unload frmVisualiza
    
End Sub
Sub P_imprRes(obj As Object)
    P_CabRelTimbr "Fechamento do Caixa: " & Right(cmbxCaixa.Text, 8) & " - Turno: " & cmbxTurno.Text & " - " & txtbxData.Text & " Func: " & frControle.caption, obj
    
    P_ImprX obj, "Histórico", 5
    P_ImprX obj, "Epécie", 100
    P_ImprX obj, "T", 115
    P_ImprX obj, "Valor", 140, False, True
    P_ImprLinha obj
    
    P_ImprX obj, "REFERENTE A VENDAS", 5
    P_ImprX obj, "DIN", 100
    P_ImprX obj, "C", 115
    P_ImprX obj, lblVendas.caption, 140, False, True
    
    With dtServ.Recordset
        .MoveFirst
        Do While Not .EOF
            If Printer.CurrentY >= Printer.ScaleHeight - 30 Then
                P_FimRel obj, "Continua"
                P_ImprX obj, "Histórico", 5
                P_ImprX obj, "Epécie", 100
                P_ImprX obj, "T", 115
                P_ImprX obj, "Valor", 140, False, True
                P_ImprLinha obj
            End If
            If !num_ven = "" Then
                P_ImprX obj, !historico_cai, 5
                P_ImprX obj, !Esp_cai, 100
                P_ImprX obj, !tipo_cai, 115
                P_ImprX obj, FormatVal(!valor_Cai), 140, False, True
            End If
            .MoveNext
        Loop
        P_ImprX obj, "REFERENTE A DEVOLUÇÕES", 5
        P_ImprX obj, "DIN", 100
        P_ImprX obj, "D", 115
        P_ImprX obj, lblDev.caption, 140, False, True
        If Printer.CurrentY >= Printer.ScaleHeight - 50 Then
            P_FimRel obj, "Continua"
        End If
        P_ImprLinha obj, "_"
        P_CentraImpressao obj, "Totais"
        P_ImprLinha obj
        
        P_ImprX obj, "Entradas: " & lblEntr.caption, 10
        P_ImprX obj, "Saídas: " & lblSai.caption, 50
        P_ImprX obj, "Saldo: " & lblSaldo.caption, 90, False
        P_ImprLinha obj, "_"
        P_CentraImpressao obj, "Vendas"
        P_ImprLinha obj
        P_ImprX obj, "A Vista: " & lblVendas.caption, 10
        P_ImprX obj, "A Prazo: " & lblAprazo.caption, 50
        P_ImprX obj, "Financeira: " & lblFinan.caption, 90
        P_ImprX obj, "Total: " & FormatVal(CCur(lblTotal.caption)), 130, False
        
        P_FimRel obj, strNome_Emp
        
    End With
    If TypeName(obj) = "Printer" Then
        Printer.EndDoc
    End If
End Sub
Sub P_imprVendas(obj As Object)
    P_CabRelTimbr "Fechamento de Caixa " & txtbxData.Text, obj
    
    P_ImprX obj, "Histórico", 5
    P_ImprX obj, "Epécie", 100
    P_ImprX obj, "T", 115
    P_ImprX obj, "Valor", 140, False, True
    P_ImprLinha obj
    With dtServ.Recordset
        .MoveFirst
        Do While Not .EOF
            If Printer.CurrentY >= Printer.ScaleHeight - 30 Then
                P_FimRel obj, "Continua"
                P_ImprX obj, "Histórico", 5
                P_ImprX obj, "Epécie", 100
                P_ImprX obj, "T", 115
                P_ImprX obj, "Valor", 140, False, True
                P_ImprLinha obj
            End If
            P_ImprX obj, !historico_cai, 5
            P_ImprX obj, !Esp_cai, 100
            P_ImprX obj, !tipo_cai, 115
            P_ImprX obj, FormatVal(!valor_Cai), 140, False, True
            
            .MoveNext
        Loop
        If Printer.CurrentY >= Printer.ScaleHeight - 50 Then
            P_FimRel obj, "Continua"
        End If
        P_ImprLinha obj, "_"
        P_CentraImpressao obj, "Totais"
        P_ImprLinha obj
        
        P_ImprX obj, "Entradas: " & lblEntr.caption, 10
        P_ImprX obj, "Saídas: " & lblSai.caption, 50
        P_ImprX obj, "Saldo: " & lblSaldo.caption, 90, False
        P_ImprLinha obj, "_"
        P_CentraImpressao obj, "Vendas"
        P_ImprLinha obj
        P_ImprX obj, "A Vista: " & lblAvista.caption, 10
        P_ImprX obj, "A Prazo: " & lblAprazo.caption, 50
        P_ImprX obj, "Financeira: " & lblFinan.caption, 90
        P_ImprX obj, "Total: " & lblTotal.caption, 130, False
        
        P_FimRel obj, strNome_Emp
        
    End With
    If TypeName(obj) = "Printer" Then
        Printer.EndDoc
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
   
    
    lblComputador.caption = strIP_Usu
    
    
    SSTab1.Tab = 0
End Sub
Sub buscaTodos()
    
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    busca 1
    buscaAprazo
    strSQL = "select tf.guerra_fun from tab_funcionarios as tf, tab_turncai as tt "
    strSQL = strSQL & " where tf.cod_fun=tt.login_usu "
    strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " and data_tur=" & F_GravaData(txtbxData.Text)
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount > 0 Then
            frControle.caption = !guerra_fun
        Else
            frControle.caption = ""
        End If
        .Close
    End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF7 Then
        On Error Resume Next
        frmEfetivaVen.Show 1
    End If
End Sub

Private Sub Form_Load()
    adc.ConnectionString = strConnect
    dtAprazo.ConnectionString = strConnect
    
    txtbxData.Text = Date
    P_PreencheCaixas cmbxCaixa
    P_PreencheTurnos cmbxTurno
    cmbxCaixa.ListIndex = 0
End Sub
Sub buscaCco()
    Dim strSQL As String
    strSQL = "select tf.*, tx.nome_con from tab_financeiro as tf,"
    strSQL = strSQL & " tab_caixas as tc,tab_contas as tx  "
    strSQL = strSQL & " where tc.cod_cai='" & strCod_Cai & "' "
    strSQL = strSQL & " and tc.cod_con=tf.cod_con "
    strSQL = strSQL & " and tx.cod_con=tf.cod_con "
    strSQL = strSQL & " and tf.data_fin=" & F_PesqData(txtbxData.Text)
    
    Dim cCred As Currency
    Dim cDeb As Currency
    dtCco.RecordSource = strSQL
    dtCco.Refresh
    With dtCco.Recordset
        If .RecordCount > 0 Then
            dbgrdCco.caption = "Lançamentos na Conta: " & !nome_con
            
        End If
        Do While Not .EOF
            If !tipo_fin = "D" Then
                cDeb = cDeb + !valor_fin
            Else
                cCred = cCred + !valor_fin
            End If
            .MoveNext
        Loop
        lblDebCco.caption = FormatVal(cDeb)
        lblCredCco.caption = FormatVal(cCred)
        lblTotCco.caption = FormatVal(cCred - cDeb)
    End With
End Sub
Private Sub txtbxData_GotFocus()
    ganhaFoco txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
    buscaTodos
    
End Sub

Private Sub txtbxFundo_Cai_GotFocus()
    ganhaFoco txtbxFundo_Cai
End Sub

Private Sub txtbxFundo_Cai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxFundo_Cai_LostFocus()
    PerdeFoco txtbxFundo_Cai
    If CCur(txtbxFundo_Cai.Text) > CCur(lblSaldo.caption) Then
        MsgBox "O Fundo de caixa está maior que o saldo!", vbExclamation
        
    End If
End Sub
