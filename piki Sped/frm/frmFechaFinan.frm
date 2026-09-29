VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFechaFinan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consolidação Diária de Movimentação Financeira"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "frmFechaFinan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frControle 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   5400
         TabIndex        =   10
         Top             =   5880
         Width           =   3375
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   1680
            Picture         =   "frmFechaFinan.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmndbtnImprREs 
            Caption         =   "&Imprimir Resumo"
            Height          =   615
            Left            =   120
            Picture         =   "frmFechaFinan.frx":058C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "&Data"
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   5880
         Width           =   1335
         Begin VB.TextBox txtbxData 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5415
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9551
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Documentos Baixados"
         TabPicture(0)   =   "frmFechaFinan.frx":06D6
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblDocBaiR"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblDocBaiP"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label6"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label7"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblDocBaiTotal"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label9"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "dbgrd"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "adc"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Documentos Gerados"
         TabPicture(1)   =   "frmFechaFinan.frx":06F2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "adcgerado"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "gridGerado"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label17"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lblDocGerTotal"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label13"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label12"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lblDocGerP"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "lblDocGerR"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).ControlCount=   8
         TabCaption(2)   =   "Movimentação C/C"
         TabPicture(2)   =   "frmFechaFinan.frx":070E
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label18"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "lblTotCco"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label16"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label15"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "lblDebCco"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "lblCredCco"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "gridMov"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "adcMov"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).ControlCount=   8
         TabCaption(3)   =   "Cheques Baixados"
         TabPicture(3)   =   "frmFechaFinan.frx":072A
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "adcChequeBx"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "gridChequeBaix"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label24"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "lblCheBaiTotal"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "Label22"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "Label21"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "lblCheBaiP"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "lblCheBaiR"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).ControlCount=   8
         TabCaption(4)   =   "Cheques Gerados"
         TabPicture(4)   =   "frmFechaFinan.frx":0746
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lblCco"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "lblCheGerR"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "lblCheGerP"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "Label27"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "Label28"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "lblCheGerTotal"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).Control(6)=   "Label30"
         Tab(4).Control(6).Enabled=   0   'False
         Tab(4).Control(7)=   "gridChGera"
         Tab(4).Control(7).Enabled=   0   'False
         Tab(4).Control(8)=   "adcChGer"
         Tab(4).Control(8).Enabled=   0   'False
         Tab(4).ControlCount=   9
         Begin MSAdodcLib.Adodc adc 
            Height          =   330
            Left            =   -73680
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
            Bindings        =   "frmFechaFinan.frx":0762
            Height          =   4215
            Left            =   -74880
            TabIndex        =   38
            Top             =   480
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7435
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
         Begin MSAdodcLib.Adodc adcgerado 
            Height          =   330
            Left            =   -73680
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
         Begin MSDataGridLib.DataGrid gridGerado 
            Bindings        =   "frmFechaFinan.frx":0774
            Height          =   4215
            Left            =   -74880
            TabIndex        =   39
            Top             =   480
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7435
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
         Begin MSAdodcLib.Adodc adcMov 
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
         Begin MSDataGridLib.DataGrid gridMov 
            Bindings        =   "frmFechaFinan.frx":0786
            Height          =   4215
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7435
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
         Begin MSAdodcLib.Adodc adcChequeBx 
            Height          =   330
            Left            =   -73680
            Top             =   4560
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
         Begin MSDataGridLib.DataGrid gridChequeBaix 
            Bindings        =   "frmFechaFinan.frx":0798
            Height          =   4215
            Left            =   -74880
            TabIndex        =   41
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7435
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
         Begin MSAdodcLib.Adodc adcChGer 
            Height          =   330
            Left            =   -73560
            Top             =   4560
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
         Begin MSDataGridLib.DataGrid gridChGera 
            Bindings        =   "frmFechaFinan.frx":07AA
            Height          =   4215
            Left            =   -74760
            TabIndex        =   42
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7435
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
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -71880
            TabIndex        =   37
            Top             =   4920
            Width           =   165
         End
         Begin VB.Label lblCheGerTotal 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -71520
            TabIndex        =   36
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -74880
            TabIndex        =   35
            Top             =   4920
            Width           =   165
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -73560
            TabIndex        =   34
            Top             =   4920
            Width           =   90
         End
         Begin VB.Label lblCheGerP 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -72840
            TabIndex        =   33
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label lblCheGerR 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -74400
            TabIndex        =   32
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -71880
            TabIndex        =   31
            Top             =   4920
            Width           =   165
         End
         Begin VB.Label lblCheBaiTotal 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -71520
            TabIndex        =   30
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -74880
            TabIndex        =   29
            Top             =   4920
            Width           =   165
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -73560
            TabIndex        =   28
            Top             =   4920
            Width           =   90
         End
         Begin VB.Label lblCheBaiP 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -72840
            TabIndex        =   27
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label lblCheBaiR 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -74400
            TabIndex        =   26
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -71880
            TabIndex        =   25
            Top             =   4920
            Width           =   165
         End
         Begin VB.Label lblDocGerTotal 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -71520
            TabIndex        =   24
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -74880
            TabIndex        =   23
            Top             =   4920
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -73560
            TabIndex        =   22
            Top             =   4920
            Width           =   90
         End
         Begin VB.Label lblDocGerP 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -72840
            TabIndex        =   21
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label lblDocGerR 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -74400
            TabIndex        =   20
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -71880
            TabIndex        =   19
            Top             =   4920
            Width           =   165
         End
         Begin VB.Label lblDocBaiTotal 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -71520
            TabIndex        =   18
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -74880
            TabIndex        =   17
            Top             =   4920
            Width           =   165
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -73560
            TabIndex        =   16
            Top             =   4920
            Width           =   90
         End
         Begin VB.Label lblDocBaiP 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -72840
            TabIndex        =   15
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label lblDocBaiR 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   -74400
            TabIndex        =   14
            Tag             =   "!"
            Top             =   5010
            Width           =   390
         End
         Begin VB.Label lblCco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conta"
            Height          =   195
            Left            =   -69480
            TabIndex        =   13
            Top             =   5040
            Width           =   420
         End
         Begin VB.Label lblCredCco 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   720
            TabIndex        =   7
            Tag             =   "!"
            Top             =   5040
            Width           =   390
         End
         Begin VB.Label lblDebCco 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   2280
            TabIndex        =   6
            Tag             =   "!"
            Top             =   5040
            Width           =   390
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   5
            Top             =   4957
            Width           =   90
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   4
            Top             =   4957
            Width           =   165
         End
         Begin VB.Label lblTotCco 
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
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   3600
            TabIndex        =   3
            Tag             =   "!"
            Top             =   5040
            Width           =   390
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3240
            TabIndex        =   2
            Top             =   4957
            Width           =   165
         End
      End
   End
End
Attribute VB_Name = "frmFechaFinan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub soma()
    Dim cCred As Currency
    Dim cDeb As Currency
    On Error Resume Next
    With dtAprazo.Recordset
        
        Do While Not .EOF
            If !tipo_doc = "R" Then
                cCred = cCred + !valor_itd
            Else
                cDeb = cCred + !valor_itd
            End If
            .MoveNext
        Loop
        lblDocGerP.caption = FormatVal(cDeb)
        lblDocGerR.caption = FormatVal(cCred)
        lblDocGerTotal.caption = FormatVal(cCred - cDeb)
        If .RecordCount > 0 Then .MoveFirst
    End With
    cCred = 0
    cDeb = 0

    With dtCco.Recordset
        
        Do While Not .EOF
            If !tipo_fin = "C" Then
                cCred = cCred + !valor_fin
            Else
                cDeb = cCred + !valor_fin
            End If
            .MoveNext
        Loop
        lblDebCco.caption = FormatVal(cDeb)
        lblCredCco.caption = FormatVal(cCred)
        lblTotCco.caption = FormatVal(cCred - cDeb)
        If .RecordCount > 0 Then .MoveFirst
    End With
    cCred = 0
    cDeb = 0
    With dtCheBai.Recordset
        
        Do While Not .EOF
            If !tipo_doc = "R" Then
                cCred = cCred + !Valor_che
            Else
                cDeb = cCred + !Valor_che
            End If
            .MoveNext
        Loop
        lblCheBaiP.caption = FormatVal(cDeb)
        lblCheBaiR.caption = FormatVal(cCred)
        lblCheBaiTotal.caption = FormatVal(cCred - cDeb)
        If .RecordCount > 0 Then .MoveFirst
    End With
    cCred = 0
    cDeb = 0
    With dtCheGer.Recordset
        
        Do While Not .EOF
            If !tipo_doc = "R" Then
                cCred = cCred + !Valor_che
            Else
                cDeb = cCred + !Valor_che
            End If
            .MoveNext
        Loop
        lblCheGerP.caption = FormatVal(cDeb)
        lblCheGerR.caption = FormatVal(cCred)
        lblCheGerTotal.caption = FormatVal(cCred - cDeb)
        If .RecordCount > 0 Then .MoveFirst
    End With
    cCred = 0
    cDeb = 0
    With dtServ.Recordset
        
        Do While Not .EOF
            If !tipo_doc = "R" Then
                cCred = cCred + !valor_itd - !reman_itd + !juros_itd - !desconto_itd
            Else
                cDeb = cCred + !valor_itd - !reman_itd + !juros_itd - !desconto_itd
            End If
            .MoveNext
        Loop
        lblDocBaiP.caption = FormatVal(cDeb)
        lblDocBaiR.caption = FormatVal(cCred)
        lblDocBaiTotal.caption = FormatVal(cCred - cDeb)
        If .RecordCount > 0 Then .MoveFirst
    End With
    
End Sub
Sub buscaCco()
    Dim strSQL As String
    strSQL = "select tc.cod_con,tc.nome_con,tf.* from  " & TAB_FINANCEIRO & "  as tf, "
    strSQL = strSQL & " tab_contas as tc where tc.cod_con=tf.cod_con "
    strSQL = strSQL & " and tf.data_fin=" & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " order by nome_Con,tipo_fin,seq_fin"
    dtCco.RecordSource = strSQL
    dtCco.Refresh
End Sub
Sub buscaDocBai()
    Dim strSQL As String
    strSQL = "select td.num_doc,td.historico_doc,td.parc_doc, "
    strSQL = strSQL & "ti.*,td.tipo_doc,ti.cod_usu,tf.* "
    strSQL = strSQL & "from  " & TAB_DOC & " 1 as td,  " & TAB_ITDOC & " 1 as ti,  " & TAB_FINANCEIRO & "  as tf "
    strSQL = strSQL & "where td.seq_doc = ti.seq_doc "
    strSQL = strSQL & "and ti.seq_itd = tf.seq_lan "
    strSQL = strSQL & "and ti.baixa_itd = " & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " order by tipo_doc, num_doc"
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub
Sub buscaDocGer()
    Dim strSQL As String
    strSQL = "select td.num_doc,td.historico_doc,td.parc_doc,ti.reman_itd, "
    strSQL = strSQL & "ti.*,td.tipo_doc,ti.cod_usu "
    strSQL = strSQL & "from  " & TAB_DOC & " 1 as td,  " & TAB_ITDOC & " 1 as ti "
    strSQL = strSQL & "where td.seq_doc = ti.seq_doc "
    strSQL = strSQL & "and td.data_doc = " & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " order by tipo_doc, num_doc"
    dtAprazo.RecordSource = strSQL
    dtAprazo.Refresh
End Sub

Sub buscaCheBai()
    Dim strSQL As String
    strSQL = "select td.num_doc,td.historico_doc,td.tipo_doc,tc.*  "
    
    strSQL = strSQL & "from tab_cheques1 as tc,  " & TAB_DOC & " 1 as td "
    strSQL = strSQL & "where td.seq_doc = tc.seq_doc "
    strSQL = strSQL & "and tc.dep_che = " & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " order by tipo_doc, num_doc"
    dtCheBai.RecordSource = strSQL
    dtCheBai.Refresh
End Sub
Sub buscaCheGer()
    Dim strSQL As String
    strSQL = "select td.num_doc,td.historico_doc,td.tipo_doc,tc.*  "
    
    strSQL = strSQL & "from tab_cheques1 as tc,  " & TAB_DOC & " 1 as td "
    strSQL = strSQL & "where tc.seq_doc = td.seq_doc "
    strSQL = strSQL & "and td.data_doc = " & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " order by tipo_doc, num_doc"
    dtCheGer.RecordSource = strSQL
    dtCheGer.Refresh
End Sub

Private Sub cmndbtnImprREs_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    Dim sPage As Integer
    P_CabRelTimbr "CONSOLIDAÇÃO DIÁRIA DIA: " & txtbxData.Text, obj
    P_ImprX obj, "DOCUMENTOS BAIXADOS", 0, False
    P_ImprLinha obj, "_"
    P_ImprX obj, "Documento", 0
    P_ImprX obj, "Valor", 30, True, True
    P_ImprX obj, "Baixado", 50, True, True
    P_ImprX obj, "Reman.", 70, True, True
    P_ImprX obj, "Juros.", 90, True, True
    P_ImprX obj, "Desconto.", 110, True, True
    P_ImprX obj, "Histórico", 115
    P_ImprX obj, "T", 200, False, True
    P_ImprLinha obj
    With dtServ.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If obj.Name = "Printer" Then
                    sPage = Printer.Page
                Else
                    sPage = sPage + 1
                End If
                P_RodapeVis obj, "", "Pág " & sPage
                
                P_ImprX obj, "Documento", 0
                P_ImprX obj, "Valor", 30, True, True
                P_ImprX obj, "Baixado", 50, True, True
                P_ImprX obj, "Reman.", 70, True, True
                P_ImprX obj, "Juros.", 90, True, True
                P_ImprX obj, "Desconto.", 110, True, True
                P_ImprX obj, "Histórico", 115
                P_ImprX obj, "T", 200, False, True
                P_ImprLinha obj
            End If
       
            P_ImprX obj, !num_doc, 0
            P_ImprX obj, FormatVal(!valor_itd), 30, True, True
            P_ImprX obj, FormatVal(!valor_fin), 50, True, True
            P_ImprX obj, FormatVal(!reman_itd), 70, True, True
            P_ImprX obj, FormatVal(!juros_itd), 90, True, True
            P_ImprX obj, FormatVal(!desconto_itd), 110, True, True
            P_ImprX obj, !historico_doc, 115
            P_ImprX obj, !tipo_doc, 200, False, True
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "A Receber: " & lblDocBaiR.caption, 60
        P_ImprX obj, "A Pagar: " & lblDocBaiP.caption, 120
        P_ImprX obj, "Total: " & lblDocBaiTotal.caption, 200, False, True
        If .RecordCount > 0 Then .MoveFirst
    
    End With
    '
    obj.Print ""
    P_ImprX obj, "DOCUMENTOS GERADOS", 0, False
    P_ImprLinha obj, "_"
    P_ImprX obj, "Documento", 0
    P_ImprX obj, "Valor", 40, True, True
    P_ImprX obj, "Reman.", 60, True, True
    P_ImprX obj, "Juros.", 80, True, True
    P_ImprX obj, "Desconto.", 100, True, True
    P_ImprX obj, "Histórico", 102
    P_ImprX obj, "T", 200, False, True
    P_ImprLinha obj
    With dtAprazo.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If obj.Name = "Printer" Then
                    sPage = Printer.Page
                Else
                    sPage = sPage + 1
                End If
                P_RodapeVis obj, "", "Pág " & sPage
                
                
                P_ImprX obj, "Documento", 0
                P_ImprX obj, "Valor", 40, True, True
                P_ImprX obj, "Reman.", 60, True, True
                P_ImprX obj, "Juros.", 80, True, True
                P_ImprX obj, "Desconto.", 100, True, True
                P_ImprX obj, "Histórico", 102
                P_ImprX obj, "T", 200, False, True
                P_ImprLinha obj
            End If
       
            P_ImprX obj, !num_doc, 0
            P_ImprX obj, FormatVal(!valor_itd), 40, True, True
            P_ImprX obj, FormatVal(!reman_itd), 60, True, True
            P_ImprX obj, FormatVal(!juros_itd), 80, True, True
            P_ImprX obj, FormatVal(!desconto_itd), 100, True, True
            P_ImprX obj, !historico_doc, 102
            P_ImprX obj, !tipo_doc, 200, False, True
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "A Receber: " & lblDocGerR.caption, 60
        P_ImprX obj, "A Pagar: " & lblDocGerP.caption, 120
        P_ImprX obj, "Total: " & lblDocGerTotal.caption, 200, False, True
        If .RecordCount > 0 Then .MoveFirst
    End With
    '
    obj.Print ""
    P_ImprX obj, "CHEQUES BAIXADOS", 0, False
    P_ImprLinha obj, "_"
    P_ImprX obj, "Nº ", 0
    P_ImprX obj, "Agência", 25
    P_ImprX obj, "Banco", 40
    P_ImprX obj, "Conta", 50
    P_ImprX obj, "Valor.", 80, True, True
    P_ImprX obj, "Histórico", 82
    P_ImprX obj, "T", 200, False, True
    P_ImprLinha obj
    With dtCheBai.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If obj.Name = "Printer" Then
                    sPage = Printer.Page
                Else
                    sPage = sPage + 1
                End If
                P_RodapeVis obj, strNomeSis, "Pág " & sPage
                P_ImprX obj, "Nº ", 0
                P_ImprX obj, "Agência", 25
                P_ImprX obj, "Banco", 40
                P_ImprX obj, "Conta", 50
                P_ImprX obj, "Valor.", 80, True, True
                P_ImprX obj, "Histórico", 82
                P_ImprX obj, "T", 200, False, True
                P_ImprLinha obj
            End If
       
            P_ImprX obj, !num_che, 0
            P_ImprX obj, !age_che, 25
            P_ImprX obj, !cod_bco, 40
            
            P_ImprX obj, IIf(!tipo_doc = "R", !cco_che, !cod_con), 50
            
            P_ImprX obj, FormatVal(!Valor_che), 80, True, True
            P_ImprX obj, !historico_doc, 82
            P_ImprX obj, !tipo_doc, 200, False, True
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "A Receber: " & lblCheBaiR.caption, 60
        P_ImprX obj, "A Pagar: " & lblCheBaiP.caption, 120
        P_ImprX obj, "Total: " & lblCheBaiTotal.caption, 200, False, True
        If .RecordCount > 0 Then .MoveFirst
    End With
    '
    obj.Print ""
    P_ImprX obj, "CHEQUES GERADOS", 0, False
    P_ImprLinha obj, "_"
    P_ImprX obj, "Nº ", 0
    P_ImprX obj, "Agência", 25
    P_ImprX obj, "Banco", 40
    P_ImprX obj, "Conta", 50
    P_ImprX obj, "Valor.", 80, True, True
    P_ImprX obj, "Histórico", 82
    P_ImprX obj, "T", 200, False, True
    P_ImprLinha obj
    With dtCheGer.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If obj.Name = "Printer" Then
                    sPage = Printer.Page
                Else
                    sPage = sPage + 1
                End If
                P_RodapeVis obj, strNomeSis, "Pág " & sPage
                P_ImprX obj, "Nº ", 0
                P_ImprX obj, "Agência", 25
                P_ImprX obj, "Banco", 40
                P_ImprX obj, "Conta", 50
                P_ImprX obj, "Valor.", 80, True, True
                P_ImprX obj, "Histórico", 82
                P_ImprX obj, "T", 200, False, True
                P_ImprLinha obj
            End If
       
            P_ImprX obj, !num_che, 0
            P_ImprX obj, !age_che, 25
            P_ImprX obj, !cod_bco, 40
            
            P_ImprX obj, IIf(!tipo_doc = "R", !cco_che, !cod_con), 50
            
            P_ImprX obj, FormatVal(!Valor_che), 80, True, True
            P_ImprX obj, !historico_doc, 82
            P_ImprX obj, !tipo_doc, 200, False, True
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "A Receber: " & lblCheGerR.caption, 60
        P_ImprX obj, "A Pagar: " & lblCheGerP.caption, 120
        P_ImprX obj, "Total: " & lblCheGerTotal.caption, 200, False, True
        If .RecordCount > 0 Then .MoveFirst
        
    End With
    '
    obj.Print ""
    P_ImprX obj, "MOVIMENTAÇÃO DE CONTAS CORRENTES", 0, False
    P_ImprLinha obj, "_"
    P_ImprX obj, "Conta ", 0
    P_ImprX obj, "Histórico", 60
    P_ImprX obj, "Esp", 160
    P_ImprX obj, "T", 170
    P_ImprX obj, "Valor.", 200, False, True
    
    P_ImprLinha obj
    With dtCco.Recordset
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                If obj.Name = "Printer" Then
                    sPage = Printer.Page
                Else
                    sPage = sPage + 1
                End If
                P_RodapeVis obj, strNomeSis, "Pág " & sPage
                P_ImprX obj, "Conta ", 0
                P_ImprX obj, "Histórico", 60
                P_ImprX obj, "Esp", 160
                P_ImprX obj, "T", 170
                P_ImprX obj, "Valor.", 200, False, True
                P_ImprLinha obj
            End If
       
            P_ImprX obj, !nome_con, 0
            P_ImprX obj, !historico_fin, 60
            P_ImprX obj, !Esp_fin, 160
            P_ImprX obj, !tipo_fin, 170
            P_ImprX obj, FormatVal(!valor_fin), 200, False, True
            .MoveNext
        Loop
        P_ImprLinha obj
        P_ImprX obj, "Crédito: " & lblCredCco.caption, 60
        P_ImprX obj, "Débito: " & lblDebCco.caption, 120
        P_ImprX obj, "Total: " & lblTotCco.caption, 200, False, True
        If .RecordCount > 0 Then .MoveFirst
    End With
        P_RodapeVis obj, strNomeSis, "(F)"
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtCheGer_Reposition()
    On Error Resume Next
    With dtCheGer.Recordset
        If !tipo_doc = "P" Then
            lblCco.caption = "Conta: " & F_BuscaUmCampo("nome_con", "tab_contas", _
                "cod_con='" & !cod_con & "'")
        Else
            lblCco.caption = "-"
        End If
    End With
End Sub

Private Sub dtServ_Validate(Action As Integer, Save As Integer)
End Sub

Private Sub Form_Load()
    'P_CentraForm me
    dtServ.ConnectionString = strConnect
    dtCheGer.ConnectionString = strConnect
    dtCheBai.ConnectionString = strConnect
    dtCco.ConnectionString = strConnect
    dtAprazo.ConnectionString = strConnect
    
    txtbxData.Text = Date
    buscaTodos
    SSTab1.Tab = 0
End Sub

Sub buscaTodos()
    buscaDocBai
    buscaDocGer
    buscaCheBai
    buscaCheGer
    buscaCco
    soma
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
