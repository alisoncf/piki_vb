VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frmSQL 
   Caption         =   "SQL - Plus"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   8775
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8535
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmSQL.frx":0000
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "frmSQL.frx":0011
      TabIndex        =   2
      Top             =   1560
      Width           =   8535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Executar: F9"
      Height          =   195
      Left            =   7680
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub busca()
    Dim iCount As Integer
    dt.RecordSource = Text1.Text
    dt.Refresh
    With dt.Recordset
        For iCount = 0 To .Fields.count - 1
            dbgrd.Columns.Add
        Next
    End With
End Sub

