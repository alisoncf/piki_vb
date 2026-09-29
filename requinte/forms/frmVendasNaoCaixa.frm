VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVendasNaoCaixa 
   Caption         =   "Vendas Efetivadas/ Lançamento no Caixa"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxDataIni 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtbxDataFin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmVendasNaoCaixa.frx":0000
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "frmVendasNaoCaixa.frx":0011
      TabIndex        =   1
      Top             =   960
      Width           =   8175
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Width           =   7020
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblDataIni 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Data Inicial:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lbldataFin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data &Final:"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmVendasNaoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sNumVen As String
    Dim rPesq As Recordset
    
    dt.Connect = Banco.Connect
    strSQL = "select distinct num_ven from tab_caixa where num_ven<>'' and data_cai >='2009-12-01'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            sNumVen = "'" & !num_ven & "'"
            .MoveNext
            If Not .EOF Then
                sNumVen = sNumVen & ","
            End If
        Loop
        .Close
    End With
    
    
    
    strSQL = "select * from tab_vendas where situacao_ven=4 and dataefet_ven >='2009-12-01'"
    strSQL = strSQL & " and num_ven not in(" & sNumVen & ")"
    strSQL = strSQL & " order by num_ven desc "
    dt.RecordSource = strSQL
    dt.Refresh
    
    
    
    
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub
