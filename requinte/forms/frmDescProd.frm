VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDescProd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Desconto sobre Produtos"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lstFormasPag 
      Height          =   1635
      ItemData        =   "frmDescProd.frx":0000
      Left            =   4800
      List            =   "frmDescProd.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   3240
      Width           =   5055
   End
   Begin VB.CommandButton cmndbtnExcluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnGravar 
      Caption         =   "Gravar"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtbxDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.ListBox lstCargos 
      Height          =   2310
      ItemData        =   "frmDescProd.frx":0004
      Left            =   4800
      List            =   "frmDescProd.frx":0006
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmDescProd.frx":0008
      Height          =   6975
      Left            =   120
      OleObjectBlob   =   "frmDescProd.frx":0019
      TabIndex        =   9
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Formas Pag:"
      Height          =   195
      Left            =   4800
      TabIndex        =   8
      Top             =   3000
      Width           =   885
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desconto:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   5160
      Width           =   4005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargos:"
      Height          =   195
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmDescProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dt_Reposition()
    On Error Resume Next
    busca
End Sub
Sub busca()
    buscaCar
End Sub
Sub buscaCar()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim iCount As Integer
    lstCargos.Clear
    strSQL = "select * from tab_cargos order by nome_car "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            lstCargos.AddItem !nome_car
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select tc.nome_car from tab_cargos as tc, tab_descontos as td where tc.cod_car=td.cod_car order by nome_car "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            For iCount = 0 To lstCargos.ListCount - 1
                lstCargos.ListIndex = iCount
                If !nome_car = lstCargos.Text Then
                    lstCargos.Selected(iCount) = True
                End If
            Next
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
Private Sub Form_Activate()
    buscaSub
End Sub
Sub buscaSub()
    Dim strSQL As String
    strSQL = "select * from tab_subsec order by nome_sub "
    dt.RecordSource = strSQL
    dt.Refresh
End Sub
Private Sub Form_Load()
    
    dt.Connect = Banco.Connect
    
End Sub

