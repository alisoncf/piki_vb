VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiArr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Arranjos"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "frmEdiArr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6990
   Begin VB.Frame frControle 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2175
         Begin VB.CommandButton cmndbtnBuscar 
            Caption         =   "&Buscar"
            Height          =   660
            Left            =   120
            Picture         =   "frmEdiArr.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1560
            Width           =   1935
         End
         Begin VB.ComboBox cmbxSitua 
            Height          =   315
            ItemData        =   "frmEdiArr.frx":0454
            Left            =   120
            List            =   "frmEdiArr.frx":045E
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtbxAte 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtbxDe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Situação:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   2175
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Novo Arranjo"
            Height          =   660
            Left            =   120
            Picture         =   "frmEdiArr.frx":0482
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   255
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   660
            Left            =   120
            Picture         =   "frmEdiArr.frx":080C
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   660
            Left            =   120
            Picture         =   "frmEdiArr.frx":0956
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2400
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "&Imprimir"
            Height          =   660
            Left            =   120
            Picture         =   "frmEdiArr.frx":0AA0
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.Data dtVendas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEdiArr.frx":0E2A
         Height          =   5415
         Left            =   2400
         OleObjectBlob   =   "frmEdiArr.frx":0E41
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmEdiArr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Busca
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_erro
    Dim fr As New frmArranjos
    With dtVendas.Recordset
        fr.Tag = !num_arr
        fr.caption = "Arranjo " & !num_arr
        fr.Show 0
        
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnImpr_Click()
    On Error GoTo trata_erro
    
    imprimeArr dtVendas.Recordset!num_arr
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmArranjos
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtVendas_Reposition()
    On Error Resume Next
    If dtVendas.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    End If
    With dtVendas.Recordset
        If !situacao_arr = "1" Then
            cmndbtnEdit.Enabled = True
            cmndbtnImpr.Enabled = False
        Else
            cmndbtnEdit.Enabled = False
            cmndbtnImpr.Enabled = True
        End If
    End With
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    Busca
End Sub

Private Sub Form_Load()
    dtVendas.Connect = banco.Connect
    P_CentraForm Me
End Sub
Sub Busca()
    Dim strSQL As String
    strSQL = "select ta.*,tf.guerra_fun from tab_arranjos as ta,"
    strSQL = strSQL & "tab_funcionarios as tf "
    strSQL = strSQL & " where ta.cod_usu=tf.cod_fun"
    If txtbxDe.Text <> "" Then
        strSQL = strSQL & " and ta.data_arr>=" & F_GravaData(txtbxDe)
    End If
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and ta.data_arr<=" & F_GravaData(txtbxAte)
    End If
    If cmbxSitua.ListIndex = 1 Then
        strSQL = strSQL & "  and situacao_arr='2'"
    Else
        strSQL = strSQL & "  and situacao_arr='1'"
    End If
    strSQL = strSQL & " order by data_arr "
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub
