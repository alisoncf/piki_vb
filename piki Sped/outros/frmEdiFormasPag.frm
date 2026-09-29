VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form frmEdiFormasPag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Formas de Pagamento"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8490
   Icon            =   "frmEdiFormasPag.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2415
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   195
            Left            =   1200
            TabIndex        =   9
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optnDup 
            Caption         =   "Duplicata"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optnCheque 
            Caption         =   "Cheque"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   2640
         TabIndex        =   1
         Top             =   120
         Width           =   5655
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   615
            Left            =   3720
            Picture         =   "frmEdiFormasPag.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   615
            Left            =   1920
            Picture         =   "frmEdiFormasPag.frx":0698
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Novo"
            Height          =   615
            Left            =   120
            Picture         =   "frmEdiFormasPag.frx":07E2
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Data dtProd 
         Caption         =   "Edição de Gêneros"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   3600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEdiFormasPag.frx":0B6C
         Height          =   3855
         Left            =   240
         OleObjectBlob   =   "frmEdiFormasPag.frx":0B81
         TabIndex        =   5
         Top             =   1200
         Width           =   8055
      End
   End
End
Attribute VB_Name = "frmEdiFormasPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    With dtProd.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há Tipo de Pagamento para editar!", vbExclamation
            Exit Sub
        End If
        Dim fr As New frmIncFormasPag
        fr.Tag = !cod_tip
        fr.txtbxDesc.Text = !Descricao_tip
        fr.txtbxInterv.Text = !Intervalo_tip
        fr.txtbxJuros.Text = FormatVal(!juros_tip)
        fr.txtbxParc.Text = !parcelas_tip
        If !Entrada_tip = "S" Then
            fr.cmbxEntr.ListIndex = 0
        Else
            fr.cmbxEntr.ListIndex = 1
        End If
        If !tipo_tip = "C" Then
            fr.optnChe.Value = True
        Else
            fr.optnDup.Value = True
        End If
        fr.Show 0
        
    End With
    busca
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmIncFormasPag
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    
    'P_CentraForm Me, 600
    dtProd.Connect = Banco.Connect
    busca
End Sub

Sub busca()
    Dim strSQL As String
    strSQL = "select tt.*,if(tipo_tip='C','Cheque','Duplicata')  as tipo from tab_tipoven  as tt "
    If optnCheque.Value Then
        strSQL = strSQL & " where tipo_tip='C'"
    ElseIf optnDup.Value Then
        strSQL = strSQL & " where tipo_tip='D'"
    End If
    strSQL = strSQL & " order by cod_tip"
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub Option1_Click()
    busca
End Sub

Private Sub optnCheque_Click()
    busca
End Sub

Private Sub optnDup_Click()
    busca
End Sub
