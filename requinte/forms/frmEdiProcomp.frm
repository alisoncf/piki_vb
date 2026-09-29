VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiProcomp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de Processos de Compra"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10050
   Icon            =   "frmEdiProcomp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnEfetiv 
      Caption         =   "E&fetivar"
      Height          =   615
      Left            =   600
      Picture         =   "frmEdiProcomp.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnCancel 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   2160
      Picture         =   "frmEdiProcomp.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnNovo 
      Caption         =   "&Novo"
      Height          =   615
      Left            =   3720
      Picture         =   "frmEdiProcomp.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnPrint 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   6840
      Picture         =   "frmEdiProcomp.frx":0CA8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Editar"
      Height          =   615
      Left            =   5280
      Picture         =   "frmEdiProcomp.frx":0DF2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   8400
      Picture         =   "frmEdiProcomp.frx":0F3C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.Data dt 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   1  'ODBCCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEdiProcomp.frx":1086
         Height          =   4935
         Left            =   120
         OleObjectBlob   =   "frmEdiProcomp.frx":1097
         TabIndex        =   16
         Top             =   720
         Width           =   9615
      End
      Begin VB.ComboBox cmbxList 
         Height          =   315
         ItemData        =   "frmEdiProcomp.frx":2100
         Left            =   6240
         List            =   "frmEdiProcomp.frx":2113
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtbxCliente 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtbxDe 
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxAte 
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   8520
         Picture         =   "frmEdiProcomp.frx":216B
         TabIndex        =   7
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação:"
         Height          =   195
         Left            =   6240
         TabIndex        =   13
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   4080
         TabIndex        =   3
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   5160
         TabIndex        =   5
         Top             =   120
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmEdiProcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    strSQL = "select tp.* from tab_procomp as tp"
    strSQL = strSQL & " where 1=1 "
    If cmbxList.ListIndex > 0 Then

        strSQL = strSQL & " and sts_pcp='" & Left(cmbxList.Text, 1) & "' "
    End If
    
    If txtbxDe.Text <> "" Then
        strSQL = strSQL & " and dti_pcp>=" & F_GravaData(txtbxDe.Text)
    End If
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and dti_pcp<=" & F_GravaData(txtbxAte.Text)
    End If
    strSQL = strSQL & " order by dti_pcp desc "
    
    dt.RecordSource = strSQL
    dt.Refresh
    Exit Sub
    abreObjetoAdo rPesq, strSQL
    With rPesq
        msflxgrd.rows = 1
        Do While Not .EOF
            strSQL = !seq_pcp & Chr(9)
            strSQL = strSQL & !dti_pcp & Chr(9)
            strSQL = strSQL & !dtf_pcp & Chr(9)
            strSQL = strSQL & !sts_pcp & Chr(9)
            strSQL = strSQL & !pse_pcp & Chr(9)
            strSQL = strSQL & !pec_pcp & Chr(9)
            strSQL = strSQL & !ppc_pcp & Chr(9)
            strSQL = strSQL & !cod_cat & Chr(9)
            strSQL = strSQL & !cod_mar & Chr(9)
            strSQL = strSQL & !nome_pro & Chr(9)
            strSQL = strSQL & !cod_usu
            msflxgrd.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    'zebraGrid msflxgrd, 1
End Sub

Private Sub cmndbtnCancel_Click()
    On Error GoTo trata_erro:
    Dim strSQL As String
    With dt.Recordset
        If !sts_pcp = "1" Then
            If MsgBox("Deseja realmente excluir o Processo de Compra: " & !seq_pcp & "?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "Delete from tab_procomp where seq_pcp='" & !seq_pcp & "' "
                Banco.Execute strSQL
                strSQL = "Delete from tab_itenspcp where seq_pcp='" & !seq_pcp & "' "
                Banco.Execute strSQL
                MsgBox "Processo Excluído!", vbCritical
            End If
        ElseIf !sts_pcp = "2" Then
            If MsgBox("Deseja realmente Cancelar o Processo de Compra: " & !seq_pcp & "?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "update tab_procomp set sts_pcp=3 where seq_pcp='" & !seq_pcp & "' "
                MsgBox "Processo Cancelado!", vbCritical
            End If
        Else
            MsgBox "Não é possível cancelar este processo!", vbExclamation
        End If
    Exit Sub

        
    cmndbtnBuscar_Click
    End With
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_erro:
    With dt.Recordset
        Dim fr As New frmSujComp
        fr.tag = !seq_pcp
        fr.Show 0
    
    End With
    Exit Sub
    
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEfetiv_Click()
    On Error GoTo trata_erro:
    Dim strSQL As String
    With dt.Recordset
        If !sts_pcp = "2" Then
            If MsgBox("Deseja realmente Concluir o Processo de Compra: " & !seq_pcp & "?", vbYesNo + vbQuestion) = vbYes Then
                strSQL = "update tab_procomp set sts_pcp=4 where seq_pcp='" & !seq_pcp & "' "
                MsgBox "Processo Cancelado!", vbCritical
            End If
        Else
            MsgBox "Não é possível Efetivar este processo!", vbExclamation
        End If
    Exit Sub

        
    cmndbtnBuscar_Click
    End With
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNovo_Click()
    Dim fr As New frmSujComp
    fr.Show 0
End Sub

Private Sub cmndbtnPrint_Click()
    On Error GoTo trata_erro:
    With dt.Recordset
       ImprimePcp !seq_pcp
    End With
    Exit Sub
trata_erro:
    P_MostraErro
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    cmndbtnBuscar_Click
    
End Sub

Private Sub Form_Load()
    cmbxList.ListIndex = 0
    Me.Top = 10
    Me.Left = 10
    dt.Connect = Banco.Connect
    
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxCliente_GotFocus()
    P_SelecTexto txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim c As cliente
        c = F_BuscaCliente(txtbxCliente.Text)
        If c.cod_cli <> "" Then
            txtbxCliente.Text = c.nome_Cli
            txtbxCliente.tag = c.cod_cli
            Sendkeys Chr(9)
        Else
            txtbxCliente.Text = ""
            txtbxCliente.tag = ""
            txtbxCliente.SetFocus
        End If
    End If
End Sub

Private Sub txtbxCliente_LostFocus()
    P_AjustaTexto txtbxCliente
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub

