VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiProcomp 
   Caption         =   "Lista de Processos de Compra"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   Icon            =   "frmEdiProcomp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   14715
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   14685
      TabIndex        =   18
      Top             =   0
      Width           =   14715
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   495
         Left            =   8400
         Picture         =   "frmEdiProcomp.frx":08CA
         TabIndex        =   8
         Top             =   75
         Width           =   1215
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
      Begin VB.ComboBox cmbxList 
         Height          =   315
         ItemData        =   "frmEdiProcomp.frx":0A14
         Left            =   6240
         List            =   "frmEdiProcomp.frx":0A27
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   5160
         TabIndex        =   4
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   4080
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação:"
         Height          =   195
         Left            =   6240
         TabIndex        =   6
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14685
      TabIndex        =   9
      Top             =   7380
      Width           =   14715
      Begin VB.CommandButton cmdHTML 
         Caption         =   "&Imprimir HTML"
         Height          =   615
         Left            =   7920
         Picture         =   "frmEdiProcomp.frx":0A7F
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   4800
         Picture         =   "frmEdiProcomp.frx":0BC9
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnPrint 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   6360
         Picture         =   "frmEdiProcomp.frx":0D13
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiProcomp.frx":0E5D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   3240
         Picture         =   "frmEdiProcomp.frx":0FA7
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEfetiv 
         Caption         =   "E&fetivar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiProcomp.frx":10F1
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   10800
         Picture         =   "frmEdiProcomp.frx":123B
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiProcomp.frx":1485
      Height          =   6255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   11033
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiProcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdHTML_Click()
     On Error GoTo trata_Erro:
    With ado
       ImprimePCPHtml !seq_pcp
    End With
    Exit Sub
trata_Erro:
    MostraErro
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
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
    strSQL = strSQL & " order by seq_pcp desc "
    
    abreObjetoAdo ado, strSQL, dtgrd
    Exit Sub
    rPesq.Open strSQL, conn_client
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
    On Error GoTo trata_Erro:
    Dim strSQL As String
    With ado
        If !sts_pcp = "1" Then
            If MsgBox("Deseja realmente excluir o Processo de Compra: " & !seq_pcp & "?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "Delete from tab_procomp where seq_pcp='" & !seq_pcp & "' "
                conn.Execute strSQL
                strSQL = "Delete from tab_itenspcp where seq_pcp='" & !seq_pcp & "' "
                conn.Execute strSQL
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
trata_Erro:
    MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_Erro:
    With ado
        Dim fr As New frmProcessoDeCompra
        fr.tag = !seq_pcp
        fr.Show 0
    
    End With
    Exit Sub
    
trata_Erro:
    MostraErro
End Sub

Private Sub cmndbtnEfetiv_Click()
    On Error GoTo trata_Erro:
    Dim strSQL As String
    With ado
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
trata_Erro:
    MostraErro
End Sub

Private Sub cmndbtnNovo_Click()
    Dim fr As New frmProcessoDeCompra
    fr.Show 0
End Sub

Private Sub cmndbtnPrint_Click()
    On Error GoTo trata_Erro:
    With ado
        dtgrd.Visible = False
       ImprimePcp !seq_pcp
       dtgrd.Visible = True
    End With
    Exit Sub
trata_Erro:
    MostraErro
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dtgrd_DblClick()
    On Error GoTo trata_Erro
    If ado!sts_pcp = "1" Then
        cmndbtnEdit_Click
    Else
    
    End If
    
    Exit Sub
trata_Erro:
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
    cmndbtnBuscar_Click
    
End Sub

Private Sub Form_Load()
    cmbxList.ListIndex = 0
    resizeForm Me
    Set ado = New ADODB.Recordset
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade dtgrd, Me, pcBot
    posicionaADireita cmndbtnSair, pcBot
End Sub

Private Sub txtbxAte_GotFocus()
    ganhaFoco txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxCliente_GotFocus()
    ganhaFoco txtbxCliente
End Sub

Private Sub txtbxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim c As fornecedor.fornecedor
        c = F_BuscaFornecedor(txtbxCliente.Text)
        If c.cod_for <> "" Then
            txtbxCliente.Text = c.nome_For
            txtbxCliente.tag = c.cod_for
            TrocaTabKey KeyCode
        Else
            txtbxCliente.Text = ""
            txtbxCliente.tag = ""
            txtbxCliente.SetFocus
        End If
    End If
End Sub

Private Sub txtbxCliente_LostFocus()
    AjustaTexto txtbxCliente
End Sub

Private Sub txtbxDe_GotFocus()
    ganhaFoco txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub

