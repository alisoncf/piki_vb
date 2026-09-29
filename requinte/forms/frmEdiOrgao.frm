VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiOrgao 
   Caption         =   "Edição de Órgãos"
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15150
   Icon            =   "frmEdiOrgao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   15150
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4215
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
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
      Caption         =   "Órgãos"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "cod_org"
         Caption         =   "Código"
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
         DataField       =   "nome_org"
         Caption         =   "Nome"
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
      BeginProperty Column02 
         DataField       =   "nome_cid"
         Caption         =   "Cidade"
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
      BeginProperty Column03 
         DataField       =   "nome_cob"
         Caption         =   "Cobrador"
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
      BeginProperty Column04 
         DataField       =   "endereco_org"
         Caption         =   "Endereco"
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
      BeginProperty Column05 
         DataField       =   "nome_set"
         Caption         =   "Setor"
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
      BeginProperty Column06 
         DataField       =   "data_org"
         Caption         =   "Atualização"
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
      BeginProperty Column07 
         DataField       =   "obs"
         Caption         =   "Obs"
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
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2775,118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2910,047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3254,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3240
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4680
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   15150
      TabIndex        =   0
      Top             =   8775
      Width           =   15150
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   6600
         Picture         =   "frmEdiOrgao.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   10320
         Picture         =   "frmEdiOrgao.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1110
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdClientes 
         Caption         =   "Clientes"
         Height          =   615
         Left            =   11520
         Picture         =   "frmEdiOrgao.frx":07DE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdMesclar 
         Caption         =   "Mesclar Órgão"
         Height          =   615
         Left            =   9000
         Picture         =   "frmEdiOrgao.frx":0928
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtbxBusca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   12720
         Picture         =   "frmEdiOrgao.frx":0A72
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   5400
         Picture         =   "frmEdiOrgao.frx":0FC0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   4200
         Picture         =   "frmEdiOrgao.frx":110A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnSelec 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   7800
         Picture         =   "frmEdiOrgao.frx":1494
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmEdiOrgao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmdClientes_Click()
    'On Error Resume Next
    Dim fr As New frmEdiClientes
    
    frmEdiClientes.txtbxOrgao.Text = ado!nome_org
    frmEdiClientes.txtbxOrgao.tag = ado!cod_org
    
    
    frmEdiClientes.Show 1
    
    
    
    
    
End Sub

Private Sub cmdDelete_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select cod_cli,cpf_cli,nome_cli from tab_clientes where cod_org='" & ado!cod_org & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
        
            MsgBox "Foram encontrados " & .RecordCount & " clientes neste órgão!", vbExclamation
            
        Else
            If MsgBox("Deseja excluir o órgão?", vbYesNo + vbQuestion) = vbYes Then
                strSQL = "delete from tab_orgaos where cod_org='" & ado!cod_org & "'"
                Conn.Execute strSQL
                ado.Requery
            End If
        End If
        
        
        
        
        .Close
    End With
    
    
End Sub

Private Sub cmdImprimir_Click()
    P_DtgrdToHtml "Lista de órgãos", grid, ado
End Sub

Private Sub cmdMesclar_Click()
    frmMesclarOrgao.Show 1
End Sub

Private Sub cmndbtnEdit_Click()
    
    With ado
        
        frmCadOrgaos.caption = !cod_org & "-" & !nome_org
        frmCadOrgaos.tag = !cod_org
        frmCadOrgaos.txtbxNome.Text = !nome_org
        frmCadOrgaos.txtbxNomeGru.Text = !nome_gru
        frmCadOrgaos.txtbxNomeGru.tag = !cod_gru
        frmCadOrgaos.txtbxEndereco.Text = !endereco_org
        frmCadOrgaos.txtbxBairro.Text = !nome_set
        frmCadOrgaos.txtbxBairro.tag = !cod_set
        frmCadOrgaos.txtbxCidade.Text = !nome_cid
        frmCadOrgaos.lblUF.caption = !uf_cid
        frmCadOrgaos.txtbxCEP.Text = !cep_org
        frmCadOrgaos.txtbxFone.Text = !fone_org
        frmCadOrgaos.txtCobrador.Text = !cobrador_org
        frmCadOrgaos.txtCobrador.tag = !cobrador_org
        frmCadOrgaos.txtObs.Text = Format(!obs)
        frmCadOrgaos.txtAtlz.Text = Format(!data_org)
        frmCadOrgaos.Show 1
        ado.Requery
    End With
    
End Sub

Private Sub cmndbtnNew_Click()
    frmCadOrgaos.Show 1
    P_Busca
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub cmndbtnSelec_Click()
    Me.Hide
End Sub

Private Sub dbgrd_DblClick()
    If Me.tag = "" Then
    
        cmndbtnEdit_Click
    Else
        cmndbtnSelec_Click
    End If
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    
    If Me.tag = "" Then
        cmndbtnEdit.Enabled = True
        cmndbtnSelec.Enabled = False
        
    Else
        cmndbtnEdit.Enabled = False
        cmndbtnSelec.Enabled = True
        
    End If
    P_Busca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sCod As String
    Dim strSQL As String
    If KeyCode = vbKeyF3 Then
        If strNome_Usu <> "ALISON" Then
            MsgBox "Apenas o Analista pode fazer isso!", vbExclamation
            Exit Sub
        End If
        If MsgBox("Deseja mesclar todos os orgãos para um só?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        sCod = InputBox("Informe o Código")
        If sCod = "" Then
            Exit Sub
        End If
        With ado
            .MoveFirst
            Do While Not .EOF
                strSQL = "update tab_clientes set cod_org='" & sCod & "' "
                strSQL = strSQL & " where cod_org='" & !cod_org & "' "
                Banco.Execute strSQL
                
                If sCod <> !cod_org Then
                    strSQL = "delete from tab_orgaos where cod_org='" & !cod_org & "' "
                    Banco.Execute strSQL
                End If
                .MoveNext
            Loop
            ado.Requery
        End With
        
    End If
    
End Sub

Private Sub Form_Load()
    
    Set ado = New ADODB.Recordset
    resizeForm Me
    
End Sub
Sub P_Busca()
    
    Dim strSQL As String
    strSQL = "select tog.*, tg.nome_gru,ts.nome_set,tc.nome_cid,tc.uf_cid,ts.cobrador_set,tb.nome_cob "
    strSQL = strSQL & " from  tab_orgaos as tog left join tab_setores as ts "
    strSQL = strSQL & " on ts.cod_set=tog.cod_set "
    strSQL = strSQL & " left join tab_cidades as tc on ts.cod_cid=tc.cod_cid "
    strSQL = strSQL & " left join tab_contacob as tb on tog.cobrador_org=tb.cod_cob "
   
  
    strSQL = strSQL & " left join tab_grupocli as tg on tg.cod_gru=tog.cod_gru  "
    
    
    strSQL = strSQL & " where "
    strSQL = strSQL & " nome_cid like '" & txtCidade.Text & "%' "
    If txtbxBusca.tag = "" Then
        strSQL = strSQL & " and nome_org like '%" & txtbxBusca.Text & "%'"
    Else
        strSQL = strSQL & " and cod_org ='" & txtbxBusca.tag & "'"
    End If
    abreObjetoAdo ado, strSQL, grid
    If ado.RecordCount = 1 Then
        If Me.tag <> "" Then
            On Error Resume Next
            Me.Hide
        End If
    Else
    
    End If
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBot
End Sub

Private Sub txtbxBusca_Change()
    P_Busca
End Sub

Private Sub txtbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_Busca
    End If
    
End Sub

Private Sub txtCidade_Change()
    P_Busca
End Sub
