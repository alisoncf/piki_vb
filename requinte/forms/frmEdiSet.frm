VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiSet 
   Caption         =   "Edição de Setores"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11370
   Icon            =   "frmEdiSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11370
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11340
      TabIndex        =   6
      Top             =   5385
      Width           =   11370
      Begin VB.TextBox txtbxBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtbxCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton cmndbtnBuscaCid 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         Picture         =   "frmEdiSet.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         Picture         =   "frmEdiSet.frx":07CC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   195
         Left            =   4320
         TabIndex        =   13
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblUF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
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
         Left            =   3720
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11370
      TabIndex        =   1
      Top             =   6000
      Width           =   11370
      Begin VB.CommandButton cmdDelete 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   4440
         Picture         =   "frmEdiSet.frx":0B56
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1350
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "Sai&r"
         Height          =   615
         Left            =   9960
         Picture         =   "frmEdiSet.frx":0EE0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSelect 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   1560
         Picture         =   "frmEdiSet.frx":0FD5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiSet.frx":10A4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEditar 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   3000
         Picture         =   "frmEdiSet.frx":1200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiSet.frx":1C02
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Registros Encontrados"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "cod_set"
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
      BeginProperty Column02 
         DataField       =   "tipo_set"
         Caption         =   "Tipo"
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
      BeginProperty Column04 
         DataField       =   "uf_cid"
         Caption         =   "UF"
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
         DataField       =   "cod_cob"
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
      BeginProperty Column06 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3555,213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2145,26
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   434,835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2940,095
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub adcgrd_Click()

End Sub

Private Sub cmdDelete_Click()
        Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select cod_cli,cpf_cli,nome_cli from tab_clientes where cod_set='" & ado!cod_set & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
        
            MsgBox "Foram encontrados " & .RecordCount & " clientes neste setor!", vbExclamation
            
        Else
            If MsgBox("Deseja excluir o setor?", vbYesNo + vbQuestion) = vbYes Then
                strSQL = "delete from tab_setores where cod_set='" & ado!cod_set & "'"
                Conn.Execute strSQL
                ado.Requery
            End If
        End If
        
        
        
        
        .Close
    End With
    
End Sub

Private Sub cmndbtnBuscaCid_Click()
    buscaCid
End Sub
Sub buscaCid()
    Dim c As Cidade
    c = Cidades.F_BuscaCidades(txtbxCidade.Text)
    
    If c.cod_cid <> "" Then
        txtbxCidade.Text = c.nome_cid
        txtbxCidade.tag = c.cod_cid
        lblUF.caption = c.uf_cid
        
        
    End If
    
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnEditar_Click()
    frmCadSetor.caption = ado!nome_set
    frmCadSetor.tag = ado!cod_set
    frmCadSetor.txtbxCidade.tag = ado!cod_cid
    frmCadSetor.txtbxCidade.Text = ado!nome_cid
    frmCadSetor.txtbxNome.Text = ado!nome_set
    frmCadSetor.cmbxTipo.Text = ado!tipo_set
    frmCadSetor.txtbxCobrador.tag = ado!cobrador_set
    frmCadSetor.txtbxCobrador.Text = F_BuscaUmCampo("nome_cob", "tab_contacob", "cod_cob='" & ado!cobrador_set & "'")
    
    
    
    frmCadSetor.Show 1
End Sub

Private Sub cmndbtnNovo_Click()
    frmCadSetor.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub cmndbtnSelect_Click()
    Me.tag = ado!cod_set
    Me.Hide
End Sub

Private Sub adcgrd_DblClick()
    edita
End Sub
Sub edita()
    If Me.tag = "" Then
        cmndbtnEditar_Click
    Else
        cmndbtnSelect_Click
    End If
End Sub

Private Sub adcgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        edita
    End If
End Sub

Private Sub dtgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEditar_Click
    Else
        cmndbtnSelect_Click
    End If
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
    
End Sub

Private Sub Form_Activate()
    
    busca
    If Me.tag = "" Then
        cmndbtnSelect.Enabled = False
    Else
        cmndbtnSelect.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Set ado = New ADODB.Recordset
End Sub

Private Sub tlbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next '
    If Button.Index = 1 Then
        With Adodc.Recordset
            
            
        End With
    ElseIf Button.Index = 2 Then
        On Error Resume Next
        frmIncArea.Show 1
    ElseIf Button.Index = 3 Then
    ElseIf Button.Index = 6 Then
        Unload Me
    End If
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    dtgrd.width = Me.width - dtgrd.Left - 200
    dtgrd.height = pcBusca.Top - dtgrd.Top - 200
    posicionaADireita cmndbtnSair, pcBusca
End Sub

Private Sub txtbxBairro_Change()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select ts.cod_set,ts.cod_cid,tcc.cod_cob,ts.cobrador_set,ts.nome_set,ts.tipo_set,tcc.nome_cob, tc.nome_cid,tc.uf_cid from tab_cidades as tc inner join "
    strSQL = strSQL & " tab_setores as ts  on "
    strSQL = strSQL & "  tc.cod_cid=ts.cod_cid "
    strSQL = strSQL & " left join tab_contacob as tcc on ts.cobrador_set=tcc.cod_cob "
    
    strSQL = strSQL & " where 1> 0 "
    
    If txtbxCidade.tag <> "" Then
        strSQL = strSQL & " and tc.cod_cid='" & txtbxCidade.tag & "' "
    ElseIf txtbxCidade.Text <> "" Then
        strSQL = strSQL & " and nome_cid like '" & txtbxCidade.Text & "%'"
    End If
    
    
    strSQL = strSQL & " and nome_set like '%" & txtbxBairro.Text & "%'"
    strSQL = strSQL & " order by tc.nome_cid,nome_set "
        
    abreObjetoAdo ado, strSQL, dtgrd
    
   
    
End Sub

Private Sub txtbxCidade_Change()
    txtbxCidade.tag = ""
End Sub
