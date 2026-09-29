VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiCid 
   Caption         =   "Edição de Cidades"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8625
   Icon            =   "frmEdiCid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ligar Cidade"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   8595
      TabIndex        =   1
      Top             =   5700
      Width           =   8625
      Begin VB.CommandButton cmdDelete 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   5160
         Picture         =   "frmEdiCid.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   1350
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ordem:"
         Height          =   615
         Left            =   6360
         TabIndex        =   8
         Top             =   0
         Width           =   2055
         Begin VB.OptionButton optnUf 
            Caption         =   "UF"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optnCid 
            Caption         =   "Cidade"
            Height          =   255
            Left            =   720
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton cmndbtnEditar 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   3480
         Picture         =   "frmEdiCid.frx":0D8C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sai&r"
         Height          =   615
         Left            =   6600
         Picture         =   "frmEdiCid.frx":0ED6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSelect 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   1800
         Picture         =   "frmEdiCid.frx":0FCB
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiCid.frx":109A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtbxNome 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   540
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiCid.frx":11F6
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "cod_cid"
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
         DataField       =   "nome_cid"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5639,812
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   569,764
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiCid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdDelete_Click()
            Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select cod_set from tab_setores as ts where cod_cid='" & ado!cod_cid & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
        
            MsgBox "Foram encontrados " & .RecordCount & " setores nesta cidade!", vbExclamation
            
        Else
            If MsgBox("Deseja excluir a cidade?", vbYesNo + vbQuestion) = vbYes Then
                strSQL = "delete from tab_cidades where cod_cid='" & ado!cod_cid & "'"
                Conn.Execute strSQL
                ado.Requery
            End If
        End If
        
        
        
        
        .Close
    End With
End Sub

Private Sub cmndbtnEditar_Click()
    On Error GoTo trata_erro
    With ado
        frmIncCid.tag = !cod_cid
        frmIncCid.txtCodigo.Text = !cod_cid
        frmIncCid.cmbxUf.Text = !uf_cid
        frmIncCid.txtbxDesc.Text = !nome_cid
        
        frmIncCid.Show 1
    End With
Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmndbtnNovo_Click()
    frmIncCid.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub cmndbtnSelect_Click()
    On Error Resume Next
    Me.tag = adc.Recordset!cod_cid
    Me.Hide
End Sub

Private Sub Command1_Click()
    Dim strSQL As String
    Dim cod As String
    With ado
        Do While Not .EOF
            cod = F_BuscaUmCampo("cod_cid", "cidades", "nome_cid='" & !nome_cid & "' and uf_cid='" & !uf_cid & "'")
            
            strSQL = "update tab_cidades set cod_mun='" & cod & "' where cod_cid='" & !cod_cid & "' "
            Conn.Execute strSQL
            
            
            .MoveNext
        Loop
    End With
    
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
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    Set ado = New ADODB.Recordset
    
    P_CentraForm Me
End Sub
Sub busca()
    On Error GoTo trata_erro
    Dim strSQL As String
    strSQL = "select * from " & TAB_CIDADES
    strSQL = strSQL & " where nome_cid like '%" & txtbxNome.Text & "%'"
    If optnCid.Value Then
        strSQL = strSQL & " order by nome_cid "
    Else
        strSQL = strSQL & " order by uf_cid,nome_cid "
    End If
    abreObjetoAdo ado, strSQL, dtgrd
    Exit Sub
trata_erro:
    P_MostraErro
    
End Sub
Private Sub tlbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        With Adodc.Recordset
        frmIncCid.tag = !cod_cid
        frmIncCid.cmbxUf.Text = !uf_cid
        frmIncCid.txtbxDesc.Text = !nome_cid
        
        frmIncCid.Show 1
        End With
    ElseIf Button.Index = 2 Then
        On Error Resume Next
        frmIncCid.Show 1
    ElseIf Button.Index = 3 Then
    ElseIf Button.Index = 6 Then
        Unload Me
    End If
    
End Sub

Private Sub optnCid_Click()
    busca
End Sub

Private Sub optnUf_Click()
    busca
End Sub

Private Sub txtbxNome_Change()
    busca
End Sub

