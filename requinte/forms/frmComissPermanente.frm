VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmComissPermanente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Comissão Permanente:"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11520
      ScaleHeight     =   705
      ScaleWidth      =   3225
      TabIndex        =   26
      Top             =   960
      Width           =   3255
      Begin VB.TextBox txtbxPagar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   31
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagar:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblSelec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         Height          =   195
         Left            =   600
         TabIndex        =   29
         Top             =   0
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   405
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5775
      Left            =   11520
      TabIndex        =   25
      Top             =   2160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   10186
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14790
      TabIndex        =   15
      Top             =   0
      Width           =   14820
      Begin VB.CommandButton cmndbtnPagarCom 
         Caption         =   "Pagar Comissão"
         Height          =   615
         Left            =   11880
         Picture         =   "frmComissPermanente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnVisComissPrev 
         Caption         =   "&Comiss Prov"
         Height          =   615
         Left            =   9000
         Picture         =   "frmComissPermanente.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   7560
         Picture         =   "frmComissPermanente.frx":09C3
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   10440
         Picture         =   "frmComissPermanente.frx":0D4D
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   13320
         Picture         =   "frmComissPermanente.frx":115A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   285
         Left            =   6840
         Picture         =   "frmComissPermanente.frx":12C5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14790
      TabIndex        =   0
      Top             =   8025
      Width           =   14820
      Begin VB.CommandButton cmndbtnTodos 
         Caption         =   "Todos"
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkAberta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Comissões Não Pagas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1080
         TabIndex        =   18
         Top             =   120
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   12600
         TabIndex        =   23
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comissões Restantes:"
         Height          =   195
         Left            =   10800
         TabIndex        =   22
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   9840
         TabIndex        =   21
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comissões Pagas:"
         Height          =   195
         Left            =   8400
         TabIndex        =   20
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comissões Geradas:"
         Height          =   195
         Left            =   5520
         TabIndex        =   14
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label lblFinal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   7080
         TabIndex        =   13
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   3840
         TabIndex        =   12
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblQtd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   4800
         TabIndex        =   11
         Top             =   120
         Width           =   120
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmComissPermanente.frx":164F
      Height          =   6975
      Left            =   0
      OleObjectBlob   =   "frmComissPermanente.frx":1660
      TabIndex        =   16
      Top             =   960
      Width           =   11415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comissões Selecionadas:"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11520
      TabIndex        =   27
      Top             =   1800
      Width           =   3255
   End
End
Attribute VB_Name = "frmComissPermanente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim sCampo As String
    Dim sTab As String
    
    If txtbxNomeCli.tag = "" Then
        MsgBox "Escolha o Representante!", vbExclamation
        txtbxNomeCli.SetFocus
        Exit Sub
    End If
    
    strSQL = "select distinct concat(parcela_itd,'/',parcelas_doc) as parcela,tc.*,"
    strSQL = strSQL & "if(status_com='2',gerado_com,0) as comiss_pg, "
    strSQL = strSQL & " tp.reman_com,ta.data_ace,tp.seq_ace,ti.valor_itd,td.num_doc,td.parcelas_doc,td.valor_doc "
    strSQL = strSQL & " from tab_doc as td, tab_comisscons as tc, tab_comissprov as tp, tab_itdoc as ti,tab_acertos as ta "
    strSQL = strSQL & " where ti.seq_doc=td.seq_doc "
    strSQL = strSQL & " and tp.seq_com=tc.seq_com "
    strSQL = strSQL & " and tc.seq_itd=ti.seq_itd "
    strSQL = strSQL & " and ta.seq_ace=tp.seq_ace "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_com<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_com>=" & F_GravaData(txtbxDataIni.Text)
    End If
    'strSQL = strSQL & " and reman_com> 0 "
    
    strSQL = strSQL & " and tp.cod_rep='" & txtbxNomeCli.tag & "' "
    
    dt.RecordSource = strSQL
    dt.Refresh
    ajustaMsf
End Sub
Private Sub cmndbtnCliente_Click()
    Dim r As representante.representante
    
    r = F_BuscaRep(txtbxNomeCli.Text)
    If r.cod_rep <> "" Then
        txtbxNomeCli.tag = r.cod_rep
        txtbxNomeCli.Text = r.nome_rep
        
        cmndbtnBuscar_Click
    End If
    
    
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DbgrdToHtml "Comissão Provisória de " & txtbxNomeCli.Text & " de " & txtbxDataIni.Text _
    & " até " & txtbxdataFin.Text, dbgrd, dt
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_DblClick()
    If dt.Recordset!status_com = "1" Then
        msflxgrd.Row = dbgrd.Row
        If msflxgrd.Text = "" Then
            msflxgrd.Text = "S"
        Else
            msflxgrd.Text = ""
        End If
    Else
    
    End If
End Sub

Private Sub Form_Load()
    dt.Connect = Banco.Connect
    resizeForm Me
End Sub
Sub ajustaMsf()
    Dim iCount As Integer
    msflxgrd.FixedCols = 0
    msflxgrd.Cols = 1
    msflxgrd.rows = 0
    On Error Resume Next

    msflxgrd.rows = dt.Recordset.RecordCount + 1
    msflxgrd.Height = dbgrd.Height
    msflxgrd.Top = dbgrd.RowTop(4)
    
    For iCount = 0 To msflxgrd.rows - 1
        msflxgrd.Row = iCount
        msflxgrd.RowHeight(iCount) = dbgrd.RowHeight
    Next
End Sub
Private Sub lblAp_Click()

End Sub
Sub soma()
    Dim cTotal As Currency
    On Error GoTo trata_erro
    
    With dt.Recordset
        
        Do While Not .EOF
            .MoveNext
        Loop
        .MoveFirst
    End With
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Private Sub Form_Resize()
    dbgrd.Width = Me.Width - dbgrd.Left - 100
    dbgrd.Height = pcBot.Top - dbgrd.Top - 100
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub
