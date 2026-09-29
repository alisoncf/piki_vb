VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVisualizarAcerto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerto:"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   685
      Left            =   8520
      Picture         =   "frmVisualizarAcerto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   685
      Left            =   10440
      Picture         =   "frmVisualizarAcerto.frx":040D
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6360
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   12480
      TabIndex        =   2
      Top             =   0
      Width           =   12510
      Begin VB.CommandButton cmdVisualizar 
         Caption         =   "Visualizar"
         Height          =   255
         Left            =   7560
         TabIndex        =   29
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblComiss 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00030ED3&
         Height          =   285
         Left            =   11520
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comiss:"
         Height          =   195
         Left            =   11520
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   6240
         TabIndex        =   22
         Top             =   600
         Width           =   525
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   21
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Representante:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblNomeRep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Pag:"
         Height          =   195
         Left            =   10320
         TabIndex        =   18
         Top             =   0
         Width           =   810
      End
      Begin VB.Label lblFormaPag 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10320
         TabIndex        =   17
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cob:"
         Height          =   195
         Left            =   9240
         TabIndex        =   16
         Top             =   0
         Width           =   690
      End
      Begin VB.Label lblCobranca 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         TabIndex        =   15
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblCapDiv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença:"
         Height          =   195
         Left            =   7560
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblDif 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acerto:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   510
      End
      Begin VB.Label lblSeqAce 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   1200
         TabIndex        =   10
         Top             =   0
         Width           =   390
      End
      Begin VB.Label lblDataAce 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   0
         Width           =   390
      End
      Begin VB.Label lblHora 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   3360
         TabIndex        =   6
         Top             =   0
         Width           =   585
      End
      Begin VB.Label lblUsu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   6240
         TabIndex        =   4
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblValorAce 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrdAcerto 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   7011
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrdComiss 
      Height          =   1215
      Left            =   0
      TabIndex        =   25
      Top             =   5880
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comissões:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   26
      Top             =   5640
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Produtos:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   12495
   End
End
Attribute VB_Name = "frmVisualizarAcerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub AjustaGrid()
    With msflxgrdAcerto
        .rows = 1
        .Cols = 7
        .col = 0: .text = "Código"
        
        .col = 1: .text = "Produto"
        .col = 2: .text = "Qtd"
        .col = 3: .text = "Unit Vendido"
        .col = 4: .text = "Total"
        .col = 5: .text = "Cons."
        .col = 6: .text = "Prç At."
        .ColWidth(0) = 1000
        
        .ColWidth(1) = 5500
        
    End With

    With msflxgrdComiss
        .rows = 1
        .Cols = 5
        .col = 0: .text = "Seq.itd"
        .col = 1: .text = "Valor"
        .col = 2: .text = "Vencimento"
        .col = 3: .text = "Status"
        .col = 4: .text = "Seq.Cai"
        
    End With
End Sub

Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ta.*, tr.nome_rep,tf.nome_for from tab_acertos as ta, tab_formaspag as tf, "
    strSQL = strSQL & "tab_representantes as tr "
    strSQL = strSQL & " where tf.cod_for=ta.cod_for"
    strSQL = strSQL & " and tr.cod_rep=ta.cod_rep "
    strSQL = strSQL & " and ta.seq_ace='" & Me.tag & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Não encontrado!", vbExclamation
            .Close
            Exit Sub
        End If
        
        lblSeqAce.caption = !seq_ace
        If !status_ace = "3" Then
            MsgBox "Este acerto foi cancelado!", vbExclamation
            
        End If
        lblDataAce.caption = !data_ace
        lblHora.caption = !hora_ace
        lblUsu.caption = F_BuscaUmCampo("guerra_fun", "tab_funcionarios", "cod_fun='" & !cod_usu & "'")
        lblValorAce.caption = FormatVal(!valor_ace)
        lblDif.caption = FormatVal(IIf(!diferenca_ace > 0, !diferenca_ace * 1, !diferenca_ace), "###,##0")
        lblCobranca.caption = !cobranca_ace
        lblFormaPag.caption = !nome_for
        lblComiss.caption = FormatVal(!comissao_ace)
        lblNomeRep.caption = !cod_Rep & "-" & !nome_rep
        If !tipo_ter = "R" Then
            lblCliente.caption = "O Próprio Representante"
        Else
            lblCliente.caption = !cod_cli & "-" & F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        End If
        
        .Close
    End With
    
    
    buscaItens
End Sub
Sub buscaItens()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    msflxgrdAcerto.rows = 1
    strSQL = "select tp.nome_pro,tp.pv_pro,ti.* from tab_produtos as tp, tab_itacertos as ti "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.seq_ace='" & Me.tag & "'"
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = !cod_pro & Chr(9)
            strSQL = strSQL & !nome_pro & Chr(9)
            strSQL = strSQL & FormatVal(!qtd_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!qtd_ite * !unit_ite) & Chr(9)
            
            strSQL = strSQL & !seq_con & Chr(9)
            strSQL = strSQL & FormatVal(!pv_pro)
            msflxgrdAcerto.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
            
    ZebraGrid msflxgrdAcerto, 1
    buscaComiss
End Sub
Sub buscaComiss()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_comissparc where seq_ace='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = !seq_itd & Chr(9) & FormatVal(!valor_com) & Chr(9)
            strSQL = strSQL & !venc_com & Chr(9)
            strSQL = strSQL & !status_com & Chr(9)
            strSQL = strSQL & !seq_cai
            msflxgrdComiss.AddItem strSQL
            
        

            .MoveNext
        Loop
        .Close
    End With
End Sub


Private Sub cmdVisualizar_Click()
    If strCadFun_Usu <> "S" Then
        MsgBox "Não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    Else
        lblDif.Visible = True
        lblCapDiv.Visible = True
    End If
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnImprimir_Click()
    imprimeacerto lblSeqAce.caption
End Sub

Private Sub Form_Activate()
    carrega
End Sub

Private Sub Form_Load()
    AjustaGrid
End Sub
