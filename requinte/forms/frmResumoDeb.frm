VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmResumoDeb 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumo de Débitos"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7860
      TabIndex        =   6
      Top             =   6435
      Width           =   7890
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7200
         Picture         =   "frmResumoDeb.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblTotal 
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
         Left            =   5400
         TabIndex        =   12
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblAVencer 
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
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Vencer"
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblVencidas 
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
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencidas:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.ComboBox cmbxTipoTer 
      Height          =   315
      ItemData        =   "frmResumoDeb.frx":016B
      Left            =   120
      List            =   "frmResumoDeb.frx":0175
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtbxTerceiro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton cmndbtnTerc 
      Height          =   285
      Left            =   6960
      Picture         =   "frmResumoDeb.frx":0191
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9763
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo de Terceiro:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terceiro:"
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   630
   End
End
Attribute VB_Name = "frmResumoDeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTerc_Click()
    If cmbxTipoTer.ListIndex = 0 Then
        buscaCLi
    Else
        buscaRep
    End If
     busca
End Sub
Sub buscaCLi()
    Dim c As cliente
    c = F_BuscaCliente(txtbxTerceiro.Text)
        If c.cod_cli <> "" Then
            txtbxTerceiro.Text = c.nome_Cli
            txtbxTerceiro.tag = c.cod_cli
        End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.Text)
        If r.cod_Rep <> "" Then
            txtbxTerceiro.Text = r.nome_rep
            txtbxTerceiro.tag = r.cod_Rep
        End If
    
End Sub

Private Sub Form_Activate()
    If Me.tag <> "" Then
        cmndbtnTerc_Click
        Me.tag = ""
    End If
End Sub

Private Sub Form_Load()
    ajustagrid
End Sub
Sub ajustagrid()
    With msflxgrd
        .rows = 1
        .Cols = 9
        .FixedCols = 2
        
        .Col = 0: .Text = "..."
        .Col = 2: .Text = "Id.Parc."
        .Col = 2: .Text = "Id.Doc."
        .Col = 3: .Text = "Nº Doc"
        .Col = 4: .Text = "Parcela"
        .Col = 5: .Text = "Vencimento"
        .Col = 6: .Text = "Valor"
        .Col = 7: .Text = "Restam"
        .Col = 8: .Text = "Histórico"
        '.Col = 9: .Text = "Cob"
        
        .ColWidth(8) = 3000
        
    End With
End Sub

Sub busca()
    Dim strSQL As String
    Dim s As setor
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sTipoTer As String
    Dim sEnd As String
    Dim cTotal As Currency
    Dim cVencidas As Currency
    Dim cAVencer As Currency
    Dim rPesq As New adodb.Recordset
    msflxgrd.rows = 1
    
        sOrderBy = "venc_Itd"
    
    
    If cmbxTipoTer.ListIndex = 0 Then
        sTab = "tab_clientes"
        sCod = "cod_cli"
        sNome = "nome_cli"
        sTipoTer = "C"
        sEnd = "endereco_cli"
    Else
        sTab = "tab_representantes"
        sNome = "nome_rep"
        sCod = "cod_rep"
        sEnd = "endereco_rep"
        sTipoTer = "R"
    End If
    
    strSQL = "select (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
    strSQL = strSQL & "tc." & sCod & " as codigo, tc." & sEnd & " as endereco,  tc.*,td.num_ven,tc." & sNome & " as nome,td.historico_doc,td.parcelas_doc, "
    strSQL = strSQL & "td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc,"
    strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
    strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, ti.* "
    
    strSQL = strSQL & "from " & sTab & " as tc, tab_doc as td, "
    strSQL = strSQL & "tab_itdoc as ti,tab_parametros as tp    "
    strSQL = strSQL & "where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
    
    
    
    
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and td.tipo_ter='" & sTipoTer & "' "
    
    
    
    
    
    strSQL = strSQL & " and ti.reman_itd > 0 "
    
    
    strSQL = strSQL & " and td.cod_ter='" & txtbxTerceiro.tag & "' "
    
    
    strSQL = strSQL & " order by " & sOrderBy
    

    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = .AbsolutePosition + 1 & Chr(9)
            
            strSQL = strSQL & !seq_itd & Chr(9)
            strSQL = strSQL & !seq_doc & Chr(9)
            strSQL = strSQL & !num_doc & Chr(9)
            strSQL = strSQL & !parcela_itd & "/" & !parcelas_doc & Chr(9)
            strSQL = strSQL & !venc_itd & Chr(9)
            strSQL = strSQL & FormatVal(!valor_itd) & Chr(9)
            strSQL = strSQL & FormatVal(!reman_itd) & Chr(9)
            strSQL = strSQL & !historico_doc
            '
            
            If CDate(rPesq!venc_itd) < CDate(Date) Then
                cVencidas = cVencidas + !reman_itd
            Else
                cAVencer = cAVencer + !reman_itd
            End If
            
            msflxgrd.AddItem strSQL
            
            .MoveNext
        Loop
        .Close
    End With
    
    ZebraGrid msflxgrd, 1
    lblVencidas.caption = FormatVal(cVencidas)
    lblAVencer.caption = FormatVal(cAVencer)
    lblTotal.caption = FormatVal(cVencidas + cAVencer)
End Sub

