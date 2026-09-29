VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHistProdConsignado 
   Caption         =   "Histórico de Produtos por Representante"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15825
   Icon            =   "frmHistProdConsignado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   15825
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   15795
      TabIndex        =   15
      Top             =   0
      Width           =   15825
      Begin VB.TextBox txtbxRepr 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   23
         Top             =   240
         Width           =   5055
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   12720
         Picture         =   "frmHistProdConsignado.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   495
         Left            =   14760
         Picture         =   "frmHistProdConsignado.frx":0D8C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   13680
         Picture         =   "frmHistProdConsignado.frx":0EF7
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   30
         Width           =   855
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10560
         MaxLength       =   10
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11640
         MaxLength       =   10
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscaPro 
         Height          =   315
         Left            =   10080
         Picture         =   "frmHistProdConsignado.frx":1304
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   315
         Left            =   4680
         Picture         =   "frmHistProdConsignado.frx":13FD
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço(R$):"
         Height          =   195
         Left            =   3360
         TabIndex        =   34
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblPv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   33
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   1320
         TabIndex        =   32
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblCodPro 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rep:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   345
      End
      Begin VB.Label lblCodRep 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   29
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblRep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblProduto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   5040
         TabIndex        =   27
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   10560
         TabIndex        =   26
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblAte 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   11640
         TabIndex        =   25
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox pcAcerto 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   10560
      ScaleHeight     =   5985
      ScaleWidth      =   5145
      TabIndex        =   10
      Top             =   960
      Width           =   5175
      Begin MSFlexGridLib.MSFlexGrid msflxgrdAcerto 
         Height          =   3495
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   6165
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMarcadosA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marcados:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lblAcerto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Acertos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   675
      End
      Begin VB.Label lblQA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3120
         TabIndex        =   12
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.PictureBox pcDev 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   5280
      ScaleHeight     =   5985
      ScaleWidth      =   5145
      TabIndex        =   5
      Top             =   960
      Width           =   5175
      Begin MSFlexGridLib.MSFlexGrid msflxgrdDev 
         Height          =   3375
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5953
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label lblMarcadosD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marcados:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lblDev 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "Devoluções"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label lblQd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3120
         TabIndex        =   7
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.PictureBox pcCons 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6105
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   960
      Width           =   5175
      Begin MSFlexGridLib.MSFlexGrid msflxgrdCons 
         Height          =   3495
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6165
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label lblEstoque 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   4680
         TabIndex        =   36
         Top             =   120
         Width           =   105
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3960
         TabIndex        =   35
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblMarcadosT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marcados:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2280
         TabIndex        =   3
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lblQt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3240
         TabIndex        =   2
         Top             =   120
         Width           =   105
      End
      Begin VB.Label lblCons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "Transferências"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   1305
      End
   End
   Begin VB.Menu mnAce 
      Caption         =   "&Acerto"
      Visible         =   0   'False
      Begin VB.Menu mnVisAce 
         Caption         =   "Visualizar Acerto"
      End
   End
End
Attribute VB_Name = "frmHistProdConsignado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmndbtnBuscaPro_Click()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxProduto.Text)
    txtbxProduto.Text = p.nome_pro
    txtbxProduto.tag = p.cod_pro
    lblCodPro.caption = p.cod_pro
    lblPv.caption = FormatVal(p.pv_pro)
    txtbxDataIni.SetFocus
    busca
    
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnCliente_Click()
    Dim r As representante.representante
    
    r = representante.F_BuscaRep(txtbxRepr.Text)
    If r.cod_Rep <> "" Then
        txtbxRepr.Text = r.nome_rep
        txtbxRepr.tag = r.cod_Rep
        lblCodRep.caption = r.cod_Rep
        txtbxProduto.SetFocus
        busca
        
    End If
    
End Sub
Sub ajustagrid()
    With msflxgrdCons
        .rows = 1
        .Cols = 6
        .Col = 0: .Text = "Nº Cons."
        .Col = 1: .Text = "Data"
        .Col = 2: .Text = "Qtd"
        .Col = 3: .Text = "Reman"
        .Col = 4: .Text = "Usuário"
        .Col = 5: .Text = "Retif"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 600
    End With
    With msflxgrdDev
        .rows = 1
        .Cols = 5
        .Col = 0: .Text = "Nº Dev."
        .Col = 1: .Text = "Data Dev."
        .Col = 2: .Text = "Qtd Dev."
        .Col = 3: .Text = "Usuário"
        .Col = 4: .Text = "Nº Cons."
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
    End With
    With msflxgrdAcerto
        .rows = 1
        .Cols = 9
        .Col = 0: .Text = "Nº Ace."
        .Col = 1: .Text = "Data Ace."
        .Col = 2: .Text = "Hora Ace."
        .Col = 3: .Text = "Usuário"
        .Col = 4: .Text = "Cliente."
        .Col = 5: .Text = "Qtd"
        .Col = 6: .Text = "Valor"
        .Col = 7: .Text = "Nº Cons"
        .Col = 8: .Text = "Status"
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 600
        .ColWidth(3) = 800
        .ColWidth(4) = 3500
        .ColWidth(5) = 800
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
    End With
    
    
    
End Sub
Sub busca()
        Dim strSQL As String
    Dim cEst As Currency
    Dim rPesq As New ADODB.Recordset
    Dim iCount As Integer
    ajustagrid
    strSQL = "select tc.*,tp.nome_pro from tab_consignacoes as tc,"
    strSQL = strSQL & " tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=tc.cod_pro "
    strSQL = strSQL & " and tc.cod_rep='" & txtbxRepr.tag & "' "
    strSQL = strSQL & " and tc.cod_pro='" & txtbxProduto.tag & "' "
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_con>=" & F_GravaDataTempo(txtbxDataIni.Text)
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_con<=" & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:00")
    End If
    
    strSQL = strSQL & " order by data_con "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = !seq_con & Chr(9)
            strSQL = strSQL & !data_con & Chr(9)
            strSQL = strSQL & !qtd_con & Chr(9)
            strSQL = strSQL & !reman_con & Chr(9)
            strSQL = strSQL & !cod_usu & Chr(9)
            strSQL = strSQL & !retif_con
            cEst = cEst + CStr(!reman_con)
            msflxgrdCons.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    lblEstoque.caption = cEst
    
    strSQL = "select ti.*,td.cod_usu,td.data_dev from tab_devconsign as td,tab_itdevconsign as ti"
    strSQL = strSQL & " where td.seq_dev=ti.seq_dev "
    strSQL = strSQL & " and td.cod_rep='" & txtbxRepr.tag & "' "
    strSQL = strSQL & " and ti.cod_pro='" & txtbxProduto.tag & "' "
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_dev>=" & F_GravaDataTempo(txtbxDataIni.Text)
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_dev>=" & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:00")
    End If
    
    strSQL = strSQL & " order by data_dev "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = !seq_dev & Chr(9)
            strSQL = strSQL & !data_dev & Chr(9)
            strSQL = strSQL & !qtd_dev & Chr(9)
            strSQL = strSQL & !cod_usu & Chr(9)
            strSQL = strSQL & !seq_con
            msflxgrdDev.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    strSQL = "select ti.qtd_ite,ti.unit_ite,ti.seq_con,ta.* "
    strSQL = strSQL & " from tab_acertos as ta,tab_itacertos as ti "
    strSQL = strSQL & " where ta.seq_ace=ti.seq_ace "
    strSQL = strSQL & " and ta.cod_rep='" & txtbxRepr.tag & "' "
    strSQL = strSQL & " and ti.cod_pro='" & txtbxProduto.tag & "' "
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_Ace>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_Ace>=" & F_GravaData(txtbxdataFin.Text)
    End If
    
    strSQL = strSQL & " order by data_ace "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = !seq_ace & Chr(9)
            strSQL = strSQL & !data_ace & Chr(9)
            strSQL = strSQL & !hora_ace & Chr(9)
            strSQL = strSQL & !cod_usu & Chr(9)
            If !tipo_ter = "R" Then
                strSQL = strSQL & txtbxRepr.Text & Chr(9)
            Else
                strSQL = strSQL & F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_cli='" & !cod_cli & "'") & Chr(9)
            End If
            strSQL = strSQL & !qtd_ite & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9)
            strSQL = strSQL & !seq_con & Chr(9)
            strSQL = strSQL & !status_ace
            
            msflxgrdAcerto.AddItem strSQL
            If !status_ace = "3" Then
                msflxgrdAcerto.row = msflxgrdAcerto.rows - 1
                
                ZebraGrid msflxgrdAcerto, msflxgrdAcerto.row, vbBlue, vbBlue ', True, True
            End If
            .MoveNext
        Loop
        .Close
    End With
    
    
    
End Sub


Private Sub cmndbtnImpr_Click()
    msflxgrdAcerto.Col = 0
    imprAcertoData txtbxRepr.tag, msflxgrdAcerto.Text
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    If Me.tag <> "" Then
        busca
    End If
    
    bActive = True
    
    
End Sub

Private Sub Form_Load()
            'dt.Connect = Banco.Connect
            resizeForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim s As Single
    pcCons.width = (Me.width / 3) - (Me.Top / 3)
    pcDev.width = pcCons.width
    pcAcerto.width = pcCons.width
    pcDev.Left = pcCons.Left + pcCons.width + 50
    pcAcerto.Left = pcDev.Left + pcDev.width + 50
    pcCons.height = Me.height - pcCons.Top - 500
    pcDev.height = pcCons.height
    pcAcerto.height = pcCons.height
    msflxgrdCons.width = pcCons.width - msflxgrdCons.Left - 150
    msflxgrdDev.width = pcDev.width - msflxgrdDev.Left - 150
    msflxgrdAcerto.width = pcAcerto.width - msflxgrdAcerto.Left - 150
    msflxgrdCons.height = pcCons.height - msflxgrdCons.Top - 100
    msflxgrdDev.height = pcDev.height - msflxgrdDev.Top - 100
    msflxgrdAcerto.height = pcAcerto.height - msflxgrdAcerto.Top - 100
    txtbxRepr.width = pcCons.width - cmndbtnBuscaPro.width - 200
    txtbxProduto.Left = pcDev.Left
    lblProduto.Left = txtbxProduto.Left
    txtbxProduto.width = txtbxRepr.width
    cmndbtnBuscaPro.Left = txtbxProduto.Left + txtbxProduto.width
    cmndbtnCliente.Left = txtbxRepr.Left + txtbxRepr.width
    
    txtbxDataIni.Left = pcAcerto.Left
    lblDe.Left = txtbxDataIni.Left
    
    txtbxdataFin.Left = txtbxDataIni.Left + txtbxDataIni.width + 100
    lblAte.Left = txtbxdataFin.Left
    
    
    cmndbtnBuscar.Left = txtbxdataFin.Left + txtbxdataFin.width + 100
    cmndbtnImpr.Left = cmndbtnBuscar.Left + cmndbtnBuscar.width + 50
    cmndbtnSair.Left = cmndbtnImpr.Left + cmndbtnImpr.width + 50
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bActive = False
End Sub


Sub marca_Cons()
    Dim iRow As Integer
    Dim sSEQ As String
    Dim iCount As Integer
    Dim iCol As Integer
    Dim Qt As Integer
    Dim Qa As Integer
    Dim Qd As Integer
    Dim sSts As String
    On Error Resume Next
    
    iRow = msflxgrdCons.row
    msflxgrdCons.Col = 2
    Qt = CInt(msflxgrdCons.Text)
    
    msflxgrdCons.Col = 0
    
    sSEQ = msflxgrdCons.Text
    ZebraGrid msflxgrdCons, 1
    ZebraGrid msflxgrdDev, 1
    ZebraGrid msflxgrdAcerto, 1
    msflxgrdCons.row = iRow
    ZebraGridPedaco msflxgrdCons, iRow, iRow, vbRed, vbRed
    lblQt.caption = Qt
    
    Qd = 0
    With msflxgrdDev
        .Col = 4
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 4
            If .Text = sSEQ Then
                .Col = 2
                Qd = Qd + CInt(.Text)
                ZebraGridPedaco msflxgrdDev, iCount, iCount, vbRed, vbRed
                
                msflxgrdAcerto.TopRow = iCount
            End If
        Next
        lblQd.caption = Qd
    End With
    Qa = 0
    With msflxgrdAcerto
        .Col = .Cols - 1
        sSts = .Text
        .Col = 7
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 7
            
            If .Text = sSEQ Then
                .Col = .Cols - 1
                sSts = .Text
                .Col = 5
                If sSts <> "3" Then
                    Qa = Qa + CInt(.Text)
                    ZebraGridPedaco msflxgrdAcerto, iCount, iCount, vbRed, vbRed
                Else
                    ZebraGridPedaco msflxgrdAcerto, iCount, iCount, vbBlue, vbBlue
                End If
                
                msflxgrdAcerto.TopRow = iCount
            End If
        Next
        lblQA.caption = Qa
    End With
    msflxgrdCons.row = iRow
    
    msflxgrdCons.SetFocus
     
End Sub

Sub acha_linha_acerto(seq_con As String)
    Dim iRow As Integer
    Dim sSEQ As String
    Dim iCount As Integer
    Dim iCol As Integer
    Dim Qt As Integer
    Dim Qa As Integer
    Dim Qd As Integer
    Dim sSts As String
   ' On Error Resume Next
        
    
    iRow = msflxgrdAcerto.row
    msflxgrdAcerto.Col = 2
   ' Qt = CInt(msflxgrdAcerto.Text)
    Qa = 0
    With msflxgrdAcerto
        .Col = .Cols - 1
        sSts = .Text
        .Col = 7
        For iCount = 1 To .rows - 1
            .row = iCount
            If .Text = seq_con Then
                .Col = 5
                If sSts <> "3" Then
                    Qa = Qa + CInt(.Text)
                    ZebraGridPedaco msflxgrdAcerto, iCount, iCount, vbRed, vbRed
                Else
                    ZebraGridPedaco msflxgrdAcerto, iCount, iCount, vbBlue, vbBlue
                End If
                msflxgrdAcerto.TopRow = iCount
            End If
        Next
        lblQA.caption = Qa
    End With
    msflxgrdAcerto.row = iRow
    
    
End Sub
Sub acha_linha_tr(seq_con As String)
    Dim iRow As Integer
    Dim sSEQ As String
    Dim iCount As Integer
    Dim iCol As Integer
    Dim Qt As Integer
    Dim Qa As Integer
    Dim Qd As Integer
    Dim cReman As Integer
    On Error Resume Next
    Qt = 0
    With msflxgrdCons
        .Col = 0
        For iCount = 1 To .rows - 1
            .row = iCount
            If .Text = seq_con Then
                .Col = 2
                Qt = Qt + CInt(.Text)
                .Col = 3
                cReman = cReman + CInt(.Text)
                ZebraGridPedaco msflxgrdCons, iCount, iCount, vbRed, vbRed
                
                msflxgrdCons.TopRow = iCount
            End If
        Next
        lblQt.caption = Qt
    End With
    
    
End Sub



Sub acha_linha_dev(seq_con As String)
    Dim iRow As Integer
    Dim sSEQ As String
    Dim iCount As Integer
    Dim iCol As Integer
    Dim Qt As Integer
    Dim Qa As Integer
    Dim Qd As Integer
    
    
    On Error Resume Next
    
    
    
    
    
    
    
    Qd = 0
    With msflxgrdDev
        .Col = 4
        For iCount = 1 To .rows - 1
            .row = iCount
            If .Text = seq_con Then
                .Col = 2
                Qd = Qd + CInt(.Text)
                ZebraGridPedaco msflxgrdDev, iCount, iCount, vbRed, vbRed
                
                msflxgrdAcerto.TopRow = iCount
            End If
        Next
        lblQd.caption = Qd
    End With
    
    
End Sub



Sub marca_Cons1(iGrid As Integer)
    Dim iRow As Integer
    Dim sSEQ As String
    Dim iCount As Integer
    Dim iCol As Integer
    Dim Qt As Integer
    Dim Qa As Integer
    Dim Qd As Integer
    Dim iColConsA As Integer
    Dim iColConsD As Integer
    Dim iColConsT As Integer
    
    If iGrid = 0 Then
    
    ElseIf iGrid = 0 Then
    
    Else
    
    End If
    
    On Error Resume Next
    
    
    
    
    
    iRow = msflxgrdCons.row
    msflxgrdCons.Col = 2
    Qt = CInt(msflxgrdCons.Text)
    
    msflxgrdCons.Col = 0
    
    sSEQ = msflxgrdCons.Text
    ZebraGrid msflxgrdCons, 1
    ZebraGrid msflxgrdDev, 1
    ZebraGrid msflxgrdAcerto, 1
    msflxgrdCons.row = iRow
    ZebraGridPedaco msflxgrdCons, iRow, iRow, vbRed, vbRed
    lblQt.caption = Qt
    
    Qd = 0
    With msflxgrdDev
        .Col = 4
        For iCount = 1 To .rows - 1
            .row = iCount
            If .Text = sSEQ Then
                .Col = 2
                Qd = Qd + CInt(.Text)
                ZebraGridPedaco msflxgrdDev, iCount, iCount, vbRed, vbRed
                
                msflxgrdAcerto.TopRow = iCount
            End If
        Next
        lblQd.caption = Qd
    End With
    Qa = 0
    With msflxgrdAcerto
        .Col = 7
        For iCount = 1 To .rows - 1
            .row = iCount
            If .Text = sSEQ Then
                .Col = 5
                Qa = Qa + CInt(.Text)
                ZebraGridPedaco msflxgrdAcerto, iCount, iCount, vbRed, vbRed
                msflxgrdAcerto.TopRow = iCount
            End If
        Next
        lblQA.caption = Qa
    End With
    msflxgrdCons.row = iRow
    
    msflxgrdCons.SetFocus
     
End Sub


Private Sub mnAce_Click()
    '
    
End Sub

Private Sub mnVisAce_Click()
    msflxgrdAcerto.Col = 0
    P_VisualizaAcerto msflxgrdAcerto.Text
End Sub

Private Sub msflxgrdAcerto_Click()
    Dim seq_con As String
    Dim iRow As Integer
    iRow = msflxgrdAcerto.row
    msflxgrdAcerto.Col = 7
    seq_con = msflxgrdAcerto.Text
    ZebraGrid msflxgrdAcerto, 1
    ZebraGrid msflxgrdCons, 1
    ZebraGrid msflxgrdDev, 1
    ZebraGridPedaco msflxgrdAcerto, iRow, iRow, vbRed
    msflxgrdAcerto.row = iRow
    
    acha_linha_dev seq_con
    acha_linha_tr seq_con
    acha_linha_acerto seq_con
    ZebraGridPedaco msflxgrdAcerto, iRow, iRow, vbRed, vbRed
End Sub

Private Sub msflxgrdAcerto_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnAce
    End If
End Sub

Private Sub msflxgrdCons_Click()
    marca_Cons
End Sub

Private Sub msflxgrdCons_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
        marca_Cons
    End If
End Sub


Private Sub msflxgrdDev_Click()
    Dim seq_con As String
    Dim iRow As Integer
    iRow = msflxgrdDev.row
    msflxgrdDev.Col = 4
    seq_con = msflxgrdDev.Text
    ZebraGrid msflxgrdAcerto, 1
    ZebraGrid msflxgrdCons, 1
    ZebraGrid msflxgrdDev, 1
    ZebraGridPedaco msflxgrdDev, iRow, iRow, vbRed, vbRed
    msflxgrdDev.row = iRow
    
    
    acha_linha_tr seq_con
    acha_linha_acerto seq_con
    acha_linha_dev seq_con
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaPro_Click
    End If
End Sub

Private Sub txtbxRepr_GotFocus()
    P_SelecTexto txtbxRepr
End Sub

Private Sub txtbxRepr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub


