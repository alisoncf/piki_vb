VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistProdAndamento 
   Caption         =   "Andamento do Produto"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16290
   Icon            =   "frmHistProdAndamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   16290
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TabStrip tbst 
      Height          =   375
      Left            =   0
      TabIndex        =   38
      Top             =   840
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Consignado"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimento Interno"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcMovCons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   7200
      ScaleHeight     =   6705
      ScaleWidth      =   8385
      TabIndex        =   25
      Top             =   1200
      Width           =   8415
      Begin VB.PictureBox pcAcerto 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3825
         ScaleWidth      =   13065
         TabIndex        =   34
         Top             =   4200
         Width           =   13095
         Begin MSFlexGridLib.MSFlexGrid msflxgrdAcerto 
            Height          =   3495
            Left            =   0
            TabIndex        =   35
            Top             =   240
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   6165
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label lblAcerto 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
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
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   0
            Width           =   660
         End
         Begin VB.Label lblQA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   1080
            TabIndex        =   36
            Top             =   0
            Width           =   90
         End
      End
      Begin VB.PictureBox pcDev 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   3600
         ScaleHeight     =   4065
         ScaleWidth      =   3345
         TabIndex        =   30
         Top             =   0
         Width           =   3375
         Begin MSFlexGridLib.MSFlexGrid msflxgrdDev 
            Height          =   3735
            Left            =   0
            TabIndex        =   31
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   6588
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label lblDev 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
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
            Left            =   120
            TabIndex        =   33
            Top             =   0
            Width           =   1020
         End
         Begin VB.Label lblQd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   1920
            TabIndex        =   32
            Top             =   0
            Width           =   90
         End
      End
      Begin VB.PictureBox pcCons 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   120
         ScaleHeight     =   3945
         ScaleWidth      =   3345
         TabIndex        =   26
         Top             =   0
         Width           =   3375
         Begin MSFlexGridLib.MSFlexGrid msflxgrdCons 
            Height          =   3735
            Left            =   0
            TabIndex        =   27
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   6588
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label lblQt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2400
            TabIndex        =   29
            Top             =   0
            Width           =   90
         End
         Begin VB.Label lblCons 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   28
            Top             =   0
            Width           =   1275
         End
      End
   End
   Begin VB.PictureBox pcMovInt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      ScaleHeight     =   6705
      ScaleWidth      =   6585
      TabIndex        =   12
      Top             =   1200
      Width           =   6615
      Begin VB.PictureBox pcCont 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   4320
         ScaleHeight     =   4185
         ScaleWidth      =   2145
         TabIndex        =   22
         Top             =   0
         Width           =   2175
         Begin MSFlexGridLib.MSFlexGrid msflxgrdCont 
            Height          =   3495
            Left            =   0
            TabIndex        =   23
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   6165
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label lblCont 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Contagens"
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
            Left            =   120
            TabIndex        =   24
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.PictureBox pcVendas 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   0
         ScaleHeight     =   3705
         ScaleWidth      =   6825
         TabIndex        =   19
         Top             =   4320
         Width           =   6855
         Begin MSFlexGridLib.MSFlexGrid msflxgrdVendas 
            Height          =   3495
            Left            =   0
            TabIndex        =   20
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   6165
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label lblVendas 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas"
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
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.PictureBox pcTrInt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   0
         ScaleHeight     =   4185
         ScaleWidth      =   1665
         TabIndex        =   16
         Top             =   0
         Width           =   1695
         Begin MSFlexGridLib.MSFlexGrid msflxgrdTrInt 
            Height          =   3735
            Left            =   0
            TabIndex        =   17
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6588
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label lblTrInt 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transferências Internas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.PictureBox pcCompras 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   2160
         ScaleHeight     =   4185
         ScaleWidth      =   1785
         TabIndex        =   13
         Top             =   0
         Width           =   1815
         Begin MSFlexGridLib.MSFlexGrid msflxgrdCompras 
            Height          =   3735
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   6588
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label lblCompras 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000040C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Compras"
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
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   16260
      TabIndex        =   11
      Top             =   0
      Width           =   16290
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnAbre 
         Caption         =   "Abrir"
         Height          =   615
         Left            =   9360
         Picture         =   "frmHistProdAndamento.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   10440
         Picture         =   "frmHistProdAndamento.frx":6C86
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   11520
         Picture         =   "frmHistProdAndamento.frx":7093
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Buscar"
         Height          =   615
         Left            =   8280
         Picture         =   "frmHistProdAndamento.frx":71FE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   975
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5535
      End
      Begin VB.CommandButton cmndbtnBuscaPro 
         Height          =   315
         Left            =   5640
         Picture         =   "frmHistProdAndamento.frx":7588
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   7080
         TabIndex        =   5
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   6000
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmHistProdAndamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmndbtnBuscaPro_Click()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxProduto.text)
    txtbxProduto.text = p.nome_pro
    txtbxProduto.tag = p.cod_pro
    busca
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub

Private Sub cmndbtnCliente_Click()
    Dim r As representante.representante
    
    r = representante.F_BuscaRep(txtbxRepr.text)
    If r.cod_Rep <> "" Then
        txtbxRepr.text = r.nome_rep
        txtbxRepr.tag = r.cod_Rep
        busca
        
    End If
    
End Sub
Sub AjustaGrid()
    With msflxgrdCons
        .rows = 1
        .Cols = 7
        .col = 0: .text = "Nº Cons."
        .col = 1: .text = "Data"
        .col = 2: .text = "Qtd"
        .col = 3: .text = "Reman"
        .col = 4: .text = "Usuário"
        .col = 5: .text = "Retif"
        .col = 6: .text = "Representante"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 600
    End With
    With msflxgrdDev
        .rows = 1
        .Cols = 6
        .col = 0: .text = "Nº Dev."
        .col = 1: .text = "Data Dev."
        .col = 2: .text = "Qtd Dev."
        .col = 3: .text = "Usuário"
        .col = 4: .text = "Nº Cons."
        .col = 5: .text = "Representante."
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
    End With
    With msflxgrdAcerto
        .rows = 1
        .Cols = 9
        .col = 0: .text = "Nº Ace."
        .col = 1: .text = "Data Ace."
        .col = 2: .text = "Hora Ace."
        .col = 3: .text = "Usuário"
        .col = 4: .text = "Cliente."
        .col = 5: .text = "Qtd"
        .col = 6: .text = "Valor"
        .col = 7: .text = "Nº Cons"
        .col = 8: .text = "Representante"
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 600
        .ColWidth(3) = 800
        .ColWidth(4) = 3500
        .ColWidth(5) = 800
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
    End With
    With msflxgrdTrInt
        .rows = 1
        .Cols = 6
        .col = 0: .text = "NºTr."
        .col = 1: .text = "Data"
        .col = 2: .text = "Usuário"
        .col = 3: .text = "Destino"
        .col = 4: .text = "Qtd"
        .col = 5: .text = "Valor"
        
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 600
        .ColWidth(3) = 2800
        .ColWidth(4) = 1000
        
    End With
    
    With msflxgrdCompras
        .rows = 1
        .Cols = 6
        .col = 0: .text = "Seq."
        .col = 1: .text = "Nota"
        .col = 2: .text = "Fornecedor"
        .col = 3: .text = "Emissao"
        .col = 4: .text = "Qtd"
        .col = 5: .text = "Valor"
        
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 3600
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        
    End With
        
    With msflxgrdVendas
        .rows = 1
        .Cols = 8
        .col = 0: .text = "NºVenda"
        .col = 1: .text = "Data Ab."
        .col = 2: .text = "Data Efet."
        .col = 3: .text = "Usuário"
        .col = 4: .text = "Cliente."
        .col = 5: .text = "Qtd"
        .col = 6: .text = "Valor"
        .col = 7: .text = "Forma Pag."
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 1000
        .ColWidth(3) = 1500
        .ColWidth(4) = 3500
        .ColWidth(5) = 800
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
    End With
    
    
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    AjustaGrid
    strSQL = "select tc.*,tp.nome_pro from tab_consignacoes as tc,"
    strSQL = strSQL & " tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=tc.cod_pro "
    'strSQL = strSQL & " and tc.cod_rep='" & txtbxRepr.tag & "' "
    strSQL = strSQL & " and tc.cod_pro='" & txtbxProduto.tag & "' "
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and data_con>=" & F_GravaDataTempo(txtbxDataIni.text)
    End If
    If txtbxdataFin.text <> "" Then
        strSQL = strSQL & " and data_con>=" & F_GravaDataTempo(txtbxdataFin.text & " 23:59:00")
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
            strSQL = strSQL & !retif_con & Chr(9) & !cod_Rep
            msflxgrdCons.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    strSQL = "select ti.*,cod_rep,td.cod_usu,td.data_dev from tab_devconsign as td,tab_itdevconsign as ti"
    strSQL = strSQL & " where td.seq_dev=ti.seq_dev "
    'strSQL = strSQL & " and td.cod_rep='" & txtbxRepr.tag & "' "
    strSQL = strSQL & " and ti.cod_pro='" & txtbxProduto.tag & "' "
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and data_dev>=" & F_GravaDataTempo(txtbxDataIni.text)
    End If
    If txtbxdataFin.text <> "" Then
        strSQL = strSQL & " and data_dev>=" & F_GravaDataTempo(txtbxdataFin.text & " 23:59:00")
    End If
    
    strSQL = strSQL & " order by data_dev "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = !seq_dev & Chr(9)
            strSQL = strSQL & !data_dev & Chr(9)
            strSQL = strSQL & !qtd_dev & Chr(9)
            strSQL = strSQL & !cod_usu & Chr(9)
            strSQL = strSQL & !seq_con & Chr(9) & !cod_Rep
            msflxgrdDev.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    strSQL = "select ti.qtd_ite,ti.unit_ite,ti.seq_con,ta.* "
    strSQL = strSQL & " from tab_acertos as ta,tab_itacertos as ti "
    strSQL = strSQL & " where ta.seq_ace=ti.seq_ace "
    'strSQL = strSQL & " and ta.cod_rep='" & txtbxRepr.tag & "' "
    strSQL = strSQL & " and ti.cod_pro='" & txtbxProduto.tag & "' "
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and data_Ace>=" & F_GravaData(txtbxDataIni.text)
    End If
    If txtbxdataFin.text <> "" Then
        strSQL = strSQL & " and data_Ace>=" & F_GravaData(txtbxdataFin.text)
    End If
    
    strSQL = strSQL & " order by data_ace "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = !seq_ace & Chr(9)
            strSQL = strSQL & !data_ace & Chr(9)
            strSQL = strSQL & !hora_ace & Chr(9)
            strSQL = strSQL & !cod_usu & Chr(9)
            If !tipo_ter = "C" Then
                strSQL = strSQL & "O Prório Representante" & Chr(9)
            Else
                strSQL = strSQL & F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_cli='" & !cod_cli & "'") & Chr(9)
            End If
            strSQL = strSQL & !qtd_ite & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9)
            strSQL = strSQL & !seq_con & Chr(9) & !cod_Rep
            msflxgrdAcerto.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select tf.nome_for,ti.qtd_ite,ti.unit_ite,ti.custo_ite,ti.custosai_ite,tf.nome_for,tc.seq_not,tc.emissao_not "
    strSQL = strSQL & " from tab_compras as tc,tab_itcompras as ti ,tab_fornecedores as tf "
    strSQL = strSQL & " where tc.seq_not=ti.seq_not "
    'strSQL = strSQL & " and ta.cod_rep='" & txtbxRepr.tag & "' "
    strSQL = strSQL & " and ti.cod_pro='" & txtbxProduto.tag & "' "
    strSQL = strSQL & " and tc.cod_for=tf.cod_for "
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and emissao_not>=" & F_GravaData(txtbxDataIni.text)
    End If
    If txtbxdataFin.text <> "" Then
        strSQL = strSQL & " and emissao_not >=" & F_GravaData(txtbxdataFin.text)
    End If
    
    strSQL = strSQL & " order by emissao_not  "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = !seq_not & Chr(9)
            strSQL = !nome_for & Chr(9)
            strSQL = strSQL & !emissao_not & Chr(9)
            strSQL = strSQL & !qtd_ite & Chr(9)
            strSQL = strSQL & !unit_ite
            msflxgrdCompras.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    strSQL = "select tv.*,ti.qtd_ite,ti.unit_ite,ti.final_ite,tf.guerra_fun "
    strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti,tab_funcionarios as tf "
    strSQL = strSQL & " where tv.num_ven=ti.num_ven "
    
    strSQL = strSQL & " and ti.cod_pro='" & txtbxProduto.tag & "' "
    strSQL = strSQL & " and tf.cod_fun=tv.abrevenda_usu "
    
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.text)
    End If
    If txtbxdataFin.text <> "" Then
        strSQL = strSQL & " and dataefet_ven >=" & F_GravaData(txtbxdataFin.text)
    End If
    strSQL = strSQL & " and situacao_ven=4 "
    
    strSQL = strSQL & " order by dataefet_ven   "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = !num_ven & Chr(9)
            strSQL = strSQL & !data_ven & Chr(9)
            strSQL = strSQL & !dataefet_ven & Chr(9)
            strSQL = strSQL & !guerra_fun & Chr(9)
            strSQL = strSQL & !nomecli_ven & Chr(9)
            strSQL = strSQL & !qtd_ite & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9)
            strSQL = strSQL & !cod_For
            msflxgrdVendas.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    
    
    
End Sub


Private Sub cmndbtnImpr_Click()
    msflxgrdAcerto.col = 0
    imprAcertoData txtbxRepr.tag, msflxgrdAcerto.text
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If txtbxProduto.tag <> "" Then
        cmndbtnBuscaPro_Click
    End If
    bActive = True
End Sub


Private Sub Form_Load()
    'adc.connectionstring = conn.connectionstring
    'sstb.Tab = 0
    resizeForm Me
    
End Sub

Private Sub Form_Resize()
    'On Error Resume Next
    If tbst.SelectedItem.Index = 1 Then
        ajusta_cons
    Else
        ajusta_externo
    End If
   ''
    
    
    
    
End Sub
Sub ajusta_cons()
    Dim iAux As Single
    On Error Resume Next
    pcMovInt.Visible = False
    pcMovCons.Visible = True
    pcMovCons.Left = tbst.Left
    pcMovCons.width = Me.width - pcMovCons.Left - 200
    pcMovCons.height = Me.height - pcMovCons.top - 400
    pcCons.width = (pcMovCons.width / 2) - (pcCons.Left / 2) - 100
    pcDev.width = pcCons.width
    pcDev.Left = pcCons.Left + pcCons.width + 50
    
    msflxgrdCons.width = pcCons.width
    msflxgrdDev.width = msflxgrdCons.width
    msflxgrdCont.width = msflxgrdDev.width
    pcCons.height = (pcMovCons.height / 2) - (pcMovCons.top / 2) - (pcCons.top / 2)
    pcDev.height = pcCons.height
    
    msflxgrdCons.height = pcCons.height - msflxgrdCons.top - 50
    msflxgrdDev.height = msflxgrdCons.height
    
    pcAcerto.Left = pcCons.Left
    pcAcerto.top = pcCons.height + pcCons.top + 50
    pcAcerto.width = (pcMovCons.width / 1) - (pcAcerto.Left / 1) - 100
    pcAcerto.height = pcMovCons.height - pcAcerto.top - 200
    msflxgrdAcerto.width = pcAcerto.width - msflxgrdAcerto.Left
    msflxgrdAcerto.height = pcAcerto.height - msflxgrdAcerto.top
End Sub
Sub ajusta_externo()
    Dim iAux As Single
    On Error Resume Next
    pcMovInt.Visible = True
    pcMovCons.Visible = False
    pcMovInt.width = Me.width - pcMovInt.Left - 200
    pcMovInt.height = Me.height - pcMovInt.top - 400
    pcTrInt.width = (pcMovInt.width / 3) - (pcTrInt.Left / 3) - 100
    pcCompras.width = pcTrInt.width
    pcCompras.Left = pcTrInt.Left + pcTrInt.width + 50
    pcCont.width = pcTrInt.width
    pcCont.Left = pcCompras.Left + pcCompras.width + 50
    msflxgrdTrInt.width = pcTrInt.width
    msflxgrdCompras.width = msflxgrdTrInt.width
    msflxgrdCont.width = msflxgrdCompras.width
    pcTrInt.height = (pcMovInt.height / 2) - (pcMovInt.top / 2) - (pcTrInt.top / 2)
    pcCompras.height = pcTrInt.height
    pcCont.height = pcTrInt.height
    msflxgrdTrInt.height = pcTrInt.height - msflxgrdTrInt.top - 50
    msflxgrdCompras.height = msflxgrdTrInt.height
    msflxgrdCont.height = msflxgrdTrInt.height
    pcVendas.Left = pcTrInt.Left
    pcVendas.top = pcTrInt.height + pcTrInt.top + 50
    pcVendas.width = (pcMovInt.width / 1) - (pcVendas.Left / 1) - 100
    pcVendas.height = pcMovInt.height - pcVendas.top - 200
    msflxgrdVendas.width = pcVendas.width - msflxgrdVendas.Left
    msflxgrdVendas.height = pcVendas.height - msflxgrdVendas.top
End Sub
Private Sub Form_Unload(Cancel As Integer)
    bActive = False
End Sub

Sub marca_Cons()
    Dim iRow As Integer
    Dim sSeq As String
    Dim iCount As Integer
    Dim iCol As Integer
    Dim Qt As Integer
    Dim Qa As Integer
    Dim Qd As Integer
    
    
    iRow = msflxgrdCons.row
    msflxgrdCons.col = 2
    If iRow > 0 Then
        Qt = CInt(msflxgrdCons.text)
    End If
    msflxgrdCons.col = 0
    
    sSeq = msflxgrdCons.text
    ZebraGrid msflxgrdCons, 1
    ZebraGrid msflxgrdDev, 1
    ZebraGrid msflxgrdAcerto, 1
    msflxgrdCons.row = iRow
    ZebraGridPedaco msflxgrdCons, iRow, iRow, vbRed, vbRed
    lblQt.caption = Qt
    
    Qd = 0
    With msflxgrdDev
        .col = 4
        For iCount = 1 To .rows - 1
            .row = iCount
            If .text = sSeq Then
                .col = 2
                Qd = Qd + CInt(.text)
                ZebraGridPedaco msflxgrdDev, iCount, iCount, vbRed, vbRed
            End If
        Next
        lblQd.caption = Qd
    End With
    Qa = 0
    With msflxgrdAcerto
        .col = 7
        For iCount = 1 To .rows - 1
            .row = iCount
            If .text = sSeq Then
                .col = 5
                Qa = Qa + CInt(.text)
                ZebraGridPedaco msflxgrdAcerto, iCount, iCount, vbRed, vbRed
            End If
        Next
        lblQA.caption = Qa
    End With
    msflxgrdCons.row = iRow
    
    msflxgrdCons.SetFocus
     
End Sub

Sub marca_Cons_old()
    Dim iRow As Integer
    Dim sSeq As String
    Dim iCount As Integer
    Dim iCol As Integer
    iRow = msflxgrdCons.row
    msflxgrdCons.col = 0
    sSeq = msflxgrdCons.text
    ZebraGrid msflxgrdCons, 1
    ZebraGrid msflxgrdDev, 1
    ZebraGrid msflxgrdAcerto, 1
    msflxgrdCons.row = iRow
    ZebraGridPedaco msflxgrdCons, iRow, iRow, vbRed, vbRed
    
    With msflxgrdDev
        .col = 4
        For iCount = 1 To .rows - 1
            .row = iCount
            If .text = sSeq Then
                ZebraGridPedaco msflxgrdDev, iCount, iCount, vbRed, vbRed
            End If
        Next
        
    End With
    
    With msflxgrdAcerto
        .col = 7
        For iCount = 1 To .rows - 1
            .row = iCount
            If .text = sSeq Then
                ZebraGridPedaco msflxgrdDev, iCount, iCount, vbRed, vbRed
            End If
        Next
        
    End With
    msflxgrdCons.row = iRow
    
    msflxgrdCons.SetFocus
     
End Sub


Private Sub msflxgrdCons_Click()
    marca_Cons
End Sub

Private Sub msflxgrdCons_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
        marca_Cons
    End If
End Sub

Private Sub sstb_DblClick()
    Form_Resize
End Sub

Private Sub tbst_Click()
    Form_Resize
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
    
End Sub

Private Sub txtbxRepr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub


