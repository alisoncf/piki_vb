VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelFiscal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Acompanhamento"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   11145
      TabIndex        =   13
      Top             =   6720
      Width           =   11175
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS"
         Height          =   195
         Left            =   2400
         TabIndex        =   18
         Top             =   120
         Width           =   390
      End
      Begin VB.Label lblIcms 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   2400
         TabIndex        =   17
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblCmp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblCompras 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compras"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   2520
      TabIndex        =   5
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   685
         Left            =   120
         Picture         =   "frmRelFiscal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   685
         Left            =   2040
         Picture         =   "frmRelFiscal.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnFechar 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   685
         Left            =   6840
         Picture         =   "frmRelFiscal.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D&e:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label6 
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
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   5175
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9128
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrdVendas 
      Height          =   5175
      Left            =   5640
      TabIndex        =   11
      Top             =   1440
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9128
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Giro"
      Height          =   195
      Left            =   5640
      TabIndex        =   12
      Top             =   1200
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Faturamento"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   885
   End
End
Attribute VB_Name = "frmRelFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    buscaECF
    buscaGiro
    buscaCompras
End Sub
Sub buscaCompras()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select sum(valor_not) as total,sum(icms_not) as icms from tab_compras as tc "
    strSQL = strSQL & " where  status_not>1 "
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and emissao_not>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and emissao_not<=" & F_GravaData(txtbxdataFin.Text)
    End If
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            lblCompras.caption = FormatVal(!total)
            lblIcms.caption = FormatVal(!icms)
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select sum(valor_not) as total from tab_cmp as tc "
    strSQL = strSQL & " where  status_not>1 "
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and emissao_not>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and emissao_not<=" & F_GravaData(txtbxdataFin.Text)
    End If
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            lblCmp.caption = FormatVal(!total)
            
            .MoveNext
        Loop
        .Close
    End With
End Sub
Sub buscaECF()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCodCai As String
    Dim cTotal As Currency
    Dim cTotalAtual As Currency
    Dim cIcms As Currency
    Dim iQtd As Integer
    Dim cTotalAtual1 As Currency
    Dim cIcms1 As Currency
    Dim iQtd1 As Integer
    Dim cTotalAtual2 As Currency
    Dim cIcms2 As Currency
    Dim iQtd2 As Integer
    
    msflxgrd.rows = 1
    If txtbxDataIni.Text = "" Then
        MsgBox "Informe a data inicial!", vbExclamation
        Exit Sub
    End If
    strSQL = "select cod_cai,count(tc.seq_cup) as qtd, sum(tc.valor_cup) as valor, sum(icms_cup) as icms "
    strSQL = strSQL & " from cupons as tc "
    strSQL = strSQL & " where cancelado_cup='N' "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and data_cup<=" & F_GravaDataTempo(txtbxdataFin.Text & "23:59")
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_cup>=" & F_GravaDataTempo(txtbxDataIni.Text)
    End If
    strSQL = strSQL & " group by cod_cai order by cod_cai "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            If sCodCai <> !cod_cai Then
                
            End If
            strSQL = !cod_cai & Chr(9)
            strSQL = strSQL & FormatVal(!valor) & Chr(9)
            strSQL = strSQL & FormatVal(!icms) & Chr(9)
            strSQL = strSQL & (!qtd)
            cTotal = cTotal + !valor
            cIcms = cIcms + !icms
            iQtd = iQtd + !qtd
            msflxgrd.AddItem strSQL
            
            .MoveNext
        Loop
        .Close
        msflxgrd.AddItem "Total ECF" & Chr(9) & FormatVal(cTotal) & Chr(9) & FormatVal(cIcms) & Chr(9) & iQtd
    End With
    strSQL = "select count(ts.seq_sai) as qtd, sum(ts.valor_sai) as valor, sum(vicms_sai) as icms "
    strSQL = strSQL & " from tab_saidasnf as ts "
    strSQL = strSQL & " where situacao_sai<>'5' "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and emissao_sai<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and emissao_sai>=" & F_GravaData(txtbxDataIni.Text)
    End If
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            ctotal1 = !valor
            cIcms1 = !icms
            iQtd1 = !qtd
            .MoveNext
        Loop
        .Close
    
    msflxgrd.AddItem "SEPD Total" & Chr(9) & FormatVal(ctotal1) & Chr(9) & FormatVal(cIcms1) & Chr(9) & iQtd1
    End With
    strSQL = "select count(ts.seq_sai) as qtd, sum(ts.valor_sai) as valor, sum(vicms_sai) as icms "
    strSQL = strSQL & " from tab_saidasnf as ts,cupons as tc  "
    strSQL = strSQL & " where tc.num_ven=ts.num_ven "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and emissao_sai<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and emissao_sai>=" & F_GravaData(txtbxDataIni.Text)
    End If
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If IsNull(!valor) Then
                ctotal2 = 0
            Else
                ctotal2 = !valor
            End If
            If IsNull(!icms) Then
                cIcms2 = 0
            Else
                cIcms2 = !icms
            End If
            If IsNull(!qtd) Then
                iQtd2 = 0
            Else
                iQtd2 = !qtd
            End If
            .MoveNext
        Loop
        .Close
        msflxgrd.AddItem "SEPD Sem Cupom " & Chr(9) & FormatVal(ctotal1 - ctotal2) & Chr(9) & FormatVal(cIcms1 - cIcms2) & Chr(9) & iQtd1 - iQtd2
    End With
    
    
    msflxgrd.AddItem "TOTAL GERAL " & Chr(9) & FormatVal(cTotal + ctotal1 - ctotal2) & Chr(9) & FormatVal(cIcms + cIcms1 - cIcms2) & Chr(9) & iQtd + iQtd1 - iQtd2
    ZebraGrid msflxgrd, 1
End Sub
Sub buscaGiro()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCodCai As String
    Dim cTotal As Currency
    Dim cTotalAtual As Currency
    Dim cIcms As Currency
    Dim iQtd As Integer
    Dim cTotalAtual1 As Currency
    Dim cIcms1 As Currency
    Dim iQtd1 As Integer
    Dim cTotalAtual2 As Currency
    Dim cIcms2 As Currency
    Dim iQtd2 As Integer
    
    msflxgrdVendas.rows = 1
    
    strSQL = "select cod_cai,count(tv.num_ven) as qtd, sum(tv.valor_ven) as valor, sum(desconto_ven) as desconto "
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv "
    strSQL = strSQL & " where situacao_ven=4 "
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    End If
    strSQL = strSQL & " group by cod_cai order by cod_cai "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            If sCodCai <> !cod_cai Then
                
            End If
            strSQL = !cod_cai & Chr(9)
            strSQL = strSQL & FormatVal(!valor) & Chr(9)
            strSQL = strSQL & FormatVal(!Desconto) & Chr(9)
            strSQL = strSQL & (!qtd)
            cTotal = cTotal + !valor
            cIcms = cIcms + !Desconto
            iQtd = iQtd + !qtd
            msflxgrdVendas.AddItem strSQL
            
            .MoveNext
        Loop
        .Close
        msflxgrdVendas.AddItem "Total Vendas" & Chr(9) & FormatVal(cTotal) & Chr(9) & FormatVal(cIcms) & Chr(9) & iQtd
    End With
    ZebraGrid msflxgrdVendas, 1
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ajustaGrid
End Sub
Sub ajustaGrid()
    
    msflxgrd.FixedCols = 0
    msflxgrd.rows = 1
    msflxgrd.Cols = 4
    msflxgrd.col = 0: msflxgrd.Text = "Caixa"
    msflxgrd.col = 1: msflxgrd.Text = "Faturado"
    msflxgrd.col = 2: msflxgrd.Text = "Icms"
    msflxgrd.col = 3: msflxgrd.Text = "Qtd"
    
    msflxgrd.ColWidth(0) = 1600
    msflxgrd.ColWidth(3) = 800
    msflxgrd.ColAlignment(0) = 1
    
    With msflxgrdVendas
        .FixedCols = 0
        .rows = 1
        .Cols = 4
        .col = 0: .Text = "Caixa"
        .col = 1: .Text = "Faturado"
        .col = 2: .Text = "desconto"
        .col = 3: .Text = "Qtd"
        
        .ColWidth(0) = 1600
        .ColWidth(3) = 800
        .ColAlignment(0) = 1
    End With
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub
