VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelFat 
   Caption         =   "Resumo de Faturamento por Vendas"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   11115
   Icon            =   "frmRelFat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   11115
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5640
      ScaleHeight     =   825
      ScaleWidth      =   5265
      TabIndex        =   23
      Top             =   4200
      Width           =   5295
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias Trabalhados:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label lblTrab 
         AutoSize        =   -1  'True
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
         Left            =   1800
         TabIndex        =   26
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias Restantes:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label lblRest 
         AutoSize        =   -1  'True
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
         Left            =   1800
         TabIndex        =   24
         Top             =   480
         Width           =   120
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2145
      ScaleWidth      =   5385
      TabIndex        =   12
      Top             =   4200
      Width           =   5415
      Begin VB.Label lblMedDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Média Diária:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previsão Mensal:"
         Height          =   195
         Left            =   3000
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblPrev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   3000
         TabIndex        =   19
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lblVlrMaior 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maior movimentação:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1515
      End
      Begin VB.Label lblValorMenor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4800
         TabIndex        =   16
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menor Movimentação:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label lblDiaMaior 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maior Dia"
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
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblMenorDia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menor Dia"
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
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   885
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   5280
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   11085
      TabIndex        =   0
      Top             =   0
      Width           =   11115
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   690
         Left            =   5760
         Picture         =   "frmRelFat.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnFechar 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   690
         Left            =   9240
         Picture         =   "frmRelFat.frx":0871
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   690
         Left            =   4080
         Picture         =   "frmRelFat.frx":0ABB
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   690
         Left            =   2520
         Picture         =   "frmRelFat.frx":0DBD
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
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
   End
   Begin VB.Frame frControle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   5640
      TabIndex        =   5
      Top             =   5040
      Width           =   5295
      Begin VB.Label lblFaturado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1650
         TabIndex        =   31
         Top             =   120
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faturado"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1650
         TabIndex        =   28
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Devoluções"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblDev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1650
         TabIndex        =   6
         Top             =   480
         Width           =   570
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5318
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
      Caption         =   "Vendas"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "esp_cai"
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
      BeginProperty Column01 
         DataField       =   "descricao_cai"
         Caption         =   "Descrição"
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
         DataField       =   "total"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
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
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4350,047
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRelFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Sub buscaMov()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sOr As String
    
    strSQL = "select  sum(valor_ven) as total ,"
    strSQL = strSQL & " dataefet_ven as dat from " & TAB_VENDAS & " as tv "
    strSQL = strSQL & " Where "
    strSQL = strSQL & " tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.Text)
    strSQL = strSQL & " and situacao_ven=4 "
    strSQL = strSQL & " group by dataefet_ven"
    
    sOrd = " order by total desc limit 1 "
    abreObjetoAdo rPesq, strSQL & sOrd
    
    With rPesq
        lblDiaMaior.caption = CStr(!dat) & " " & WeekdayName(Weekday(CStr(!dat)))
        lblVlrMaior.caption = FormatVal(!total)
        .Close
    End With
    sOrd = " order by total asc limit 1 "
    abreObjetoAdo rPesq, strSQL & sOrd
    
    With rPesq
        lblMenorDia.caption = !dat & " " & WeekdayName(Weekday(!dat))
        lblValorMenor.caption = FormatVal(!total)
        .Close
    End With
End Sub
Sub buscaDias()
    Dim sAux As String
    Dim sMes As String
    Dim strSQL As String
    Dim sDia As String
    strSQL = F_BuscaUmCampo("count(distinct dataefet_ven)", "" & TAB_VENDAS & "", _
        "dataefet_ven>=" & F_GravaData(txtbxDataIni.Text) & " and " & _
        "dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.Text))
    lblTrab.caption = strSQL
    
    
    sAux = DateDiff("D", Date, txtbxDAtaFIn.Text)
    If CCur(sAux) < 0 Then sAux = "0"
    lblRest.caption = sAux
    
    
    
    sAux = Int(sAux - (Int(sAux) / 7))
    lblRest.caption = sAux
    
    
    
    
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV grid, ado
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_erro
    buscaDias
    strSQL = " select sum(tf.valor_ven) as total,tf.esp_cai,descricao_cai  "
    
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv inner join " & TAB_VENFEC & "  as tf "
    strSQL = strSQL & " on tv.num_ven=tf.num_ven "
    strSQL = strSQL & " inner join tab_espcai as te on te.sigla_cai=tf.esp_cai "
    strSQL = strSQL & " where 1>0 "
    
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.Text)
    strSQL = strSQL & " and tv.situacao_ven=4  "
    strSQL = strSQL & " group by tf.esp_cai "
    abreObjetoAdo ado, strSQL, grid
    
    buscaMov

    strSQL = "select sum(valor_dev) as total from " & TAB_DEVOLUCOES
    strSQL = strSQL & " where left(data_dev,10)>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and left(data_Dev,10)<=" & F_GravaData(txtbxDAtaFIn.Text)
    abreObjetoAdo rPesq, strSQL
     soma
    With rPesq
            lblDev.caption = FormatVal(!total)
            If lblDev.caption = "" Then lblDev.caption = "0,00"
            lblTotal.caption = FormatVal(CCur(lblFaturado.caption) - CCur(lblDev.caption))
        .Close
    End With
    
   
    
    lblMedDia.caption = FormatVal(CCur(lblTotal.caption / CCur(lblTrab.caption)))
    lblPrev.caption = FormatVal(CCur(lblMedDia.caption) * (CCur(lblRest.caption) + CCur(lblTrab.caption)))
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Sub soma()
    With ado
        Dim cTotal As Currency
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !total
            .MoveNext
        Loop
        lblFaturado.caption = FormatVal(cTotal)
    End With
End Sub

Private Sub Command1_Click()

End Sub



Private Sub cmndbtnImprimir_Click()
    imprime frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprime Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub
Sub imprime(obj As Object)
    P_CabRelTimbr "Resumo de Faturamento de " & txtbxDataIni.Text & " até " & txtbxDAtaFIn.Text, obj
    P_ImprLinha obj
    obj.fontsize = 14
    P_ImprX obj, "Venda à Vista............................: ", 10
    P_ImprX obj, lblVendaAvista.caption, 160, False, True
    P_ImprX obj, "Venda a Prazo...........................: ", 10
    P_ImprX obj, lblVendaAp.caption, 160, False, True
    P_ImprX obj, "Venda por Financeiras..............: ", 10
    P_ImprX obj, lblVendaFin.caption, 160, False, True
    P_ImprLinha obj, "_"
    P_ImprX obj, "Total............................................: ", 10
    P_ImprX obj, lblTotal.caption, 160, False, True
    obj.fontsize = 10
    P_RodapeVis obj, strNomeSis & " " & constVersao_Sis
End Sub
Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub
