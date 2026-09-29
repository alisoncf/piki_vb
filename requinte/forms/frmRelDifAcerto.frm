VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelDifAcerto 
   Caption         =   "Relatório de Diferenças de Acertos"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid grid 
      Height          =   6495
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   11456
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "nome_rep"
         Caption         =   "Representante"
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
         DataField       =   "nome_cli"
         Caption         =   "Cliente"
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
         DataField       =   "num_doc"
         Caption         =   "Carnê"
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
         DataField       =   "dif"
         Caption         =   "Diferença"
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
      BeginProperty Column04 
         DataField       =   "reman"
         Caption         =   "Reman"
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
            ColumnWidth     =   4155,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4020,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1110,047
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   13065
      TabIndex        =   12
      Top             =   7800
      Width           =   13095
      Begin VB.CheckBox chkNeg 
         Caption         =   "Considerar negativas"
         Height          =   255
         Left            =   10680
         TabIndex        =   16
         Top             =   120
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label lblVendas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   720
         TabIndex        =   14
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   6480
         Picture         =   "frmRelDifAcerto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   8160
         Picture         =   "frmRelDifAcerto.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   11520
         Picture         =   "frmRelDifAcerto.frx":0797
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico"
         Height          =   615
         Left            =   9840
         Picture         =   "frmRelDifAcerto.frx":088C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   285
         Left            =   5880
         Picture         =   "frmRelDifAcerto.frx":09D6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRelDifAcerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub chkNeg_Click()
    busca
End Sub

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select nome_rep,nome_cli, td.num_doc,td.seq_doc,sum(ti.reman_itd) as reman, valor_doc,valor_ace,if(diferenca_ace>0,(diferenca_ace*1),diferenca_ace) as dif,data_ace "
    strSQL = strSQL & " from tab_acertos as ta inner join tab_doc as td on ta.seq_ace=td.num_doc and td.esp_doc='A'"
    strSQL = strSQL & " inner join tab_itdoc as ti on td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " inner join tab_clientes as tc on tc.cod_cli=ta.cod_cli "
    strSQL = strSQL & " inner join tab_representantes as tr on tr.cod_rep=ta.cod_rep "
    strSQL = strSQL & " where status_ace='2' "
    
    If chkNeg.Value = 0 Then
        strSQL = strSQL & " and diferenca_ace>0 "
    End If
    
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and vencini_itd<=" & F_GravaData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and vencini_itd>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxRepresentante.tag <> "" Then
        strSQL = strSQL & " and ta.cod_rep='" & txtbxRepresentante.tag & "' "
    End If
    strSQL = strSQL & " and status_ace=2  and diferenca_ace <> 0 "
    
     
    strSQL = strSQL & " group by ta.seq_ace "
    strSQL = strSQL & " having reman>0"
    
    abreObjetoAdo ado, strSQL, grid
    
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    On Error GoTo trata_Erro:
    grid.Visible = False
    With ado
        Do While Not .EOF
            'cTotal = cTotal + !diferenca_ace
            cTotal = cTotal + !dif
            
            .MoveNext
        Loop
        lblVendas.caption = FormatVal(cTotal)
        .MoveFirst
        grid.Visible = True
    End With
trata_Erro:
End Sub

Private Sub cmndbtnBuscaRep_Click()
    buscaRep
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Relatório de Deferenças de Acertos de " & txtbxDataIni.Text & " até " & txtbxdataFin.Text, grid, ado, txtbxRepresentante.Text, "4"
    
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
End Sub

Private Sub Text1_Change()

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

Private Sub txtbxRepresentante_GotFocus()
    P_SelecTexto txtbxRepresentante
End Sub

Private Sub txtbxRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaRep
    End If
End Sub

Sub buscaRep()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxRepresentante.Text)
    txtbxRepresentante.Text = r.nome_rep
    txtbxRepresentante.tag = r.cod_Rep
    busca
End Sub

