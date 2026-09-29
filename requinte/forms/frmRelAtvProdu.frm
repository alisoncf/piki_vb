VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelAtvProdu 
   Caption         =   "Relatório de Atividades"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid grid 
      Height          =   7695
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   13573
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
         DataField       =   "cod_pro"
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
         DataField       =   "nome_pro"
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
         DataField       =   "nome_sub"
         Caption         =   "Subseção"
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
         DataField       =   "qv"
         Caption         =   "Qtd. Vendida"
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
         DataField       =   "tv"
         Caption         =   "Total Vendido"
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
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5520,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2624,882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1200,189
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12975
      TabIndex        =   0
      Top             =   0
      Width           =   13005
      Begin VB.TextBox txtLimite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optnAce 
         Caption         =   "Acertos"
         Height          =   255
         Left            =   4800
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optnVendas 
         Caption         =   "Vendas"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   7560
         Picture         =   "frmRelAtvProdu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   9000
         Picture         =   "frmRelAtvProdu.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   11400
         Picture         =   "frmRelAtvProdu.frx":0797
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite:"
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmRelAtvProdu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim tabv As String
    Dim tabi As String
    Dim stat As String
    Dim opt As String
    If optnVendas.Value Then
        tabv = "tab_vendas"
        tabi = "tab_itvendas"
        stat = "situacao_ven = '4'"
        opt = "on tv.num_ven=tiv.nuM_ven "
    Else
        tabv = "tab_acertos"
        tabi = "tab_itace"
        stat = "status_ace = '2'"
        opt = "on tv.seq_ace=tiv.seq_ace "
    End If
    
    strSQL = "select tp.cod_pro, tp.nome_pro,tsu.nome_sub,  sum(tiv.qtd_ite) as qv  ,sum(tiv.final_ite) as tv"
    strSQL = strSQL & " from "
    strSQL = strSQL & tabv & " as tv inner join " & tabi & " as tiv " & opt
    strSQL = strSQL & " inner join tab_produtos as tp on tp.cod_pro=tiv.cod_pro"
    strSQL = strSQL & " inner join tab_subsec as tsu on tp.cod_sub=tsu.cod_sub "
    strSQL = strSQL & " Where " & stat
    
    
    If txtDe.text <> "" Then
        strSQL = strSQL & " and  tv.dataefet_ven>=" & F_GravaData(txtDe.text)
        'strSQL = strSQL & " and  tv.dataefet_ven>=" & F_GravaData(txtAte.Text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and  tv.dataefet_ven<=" & F_GravaData(txtAte.text)
        'strSQL = strSQL & " and  tv.dataefet_ven>=" & F_GravaData(txtAte.Text)
    End If
    strSQL = strSQL & " and tipo_operacao=1"
    strSQL = strSQL & " group by tp.cod_pro "
    strSQL = strSQL & " Having qv > 0  order by qv desc  "
    If txtLimite.text <> "0" Then
        strSQL = strSQL & " limit " & txtLimite.text
    End If
    abreObjetoAdo ado, strSQL, grid
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Listagem de Produtos vendidos de: " & txtDe.text & "- até: " & txtAte.text, grid, ado
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub

Private Sub txtLimite_LostFocus()
    P_ValorLostFocus txtLimite, "###"
End Sub
