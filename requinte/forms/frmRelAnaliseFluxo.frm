VERSION 5.00
Begin VB.Form frmRelAnaliseFluxo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Análise Direta de Fluxo"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2415
         Begin VB.TextBox txtbxDAtaFIn 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1320
            TabIndex        =   4
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D&e:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2415
         Begin VB.CommandButton cmndbtnGraf 
            Caption         =   "&Gráfico Comparativo"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelAnaliseFluxo.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1680
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnFechar 
            Cancel          =   -1  'True
            Caption         =   "&Sair"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelAnaliseFluxo.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnImprimir 
            Caption         =   "Im&primir"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelAnaliseFluxo.frx":0D0C
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton cmndbtnBuscar 
            Caption         =   "&Buscar"
            Height          =   685
            Left            =   120
            Picture         =   "frmRelAnaliseFluxo.frx":1016
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Line Line4 
         X1              =   2880
         X2              =   8040
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line3 
         X1              =   2880
         X2              =   8040
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentual F:"
         Height          =   195
         Left            =   6000
         TabIndex        =   25
         Top             =   3960
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença:"
         Height          =   195
         Left            =   2880
         TabIndex        =   23
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentual F:"
         Height          =   195
         Left            =   6000
         TabIndex        =   21
         Top             =   1800
         Width           =   945
      End
      Begin VB.Line Line2 
         X1              =   2880
         X2              =   8040
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença:"
         Height          =   195
         Left            =   2880
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada E:"
         Height          =   195
         Left            =   2880
         TabIndex        =   18
         Top             =   480
         Width           =   750
      End
      Begin VB.Label lblEntradaE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6840
         TabIndex        =   17
         Top             =   165
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada F:"
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblEntradaF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6840
         TabIndex        =   15
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Giro E:"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblGiroE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6840
         TabIndex        =   13
         Top             =   2205
         Width           =   1260
      End
      Begin VB.Line Line1 
         X1              =   2880
         X2              =   8040
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label lblGiroF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6840
         TabIndex        =   12
         Top             =   2880
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Giro F:"
         Height          =   195
         Left            =   2880
         TabIndex        =   11
         Top             =   3195
         Width           =   465
      End
      Begin VB.Label lblPercEnt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   7080
         TabIndex        =   22
         Top             =   1725
         Width           =   510
      End
      Begin VB.Label lblDifEnt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3720
         TabIndex        =   20
         Top             =   1725
         Width           =   570
      End
      Begin VB.Label lblPerSai 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   7020
         TabIndex        =   26
         Top             =   3885
         Width           =   510
      End
      Begin VB.Label lblDifSai 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3690
         TabIndex        =   24
         Top             =   3885
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmRelAnaliseFluxo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    strSQL = "select sum(valor_not) as total from tab_Compras where entrada_not>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and entrada_not<=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " and status_not>=3 "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            lblEntradaF.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    
    strSQL = "select sum(valor_not) as total from tab_Cmp where entrada_not>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and entrada_not<=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " and status_not>=3 "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblEntradaE.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    lblDifEnt.caption = FormatVal(CCur(lblEntradaE.caption) - CCur(lblEntradaF.caption))
    lblPercEnt.caption = FormatVal(CCur(lblEntradaF.caption) / CCur(lblEntradaE.caption) * 100) & "%"
    
    strSQL = "select sum(valor_sai) as total from tab_Saidas where emissao_sai>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and emissao_sai<=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " and serie_sai <>'CANC' "
    strSQL = strSQL & " and serie_sai <>'CNC' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblGiroF.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    
    strSQL = "select sum(valorfinal_ven) as total from tab_vendas as tv, "
    strSQL = strSQL & " tab_venfec as tf where dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxdataFin.Text)
    strSQL = strSQL & " and tv.num_ven=tf.num_ven "
    strSQL = strSQL & " and tv.situacao_ven=4"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblGiroE.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    lblDifSai.caption = FormatVal(CCur(lblGiroE.caption) - CCur(lblGiroF.caption))
    If CCur(lblGiroE.caption) > 0 Then
    lblPerSai.caption = FormatVal(CCur(lblGiroF.caption) / CCur(lblGiroE.caption) * 100) & "%"
    End If
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
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
