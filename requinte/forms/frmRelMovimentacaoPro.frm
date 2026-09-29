VERSION 5.00
Begin VB.Form frmRelMovimentacaoPro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Movimentação de Produtos"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11850
   Icon            =   "frmRelMovimentacaoPro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   11655
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cmbxOrderBY 
         Height          =   315
         ItemData        =   "frmRelMovimentacaoPro.frx":014A
         Left            =   8640
         List            =   "frmRelMovimentacaoPro.frx":0157
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   10440
         Picture         =   "frmRelMovimentacaoPro.frx":0185
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
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
      Begin VB.CommandButton cmndbtnBuscaSec 
         Caption         =   "..."
         Height          =   285
         Left            =   5160
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   8280
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblSubCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Subseção:"
         Height          =   195
         Left            =   5520
         TabIndex        =   12
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lblSecCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seção:"
         Height          =   195
         Left            =   2400
         TabIndex        =   11
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   8640
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   9840
      Picture         =   "frmRelMovimentacaoPro.frx":050F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Im&primir"
      Height          =   735
      Left            =   7800
      Picture         =   "frmRelMovimentacaoPro.frx":067A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnGraf 
      Caption         =   "&Gráfico"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5640
      Picture         =   "frmRelMovimentacaoPro.frx":0A87
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   1815
   End
End
Attribute VB_Name = "frmRelMovimentacaoPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sTab As String
    sTab = "tab_relmovprod_tmp"
    
    strSQL = "drop table if exists " & sTab
    Banco.Execute strSQL
    strSQL = "create table " & sTab & " (cod_pro char(8) not null, "
    strSQL = strSQL & "qc decimal(14,2) not null, qv decimal(14,2) not null,qa decimal(14,2) not null, "
    strSQL = strSQL & " vc decimal(14,2) not null, vv decimal(14,2) not null, va decimal(14,2) not null,"
    strSQL = strSQL & "tipo char(1) not null)"
    Banco.Execute strSQL
    strSQL = "delete from " & sTab
    Banco.Execute strSQL
    
    
    strSQL = "insert into " & sTab & "(cod_pro,qc,qv,qa,vc,vv,va)values("
    strSQL = strSQL & " select cod_pro,0,sum(ti.qtd_ite) as qtd,0,0,sum(final_ite) as final "
    
    strSQL = strSQL & " from tab_produtos as tp,tab_itvendas as ti, tab_vendas as tv,tab_secoes as ts, "
    strSQL = strSQL & " tab_subsec as tu "
    strSQL = strSQL & " where ts.cod_sec=tu.cod_sec "
    strSQL = strSQL & " and tu.cod_sub=tp.cod_sub "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tv.num_ven=ti.num_ven "

    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.Text)
    End If
    strSQL = strSQL & " group by tp.cod_pro )"
    Banco.Execute strSQL
    
End Sub

Private Sub txtbxDatafin_GotFocus()
    P_SelecTexto txtbxDAtaFIn
End Sub

Private Sub txtbxDatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDatafin_LostFocus()
    P_AjustaData txtbxDAtaFIn
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
