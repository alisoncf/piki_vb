VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelAcertosFormasPag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Acertos"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   13065
      TabIndex        =   12
      Top             =   7680
      Width           =   13095
      Begin VB.OptionButton optnEsp 
         Caption         =   "Por Espécie"
         Height          =   195
         Left            =   10920
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optnFormaP 
         Caption         =   "Por Forma de Pagamento"
         Height          =   195
         Left            =   8640
         TabIndex        =   16
         Top             =   120
         Width           =   2175
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
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "&Gráfico"
         Height          =   615
         Left            =   9840
         Picture         =   "frmRelAcertosFormasPag.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   11520
         Picture         =   "frmRelAcertosFormasPag.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   8160
         Picture         =   "frmRelAcertosFormasPag.frx":023F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   6480
         Picture         =   "frmRelAcertosFormasPag.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   285
         Left            =   5880
         Picture         =   "frmRelAcertosFormasPag.frx":09D6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   120
         Width           =   1095
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
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelAcertosFormasPag.frx":0D60
      Height          =   6495
      Left            =   120
      OleObjectBlob   =   "frmRelAcertosFormasPag.frx":0D71
      TabIndex        =   9
      Top             =   960
      Width           =   13095
   End
End
Attribute VB_Name = "frmRelAcertosFormasPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim sCampo As String
    Dim sCod As String
    If optnFormaP.Value Then
        sCampo = "tf.nome_for"
        sCod = "tf.cod_for"
    Else
        sCampo = "tf.tipo_cai as nome_for"
        sCod = "tf.tipo_cai"
    End If
    
    strSQL = "select " & sCampo & ",sum(valor_ace) as total,sum(diferenca_ace) as dif "
    strSQL = strSQL & " from tab_acertos as ta, tab_formaspag as tf "
    strSQL = strSQL & " where tf.cod_for=ta.cod_for "
    strSQL = strSQL & " and status_ace=2 "
    If txtbxDataFin.Text <> "" Then
        strSQL = strSQL & " and data_ace<=" & F_GravaData(txtbxDataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and data_ace>=" & F_GravaData(txtbxDataIni.Text)
    End If
    If txtbxRepresentante.tag <> "" Then
        strSQL = strSQL & " and ta.cod_rep='" & txtbxRepresentante.tag & "' "
    End If
    
    strSQL = strSQL & " group by " & sCod
    
    dt.RecordSource = strSQL
    dt.Refresh
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    With dt.Recordset
        Do While Not .EOF
            cTotal = cTotal + !total
            
            .MoveNext
        Loop
        lblVendas.caption = FormatVal(cTotal)
    End With
End Sub

Private Sub cmndbtnBuscaRep_Click()
    buscaRep
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DbgrdToHtml "Relatório de Acertos de " & txtbxDataIni.Text & " até " & _
        txtbxDataFin.Text, dbgrd, dt, txtbxRepresentante.Text, "1"
    
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dt.Connect = Banco.Connect
    
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDataFin
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
    txtbxRepresentante.tag = r.cod_rep
    busca
End Sub
