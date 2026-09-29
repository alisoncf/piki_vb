VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiLoj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Lojas"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmEdiLoj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6915
   Begin VB.Data dtServ 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   6855
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   5160
         Picture         =   "frmEdiLoj.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1800
         Picture         =   "frmEdiLoj.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiLoj.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDesab 
         Caption         =   "&Desab. / Hab."
         Enabled         =   0   'False
         Height          =   615
         Left            =   3480
         Picture         =   "frmEdiLoj.frx":0928
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiLoj.frx":0A72
      Height          =   3015
      Left            =   0
      OleObjectBlob   =   "frmEdiLoj.frx":0A87
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmEdiLoj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnDesab_Click()
'    Dim strSQL As String
'    With dtServ.Recordset
'        If MsgBox("Deseja Realmente mudat status da Conta " & _
'            !Nome_Con & "?", vbYesNo + vbCritical) = vbNo Then
'            Exit Sub
'        End If
'        strSQL = "update tab_contas set status_con='"
'        If !Status_Con = "D" Then
'            strSQL = strSQL & "H'"
'        Else
'            strSQL = strSQL & "D'"
'        End If
'        strSQL = strSQL & " where cod_Con='" & !cod_Con & "'"
'        banco.Execute strSQL
'        dtServ.Refresh
'    End With
    
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadLoj
    With dtServ.Recordset
    
        fr.Tag = !cod_loj
        fr.txtbxCidade.Text = !Cidade_Loj
        fr.txtbxRazao.Text = !nome_loj
        fr.txtbxCNPJ.Text = !Cnpj_loj
        fr.txtbxEndereco.Text = !endereco_loj
        fr.txtbxTel.Text = !fone_loj
        fr.cmbxUF.Text = !uf_loj
    End With
    fr.Show 0
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadLoj
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Load()
    Me.Height = 4500
    Me.Width = 6975
    Dim strSQL As String
    dtServ.Connect = banco.Connect
    strSQL = "select * from tab_lojas order by cnpj_loj"
    
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub


