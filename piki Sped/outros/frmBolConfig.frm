VERSION 5.00
Begin VB.Form frmBolConfig 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Coordenadas de Impressão"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbx 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmBolConfig.frx":0000
      Left            =   120
      List            =   "frmBolConfig.frx":0007
      TabIndex        =   42
      Text            =   "Banco do Brasil"
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmndbtnInstr 
      Caption         =   "Instrução"
      Height          =   615
      Left            =   1040
      Picture         =   "frmBolConfig.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   39
      Text            =   "170"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   38
      Text            =   "102"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txtbxEndy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   34
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtbxEndx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   33
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtbxSacadoy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   31
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtbxSacadox 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   30
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtbxValory 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtbxValorx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   27
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtbxInstry 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   25
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txtbxinstrx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   24
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txtbxcnpjy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   37
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtbxcnpjx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   36
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtbxAceitey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   22
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtbxAceitex 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtbxEspy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   19
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtbxespx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   18
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtbxdocy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtbxdocx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtbxEmissy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtbxEmissx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtbxVency 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtbxVencx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtbxParcy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtbxParcx 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   1940
      Picture         =   "frmBolConfig.frx":0166
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Salvar"
      Height          =   615
      Left            =   120
      Picture         =   "frmBolConfig.frx":02B0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   0
      Width           =   510
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Medidas da Folha"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   40
      Top             =   4800
      Width           =   1890
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   32
      Top             =   4080
      Width           =   1890
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sacado"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   29
      Top             =   3720
      Width           =   1890
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   1890
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Instrução"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   1890
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CNPJ/CPF && CEP:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   35
      Top             =   4440
      Width           =   1890
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aceite"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   1890
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Espécie"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1890
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nº Documento"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1890
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Emissão"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1890
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vencimento"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parcela:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1890
   End
End
Attribute VB_Name = "frmBolConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnInstr_Click()
    frmBolInstr.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    grava
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    'On Error Resume Next
    strSQL = "select * from tab_bolconfig "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        txtbxAceitex.Text = !aceitex_bol
        txtbxAceitey.Text = !aceiteY_bol
        txtbxcnpjx.Text = Format(!cnpjx_bol)
        txtbxcnpjy.Text = Format(!cnpjy_bol)
        txtbxdocx.Text = !numdocx_bol
        txtbxdocy.Text = !numdocY_boL
        txtbxEmissx.Text = !emissx_bol
        txtbxEmissy.Text = !emissy_bol
        txtbxEndx.Text = !endx_bol
        txtbxEndy.Text = !endy_bol
        txtbxespx.Text = !espDOCx_BOL
        txtbxEspy.Text = !espDOCy_BOL
        txtbxinstrx.Text = !instrx_bol
        txtbxInstry.Text = !instry_bol
        txtbxParcx.Text = Format(!parcx_bol)
        txtbxParcy.Text = Format(!parcy_bol)
        txtbxSacadox.Text = !sacadox_bol
        txtbxSacadoy.Text = !sacadoy_bol
        txtbxValorx.Text = !valorx_bol
        txtbxValory.Text = !valory_bol
        txtbxVencx.Text = !vencx_bol
        txtbxVency.Text = !vency_bol
        .Close
    End With
End Sub
Sub grava()
    strSQL = "delete from tab_bolconfig"
    Banco.Execute strSQL
    strSQL = "INSERT INTO tab_bolconfig (parcx_bol, parcy_bol, vencx_bol, "
    strSQL = strSQL & "vency_bol, emissx_bol, emissy_bol, numdocx_bol, "
    strSQL = strSQL & " numdocy_bol, espdocx_bol, espdocy_bol, aceitex_bol, "
    strSQL = strSQL & "aceitey_bol, cnpjx_bol, cnpjy_bol, instrx_bol,"
    strSQL = strSQL & "instry_bol, valorx_bol, valory_bol, sacadox_bol, "
    strSQL = strSQL & "sacadoy_bol, endx_bol, endy_bol) VALUES ("
    strSQL = strSQL & txtbxParcx.Text & ","
    strSQL = strSQL & txtbxParcy.Text & ","
    strSQL = strSQL & txtbxVencx.Text & ","
    strSQL = strSQL & txtbxVency.Text & ","
    strSQL = strSQL & txtbxEmissx.Text & ","
    strSQL = strSQL & txtbxEmissy.Text & ","
    strSQL = strSQL & txtbxdocx.Text & ","
    strSQL = strSQL & txtbxdocy.Text & ","
    strSQL = strSQL & txtbxespx.Text & ","
    strSQL = strSQL & txtbxEspy.Text & ","
    strSQL = strSQL & txtbxAceitex.Text & ","
    strSQL = strSQL & txtbxAceitey.Text & ","
    strSQL = strSQL & txtbxcnpjx.Text & ","
    strSQL = strSQL & txtbxcnpjy.Text & ","
    strSQL = strSQL & txtbxinstrx.Text & ","
    strSQL = strSQL & txtbxInstry.Text & ","
    strSQL = strSQL & txtbxValorx.Text & ","
    strSQL = strSQL & txtbxValory.Text & ","
    strSQL = strSQL & txtbxSacadox.Text & ","
    strSQL = strSQL & txtbxSacadoy.Text & ","
    strSQL = strSQL & txtbxEndx.Text & ","
    strSQL = strSQL & txtbxEndy.Text & ")"
    Banco.Execute strSQL
End Sub
Private Sub Form_Load()
'    With msflxGrd
'        .Cols = 4
'        .rows = 1
'        .Col = 0: .Text = "Descrição"
'        .Col = 1: .Text = "Eixo X"
'        .Col = 2: .Text = "Eixo Y"
'        .Col = 3: .Text = "Chave"
'        .ColWidth(0) = 2000
'
'    End With
    busca
End Sub

