VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEfetivaVen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Efetivação de Vendas"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9885
   Begin VB.Frame Frame1 
      Caption         =   "&Vendas a Efetivar:"
      Height          =   5655
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9855
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   4440
         Top             =   1200
      End
      Begin MSAdodcLib.Adodc adc 
         Height          =   375
         Left            =   1320
         Top             =   6120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dtgrd 
         Bindings        =   "frmEfetivaVen.frx":0000
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8705
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   9855
      Begin VB.CommandButton cmndbtnReabrir 
         Caption         =   "&Reabrir"
         Height          =   660
         Left            =   3960
         Picture         =   "frmEfetivaVen.frx":0012
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Nota Fiscal"
         Enabled         =   0   'False
         Height          =   660
         Left            =   2040
         Picture         =   "frmEfetivaVen.frx":015C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   660
         Left            =   5880
         Picture         =   "frmEfetivaVen.frx":06E6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   660
         Left            =   7800
         Picture         =   "frmEfetivaVen.frx":0A70
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnEfetivar 
         Caption         =   "&Efetivar"
         Height          =   660
         Left            =   120
         Picture         =   "frmEfetivaVen.frx":0BBA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmEfetivaVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub busca()
     Dim strSQL As String
    
    
    
    
    
    strSQL = "select tc.cod_cli,tc.nome_cli,tu.guerra_fun,tv.num_ven,tv.data_ven,tv.situacao_ven,tv.valor_ven, "
    strSQL = strSQL & " tf.*,(tv.desconto_ven*tv.valor_ven)/100 as desconto"
    strSQL = strSQL & " from " & TAB_CLIENTES & "  as tc, "
    strSQL = strSQL & " " & TAB_VENDAS & " as tv, " & TAB_VENFEC & "  as tf, tab_funcionarios as tu "
    strSQL = strSQL & " where tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tv.situacao_ven='2' "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    strSQL = strSQL & " and tv.abrevenda_usu=tu.cod_fun "
    If strCod_Cai <> "00" And strCod_Cai <> "" Then
        strSQL = strSQL & " and tf.cod_cai='" & strCod_Cai & "' "
    End If
    strSQL = strSQL & " order by tf.num_ven"
    adc.RecordSource = strSQL
    adc.Refresh
End Sub
    
Private Sub cmndbtnEfetivar_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If dtgrd.row = -1 Then
        MsgBox "selecione a Venda!", vbInformation
        dtgrd.SetFocus
        Exit Sub
    End If
    With adc.Recordset
        If MsgBox("Deseja Realmente Efetivar a Venda?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
            
        End If
    End With

    strSQL = "select * from  " & TAB_VENFEC & "  where num_ven='" & adc.Recordset!num_ven & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
            frmFecharVenda.operacao = "venda"
            frmFecharVenda.lblvlrVenda.caption = FormatVal(!valorfinal_ven)
            frmFecharVenda.lblDesconto.caption = FormatVal(!desconto_ven)
            If !aprazo_ven > 0 Then
                frmFecharVenda.frAp.Enabled = True
                frmFecharVenda.frForm.Enabled = True
                frmFecharVenda.frVal.Enabled = True
                If Not IsNull(!cod_tip) Then
                    frmFecharVenda.cmbxFormas_Pag.tag = !cod_tip
                End If
            End If
            If !tipo_ven = "C" Then
                frmFecharVenda.optnCheque.Value = True
            Else
                frmFecharVenda.otpnCart.Value = True
            End If
            frmFecharVenda.lblAPrazo.caption = FormatVal(!aprazo_ven)
            frmFecharVenda.lblAPrazo.tag = FormatVal(!aprazo_ven)
            frmFecharVenda.lblAvIsta.caption = FormatVal(!avista_ven)
            frmFecharVenda.lblFinan.caption = FormatVal(!finan_ven)
            
            frmFecharVenda.lblvlrVenda.caption = FormatVal(!valorfinal_ven + (!valorfinal_ven * !desconto_ven / 100))
            frmFecharVenda.lblTotal.caption = FormatVal(!valorfinal_ven)
            frmFecharVenda.lblNome_Cli.caption = adc.Recordset!nome_Cli
            frmFecharVenda.lblNome_Cli.tag = adc.Recordset!cod_cli
            frmFecharVenda.tag = !num_ven
            frmFecharVenda.caption = "Fechar Venda: " & !num_ven & " de " & adc.Recordset!nome_Cli
            
            frmFecharVenda.Show 1
            
            adc.Refresh
            rPesq.Close
         
        End With
End Sub


Private Sub cmndbtnReabrir_Click()
    Dim strSQL As String
    If dtgrd.row = -1 Then
        MsgBox "selecione a Venda!", vbInformation
        dtgrd.SetFocus
        Exit Sub
    End If
    With adc.Recordset
        If MsgBox("Deseja Realmente Reabrir a Venda?", vbYesNo + vbQuestion) = vbYes Then
            
            strSQL = strSQL & " update " & TAB_VENDAS & " set situacao_ven='1'"
            strSQL = strSQL & " where num_ven='" & !num_ven & "' "
            
            
            conn.Execute strSQL
            strSQL = "delete from  " & TAB_VENFEC & "  where num_ven='" & !num_ven & "' "
            conn.Execute strSQL
            busca
        End If
    End With
End Sub

Private Sub cmndbtnSair_Click()
        Unload Me
End Sub

Private Sub dtVendas_Reposition()
    If adc.Recordset.RecordCount = 0 Then
        cmndbtnCancelar.Enabled = False
        cmndbtnEfetivar.Enabled = False
        cmndbtnImpr.Enabled = False
        cmndbtnReabrir.Enabled = False
    Else
        cmndbtnCancelar.Enabled = True
        cmndbtnEfetivar.Enabled = True
        cmndbtnImpr.Enabled = True
        cmndbtnReabrir.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    
    busca
    Me.caption = "Vendas a Efetivar  - Usuário: " & strNome_Usu
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    adc.ConnectionString = strConnect
End Sub

Private Sub Timer1_Timer()
    
    Dim i As Integer
    On Error Resume Next
    i = adc.Recordset.AbsolutePosition
    busca
    adc.Recordset.Move i
End Sub
