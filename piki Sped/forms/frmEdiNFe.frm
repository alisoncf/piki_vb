VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiNFe 
   Caption         =   "NFes"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar NFe"
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdAtlzRetorno 
      Caption         =   "Atualizar Retorno"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   11760
      TabIndex        =   4
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquisar"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid grade 
      Bindings        =   "frmEdiNFe.frx":0000
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   11668
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
      Caption         =   "Lista de NFe"
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
Attribute VB_Name = "frmEdiNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmdAtlzRetorno_Click()
    Dim s As String
    Dim chNfe As String
    Dim autorizacao As String
    s = XmlRetornoNFe(ado!chave)
    If s = "" Then
        If MsgBox("Arquivo não encontrado!, Deseja localizar manualmente?", vbYesNo + vbQuestion) = vbYes Then
            mdiFrm.cmndlg.Filter = "*.xml"
            mdiFrm.cmndlg.ShowOpen
            
            s = mdiFrm.cmndlg.FileName
            s = XmlRetornoNFeCaminho(s)
            If s = "" Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    
    autorizacao = RetornaValorTagXML(s, "<cStat>")
    chNfe = RetornaValorTagXML(s, "<chNFe>")
    If autorizacao = "100" Then
        MsgBox "Nota fiscal autorizada", vbInformation
    End If
    
    
    MsgBox s
End Sub

Private Sub cmdCancelar_Click()
    AcBrMonitor.CancelarNFe ado!chave, "teste de cancelamento", "", ""
End Sub

Private Sub cmdConsultar_Click()
    Dim s As String
    s = ConsultarNFe(ado!chave)
    MsgBox s
    
    
End Sub

Private Sub cmdEnviar_Click()
    Dim strSQL As String
    
    strSQL = TransmitirNFcE(ado!id)
    
    If InStr(1, "CStat=100", strSQL) > 0 Then
        AlteraStatusNfe ado!id, "2"
        'MsgBox strSQL
    Else
       ' MsgBox strSQL
    End If
    
    
    
        
    


End Sub

Private Sub cmdExcluir_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    
    With ado
        
        If !status > "1" Then
            MsgBox "Não pode ser excluída!", vbExclamation
            Exit Sub
        End If
        
        strSQL = "delete from nfe where id=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_item where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_dest where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_emit where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_enderdest where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_enderemit where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_dest where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from icms_tot where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_infadic where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_dest where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_item_imposto where id_nf=" & ado!id
        conn.Execute strSQL
        strSQL = "delete from nfe_dest where id_nf=" & ado!id
        conn.Execute strSQL
        
        
        
    End With
    Exit Sub
trata_Erro:
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub

Private Sub cmdXml_Click()
    GeraXML_NFcE_Dados ado!id
End Sub

Private Sub Command1_Click()
    Dim arquivo As String
    arquivo = "C:\ACBrMonitorPLUS\Arqs\05063223000111\NFCe\201711\NFCe\52171105063223000111650010000000071289562461-nfe.xml"
    AcBrMonitor.ImprimirDanfe arquivo, ""
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select * from nfe order by id desc"
    abreObjetoAdo ado, strSQL, grade
    
    
    
End Sub
