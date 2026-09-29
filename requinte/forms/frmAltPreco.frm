VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAltPreco 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listagem de Alteração de Preços"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9763
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
      ColumnCount     =   7
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
         DataField       =   "codfab_pro"
         Caption         =   "Ref."
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
      BeginProperty Column03 
         DataField       =   "antigo_pro"
         Caption         =   "Pç Antigo"
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
         DataField       =   "novo_pro"
         Caption         =   "Pç Novo"
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
      BeginProperty Column05 
         DataField       =   "data_alt"
         Caption         =   "em"
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
      BeginProperty Column06 
         DataField       =   "cod_usu"
         Caption         =   "Por"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   7680
      Picture         =   "frmAltPreco.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   5880
      Picture         =   "frmAltPreco.frx":054E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtbxDe 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apartir de:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   720
   End
End
Attribute VB_Name = "frmAltPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
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
    P_CabRelTimbr "Listagem de Alteração de Preços ", obj
    
    obj.Print ""
    Dim yIni As Single, yFin As Single
    yIni = obj.CurrentY
    P_ImprX obj, "Código", 5
    P_ImprX obj, "Produto", 25
    P_ImprX obj, "Pç. Antigo", 115, True, True
    P_ImprX obj, "Novo Preço", 140, True, True
    P_ImprX obj, "Preço Atual", 160, True, True
    P_ImprX obj, "Alterado Em", 184, False, True
    P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY
    With dtPro.Recordset
        .MoveFirst
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                
                If TypeName(obj) = "Printer" Then
                    Printer.NewPage
                    yIni = obj.CurrentY
                End If
            End If
            
            P_ImprX obj, !cod_pro, 5
            P_ImprX obj, !nome_pro, 25
            P_ImprX obj, FormatVal(!antigo_pro), 115, True, True
            P_ImprX obj, FormatVal(!novo_pro), 140, True, True
            P_ImprX obj, FormatVal(!pv_pro), 160, True, True
            P_ImprX obj, !data_alt, 164, False
            P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY, vbDot
            
            
            
            .MoveNext
        Loop
        yFin = obj.CurrentY
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY
        P_Line obj, 4, yIni, 4, yFin
        P_Line obj, 200, yIni, 200, yFin
        
    End With
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
    
End Sub

Private Sub Form_Load()
    
    txtbxDe.Text = DateAdd("D", -1, Date)
    Set ado = New ADODB.Recordset
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tp.nome_pro,tp.pv_pro,ta.* from tab_produtos as tp, tab_altpreco as ta "
    strSQL = strSQL & " where tp.cod_pro=ta.cod_pro "
    strSQL = strSQL & " and left(ta.data_alt,10) >=" & F_GravaData(txtbxDe.Text)
    strSQL = strSQL & " order by data_alt "
    abreObjetoAdo ado, strSQL, grid
End Sub
