VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEscolheCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caixa"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmEscolheCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   3015
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   360
      ScaleHeight     =   2985
      ScaleWidth      =   3105
      TabIndex        =   0
      Top             =   360
      Width           =   3135
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   630
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1111
         ButtonWidth     =   1667
         ButtonHeight    =   953
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "01 - Caixa 1"
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "02 - Caixa 2"
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
End
Attribute VB_Name = "frmEscolheCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_contas where tipo_con=1 order by cod_con "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            .MoveNext
        Loop
    End With
End Sub

