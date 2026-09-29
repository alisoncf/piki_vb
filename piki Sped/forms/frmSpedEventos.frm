VERSION 5.00
Begin VB.Form frmSpedEventos 
   Caption         =   "Eventos SPED"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmSpedEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    
    
End Sub
