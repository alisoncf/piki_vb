VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSituacaoDoc 
   Caption         =   "sped_situacao_doc"
   ClientHeight    =   5025
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   6555
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   300
      Left            =   5340
      TabIndex        =   0
      Top             =   3960
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      DragIcon        =   "frmSituacaoDoc.frx":0000
      Height          =   3840
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   6773
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   9
      FixedCols       =   0
      BackColorFixed  =   8421376
      ForeColorFixed  =   16777215
      GridColor       =   8421504
      GridColorFixed  =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      MergeCells      =   4
      AllowUserResizing=   1
      FormatString    =   "00|Documento regular"
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
End
Attribute VB_Name = "frmSituacaoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MARGIN_SIZE = 60      ' in Twips
' variables for data binding
Private datPrimaryRS As ADODB.Recordset

' variables for column dragging
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer

Private Sub Form_Load()

    Dim sConnect As String
    Dim sSQL As String
    Dim dfwConn As ADODB.Connection
    Dim i As Integer

    ' set strings
    sConnect = "Provider=MSDASQL.1;"
    sSQL = "select codigo,descricao from sped_situacao_doc"

    ' open connection
    Set dfwConn = New Connection
    Set dfwConn = conn_client
    
    ' create a recordset using the provided collection
    Set datPrimaryRS = New Recordset
    datPrimaryRS.CursorLocation = adUseClient
    datPrimaryRS.Open sSQL, dfwConn, adOpenForwardOnly, adLockReadOnly

    Set MSHFlexGrid1.DataSource = datPrimaryRS

    With MSHFlexGrid1

        .Redraw = False
        ' set grid's column widths
        .ColWidth(0) = -1
        .ColWidth(1) = -1

        ' set grid's column merging and sorting
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next i

        .Sort = flexSortGenericAscending

        ' set grid's style
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' make header bold
        .Row = 0
        .Col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True

        ' grey every other column
        For i = .FixedCols To .Cols() - 1 Step 2
            .Col = i
            .Row = .FixedRows
            .RowSel = .Rows - 1
            .CellBackColor = &HC0C0C0   ' light grey
        Next i

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

        .ColWidth(1) = 3500
    End With

End Sub

Private Sub MSHFlexGrid1_DragDrop(Source As Control, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    If m_iDragCol = -1 Then Exit Sub    ' we weren't dragging
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub

    With MSHFlexGrid1
        .Redraw = False
        .ColPosition(m_iDragCol) = .MouseCol

        .FillStyle = flexFillRepeat
        .Col = 0
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = &HFFFFFF
        Dim iLoop As Integer
        For iLoop = .FixedCols To .Cols() - 1 Step 2
            .Col = iLoop
            .Row = .FixedRows
            .RowSel = .Rows - 1
            .CellBackColor = &HC0C0C0
        Next iLoop
        .FillStyle = flexFillSingle

        DoSort
        .Redraw = True
    End With

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub

    xdn = X
    ydn = Y
    m_iDragCol = -1     ' clear drag flag
    m_bDragOK = True

End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    ' test to see if we should start drag
    If Not m_bDragOK Then Exit Sub
    If Button <> 1 Then Exit Sub                        ' wrong button
    If m_iDragCol <> -1 Then Exit Sub                   ' already dragging
    If Abs(xdn - X) + Abs(ydn - Y) < 50 Then Exit Sub   ' didn't move enough yet
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub         ' must drag header

    ' if got to here then start the drag
    m_iDragCol = MSHFlexGrid1.MouseCol
    MSHFlexGrid1.Drag vbBeginDrag

End Sub

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    m_bDragOK = False

End Sub

Sub DoSort()

    With MSHFlexGrid1
        .Redraw = False
        .Col = 0
        .Row = 1
        .RowSel = .Rows - 1
        .Sort = flexSortGenericAscending
        .Redraw = True
    End With

End Sub

Private Sub Form_Resize()

    Dim sngButtonTop As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single

    On Error GoTo Form_Resize_Error
    With Me
        sngScaleWidth = .ScaleWidth
        sngScaleHeight = .ScaleHeight

        ' move Close button to the lower right corner
        With .cmdClose
                sngButtonTop = sngScaleHeight - (.Height + MARGIN_SIZE)
                .Move sngScaleWidth - (.Width + MARGIN_SIZE), sngButtonTop
        End With

        .MSHFlexGrid1.Move MARGIN_SIZE, _
            MARGIN_SIZE, _
            sngScaleWidth - (2 * MARGIN_SIZE), _
            sngButtonTop - (2 * MARGIN_SIZE)

    End With
    Exit Sub

Form_Resize_Error:
    ' avoid error on negative values
    Resume Next

End Sub
Private Sub cmdClose_Click()

    Unload Me

End Sub


