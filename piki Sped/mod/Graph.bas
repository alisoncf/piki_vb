Attribute VB_Name = "Graph"
Sub P_Graph(bl() As valChart, qtd As Integer, Optional tipo As Long = 1, _
Optional visLeg As Boolean = True, Optional Tit As String = "", _
    Optional Bottom As String = "", Optional toolVis As Boolean = False, _
    Optional rows As Integer = 1, Optional porcentagem As Boolean = True)

    

'0   No graph               12  Bubble
'1   2D pie                 13  Tape
'2   3D pie                 14  3D area
'3   2D bar                 15  Log/log
'4   3D bar (default)       16  Lin/log
'5   Gantt                  17  Box-whisker
'6   Line                   18  Open-high-low-close
'7   Log/lin                19  Candlestick
'8   2D area                20  True3D surface
'9   2D scatter             21  True3D scatter
'10  Polar                  22  Time series
'11  High-low-close
'    Dim iRow As Integer
'    Dim iCount As Integer
'    With frmGraph.grph
'        .NumSets = rows
'        .GraphTitle = Tit
'        .BottomTitle = Bottom
'        .Labels = 1
'        .NumPoints = qtd
'
'        If porcentagem Then
'            .GraphStyle = 4
'        Else
'            .GraphStyle = 0
'        End If
'        If toolVis = True Then
'            .Toolbar = graphRunTime
'        End If
'        For iRow = 1 To rows
'            For iCount = 1 To qtd
'                .ThisSet = iRow
'                .data(iCount) = bl(iCount - 1).Value
'                If visLeg = True Then
'
'                    .Legend(iCount) = bl(iCount - 1).caption
'                    If .Legend(iCount) = "" Then
'                        .Legend(iCount) = "---"
'                    End If
'                    .LegendStyle = 0
'                End If
'
'            Next
'        Next
'    .GraphType = tipo
'    .DrawMode = 3
'    End With
'    frmGraph.Show 0
End Sub


