Attribute VB_Name = "mDetailGraphics"
Sub FormatHeadingDetail(curRow As Integer)
    'item #
    With Range("A" & curRow)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Bold = True
    End With
    
    'description
    With Range("B" & curRow & ":" & "C" & curRow)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    'qty
    With Range("F" & curRow)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    'total dollar amount
    With Range("G" & curRow)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
End Sub

Sub DetailAddBottomBorders(curRow As Integer)
    With Range("A" & curRow & ":" & "G" & curRow).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
End Sub

Sub AddLotSeparator(curRow As Integer)
    With Range("A" & curRow & ":" & "G" & curRow).Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
    End With
End Sub
