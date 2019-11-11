Attribute VB_Name = "mGraphicInterface"
Sub AddBulkHeader(curRow As Integer)
    Range("B" & curRow).Value = "Item #"
    Range("C" & curRow).Value = "Lot"
    Range("D" & curRow).Value = "Lot Status"

    With Range("B" & curRow & ":" & "D" & curRow)
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 14470546
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

End Sub

Sub AddBulkText(curRow As Integer, strItem As Variant, strLot As Variant, strStatus As Variant)
    Range("A" & curRow).Value = "Bulk:"
    With Range("A" & curRow)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Color = -3109376
        .Font.TintAndShade = 0
        .Font.Bold = True
    End With
    
    Range("B" & curRow).Value = strItem
    Range("C" & curRow).Value = strLot
    Range("D" & curRow).Value = strStatus
End Sub

Sub SetBackgroundBulk(curRow As Integer)
    With Range("B" & curRow & ":" & "D" & curRow)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 14277081
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
End Sub

Sub AddBottomBorders(curRow As Integer)
    With Range("A" & curRow & ":" & "L" & curRow).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
End Sub

'------------------------------------------------------------------------
'
' Button menu interface defenitions
'
'------------------------------------------------------------------------

Sub Clear_Click()
    Dim vntTopType As Variant
    Dim intTopInset As Integer
    Dim intTopDepth As Integer
    
    'Record original button properties
    With ActiveSheet.Shapes(Application.Caller).ThreeD
        vntTopType = .BevelTopType
        intTopInset = .BevelTopInset
        intTopDepth = .BevelTopDepth
    End With

    'Button Down
    With ActiveSheet.Shapes(Application.Caller).ThreeD
        .BevelTopType = msoBevelSoftRound
        .BevelTopInset = 12
        .BevelTopDepth = 4
    End With
    Application.ScreenUpdating = True
    
    'Pause while Button is Down
    Sleep 250
    Application.ScreenUpdating = True
        
    'Button Up - set back to original values
    With ActiveSheet.Shapes(Application.Caller).ThreeD
        .BevelTopType = vntTopType
        .BevelTopInset = intTopInset
        .BevelTopDepth = intTopDepth
    End With
    mUtility.ClearReport
End Sub

Sub Update_Click()
    Dim vntTopType As Variant
    Dim intTopInset As Integer
    Dim intTopDepth As Integer
    
    'Record original button properties
    With ActiveSheet.Shapes(Application.Caller).ThreeD
        vntTopType = .BevelTopType
        intTopInset = .BevelTopInset
        intTopDepth = .BevelTopDepth
    End With

    'Button Down
    With ActiveSheet.Shapes(Application.Caller).ThreeD
        .BevelTopType = msoBevelSoftRound
        .BevelTopInset = 12
        .BevelTopDepth = 4
    End With
    Application.ScreenUpdating = True
    
    'Pause while Button is Down
    Sleep 250
    Application.ScreenUpdating = True
        
    'Button Up - set back to original values
    With ActiveSheet.Shapes(Application.Caller).ThreeD
        .BevelTopType = vntTopType
        .BevelTopInset = intTopInset
        .BevelTopDepth = intTopDepth
    End With
        
    ProgressBarPopulateForm.Show
End Sub

Sub Execute_Click()
    Dim vntTopType As Variant
    Dim intTopInset As Integer
    Dim intTopDepth As Integer
    
    'Record original button properties
    With ActiveSheet.Shapes(Application.Caller).ThreeD
        vntTopType = .BevelTopType
        intTopInset = .BevelTopInset
        intTopDepth = .BevelTopDepth
    End With

    'Button Down
    With ActiveSheet.Shapes(Application.Caller).ThreeD
        .BevelTopType = msoBevelSoftRound
        .BevelTopInset = 12
        .BevelTopDepth = 4
    End With
    Application.ScreenUpdating = True
    
    'Pause while Button is Down
    Sleep 250
    Application.ScreenUpdating = True
        
    'Button Up - set back to original values
    With ActiveSheet.Shapes(Application.Caller).ThreeD
        .BevelTopType = vntTopType
        .BevelTopInset = intTopInset
        .BevelTopDepth = intTopDepth
    End With
        
    Execute_Options
End Sub
