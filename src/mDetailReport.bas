Attribute VB_Name = "mDetailReport"
Sub DoWork()
    'progress bar
    Dim i As Integer
    i = 1
    Call PopulateSalesData
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("Seroyal")
    Call DetailWorksheet("SER")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("Trophic")
    Call DetailWorksheet("TRO")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("GOL")
    Call DetailWorksheet("GOL")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("MCO")
    Call DetailWorksheet("MCO")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("DLC")
    Call DetailWorksheet("DLC")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("Iovate")
    Call DetailWorksheet("IOV")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("House")
    Call DetailWorksheet("DLU")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("BS")
    Call DetailWorksheet("BS")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("PL")
    Call DetailWorksheet("PL")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("Factor")
    Call DetailWorksheet("FACT")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Call mUtility.ClearDetailReport("Misc")
    Call DetailWorksheet("MISC")
    
    Call mUtility.UpdateProgress(i, 12, ProgressBarDetailForm)
    i = i + 1
    
    Unload ProgressBarDetailForm
    Sheets("Cover Page").Range("E18") = Now()
    MsgBox "Detailed report generated"
End Sub

Private Sub PopulateSalesData()
    '--------------------------------------------------------------------------
    ' This subroutine gets the customer name from sales order data.
    '--------------------------------------------------------------------------
    Dim lngCount As Long
    Dim strItemRange As String
    Dim intFirstEmptyCell As Integer
    Dim rng As Range
    
    Dim itemNumber As Range
    
    Set wsData = Sheets("Data")
    Set wsReport = Sheets("SalesData")
    
    lngCount = Application.WorksheetFunction.CountA(wsData.Range("Table_Query_from_E1[IOLITM]"))
    strItemRange = "C2:C" & CStr(lngCount + 1)
    Set rng = wsData.Range(strItemRange)
        
    'disable screen updating
    Application.ScreenUpdating = False
    
    wsReport.Activate
    Range("A:A").Select
    selection.EntireRow.Delete
    intFirstCell = Range("A1").row
        
    For Each rngCurrentCel In rng.SpecialCells(xlCellTypeVisible)
        Set itemNumber = wsData.Range("A" & rngCurrentCel.row)
        If Not (mUtility.IsBulkItem(itemNumber) Or mUtility.IsPLItem(itemNumber) Or _
        mUtility.IsBSItem(itemNumber) Or mUtility.IsHouseItem(itemNumber)) Then
            wsReport.Range("A" & intFirstCell) = wsData.Range("A" & rngCurrentCel.row)
            wsReport.Range("B" & intFirstCell) = mDatabaseOperations.GetCustomerName(itemNumber.Value)
            intFirstCell = intFirstCell + 1
        ElseIf mUtility.IsHouseItem(itemNumber) Then
            wsReport.Range("A" & intFirstCell) = wsData.Range("A" & rngCurrentCel.row)
            wsReport.Range("B" & intFirstCell) = "House"
            intFirstCell = intFirstCell + 1
        ElseIf mUtility.IsBSItem(itemNumber) Then
            wsReport.Range("A" & intFirstCell) = wsData.Range("A" & rngCurrentCel.row)
            wsReport.Range("B" & intFirstCell) = "BS"
            intFirstCell = intFirstCell + 1
        ElseIf mUtility.IsPLItem(itemNumber) Then
            wsReport.Range("A" & intFirstCell) = wsData.Range("A" & rngCurrentCel.row)
            wsReport.Range("B" & intFirstCell) = "PL"
            intFirstCell = intFirstCell + 1
        End If
    Next rngCurrentCel
    
    Sheets("Report").Select
    'enable screen updating
    Application.ScreenUpdating = True
End Sub

Private Sub DetailWorksheet(category As String)
    '--------------------------------------------------------------------------
    ' This subroutine gets the data that can be used to generate
    ' report with 'category' items on separate worksheet.
    '--------------------------------------------------------------------------
    
    Dim lngCount As Long
    Dim strItemRange As String
    Dim dict As Scripting.dictionary
    Dim rng As Range

    Set wsData = Sheets("DataDetail")

    lngCount = Application.WorksheetFunction.CountA(wsData.Range("Table_Query_from_E13[IOLITM]"))
    strItemRange = "A2:A" & CStr(lngCount + 1)
    Set rng = wsData.Range(strItemRange)
    
    Set dict = New Scripting.dictionary
    
    'disable screen updating
    Application.ScreenUpdating = False

    For Each rngCurrentCel In rng.SpecialCells(xlCellTypeVisible)
        If (Not mUtility.IsBulkItem(wsData.Range("A" & rngCurrentCel.row))) And (mUtility.GetCustomerName(wsData.Range("A" & rngCurrentCel.row).Value) = category) Then
             If Not dict.Exists(rngCurrentCel.Value) Then
                Dim Item As New clsProductLotList
                Call Item.InitProductLotList(rngCurrentCel.Value, wsData.Range("B" & rngCurrentCel.row).Value)

                lot = wsData.Range("C" & rngCurrentCel.row)
                Set LotStatus = wsData.Range("D" & rngCurrentCel.row)
                Set locQty = wsData.Range("F" & rngCurrentCel.row)
                Set onHand = wsData.Range("E" & rngCurrentCel.row)
                Set UnitPrice = wsData.Range("I" & rngCurrentCel.row)
                Set loctn = wsData.Range("J" & rngCurrentCel.row)
                Set WrkOrder = wsData.Range("H" & rngCurrentCel.row)
                
                Call Item.InsertLot(CStr(lot), LotStatus, loctn, locQty, onHand, UnitPrice, WrkOrder)
                
                Item.AddQuantity (CLng(wsData.Range("F" & rngCurrentCel.row)))
                dict.Add rngCurrentCel.Value, Item
                Set Item = Nothing

             Else
                Dim getItem As New clsProductLotList
                Set getItem = dict(rngCurrentCel.Value)

                lot = wsData.Range("C" & rngCurrentCel.row)
                Set LotStatus = wsData.Range("D" & rngCurrentCel.row)
                Set locQty = wsData.Range("F" & rngCurrentCel.row)
                Set onHand = wsData.Range("E" & rngCurrentCel.row)
                Set UnitPrice = wsData.Range("I" & rngCurrentCel.row)
                Set loctn = wsData.Range("J" & rngCurrentCel.row)
                Set WrkOrder = wsData.Range("H" & rngCurrentCel.row)
                
                Call getItem.InsertLot(CStr(lot), LotStatus, loctn, locQty, onHand, UnitPrice, WrkOrder)
                
                getItem.AddQuantity (CLng(wsData.Range("F" & rngCurrentCel.row)))
                dict.Remove rngCurrentCel.Value
                dict.Add rngCurrentCel.Value, getItem
                Set getItem = Nothing
            End If
        End If

    Next rngCurrentCel
    
    If (category = "GOL") Then
        Call FormatSearchAndInsert(dict, "GOL")
    ElseIf (category = "SER") Then
        Call FormatSearchAndInsert(dict, "Seroyal")
    ElseIf (category = "TRO") Then
        Call FormatSearchAndInsert(dict, "Trophic")
    ElseIf (category = "DLC") Then
        Call FormatSearchAndInsert(dict, "DLC")
    ElseIf (category = "MCO") Then
        Call FormatSearchAndInsert(dict, "MCO")
    ElseIf (category = "IOV") Then
        Call FormatSearchAndInsert(dict, "Iovate")
    ElseIf (category = "DLU") Then
        Call FormatSearchAndInsert(dict, "House")
    ElseIf (category = "BS") Then
        Call FormatSearchAndInsert(dict, "BS")
    ElseIf (category = "FACT") Then
        Call FormatSearchAndInsert(dict, "Factor")
    ElseIf (category = "MISC") Then
        Call FormatSearchAndInsert(dict, "Misc")
    ElseIf (category = "PL") Then
        Call FormatSearchAndInsert(dict, "PL")
    End If
    
    'enable screen updating
    Application.ScreenUpdating = True
End Sub

Private Sub FormatSearchAndInsert(data As Scripting.dictionary, report As String)
    '--------------------------------------------------------------------------
    ' This subroutine inserts data in appropriate category and adds formating for each insertion.
    '--------------------------------------------------------------------------
    Dim intFirstEmptyCell As Integer
    Dim rng As Range
    
    Dim rngSearchResult As Range
    Dim product As Variant

    Set wsData = Sheets("Data")
    Set wsReport = Sheets(report)
    
    'disable screen updating
    Application.ScreenUpdating = False
    
    wsReport.Activate
    'OUTER LOOP = go throught products and get item
    'INNER LOOP = go through lots for the product
    For Each product In data.Keys
        Dim dict As ArrayList
        Dim getItems As New clsProductLotList
        Set getItems = data(product)
        
        Set dict = getItems.LotListArray

        Dim key As Variant
        
        intFirstEmptyCell = Range("B1").End(xlDown).row + 1
            
        mDetailGraphics.FormatHeadingDetail (intFirstEmptyCell)
        
        'item #
        If Len(getItems.itemNumber) > 13 Then
            wsReport.Range("A" & intFirstEmptyCell).WrapText = True
            wsReport.Range("A" & intFirstEmptyCell) = getItems.itemNumber
        Else
            wsReport.Range("A" & intFirstEmptyCell).WrapText = False
            wsReport.Range("A" & intFirstEmptyCell) = getItems.itemNumber
        End If
        
        'desc
        wsReport.Range("B" & intFirstEmptyCell) = getItems.Description
        'total qty F
        wsReport.Range("F" & intFirstEmptyCell) = getItems.TotalQuantity
        'total cost
        wsReport.Range("G" & intFirstEmptyCell) = getItems.TotalQuantity * getItems.UnitPrice
        
        For Each lot In dict
            Dim lotInfo As New clsProductLotInformation
            Dim z As Variant
            intFirstEmptyCell = Range("B1").End(xlDown).row + 1
            Set lotInfo = lot
            ' Lot bumber (A), location (B), lot status (C), Quantity (D), OnHandDate (E) , WrkOrder (F)
            wsReport.Range("A" & intFirstEmptyCell) = lotInfo.LotNumber
            wsReport.Range("B" & intFirstEmptyCell) = lotInfo.Location
            wsReport.Range("C" & intFirstEmptyCell) = lotInfo.LotStatus
            wsReport.Range("D" & intFirstEmptyCell) = lotInfo.LocationQuantity
            wsReport.Range("E" & intFirstEmptyCell) = lotInfo.OnHandDate
            wsReport.Range("F" & intFirstEmptyCell) = lotInfo.WrkOrder
            
            ' insert BULK information
            intFirstEmptyCell = mUtility.DetailInsertBulkInformation(intFirstEmptyCell, lotInfo.LotNumber)
            mDetailGraphics.AddLotSeparator (intFirstEmptyCell - 1)
        Next lot
        mDetailGraphics.DetailAddBottomBorders (intFirstEmptyCell - 1)
    Next product
    
    'enable screen updating
    Application.ScreenUpdating = True
End Sub
