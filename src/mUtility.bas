Attribute VB_Name = "mUtility"
 Function TInsertBulkInformation(curRow As Integer, strLot As String)
    ''' TEST METHOD
    Dim row As Integer
    Dim x As Integer, y As Integer
    Dim vBulkArray As Variant
    
    row = curRow + 1
    Call mGraphicInterface.AddBulkHeader(row)
    row = row + 1
    For x = 0 To 2
        Call mGraphicInterface.AddBulkText(row, "7515-9999", "50087656", M)
        Call mGraphicInterface.SetBackgroundBulk(row)
        row = row + 1
    Next x
    InsertBulkInformation = row
End Function
 
Sub InsertBulkInformation(curRow As Integer, strLot As String)
    '------------------------------------------------------------------------------
    ' This subroutine inserts information about bulk used to create lot.
    ' Used in main report.
    '------------------------------------------------------------------------------
    Dim row As Integer
    Dim x As Integer, y As Integer
    Dim vBulkArray As Variant
    
    row = curRow + 1
    vBulkArray = mDatabaseOperations.GetActiveIngredients(strLot)
    
    If (Not IsEmpty(vBulkArray)) Then
        Call mGraphicInterface.AddBulkHeader(row)
        row = row + 1
        
        For x = 0 To UBound(vBulkArray, 2)
            Call mGraphicInterface.AddBulkText(row, vBulkArray(0, x), _
            vBulkArray(1, x), vBulkArray(2, x))
            Call mGraphicInterface.SetBackgroundBulk(row)
            row = row + 1
        Next x
        Call mGraphicInterface.AddBottomBorders(row - 1)
    End If
End Sub

Function DetailInsertBulkInformation(curRow As Integer, strLot As String) As Integer
    '------------------------------------------------------------------------------
    ' This subroutine inserts information about bulk used to create lot.
    ' Used in detailed report.
    '------------------------------------------------------------------------------
    Dim row As Integer
    Dim x As Integer, y As Integer
    Dim vBulkArray As Variant
    
    row = curRow + 1
    vBulkArray = mDatabaseOperations.GetActiveIngredients(strLot)
    
    If (Not IsEmpty(vBulkArray)) Then
        Call mGraphicInterface.AddBulkHeader(row)
        row = row + 1
        
        For x = 0 To UBound(vBulkArray, 2)
            Call mGraphicInterface.AddBulkText(row, vBulkArray(0, x), _
            vBulkArray(1, x), vBulkArray(2, x))
            Call mGraphicInterface.SetBackgroundBulk(row)
            row = row + 1
        Next x
    End If
    DetailInsertBulkInformation = row
End Function

Sub ClearReport()
    '------------------------------------------------------------------------------
    ' This subroutine clears information from the main report.
    '------------------------------------------------------------------------------
    If MsgBox("This will erase ALL items from the Report worksheet! Are you sure?", _
    vbYesNo + vbExclamation, "Confirm") = vbNo Then Exit Sub
    Range("B6").Select
    Range(selection, selection.End(xlDown)).Select
    selection.EntireRow.Delete
End Sub

Sub ClearDetailReport(report As String)
    '------------------------------------------------------------------------------
    ' This subroutine clears information from the detailed report.
    '------------------------------------------------------------------------------
    Application.ScreenUpdating = False
    Set wsReport = Sheets(report)
    wsReport.Activate
    Range("A4:A1048576").Select
    selection.EntireRow.Delete
    Application.ScreenUpdating = True
End Sub

Sub UpdateProgress(i As Integer, totalCount As Integer, form As UserForm)
    '------------------------------------------------------------------------------
    ' This subroutine updates progress for the progress bar.
    '------------------------------------------------------------------------------
    Dim pctCompl As Single
    pctCompl = Round(((i / totalCount) * 100), 0)
    form.Text.Caption = pctCompl & "% Completed"
    form.Bar.Width = pctCompl * 2
    DoEvents
End Sub

Function GetCustomerName(itemNumber As String) As String
    '------------------------------------------------------------------------------
    ' This function gets the customer name from SalesData worksheet.
    '------------------------------------------------------------------------------
    Set rngSearchResult = Sheets("SalesData").Cells.Find(What:=itemNumber, _
    LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngSearchResult Is Nothing Then
        'data is found
        rng = Sheets("SalesData").Range("B" & rngSearchResult.row)
        If StrComp("PGH TO DL CANADA", rng) = 0 Then
            GetCustomerName = "DLC"
        ElseIf StrComp("PGH TO GARDEN OF LIFE, LLC", rng) = 0 Then
            GetCustomerName = "GOL"
        ElseIf rng Like "*PGH TO MCO*" Then
            GetCustomerName = "MCO"
        ElseIf rng Like "*PGH TO SEROYAL*" Then
            GetCustomerName = "SER"
        ElseIf rng Like "*IOVATE HEALTH*" Then
            GetCustomerName = "IOV"
        ElseIf rng Like "*PGH TO TROPHIC*" Then
            GetCustomerName = "TRO"
        ElseIf rng Like "*FACTOR NUTRITION LABS  LLC*" Then
            GetCustomerName = "FACT"
        ElseIf rng Like "House" Then
            GetCustomerName = "DLU"
        ElseIf rng Like "BS" Then
            GetCustomerName = "BS"
        ElseIf rng Like "PL" Then
            GetCustomerName = "PL"
        Else
            GetCustomerName = "MISC"
        End If
    End If
End Function


'------------------------------------------------------------------------------
' Functions below used for determining product types
'------------------------------------------------------------------------------
    
Function IsBulkItem(rng As Range) As Boolean
    IsBulkItem = ((rng.Value Like "*-8888") Or (rng.Value Like "*-9999"))
End Function

Function IsMcoItem(rng As Range) As Boolean
    IsMcoItem = (rng.Value Like "*MCO")
End Function

Function IsDlcItem(rng As Range) As Boolean
    IsDlcItem = (rng.Value Like "*HYC")
End Function

Function IsPLItem(rng As Range) As Boolean
    IsPLItem = (rng.Value Like "*-PL")
End Function

Function IsBSItem(rng As Range) As Boolean
    IsBSItem = (rng.Value Like "*-BS")
End Function

Function IsHouseItem(rng As Range) As Boolean
    IsHouseItem = (rng.Value Like "*X")
End Function

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
