Attribute VB_Name = "mMainReport"
#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Sub Execute_Options()
    Dim selection As Integer
    selection = Sheets("Report").ComboBox1.ListIndex
    Select Case selection
        Case 0
            ProgressBarPopulateForm.Show
        Case 1
            ProgressBarWoForm.Show
        Case 2
            ProgressBarLotForm.Show
        Case 3
            ProgressBarDetailForm.Show
    End Select
End Sub

Sub UpdateReport()
    '--------------------------------------------------------------------------
    '   This subroutine gets the filtered data and puts it into
    '   a more readable format on a Report page. Any new data is
    '   placed in the first empty cell to prevent overwriting data,
    '   so that the user notes/modifications for existing data are kept.
    '--------------------------------------------------------------------------
    Dim lngCount As Long
    Dim strItemRange As String
    Dim intFirstEmptyCell As Integer
    Dim rng As Range
    
    Set wsData = Sheets("Data")
    Set wsReport = Sheets("Report")
    'progress bar
    Dim i As Integer
    i = 1
    
    lngCount = Application.WorksheetFunction.CountA(wsData.Range("Table_Query_from_E1[IOLITM]"))
    strItemRange = "C2:C" & CStr(lngCount + 5)
    Set rng = wsData.Range(strItemRange)
    
    'disable screen updating
    Application.ScreenUpdating = False
    wsReport.Activate
    For Each rngCurrentCel In rng.SpecialCells(xlCellTypeVisible)
        Set rngSearchResult = Cells.Find(What:=rngCurrentCel, LookIn:=xlValues, LookAt:=xlWhole)
        If Not rngSearchResult Is Nothing Then
            'data is found, updating lot status, work order status and quantity
            With wsReport
                .Range("D" & rngSearchResult.row) = wsData.Range("D" & rngCurrentCel.row)
                .Range("H" & rngSearchResult.row) = wsData.Range("H" & rngCurrentCel.row)
                .Range("F" & rngSearchResult.row) = wsData.Range("F" & rngCurrentCel.row)
            End With
        Else
            If Not mUtility.IsBulkItem(wsData.Range("A" & rngCurrentCel.row)) Then
                'data not found, paste into first empty row
                intFirstEmptyCell = Range("B1").End(xlDown).row + 1
                ' Work order status is 'N/A', just do regular insertion
                If IsEmpty(wsData.Range("H" & rngCurrentCel.row)) Then
                    wsData.Range("A" & rngCurrentCel.row & ":" & "H" & rngCurrentCel.row).SpecialCells(xlCellTypeVisible).Copy
                    wsReport.Range("A" & intFirstEmptyCell & ":" & "H" & intFirstEmptyCell).PasteSpecial Paste:=xlPasteValues
                    wsReport.Range("I" & intFirstEmptyCell) = (wsData.Range("I" & rngCurrentCel.row) * wsData.Range("F" & rngCurrentCel.row))
                Else
                ' If work order status is NOT 'N/A' do regular insertion then list all bulks
                    wsData.Range("A" & rngCurrentCel.row & ":" & "H" & rngCurrentCel.row).SpecialCells(xlCellTypeVisible).Copy
                    wsReport.Range("A" & intFirstEmptyCell & ":" & "H" & intFirstEmptyCell).PasteSpecial Paste:=xlPasteValues
                    wsReport.Range("I" & intFirstEmptyCell) = (wsData.Range("I" & rngCurrentCel.row) * wsData.Range("F" & rngCurrentCel.row))
                    Call mUtility.InsertBulkInformation(intFirstEmptyCell, rngCurrentCel.Value)
                End If
            End If
        End If
        'update progress bar
        Call mUtility.UpdateProgress(i, rng.SpecialCells(xlCellTypeVisible).Count, ProgressBarPopulateForm)
        i = i + 1
    Next rngCurrentCel
    
    'enable screen updating
    Application.ScreenUpdating = True
    Unload ProgressBarPopulateForm
    
    'update date and time
    wsReport.OLEObjects("LastUpdateLbl").Object.Caption = Now()
    OutPut = MsgBox("Updating Complete", vbInformation, "Update")
End Sub

Sub Highlight()
    '--------------------------------------------------------------------------
    '   This subroutine highlights the cells in the Report worksheet that
    '   are no longer in the Data worksheet because their status was changed.
    '--------------------------------------------------------------------------
    Dim lngCount As Long
    Dim strLotRange As String
    Dim rng As Range
    
    Set wsData = Sheets("Data")
    Set wsReport = Sheets("Report")

    'progress bar
    Dim i As Integer
    i = 1

    lngCount = Application.WorksheetFunction.CountA(wsReport.Range("Report!C:C"))
    strLotRange = "C6:C" & CStr(lngCount)
    Set rng = wsReport.Range(strLotRange)
    
    For Each rngCurrentCel In rng.SpecialCells(xlCellTypeVisible)
        If rngCurrentCel.Interior.ColorIndex = xlNone Then
            Set rngSearchResult = wsData.Cells.Find(What:=rngCurrentCel.Value, LookIn:=xlValues, LookAt:=xlWhole)
            If rngSearchResult Is Nothing Then
                'data not found, highlight row in Report worksheet
                LotStatus = mDatabaseOperations.Get_Lot_Status(rngCurrentCel.Value)
                If LotStatus = "blank" Then
                    wsReport.Range("A" & rngCurrentCel.row & ":" & "I" & rngCurrentCel.row).Interior.Color = RGB(175, 220, 126)
                    wsReport.Range("D" & rngCurrentCel.row) = LotStatus
                Else
                    wsReport.Range("A" & rngCurrentCel.row & ":" & "I" & rngCurrentCel.row).Interior.Color = 13434879
                    wsReport.Range("D" & rngCurrentCel.row) = LotStatus
                End If
            End If
        End If
        
        'update progress bar
        Call mUtility.UpdateProgress(i, rng.SpecialCells(xlCellTypeVisible).Count, ProgressBarLotForm)
        i = i + 1
    Next rngCurrentCel
    
    Unload ProgressBarLotForm
    wsReport.OLEObjects("LastUpdateLbl").Object.Caption = Now()
    OutPut = MsgBox("Cells highlighted.", vbInformation, "Update")
End Sub

Sub Get_WO_Status()
    '--------------------------------------------------------------------------
    '   This subroutine tries to get the work order status for each lot number,
    '   otherwise N/A is placed for W/O status.
    '--------------------------------------------------------------------------
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim adoRecordset As ADODB.Recordset
    
    Dim lngCount As Long
    Dim strLotRange As String
    Dim rng As Range
    
    'progress bar
    Dim i As Integer
    i = 1
    
    Set wsReport = Sheets("Report")
    
    lngCount = Application.WorksheetFunction.CountA(wsReport.Range("Report!C:C"))
    strLotRange = "C6:C" & CStr(lngCount)
    
    Set rng = wsReport.Range(strLotRange)
    
    Set adoConnection = New ADODB.Connection
    adoConnection.Open "E1", "JDEREAD", "JDEREAD1"
    
   'Find out if the connection attempt worked worked.
   If adoConnection.State = adStateOpen Then
   
        For Each rngCurrentCel In rng.SpecialCells(xlCellTypeVisible)
            If rngCurrentCel.Interior.ColorIndex = xlNone Then
                Set adoRecordset = New ADODB.Recordset
                Set adoCommand = New ADODB.Command
                With adoCommand
                    'set active connection for adoCommand
                    .ActiveConnection = adoConnection
                    'set adoCommand to SQL statement
                    'using prepared statement ? to bind parameters
                    .CommandText = "SELECT F4801.WASRST FROM ATJDENT1.PRODDTA.F4801 F4801 WHERE F4801.WADOCO=?"
                    'in this case adoCommand type is adCmdText
                    .CommandType = adCmdText
                    'setting parameter options
                    Set adoParameter = .CreateParameter("workOrder", adBSTR, adParamInput)
                    'parameter value
                    adoParameter.Value = rngCurrentCel.Value
                    'bind the parameter for prepared statement
                    .Parameters.Append adoParameter
                    'executes the command and store results in record set
                    Set adoRecordset = .Execute
                End With
    
                If adoRecordset.EOF Then
                    'recordset is empty, so no W/O status found
                    wsReport.Range("H" & rngCurrentCel.row) = "N/A"
                Else
                    wsReport.Range("H" & rngCurrentCel.row) = adoRecordset.Fields(0).Value
                End If
                
                'must close recordset and set to nothing for next loop iteration
                adoRecordset.Close
                Set adoRecordset = Nothing
              End If
              'update progress bar
              Call mUtility.UpdateProgress(i, rng.SpecialCells(xlCellTypeVisible).Count, ProgressBarWoForm)
              i = i + 1
        Next rngCurrentCel
        
   Else
      MsgBox "Sorry. Can't connect to JDEREAD."
   End If
   
   Unload ProgressBarWoForm
   'Close the connection.
   adoConnection.Close
   wsReport.OLEObjects("LastUpdateLbl").Object.Caption = Now()
   OutPut = MsgBox("W/O status updated. Connection is closed.", vbInformation, "Update")
End Sub
