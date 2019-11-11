Attribute VB_Name = "mDatabaseOperations"
Function Get_Lot_Status(lot As String) As String
    '--------------------------------------------------------------------------
    '   This function gets the lot status from JDE database.
    '--------------------------------------------------------------------------
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim adoRecordset As ADODB.Recordset
        
    Dim strLotStatus As String
    
    Set adoConnection = New ADODB.Connection
    adoConnection.Open "E1", "JDEREAD", "JDEREAD1"
    
    If adoConnection.State = adStateOpen Then
        Set adoRecordset = New ADODB.Recordset
        Set adoCommand = New ADODB.Command
        
        With adoCommand
            .ActiveConnection = adoConnection
            .CommandText = "SELECT F4108.IOLOTS FROM ATJDENT1.PRODDTA.F4108 F4108 WHERE F4108.IOLOTN=?"
            .CommandType = adCmdText
            Set adoParameter = .CreateParameter("lotNumber", adBSTR, adParamInput)
            adoParameter.Value = lot
            .Parameters.Append adoParameter
            Set adoRecordset = .Execute
        End With
        
        If adoRecordset.EOF Then
            'recordset is empty, so no W/O status found
        Else
            If adoRecordset.Fields(0) = " " Then
                strLotStatus = "blank"
            Else
                strLotStatus = adoRecordset.Fields(0).Value
            End If
        End If
    Else
      MsgBox "Sorry. Can't connect to JDEREAD."
    End If
    
    adoRecordset.Close
    Set adoRecordset = Nothing
    adoConnection.Close
    
    Get_Lot_Status = strLotStatus
End Function

Function GetActiveIngredients(strWoNumber As String)
    '--------------------------------------------------------------------------
    '   This function gets active ingredient. Bulk, powders, etc.
    '--------------------------------------------------------------------------
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim adoRecordset As ADODB.Recordset
        
    Dim strPartsRow As String
    
    Dim vBulkArray As Variant
    
    Set adoConnection = New ADODB.Connection
    adoConnection.Open "E1", "JDEREAD", "JDEREAD1"
    
    If adoConnection.State = adStateOpen Then
        Set adoRecordset = New ADODB.Recordset
        Set adoCommand = New ADODB.Command
        
        With adoCommand
            .ActiveConnection = adoConnection
            .CommandText = "SELECT DISTINCT wmcpil, wmlotn, iolots " & _
            "FROM atjdent1.proddta.f3111, atjdent1.proddta.f4108" & _
            " WHERE wmlotn = iolotn AND wmaing = 1 AND wmcpil = iolitm AND wmdoco=?"
            .CommandType = adCmdText
            Set adoParameter = .CreateParameter("woNumber", adBSTR, adParamInput)
            adoParameter.Value = strWoNumber
            .Parameters.Append adoParameter
            Set adoRecordset = .Execute
        End With
        
        If adoRecordset.EOF Then
        Else
            vBulkArray = adoRecordset.GetRows()
        End If
    Else
      MsgBox "Sorry. Can't connect to JDEREAD."
    End If
    
    adoRecordset.Close
    Set adoRecordset = Nothing
    adoConnection.Close
    
    GetActiveIngredients = vBulkArray
End Function

Function GetBulkItems2(strWoNumber As String)
    '--------------------------------------------------------------------------
    '   This function gets the bulk items from JDE database.
    '--------------------------------------------------------------------------
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim adoRecordset As ADODB.Recordset
        
    Dim strPartsRow As String
    
    Dim vBulkArray As Variant
    
    Set adoConnection = New ADODB.Connection
    adoConnection.Open "E1", "JDEREAD", "JDEREAD1"
    
    If adoConnection.State = adStateOpen Then
        Set adoRecordset = New ADODB.Recordset
        Set adoCommand = New ADODB.Command
        
        With adoCommand
            .ActiveConnection = adoConnection
            .CommandText = "SELECT DISTINCT wmcpil, wmlotn, iolots " & _
            "FROM atjdent1.proddta.f3111, atjdent1.proddta.f4108" & _
            " WHERE wmlotn = iolotn AND (wmcpil like '%-8888%' or wmcpil like '%-9999%') AND wmcpil = iolitm AND wmdoco=?"
            .CommandType = adCmdText
            Set adoParameter = .CreateParameter("woNumber", adBSTR, adParamInput)
            adoParameter.Value = strWoNumber
            .Parameters.Append adoParameter
            Set adoRecordset = .Execute
        End With
        
        If adoRecordset.EOF Then
         ' WORKING ON powders
         'MsgBox "Nada"
         'MsgBox strWoNumber
        Else
            vBulkArray = adoRecordset.GetRows()
        End If
    Else
      MsgBox "Sorry. Can't connect to JDEREAD."
    End If
    
    adoRecordset.Close
    Set adoRecordset = Nothing
    adoConnection.Close
    
    GetBulkItems2 = vBulkArray
End Function

Function GetCustomerName(itemNumber As String) As String
    '--------------------------------------------------------------------------
    '   This function gets the customer name from JDE database.
    '--------------------------------------------------------------------------
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    Dim adoRecordset As ADODB.Recordset
        
    Dim strCustomerName As String
    
    Set adoConnection = New ADODB.Connection
    adoConnection.Open "E1", "JDEREAD", "JDEREAD1"
    
    If adoConnection.State = adStateOpen Then
        Set adoRecordset = New ADODB.Recordset
        Set adoCommand = New ADODB.Command
        
        With adoCommand
            .ActiveConnection = adoConnection
            .CommandText = "SELECT wwalph FROM F4211 INNER JOIN f0111 ON f0111.wwan8 = f4211.sdan8 WHERE sdlitm = ? AND wwidln = 0 FETCH FIRST 1 ROWS ONLY"
            .CommandType = adCmdText
            Set adoParameter = .CreateParameter("itemNumber", adBSTR, adParamInput)
            adoParameter.Value = itemNumber
            .Parameters.Append adoParameter
            Set adoRecordset = .Execute
        End With
        
        If adoRecordset.EOF Then
            'recordset is empty, so no customer name found
        Else
            strCustomerName = Trim(adoRecordset.Fields(0).Value)
        End If
    Else
      MsgBox "Sorry. Can't connect to JDEREAD."
    End If
    
    adoRecordset.Close
    Set adoRecordset = Nothing
    adoConnection.Close
    
    GetCustomerName = strCustomerName
End Function
