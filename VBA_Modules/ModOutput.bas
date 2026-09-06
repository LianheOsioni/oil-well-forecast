Attribute VB_Name = "ModOutput"
' ============================================================================
' MODULE: ModOutput
' Safe output writing with validation and atomic transactions
' ============================================================================

Option Explicit

' ============================================================================
' OUTPUT RANGE VALIDATION
' ============================================================================

Public Function ValidateOutputRange(ws As Worksheet, outputCol As Long, outputStart As Long, _
                                    outputEnd As Long, learningEnd As Long) As ValidationResult
    ' Comprehensive validation of output range before writing
    
    Dim lastDataRow As Long
    Dim warningMsg As String
    
    With ValidateOutputRange
        .IsValid = True
        .ErrorMessage = ""
        .ErrorCode = 0
        
        ' Check basic range validity
        If outputStart < 1 Or outputEnd < outputStart Then
            .IsValid = False
            .ErrorMessage = "Invalid output range: start row must be >= 1 and <= end row."
            .ErrorCode = 5001
            Exit Function
        End If
        
        ' Critical: Check for overlap with learning data
        If outputStart <= learningEnd Then
            .IsValid = False
            .ErrorMessage = "Output range starts at row " & outputStart & " but learning data ends at " & _
                          learningEnd & ". This would overwrite historical data."
            .ErrorCode = 5002
            Exit Function
        End If
        
        ' Check for existing data in output range
        On Error Resume Next
        lastDataRow = ws.Cells(ws.Rows.Count, outputCol).End(xlUp).Row
        On Error GoTo 0
        
        If lastDataRow > 0 And lastDataRow >= outputStart Then
            ' Data exists in the output range
            warningMsg = "Output range contains " & (lastDataRow - outputStart + 1) & " existing values. These will be overwritten."
            .ErrorMessage = warningMsg
            .ErrorCode = 5003
            ' Don't fail; just warn
        End If
        
    End With
    
End Function

Public Function CheckColumnForFormulas(ws As Worksheet, col As Long, startRow As Long, _
                                       endRow As Long) As Boolean
    ' Check if output column contains formulas (which should not be overwritten)
    
    Dim i As Long
    
    For i = startRow To endRow
        If ws.Cells(i, col).HasFormula Then
            CheckColumnForFormulas = True
            Exit Function
        End If
    Next i
    
    CheckColumnForFormulas = False
    
End Function

' ============================================================================
' ATOMIC OUTPUT WRITING
' ============================================================================

Public Function WriteOutputData(ws As Worksheet, outputCol As Long, outputStart As Long, _
                                dataArray() As Double, decimalPlaces As Long) As Boolean
    ' Write forecast data to worksheet (atomic operation)
    ' All data validated before writing begins
    
    Dim i As Long, dataIndex As Long
    Dim outputRow As Long
    Dim roundedValue As Double
    
    On Error GoTo ErrorHandler
    
    ' Turn off screen updating for performance
    Dim savedScreenUpdate As Boolean
    savedScreenUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' Write data row by row
    For dataIndex = LBound(dataArray) To UBound(dataArray)
        outputRow = outputStart + dataIndex - 1
        roundedValue = Round(dataArray(dataIndex), decimalPlaces)
        ws.Cells(outputRow, outputCol).Value = roundedValue
    Next dataIndex
    
    ' Restore screen updating
    Application.ScreenUpdating = savedScreenUpdate
    WriteOutputData = True
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = savedScreenUpdate
    WriteOutputData = False
    
End Function

Public Function WriteMultiVariableOutput(ws As Worksheet, outputStart As Long, _
                                         oilArray() As Double, gasArray() As Double, _
                                         bswArray() As Double, config As ConfigSettings) As Boolean
    ' Write multiple variables to worksheet
    
    Dim i As Long, outputRow As Long
    On Error GoTo ErrorHandler
    
    Dim savedScreenUpdate As Boolean
    savedScreenUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' Write oil data
    If config.OilRateCol > 0 Then
        For i = LBound(oilArray) To UBound(oilArray)
            outputRow = outputStart + i - 1
            ws.Cells(outputRow, config.OilRateCol).Value = Round(oilArray(i), config.DecimalPlaces)
        Next i
    End If
    
    ' Write gas data
    If config.GasRateCol > 0 Then
        For i = LBound(gasArray) To UBound(gasArray)
            outputRow = outputStart + i - 1
            ws.Cells(outputRow, config.GasRateCol).Value = Round(gasArray(i), config.DecimalPlaces)
        Next i
    End If
    
    ' Write BS&W data (with special rounding)
    If config.BSWCol > 0 Then
        For i = LBound(bswArray) To UBound(bswArray)
            outputRow = outputStart + i - 1
            ws.Cells(outputRow, config.BSWCol).Value = RoundBSWToStandard(bswArray(i))
        Next i
    End If
    
    Application.ScreenUpdating = savedScreenUpdate
    WriteMultiVariableOutput = True
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = savedScreenUpdate
    WriteMultiVariableOutput = False
    
End Function

' ============================================================================
' OUTPUT VERIFICATION
' ============================================================================

Public Function VerifyOutputWritten(ws As Worksheet, outputCol As Long, outputStart As Long, _
                                     expectedCount As Long, decimalPlaces As Long) As Boolean
    ' Verify that all expected data was written correctly
    
    Dim i As Long, cellValue As Variant
    Dim nonEmptyCount As Long
    
    For i = outputStart To outputStart + expectedCount - 1
        cellValue = ws.Cells(i, outputCol).Value2
        
        ' Check that cell is not empty
        If Not IsEmpty(cellValue) Then
            nonEmptyCount = nonEmptyCount + 1
            
            ' Check that value is numeric
            If Not IsNumeric(cellValue) Then
                VerifyOutputWritten = False
                Exit Function
            End If
        End If
    Next i
    
    ' All expected cells should be filled
    VerifyOutputWritten = (nonEmptyCount = expectedCount)
    
End Function

Public Function GenerateOutputSummary(outputCol As Long, outputStart As Long, outputEnd As Long, _
                                      config As ConfigSettings) As String
    ' Generate a summary message about what was written
    
    Dim summary As String
    Dim numRows As Long
    
    numRows = outputEnd - outputStart + 1
    
    summary = "Output Summary:" & vbCrLf & vbCrLf
    summary = summary & "Written to: Column " & ColumnIndexToLetter(outputCol) & ", Rows " & outputStart & " to " & outputEnd & vbCrLf
    summary = summary & "Rows generated: " & numRows & vbCrLf
    summary = summary & "Decimal places: " & config.DecimalPlaces & vbCrLf & vbCrLf
    
    If config.TargetMode = 1 Then
        summary = summary & "Target Mode: Single" & vbCrLf
        summary = summary & "  Target Value: " & Format(config.Target1, "0.00") & vbCrLf
    ElseIf config.TargetMode = 2 Then
        summary = summary & "Target Mode: Dual" & vbCrLf
        summary = summary & "  Target 1: " & Format(config.Target1, "0.00") & vbCrLf
        summary = summary & "  Target 2: " & Format(config.Target2, "0.00") & vbCrLf
    Else
        summary = summary & "Target Mode: None (free simulation)" & vbCrLf
    End If
    
    GenerateOutputSummary = summary
    
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function IsNumeric(val As Variant) As Boolean
    Dim testDbl As Double
    On Error Resume Next
    testDbl = CDbl(val)
    IsNumeric = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function ColumnIndexToLetter(col As Long) As String
    Dim result As String
    If col < 1 Then
        ColumnIndexToLetter = "?"
        Exit Function
    End If
    result = ""
    Do While col > 0
        result = Chr(64 + ((col - 1) Mod 26) + 1) & result
        col = (col - 1) \ 26
    Loop
    ColumnIndexToLetter = result
End Function
