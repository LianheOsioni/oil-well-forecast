Attribute VB_Name = "ModValidation"
' ============================================================================
' MODULE: ModValidation
' Input validation, range checking, and error handling
' ============================================================================

Option Explicit

Private Const MAX_COLUMN_INDEX As Long = 16384
Private Const MAX_ROW_INDEX As Long = 1048576

' Result type for validation operations
Public Type ValidationResult
    IsValid As Boolean
    ErrorMessage As String
    ErrorCode As Long
End Type

' ============================================================================
' COLUMN VALIDATION
' ============================================================================

Public Function ValidateColumnLetter(colLetter As String) As Boolean
    ' Validate that a string is a valid Excel column letter (A, B, ..., XFD)
    
    Dim i As Long, col As Long
    
    colLetter = UCase(Trim(colLetter))
    
    ' Check length (1-3 characters)
    If Len(colLetter) < 1 Or Len(colLetter) > 3 Then
        ValidateColumnLetter = False
        Exit Function
    End If
    
    ' Check that all characters are uppercase letters
    For i = 1 To Len(colLetter)
        If Not (Mid(colLetter, i, 1) >= "A" And Mid(colLetter, i, 1) <= "Z") Then
            ValidateColumnLetter = False
            Exit Function
        End If
    Next i
    
    ' Convert to column index and check range
    col = ColumnLetterToIndex(colLetter)
    ValidateColumnLetter = (col >= 1 And col <= MAX_COLUMN_INDEX)
    
End Function

Public Function ValidateColumnIndex(col As Long) As Boolean
    ' Validate that a numeric column index is in valid range
    ValidateColumnIndex = (col >= 1 And col <= MAX_COLUMN_INDEX)
End Function

' ============================================================================
' ROW VALIDATION
' ============================================================================

Public Function ValidateRowNumber(rowNum As Variant) As Boolean
    ' Check that a value is a valid row number (positive integer within limits)
    
    If VarType(rowNum) = vbBoolean Then
        ValidateRowNumber = False
        Exit Function
    End If
    
    On Error Resume Next
    Dim rNum As Long
    rNum = CLng(rowNum)
    On Error GoTo 0
    
    ValidateRowNumber = (rNum >= 1 And rNum <= MAX_ROW_INDEX)
    
End Function

Public Function ValidateLearningRange(startRow As Long, endRow As Long, minObservations As Long) As ValidationResult
    ' Validate that learning range has sufficient observations
    
    With ValidateLearningRange
        .IsValid = True
        .ErrorMessage = ""
        .ErrorCode = 0
        
        If startRow < 1 Then
            .IsValid = False
            .ErrorMessage = "Learning start row must be at least 1."
            .ErrorCode = 1001
            Exit Function
        End If
        
        If endRow < startRow Then
            .IsValid = False
            .ErrorMessage = "Learning end row must be after start row."
            .ErrorCode = 1002
            Exit Function
        End If
        
        If (endRow - startRow + 1) < minObservations Then
            .IsValid = False
            .ErrorMessage = "Learning range has only " & (endRow - startRow + 1) & " observations. " & _
                          "Minimum required: " & minObservations & "."
            .ErrorCode = 1003
        End If
    End With
    
End Function

' ============================================================================
' NUMERIC DATA VALIDATION
' ============================================================================

Public Function ValidateNumericRange(ws As Worksheet, col As Long, _
                                     startRow As Long, endRow As Long, _
                                     Optional allowEmpty As Boolean = False) As ValidationResult
    ' Check that a range contains only numeric or empty cells
    
    Dim i As Long, nonNumericCount As Long, emptyCount As Long
    Dim cellValue As Variant
    
    With ValidateNumericRange
        .IsValid = True
        .ErrorMessage = ""
        .ErrorCode = 0
        
        For i = startRow To endRow
            cellValue = ws.Cells(i, col).Value2
            
            If IsEmpty(cellValue) Then
                emptyCount = emptyCount + 1
                If Not allowEmpty Then
                    .IsValid = False
                    .ErrorMessage = "Range contains empty cell at row " & i & "."
                    .ErrorCode = 2001
                    Exit Function
                End If
            ElseIf IsError(cellValue) Then
                .IsValid = False
                .ErrorMessage = "Range contains error at row " & i & "."
                .ErrorCode = 2002
                Exit Function
            ElseIf Not IsNumericValue(cellValue) Then
                nonNumericCount = nonNumericCount + 1
                .IsValid = False
                .ErrorMessage = "Range contains non-numeric value '" & CStr(cellValue) & "' at row " & i & "."
                .ErrorCode = 2003
                Exit Function
            End If
        Next i
    End With
    
End Function

Public Function IsNumericValue(val As Variant) As Boolean
    ' Check if a value is numeric
    
    If IsEmpty(val) Then
        IsNumericValue = False
        Exit Function
    End If
    
    If IsError(val) Then
        IsNumericValue = False
        Exit Function
    End If
    
    ' Try to convert to double
    Dim testDbl As Double
    On Error Resume Next
    testDbl = CDbl(val)
    IsNumericValue = (Err.Number = 0)
    On Error GoTo 0
    
End Function

Public Function GetNumericValue(val As Variant, ByRef numVal As Double) As Boolean
    ' Try to extract a numeric value from a variant
    
    If Not IsNumericValue(val) Then
        GetNumericValue = False
        Exit Function
    End If
    
    On Error Resume Next
    numVal = CDbl(val)
    On Error GoTo 0
    
    GetNumericValue = True
    
End Function

' ============================================================================
' RANGE OVERLAP DETECTION
' ============================================================================

Public Function DetectRangeOverlap(col1 As Long, start1 As Long, end1 As Long, _
                                   col2 As Long, start2 As Long, end2 As Long) As Boolean
    ' Check if two ranges overlap (same column and overlapping rows)
    
    If col1 <> col2 Then
        DetectRangeOverlap = False
        Exit Function
    End If
    
    ' Check for row overlap
    DetectRangeOverlap = Not ((end1 < start2) Or (end2 < start1))
    
End Function

Public Function CheckOutputOverwrite(ws As Worksheet, outputCol As Long, outputStart As Long, _
                                     outputEnd As Long, learningEnd As Long) As ValidationResult
    ' Check if output range contains existing data and warn if it overlaps learning data
    
    Dim lastRow As Long
    
    With CheckOutputOverwrite
        .IsValid = True
        .ErrorMessage = ""
        .ErrorCode = 0
        
        ' Check if output overlaps learning data
        If outputStart <= learningEnd Then
            .IsValid = False
            .ErrorMessage = "Output start row (" & outputStart & ") is within or before learning range end (" & _
                          learningEnd & "). This would overwrite historical data."
            .ErrorCode = 3001
            Exit Function
        End If
        
        ' Check if output range contains existing data
        lastRow = ws.Cells(ws.Rows.Count, outputCol).End(xlUp).Row
        
        If lastRow > outputEnd Then
            ' Data exists beyond the output range (not our concern here)
        ElseIf lastRow >= outputStart Then
            ' Data exists in the output range
            .ErrorMessage = "Output range " & ColumnIndexToLetter(outputCol) & outputStart & ":" & _
                          ColumnIndexToLetter(outputCol) & outputEnd & " contains " & _
                          (lastRow - outputStart + 1) & " existing values."
            .ErrorCode = 3002
            ' Don't set IsValid = False here; just warn the user
        End If
    End With
    
End Function

' ============================================================================
' TARGET VALIDATION
' ============================================================================

Public Function ValidateTarget(targetValue As Variant, currentCumulative As Double, _
                               mean As Double, stdev As Double, numRows As Long) As ValidationResult
    ' Validate that a target is numerically feasible
    
    Dim maxPossible As Double, minPossible As Double
    Dim targetDbl As Double
    
    With ValidateTarget
        .IsValid = True
        .ErrorMessage = ""
        .ErrorCode = 0
        
        ' Check if target is numeric
        If Not IsNumericValue(targetValue) Then
            .IsValid = False
            .ErrorMessage = "Target value must be numeric."
            .ErrorCode = 4001
            Exit Function
        End If
        
        targetDbl = CDbl(targetValue)
        
        ' Check if target is physically achievable
        maxPossible = currentCumulative + numRows * (mean + 2 * stdev)
        minPossible = currentCumulative + numRows * (mean - 2 * stdev)
        
        If targetDbl > maxPossible Then
            .ErrorMessage = "Target " & Format(targetDbl, "0.00") & " may be unrealistic. " & _
                          "Maximum achievable (with normal volatility): " & Format(maxPossible, "0.00") & "."
            .ErrorCode = 4002
            ' Not fatal; proceed with warning
        End If
        
        If targetDbl < minPossible Then
            .ErrorMessage = "Target " & Format(targetDbl, "0.00") & " may be unrealistic. " & _
                          "Minimum achievable (with normal volatility): " & Format(minPossible, "0.00") & "."
            .ErrorCode = 4003
            ' Not fatal; proceed with warning
        End If
    End With
    
End Function

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

Public Function ColumnLetterToIndex(colLetter As String) As Long
    ' Convert column letter (e.g., "AB") to numeric column index (e.g., 28)
    Dim i As Long, col As Long
    
    colLetter = UCase(Trim(colLetter))
    col = 0
    
    For i = 1 To Len(colLetter)
        col = col * 26 + (Asc(Mid(colLetter, i, 1)) - Asc("A") + 1)
    Next i
    
    ColumnLetterToIndex = col
End Function

Public Function ColumnIndexToLetter(col As Long) As String
    ' Convert numeric column index to letter(s)
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
