Attribute VB_Name = "ModDataExtraction"
' ============================================================================
' MODULE: ModDataExtraction
' Load data from worksheet, handle rates vs. cumulative values
' ============================================================================

Option Explicit

' Data extraction result type
Public Type DataExtractionResult
    Success As Boolean
    ErrorMessage As String
    DataArray() As Double
    ObservationCount As Long
    ContainsEmpty As Boolean
End Type

' ============================================================================
' PRIMARY DATA EXTRACTION
' ============================================================================

Public Function ExtractDataRange(ws As Worksheet, col As Long, startRow As Long, _
                                 endRow As Long) As DataExtractionResult
    ' Load a range of data from the worksheet into an array
    ' Handle empty cells and non-numeric values
    
    Dim i As Long, idx As Long
    Dim cellValue As Variant
    Dim tempArray() As Double
    Dim tempCount As Long
    
    With ExtractDataRange
        .Success = True
        .ErrorMessage = ""
        .ContainsEmpty = False
        
        ' Validate inputs
        If startRow >= endRow Then
            .Success = False
            .ErrorMessage = "Start row must be less than end row."
            Exit Function
        End If
        
        ' Allocate temporary array
        tempCount = endRow - startRow + 1
        ReDim tempArray(1 To tempCount)
        
        ' Load data
        For i = startRow To endRow
            cellValue = ws.Cells(i, col).Value2
            idx = i - startRow + 1
            
            If IsEmpty(cellValue) Then
                .ContainsEmpty = True
                tempArray(idx) = 0  ' Treat as zero
            ElseIf IsError(cellValue) Then
                .Success = False
                .ErrorMessage = "Column contains error at row " & i & "."
                Exit Function
            ElseIf IsNumericValue(cellValue) Then
                tempArray(idx) = CDbl(cellValue)
            Else
                .Success = False
                .ErrorMessage = "Column contains non-numeric value '" & CStr(cellValue) & "' at row " & i & "."
                Exit Function
            End If
        Next i
        
        ' Copy to result
        .DataArray = tempArray
        .ObservationCount = tempCount
    End With
    
End Function

' ============================================================================
' CUMULATIVE TO RATE CONVERSION
' ============================================================================

Public Function ConvertCumulativeToRates(cumulativeData() As Double) As DataExtractionResult
    ' Convert cumulative values to 15-minute rate increments
    ' Rate(i) = Cumulative(i) - Cumulative(i-1)
    
    Dim i As Long
    Dim rateArray() As Double
    Dim n As Long
    
    With ConvertCumulativeToRates
        .Success = True
        .ErrorMessage = ""
        .ContainsEmpty = False
        
        n = UBound(cumulativeData) - LBound(cumulativeData) + 1
        
        If n < 2 Then
            .Success = False
            .ErrorMessage = "Cumulative data must have at least 2 observations to calculate rates."
            Exit Function
        End If
        
        ' Allocate rate array (one fewer than cumulative)
        ReDim rateArray(1 To n - 1)
        
        ' Calculate increments
        For i = 1 To n - 1
            rateArray(i) = cumulativeData(LBound(cumulativeData) + i) - _
                          cumulativeData(LBound(cumulativeData) + i - 1)
        Next i
        
        .DataArray = rateArray
        .ObservationCount = UBound(rateArray) - LBound(rateArray) + 1
    End With
    
End Function

Public Function ConvertRatesToCumulative(rateData() As Double, startingValue As Double) As Double()
    ' Convert rate increments to cumulative values
    ' Cumulative(i) = StartingValue + Sum(Rate(1..i))
    
    Dim cumulativeArray() As Double
    Dim i As Long
    Dim n As Long
    
    n = UBound(rateData) - LBound(rateData) + 1
    ReDim cumulativeArray(1 To n)
    
    cumulativeArray(1) = startingValue + rateData(LBound(rateData))
    
    For i = 2 To n
        cumulativeArray(i) = cumulativeArray(i - 1) + rateData(LBound(rateData) + i - 1)
    Next i
    
    ConvertRatesToCumulative = cumulativeArray
    
End Function

' ============================================================================
' DATA VALIDATION AND CLEANING
' ============================================================================

Public Function ValidateDataArray(dataArray() As Double) As DataExtractionResult
    ' Check that data array has reasonable properties
    
    Dim i As Long
    Dim n As Long
    Dim hasNegative As Boolean, hasZero As Boolean, hasPositive As Boolean
    Dim minVal As Double, maxVal As Double
    
    With ValidateDataArray
        .Success = True
        .ErrorMessage = ""
        .ContainsEmpty = False
        .ObservationCount = UBound(dataArray) - LBound(dataArray) + 1
        
        If .ObservationCount < 2 Then
            .Success = False
            .ErrorMessage = "Data array must contain at least 2 observations."
            Exit Function
        End If
        
        ' Analyze data
        minVal = dataArray(LBound(dataArray))
        maxVal = dataArray(LBound(dataArray))
        
        For i = LBound(dataArray) To UBound(dataArray)
            If dataArray(i) < 0 Then hasNegative = True
            If dataArray(i) = 0 Then hasZero = True
            If dataArray(i) > 0 Then hasPositive = True
            
            If dataArray(i) < minVal Then minVal = dataArray(i)
            If dataArray(i) > maxVal Then maxVal = dataArray(i)
        Next i
        
        ' Warn about potentially problematic data
        If hasNegative And hasPositive Then
            .ErrorMessage = "Data contains both positive and negative values. " & _
                          "This may indicate mixed directions or errors. "
        End If
        
        If minVal = maxVal Then
            .ErrorMessage = .ErrorMessage & "All data values are identical. " & _
                          "Historical variation analysis will be limited."
        End If
        
    End With
    
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function IsNumericValue(val As Variant) As Boolean
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
