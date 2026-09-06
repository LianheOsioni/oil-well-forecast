Attribute VB_Name = "ModBSW"
' ============================================================================
' MODULE: ModBSW
' BS&W (Basic Sediment & Water) specific handling and rounding
' ============================================================================

Option Explicit

' BS&W rounding standard
' From 0 to <0.1: increments of 0.025
' From 0.1 to <1.0: increments of 0.05
' From 1.0 to 10.0: increments of 1.0
' Above 10.0: treated as error or maximum

Private Const BSW_TIER1_MAX As Double = 0.1
Private Const BSW_TIER2_MAX As Double = 1.0
Private Const BSW_TIER3_MAX As Double = 10.0
Private Const BSW_TIER1_INCREMENT As Double = 0.025
Private Const BSW_TIER2_INCREMENT As Double = 0.05
Private Const BSW_TIER3_INCREMENT As Double = 1.0

' ============================================================================
' BS&W VALUE GENERATION
' ============================================================================

Public Function GenerateBSWValue(historicalMean As Double, historicalStdev As Double, _
                                 historicalMin As Double, historicalMax As Double) As Double
    ' Generate a plausible BS&W value based on historical behavior
    
    Dim generatedValue As Double
    Dim noise As Double
    
    ' Generate normally distributed value around historical mean
    noise = (Rnd + Rnd + Rnd + Rnd - 2) * historicalStdev ' Approximate normal using sum of uniforms
    generatedValue = historicalMean + noise
    
    ' Constrain within historical range
    If generatedValue < historicalMin Then
        generatedValue = historicalMin
    End If
    If generatedValue > historicalMax Then
        generatedValue = historicalMax
    End If
    
    ' Ensure non-negative
    If generatedValue < 0 Then
        generatedValue = 0
    End If
    
    GenerateBSWValue = generatedValue
    
End Function

' ============================================================================
' BS&W ROUNDING
' ============================================================================

Public Function RoundBSWToStandard(value As Double) As Double
    ' Round BS&W value to the nearest standard increment
    
    If value < 0 Then
        RoundBSWToStandard = 0
        Exit Function
    End If
    
    If value >= BSW_TIER3_MAX Then
        ' Cap at 10.0
        RoundBSWToStandard = BSW_TIER3_MAX
        Exit Function
    End If
    
    If value >= BSW_TIER2_MAX Then
        ' Tier 3: 1.0 increments (1.0, 2.0, 3.0, ..., 10.0)
        RoundBSWToStandard = Round(value / BSW_TIER3_INCREMENT) * BSW_TIER3_INCREMENT
        Exit Function
    End If
    
    If value >= BSW_TIER1_MAX Then
        ' Tier 2: 0.05 increments (0.1, 0.15, 0.2, ..., 0.95)
        RoundBSWToStandard = Round(value / BSW_TIER2_INCREMENT) * BSW_TIER2_INCREMENT
        Exit Function
    End If
    
    ' Tier 1: 0.025 increments (0, 0.025, 0.05, 0.075)
    RoundBSWToStandard = Round(value / BSW_TIER1_INCREMENT) * BSW_TIER1_INCREMENT
    
End Function

Public Function ValidateBSWValue(value As Double) As Boolean
    ' Check if a BS&W value conforms to standard increments
    
    Dim rounded As Double
    
    If value < 0 Or value > BSW_TIER3_MAX Then
        ValidateBSWValue = False
        Exit Function
    End If
    
    rounded = RoundBSWToStandard(value)
    
    ' Allow small rounding error (due to floating point)
    ValidateBSWValue = (Abs(value - rounded) < 0.001)
    
End Function

' ============================================================================
' BS&W STATISTICS AND ANALYSIS
' ============================================================================

Public Function AnalyzeBSWData(dataArray() As Double) As BSWAnalysisResult
    ' Analyze historical BS&W data
    
    Dim i As Long, n As Long
    Dim sum As Double, sumSq As Double
    Dim mean As Double, stdev As Double
    Dim minVal As Double, maxVal As Double
    
    With AnalyzeBSWData
        n = UBound(dataArray) - LBound(dataArray) + 1
        
        If n < 1 Then Exit Function
        
        ' Calculate basic stats
        minVal = dataArray(LBound(dataArray))
        maxVal = dataArray(LBound(dataArray))
        
        For i = LBound(dataArray) To UBound(dataArray)
            sum = sum + dataArray(i)
            If dataArray(i) < minVal Then minVal = dataArray(i)
            If dataArray(i) > maxVal Then maxVal = dataArray(i)
        Next i
        
        mean = sum / n
        
        If n < 2 Then
            stdev = 0
        Else
            For i = LBound(dataArray) To UBound(dataArray)
                sumSq = sumSq + (dataArray(i) - mean) ^ 2
            Next i
            stdev = Sqr(sumSq / (n - 1))
        End If
        
        .Mean = mean
        .StDev = stdev
        .MinValue = minVal
        .MaxValue = maxVal
        .ObservationCount = n
        
    End With
    
End Function

Public Type BSWAnalysisResult
    Mean As Double
    StDev As Double
    MinValue As Double
    MaxValue As Double
    ObservationCount As Long
End Type

' ============================================================================
' BS&W TREND AND BEHAVIOR
' ============================================================================

Public Function AnalyzeBSWTrend(dataArray() As Double) As Boolean
    ' Detect if BS&W is trending up or down
    ' Returns True if trend is upward
    
    Dim n As Long, i As Long
    Dim earlyAverage As Double, lateAverage As Double
    Dim earlyCount As Long, lateCount As Long
    
    n = UBound(dataArray) - LBound(dataArray) + 1
    
    If n < 4 Then
        AnalyzeBSWTrend = False
        Exit Function
    End If
    
    ' Compare early and late periods
    For i = LBound(dataArray) To LBound(dataArray) + (n \ 2) - 1
        earlyAverage = earlyAverage + dataArray(i)
        earlyCount = earlyCount + 1
    Next i
    
    For i = LBound(dataArray) + (n \ 2) To UBound(dataArray)
        lateAverage = lateAverage + dataArray(i)
        lateCount = lateCount + 1
    Next i
    
    If earlyCount > 0 Then earlyAverage = earlyAverage / earlyCount
    If lateCount > 0 Then lateAverage = lateAverage / lateCount
    
    AnalyzeBSWTrend = (lateAverage > earlyAverage)
    
End Function

Public Function GetBSWStatusMessage(mean As Double, stdev As Double, _
                                    minVal As Double, maxVal As Double) As String
    ' Generate human-readable status message about BS&W behavior
    
    Dim msg As String
    
    msg = "BS&W Analysis:" & vbCrLf
    msg = msg & "  Mean: " & Format(mean, "0.000") & vbCrLf
    msg = msg & "  Range: " & Format(minVal, "0.000") & " to " & Format(maxVal, "0.000") & vbCrLf
    msg = msg & "  Volatility: "
    
    If stdev < 0.1 Then
        msg = msg & "Low (stable)" & vbCrLf
    ElseIf stdev < 0.5 Then
        msg = msg & "Moderate" & vbCrLf
    Else
        msg = msg & "High (variable)" & vbCrLf
    End If
    
    GetBSWStatusMessage = msg
    
End Function
