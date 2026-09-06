Attribute VB_Name = "ModStatistics"
' ============================================================================
' MODULE: ModStatistics
' Comprehensive statistical analysis of historical data
' ============================================================================

Option Explicit

' Statistics result type
Public Type StatisticsResult
    Count As Long
    Mean As Double
    Median As Double
    StDev As Double
    Variance As Double
    MinValue As Double
    MaxValue As Double
    Range As Double
    SkewnessValue As Double
    KurtosisValue As Double
    Q1 As Double  ' 25th percentile
    Q3 As Double  ' 75th percentile
    IQR As Double ' Interquartile range
End Type

' Trend detection result type
Public Type TrendResult
    Direction As Long ' -1 = down, 0 = flat, 1 = up
    Strength As Double ' 0 to 1, where 1 is perfect linear trend
    Slope As Double
    Intercept As Double
End Type

' Autocorrelation result type
Public Type AutocorrelationResult
    Lag1 As Double
    Lag2 As Double
    Lag3 As Double
    PersistenceScore As Double ' 0 to 1
End Type

' ============================================================================
' BASIC STATISTICS
' ============================================================================

Public Function CalculateStatistics(dataArray() As Double) As StatisticsResult
    ' Calculate comprehensive statistics from data array
    
    Dim i As Long, n As Long
    Dim sum As Double, sumSq As Double, sumCubed As Double, sumQuartic As Double
    Dim mean As Double, variance As Double, stdev As Double
    Dim minVal As Double, maxVal As Double
    Dim sortedArray() As Double
    
    With CalculateStatistics
        n = UBound(dataArray) - LBound(dataArray) + 1
        .Count = n
        
        If n < 1 Then Exit Function
        
        ' Calculate sum and basic stats
        minVal = dataArray(LBound(dataArray))
        maxVal = dataArray(LBound(dataArray))
        sum = 0
        
        For i = LBound(dataArray) To UBound(dataArray)
            sum = sum + dataArray(i)
            If dataArray(i) < minVal Then minVal = dataArray(i)
            If dataArray(i) > maxVal Then maxVal = dataArray(i)
        Next i
        
        mean = sum / n
        .Mean = mean
        .MinValue = minVal
        .MaxValue = maxVal
        .Range = maxVal - minVal
        
        ' Calculate variance and standard deviation
        If n < 2 Then
            .Variance = 0
            .StDev = 0
        Else
            For i = LBound(dataArray) To UBound(dataArray)
                sumSq = sumSq + (dataArray(i) - mean) ^ 2
                sumCubed = sumCubed + (dataArray(i) - mean) ^ 3
                sumQuartic = sumQuartic + (dataArray(i) - mean) ^ 4
            Next i
            
            variance = sumSq / (n - 1)
            .Variance = variance
            .StDev = Sqr(variance)
            
            ' Calculate skewness and kurtosis
            If .StDev > 0 Then
                .SkewnessValue = (sumCubed / n) / (.StDev ^ 3)
                .KurtosisValue = (sumQuartic / n) / (.StDev ^ 4) - 3
            End If
        End If
        
        ' Calculate quartiles
        sortedArray = SortArray(dataArray)
        .Median = CalculatePercentile(sortedArray, 0.5)
        .Q1 = CalculatePercentile(sortedArray, 0.25)
        .Q3 = CalculatePercentile(sortedArray, 0.75)
        .IQR = .Q3 - .Q1
        
    End With
    
End Function

' ============================================================================
' PERCENTILE CALCULATION
' ============================================================================

Private Function CalculatePercentile(sortedArray() As Double, percentile As Double) As Double
    ' Calculate a percentile value from sorted array
    ' percentile: 0 to 1 (e.g., 0.5 = median)
    
    Dim n As Long, position As Double
    Dim lower As Long, upper As Long
    Dim lowerVal As Double, upperVal As Double
    
    n = UBound(sortedArray) - LBound(sortedArray) + 1
    position = (n - 1) * percentile + 1
    
    lower = Int(position)
    upper = lower + 1
    
    If lower < LBound(sortedArray) Then lower = LBound(sortedArray)
    If upper > UBound(sortedArray) Then upper = UBound(sortedArray)
    
    lowerVal = sortedArray(lower)
    upperVal = sortedArray(upper)
    
    CalculatePercentile = lowerVal + (upperVal - lowerVal) * (position - lower)
    
End Function

Private Function SortArray(dataArray() As Double) As Double()
    ' Return a sorted copy of the array (ascending)
    
    Dim i As Long, j As Long, n As Long
    Dim sortedArray() As Double
    Dim temp As Double
    
    n = UBound(dataArray) - LBound(dataArray) + 1
    ReDim sortedArray(1 To n)
    
    ' Copy data
    For i = 1 To n
        sortedArray(i) = dataArray(LBound(dataArray) + i - 1)
    Next i
    
    ' Bubble sort (simple for small arrays)
    For i = 1 To n - 1
        For j = i + 1 To n
            If sortedArray(i) > sortedArray(j) Then
                temp = sortedArray(i)
                sortedArray(i) = sortedArray(j)
                sortedArray(j) = temp
            End If
        Next j
    Next i
    
    SortArray = sortedArray
    
End Function

' ============================================================================
' TREND ANALYSIS
' ============================================================================

Public Function AnalyzeTrend(dataArray() As Double) As TrendResult
    ' Analyze linear trend using simple linear regression
    
    Dim i As Long, n As Long
    Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double
    Dim meanX As Double, meanY As Double
    Dim slope As Double, intercept As Double
    Dim correlation As Double, slopeSignificance As Double
    
    With AnalyzeTrend
        n = UBound(dataArray) - LBound(dataArray) + 1
        
        If n < 2 Then
            .Direction = 0
            .Strength = 0
            .Slope = 0
            .Intercept = 0
            Exit Function
        End If
        
        ' Calculate sums for linear regression
        For i = LBound(dataArray) To UBound(dataArray)
            sumX = sumX + (i - LBound(dataArray))
            sumY = sumY + dataArray(i)
            sumXY = sumXY + (i - LBound(dataArray)) * dataArray(i)
            sumX2 = sumX2 + (i - LBound(dataArray)) ^ 2
        Next i
        
        meanX = sumX / n
        meanY = sumY / n
        
        ' Calculate slope and intercept
        If sumX2 <> meanX * n Then
            slope = (sumXY - meanX * sumY) / (sumX2 - meanX * n)
            intercept = meanY - slope * meanX
            .Slope = slope
            .Intercept = intercept
            
            ' Determine trend direction
            If slope > 0.000001 Then
                .Direction = 1  ' Upward trend
            ElseIf slope < -0.000001 Then
                .Direction = -1  ' Downward trend
            Else
                .Direction = 0  ' Flat
            End If
            
            ' Calculate strength (correlation coefficient)
            Dim sumRsq As Double, sumYdiff As Double
            For i = LBound(dataArray) To UBound(dataArray)
                sumRsq = sumRsq + (dataArray(i) - (intercept + slope * (i - LBound(dataArray)))) ^ 2
                sumYdiff = sumYdiff + (dataArray(i) - meanY) ^ 2
            Next i
            
            If sumYdiff > 0 Then
                .Strength = 1 - (sumRsq / sumYdiff)
                If .Strength < 0 Then .Strength = 0
            Else
                .Strength = 0
            End If
        End If
    End With
    
End Function

' ============================================================================
' AUTOCORRELATION ANALYSIS
' ============================================================================

Public Function AnalyzeAutocorrelation(dataArray() As Double) As AutocorrelationResult
    ' Calculate lag-1, lag-2, and lag-3 autocorrelation
    
    Dim i As Long, n As Long
    Dim mean As Double, sum As Double
    Dim cov1 As Double, cov2 As Double, cov3 As Double
    Dim var As Double
    Dim lag1 As Double, lag2 As Double, lag3 As Double
    
    With AnalyzeAutocorrelation
        n = UBound(dataArray) - LBound(dataArray) + 1
        
        If n < 4 Then
            .Lag1 = 0.3
            .Lag2 = 0.1
            .Lag3 = 0.05
            .PersistenceScore = 0.15
            Exit Function
        End If
        
        ' Calculate mean
        For i = LBound(dataArray) To UBound(dataArray)
            sum = sum + dataArray(i)
        Next i
        mean = sum / n
        
        ' Calculate variances and covariances
        For i = LBound(dataArray) To UBound(dataArray) - 3
            var = var + (dataArray(i) - mean) ^ 2
            
            If i <= UBound(dataArray) - 1 Then
                cov1 = cov1 + (dataArray(i) - mean) * (dataArray(i + 1) - mean)
            End If
            
            If i <= UBound(dataArray) - 2 Then
                cov2 = cov2 + (dataArray(i) - mean) * (dataArray(i + 2) - mean)
            End If
            
            If i <= UBound(dataArray) - 3 Then
                cov3 = cov3 + (dataArray(i) - mean) * (dataArray(i + 3) - mean)
            End If
        Next i
        
        ' Avoid division by zero
        If var > 0 Then
            lag1 = cov1 / var
            lag2 = cov2 / var
            lag3 = cov3 / var
        End If
        
        ' Clip to valid range [-1, 1]
        If lag1 < -1 Then lag1 = -1
        If lag1 > 1 Then lag1 = 1
        If lag2 < -1 Then lag2 = -1
        If lag2 > 1 Then lag2 = 1
        If lag3 < -1 Then lag3 = -1
        If lag3 > 1 Then lag3 = 1
        
        .Lag1 = lag1
        .Lag2 = lag2
        .Lag3 = lag3
        
        ' Persistence score: how much past values influence current
        .PersistenceScore = Abs(lag1) * 0.6 + Abs(lag2) * 0.3 + Abs(lag3) * 0.1
        If .PersistenceScore > 1 Then .PersistenceScore = 1
        
    End With
    
End Function

' ============================================================================
' EVENT DETECTION
' ============================================================================

Public Function DetectPlunges(dataArray() As Double, mean As Double, stdev As Double) As Long
    ' Count number of plunge events (sudden downward deviations)
    ' Plunge defined as: value < Mean - 1.5 * StDev
    
    Dim i As Long, count As Long
    Dim threshold As Double
    
    If stdev <= 0 Then
        DetectPlunges = 0
        Exit Function
    End If
    
    threshold = mean - 1.5 * stdev
    
    For i = LBound(dataArray) To UBound(dataArray)
        If dataArray(i) < threshold Then
            count = count + 1
        End If
    Next i
    
    DetectPlunges = count
    
End Function

Public Function DetectSpikes(dataArray() As Double, mean As Double, stdev As Double) As Long
    ' Count number of spike events (sudden upward deviations)
    ' Spike defined as: value > Mean + 1.5 * StDev
    
    Dim i As Long, count As Long
    Dim threshold As Double
    
    If stdev <= 0 Then
        DetectSpikes = 0
        Exit Function
    End If
    
    threshold = mean + 1.5 * stdev
    
    For i = LBound(dataArray) To UBound(dataArray)
        If dataArray(i) > threshold Then
            count = count + 1
        End If
    Next i
    
    DetectSpikes = count
    
End Function

Public Function CalculateAverageEventDepth(dataArray() As Double, mean As Double, stdev As Double, _
                                            isPlunge As Boolean) As Double
    ' Calculate average depth/height of plunges or spikes
    
    Dim i As Long, count As Long
    Dim threshold As Double, sumDepth As Double
    
    If stdev <= 0 Then
        CalculateAverageEventDepth = stdev
        Exit Function
    End If
    
    If isPlunge Then
        threshold = mean - 1.5 * stdev
        For i = LBound(dataArray) To UBound(dataArray)
            If dataArray(i) < threshold Then
                sumDepth = sumDepth + (mean - dataArray(i))
                count = count + 1
            End If
        Next i
    Else
        threshold = mean + 1.5 * stdev
        For i = LBound(dataArray) To UBound(dataArray)
            If dataArray(i) > threshold Then
                sumDepth = sumDepth + (dataArray(i) - mean)
                count = count + 1
            End If
        Next i
    End If
    
    If count > 0 Then
        CalculateAverageEventDepth = sumDepth / count
    Else
        CalculateAverageEventDepth = stdev  ' Fallback
    End If
    
End Function

' ============================================================================
' VOLATILITY AND CHANGE ANALYSIS
' ============================================================================

Public Function CalculateVolatility(dataArray() As Double) As Double
    ' Calculate volatility as standard deviation of period-to-period changes
    
    Dim i As Long, n As Long
    Dim changes() As Double
    Dim sum As Double, mean As Double, sumSq As Double
    
    n = UBound(dataArray) - LBound(dataArray)
    
    If n < 2 Then
        CalculateVolatility = 0
        Exit Function
    End If
    
    ReDim changes(1 To n)
    
    ' Calculate period-to-period changes
    For i = 1 To n
        changes(i) = dataArray(LBound(dataArray) + i) - dataArray(LBound(dataArray) + i - 1)
        sum = sum + changes(i)
    Next i
    
    mean = sum / n
    
    ' Calculate standard deviation of changes
    For i = 1 To n
        sumSq = sumSq + (changes(i) - mean) ^ 2
    Next i
    
    If n > 1 Then
        CalculateVolatility = Sqr(sumSq / (n - 1))
    Else
        CalculateVolatility = 0
    End If
    
End Function

Public Function CalculateConsecutiveChange(dataArray() As Double, numPeriods As Long) As Double
    ' Calculate the typical change over N consecutive periods
    
    Dim i As Long, n As Long
    Dim sum As Double, count As Long
    
    n = UBound(dataArray) - LBound(dataArray) + 1
    
    If n <= numPeriods Then
        CalculateConsecutiveChange = 0
        Exit Function
    End If
    
    For i = LBound(dataArray) To UBound(dataArray) - numPeriods
        sum = sum + (dataArray(i + numPeriods) - dataArray(i))
        count = count + 1
    Next i
    
    If count > 0 Then
        CalculateConsecutiveChange = sum / count / numPeriods
    Else
        CalculateConsecutiveChange = 0
    End If
    
End Function
