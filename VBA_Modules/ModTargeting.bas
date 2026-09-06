Attribute VB_Name = "ModTargeting"
' ============================================================================
' MODULE: ModTargeting
' Target-based trajectory generation and constraint handling
' ============================================================================

Option Explicit

' Target trajectory type
Public Type TargetTrajectory
    Mode As Long ' 0=None, 1=Single target, 2=Dual target
    Target1 As Double
    Target1Row As Long
    Target2 As Double
    Target2Row As Long
    CurrentValue As Double
    RequiredAverage As Double
    IsReachable As Boolean
    WarningMessage As String
End Type

' ============================================================================
' TARGET SETUP AND VALIDATION
' ============================================================================

Public Function SetupTargetTrajectory(mode As Long, target1 As Double, target1Row As Long, _
                                       target2 As Double, target2Row As Long, _
                                       currentValue As Double, mean As Double, stdev As Double) As TargetTrajectory
    ' Set up target trajectory and check feasibility
    
    With SetupTargetTrajectory
        .Mode = mode
        .Target1 = target1
        .Target1Row = target1Row
        .Target2 = target2
        .Target2Row = target2Row
        .CurrentValue = currentValue
        .IsReachable = True
        .WarningMessage = ""
        
        Select Case mode
            Case 0
                ' No target
                .RequiredAverage = mean
                
            Case 1
                ' Single target
                If target1Row > 0 Then
                    .RequiredAverage = (target1 - currentValue) / target1Row
                Else
                    .RequiredAverage = mean
                End If
                
                ' Check feasibility
                Dim requiredPerPeriod As Double
                requiredPerPeriod = (target1 - currentValue) / target1Row
                
                If Abs(requiredPerPeriod) > mean + 3 * stdev Then
                    .WarningMessage = "Target requires unusual average increment " & _
                                    Format(requiredPerPeriod, "0.00") & ". " & _
                                    "Historical average: " & Format(mean, "0.00") & "."
                End If
                
            Case 2
                ' Dual target
                If target2Row > 0 Then
                    .RequiredAverage = (target2 - currentValue) / target2Row
                Else
                    .RequiredAverage = mean
                End If
                
                ' Validate both targets are feasible
                If target2Row <= target1Row Then
                    .IsReachable = False
                    .WarningMessage = "Target 2 row must be after Target 1 row."
                End If
        End Select
    End With
    
End Function

' ============================================================================
' TRAJECTORY-CONSTRAINED INCREMENT GENERATION
' ============================================================================

Public Function ApplyTargetConstraint(baseIncrement As Double, currentCumulative As Double, _
                                      rowIndex As Long, trajectory As TargetTrajectory) As Double
    ' Adjust increment to stay on track toward target(s)
    
    Dim remainingRows As Long
    Dim remainingBudget As Double
    Dim requiredIncrement As Double
    Dim blendFactor As Double
    Dim targetValue As Double
    
    If trajectory.Mode = 0 Then
        ' No target; return base increment unchanged
        ApplyTargetConstraint = baseIncrement
        Exit Function
    End If
    
    ' Determine which target we're tracking toward
    If trajectory.Mode = 1 Then
        targetValue = trajectory.Target1
        remainingRows = trajectory.Target1Row - rowIndex + 1
    ElseIf trajectory.Mode = 2 Then
        ' Choose target based on current row position
        If rowIndex <= trajectory.Target2Row Then
            targetValue = trajectory.Target2
            remainingRows = trajectory.Target2Row - rowIndex + 1
        Else
            targetValue = trajectory.Target1
            remainingRows = trajectory.Target1Row - rowIndex + 1
        End If
    Else
        ApplyTargetConstraint = baseIncrement
        Exit Function
    End If
    
    ' Calculate required increment
    If remainingRows > 0 Then
        remainingBudget = targetValue - currentCumulative
        requiredIncrement = remainingBudget / remainingRows
    Else
        ApplyTargetConstraint = baseIncrement
        Exit Function
    End If
    
    ' Blend base increment toward required increment
    ' Increase blend factor as we approach the target row
    blendFactor = 0.15 + 0.5 * (1 - remainingRows / trajectory.Target1Row)
    If blendFactor > 0.9 Then blendFactor = 0.9
    
    ApplyTargetConstraint = baseIncrement * (1 - blendFactor) + requiredIncrement * blendFactor
    
End Function

' ============================================================================
' TRAJECTORY CALCULATION AND VISUALIZATION
' ============================================================================

Public Function CalculateIdealTrajectory(startValue As Double, target1 As Double, target1Row As Long, _
                                         Optional target2 As Double = 0, _
                                         Optional target2Row As Long = 0) As Double()
    ' Calculate ideal (linear interpolation) trajectory between targets
    
    Dim i As Long
    Dim trajectory() As Double
    Dim numRows As Long
    Dim segment1Rows As Long, segment2Rows As Long
    Dim slope1 As Double, slope2 As Double
    
    If target1Row < 1 Then
        ReDim trajectory(1 To 1)
        trajectory(1) = startValue
        CalculateIdealTrajectory = trajectory
        Exit Function
    End If
    
    ' Determine trajectory mode
    If target2Row > 0 And target2Row < target1Row Then
        ' Two-segment trajectory
        segment1Rows = target2Row
        segment2Rows = target1Row - target2Row
        slope1 = (target2 - startValue) / segment1Rows
        slope2 = (target1 - target2) / segment2Rows
        
        ReDim trajectory(1 To target1Row)
        
        ' First segment
        For i = 1 To segment1Rows
            trajectory(i) = startValue + slope1 * i
        Next i
        
        ' Second segment
        For i = segment1Rows + 1 To target1Row
            trajectory(i) = target2 + slope2 * (i - segment1Rows)
        Next i
    Else
        ' Single segment trajectory
        slope1 = (target1 - startValue) / target1Row
        
        ReDim trajectory(1 To target1Row)
        
        For i = 1 To target1Row
            trajectory(i) = startValue + slope1 * i
        Next i
    End If
    
    CalculateIdealTrajectory = trajectory
    
End Function

Public Function CalculateTargetDeviation(actualValue As Double, targetValue As Double, _
                                         decimalPlaces As Long) As Double
    ' Calculate absolute deviation from target after rounding
    
    Dim roundedActual As Double
    roundedActual = Round(actualValue, decimalPlaces)
    CalculateTargetDeviation = Abs(roundedActual - targetValue)
    
End Function

' ============================================================================
' FEASIBILITY CHECKING
' ============================================================================

Public Function CheckTargetFeasibility(startValue As Double, targetValue As Double, _
                                       numRows As Long, mean As Double, stdev As Double) As Boolean
    ' Quick check: is target achievable with reasonable increments?
    
    Dim requiredAverage As Double
    Dim maxInc As Double, minInc As Double
    
    If numRows < 1 Then
        CheckTargetFeasibility = False
        Exit Function
    End If
    
    requiredAverage = (targetValue - startValue) / numRows
    
    ' Allow up to ±5 standard deviations variation
    maxInc = mean + 5 * stdev
    minInc = mean - 5 * stdev
    
    CheckTargetFeasibility = (requiredAverage >= minInc And requiredAverage <= maxInc)
    
End Function

Public Function EstimateTargetReachability(startValue As Double, targetValue As Double, _
                                           numRows As Long, mean As Double, stdev As Double) As String
    ' Provide human-readable assessment of target reachability
    
    Dim requiredAverage As Double
    Dim assessment As String
    Dim deviationFromMean As Double
    
    If numRows < 1 Then
        EstimateTargetReachability = "Invalid row count."
        Exit Function
    End If
    
    requiredAverage = (targetValue - startValue) / numRows
    deviationFromMean = requiredAverage - mean
    
    If Abs(deviationFromMean) < stdev Then
        assessment = "Very feasible (within 1 SD of normal)."
    ElseIf Abs(deviationFromMean) < 2 * stdev Then
        assessment = "Feasible (within 2 SD of normal)."
    ElseIf Abs(deviationFromMean) < 3 * stdev Then
        assessment = "Challenging (within 3 SD of normal)."
    ElseIf Abs(deviationFromMean) < 5 * stdev Then
        assessment = "Very challenging (within 5 SD of normal)."
    Else
        assessment = "Likely unrealistic (beyond 5 SD of normal)."
    End If
    
    EstimateTargetReachability = assessment & " Required average: " & _
                               Format(requiredAverage, "0.00") & ". Historical: " & _
                               Format(mean, "0.00") & " ± " & Format(stdev, "0.00") & "."
    
End Function
