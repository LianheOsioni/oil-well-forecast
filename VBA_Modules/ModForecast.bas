Attribute VB_Name = "ModForecast"
' ============================================================================
' MODULE: ModForecast
' Core forecasting and simulation engine
' ============================================================================

Option Explicit

' Forecast session state type
Public Type ForecastSession
    ' Learning statistics
    Mean As Double
    StDev As Double
    AutoCorr As Double
    MinHistorical As Double
    MaxHistorical As Double
    
    ' Band definitions
    LowBand As Double
    HighBand As Double
    ProbLow As Double
    ProbMid As Double
    ProbHigh As Double
    
    ' Event parameters
    PlungeFrequency As Double
    PlungeAvgDepth As Double
    SpikeFrequency As Double
    SpikeAvgHeight As Double
    
    ' Target tracking
    TargetMode As Long ' 0=None, 1=One, 2=Two
    Target1 As Double
    Target2 As Double
    Target1Row As Long
    Target2Row As Long
    
    ' Generation options
    AllowExtrapolation As Boolean
    MaxDeviationFactor As Double ' Limit extreme deviations
    
    ' State tracking
    PreviousIncrement As Double
    CurrentCumulative As Double
    RowIndex As Long
    
End Type

Private session As ForecastSession
Private useFixedSeed As Boolean
Private fixedSeed As Long

' ============================================================================
' SESSION INITIALIZATION
' ============================================================================

Public Sub InitializeSession(stats As StatisticsResult, trendResult As TrendResult, _
                             autoCorr As AutocorrelationResult, plungeCount As Long, _
                             spikeCount As Long, plungeDepth As Double, spikeHeight As Double)
    ' Initialize forecast session with learned statistics
    
    Dim totalObs As Long, plungeFreq As Double, spikeFreq As Double
    Dim cLow As Long, cMid As Long, cHigh As Long, i As Long
    
    With session
        .Mean = stats.Mean
        .StDev = stats.StDev
        .AutoCorr = autoCorr.Lag1
        .MinHistorical = stats.MinValue
        .MaxHistorical = stats.MaxValue
        
        ' Clip autocorrelation to reasonable range
        If .AutoCorr < 0.05 Then .AutoCorr = 0.05
        If .AutoCorr > 0.95 Then .AutoCorr = 0.95
        
        ' Define bands
        .LowBand = .Mean - 0.5 * .StDev
        .HighBand = .Mean + 0.5 * .StDev
        
        ' Band probabilities (simplified three-regime model)
        .ProbLow = 0.25
        .ProbMid = 0.5
        .ProbHigh = 0.25
        
        ' Event parameters
        .PlungeFrequency = plungeFreq
        .PlungeAvgDepth = plungeDepth
        .SpikeFrequency = spikeFreq
        .SpikeAvgHeight = spikeHeight
        
        .TargetMode = 0
        .AllowExtrapolation = True
        .MaxDeviationFactor = 3  ' Limit to ±3 standard deviations
        
    End With
    
End Sub

' ============================================================================
' INDIVIDUAL INCREMENT GENERATION
' ============================================================================

Public Function GenerateIncrement(prevIncrement As Double, ByRef state As ForecastSession) As Double
    ' Generate a single 15-minute increment based on learned model
    
    Dim increment As Double
    Dim r As Double
    Dim mean As Double, stdev As Double, autoCorr As Double
    
    mean = state.Mean
    stdev = state.StDev
    autoCorr = state.AutoCorr
    
    If stdev <= 0 Then stdev = mean * 0.1  ' Fallback for constant data
    
    ' Step 1: Select regime based on probabilities
    r = Rnd
    
    If r < state.ProbLow Then
        ' Low regime: below mean
        increment = state.LowBand - Rnd * 0.35 * stdev
    ElseIf r < state.ProbLow + state.ProbMid Then
        ' Mid regime: around mean
        increment = mean + (Rnd - 0.5) * 0.7 * stdev
    Else
        ' High regime: above mean
        increment = state.HighBand + Rnd * 0.35 * stdev
    End If
    
    ' Step 2: Apply lag-1 autocorrelation
    ' Formula: Increment(t) = Mean + AutoCorr * (Prev - Mean) + Noise
    Dim noise As Double
    noise = (Rnd - 0.5) * stdev
    increment = mean + autoCorr * (prevIncrement - mean) + noise
    
    ' Step 3: Add mean-reversion bias
    Dim drift As Double
    drift = (prevIncrement - mean) * 0.05
    increment = increment - drift
    
    ' Step 4: Inject plunge if triggered
    If Rnd < state.PlungeFrequency Then
        Dim plungeDepth As Double
        plungeDepth = (Rnd * 0.6 + 0.4) * state.PlungeAvgDepth
        increment = increment - plungeDepth
    End If
    
    ' Step 5: Inject spike if triggered
    If Rnd < state.SpikeFrequency Then
        Dim spikeHeight As Double
        spikeHeight = (Rnd * 0.6 + 0.4) * state.SpikeAvgHeight
        increment = increment + spikeHeight
    End If
    
    GenerateIncrement = increment
    
End Function

' ============================================================================
' VALIDATION AND BOUNDS CHECKING
' ============================================================================

Public Function ValidateIncrement(increment As Double, prevCumulative As Double, _
                                   state As ForecastSession, Optional allowNegative As Boolean = False) As Boolean
    ' Check if generated increment is plausible
    
    ' Check physical bounds
    If Not allowNegative And increment < 0 Then
        ValidateIncrement = False
        Exit Function
    End If
    
    ' Check deviation from historical range
    If Not state.AllowExtrapolation Then
        If increment > state.MaxHistorical * 1.2 Or increment < state.MinHistorical * 0.8 Then
            ValidateIncrement = False
            Exit Function
        End If
    End If
    
    ' Check extreme deviation from mean
    Dim maxDeviation As Double
    maxDeviation = state.Mean + state.MaxDeviationFactor * state.StDev
    
    If increment > maxDeviation Then
        ValidateIncrement = False
        Exit Function
    End If
    
    ' Passed all checks
    ValidateIncrement = True
    
End Function

' ============================================================================
' BATCH GENERATION
' ============================================================================

Public Function GenerateForecastBatch(numRows As Long, startingCumulative As Double, _
                                       startingIncrement As Double, state As ForecastSession, _
                                       ByRef outputArray() As Double) As Boolean
    ' Generate a batch of forecast increments
    
    Dim i As Long, attemptCount As Long
    Dim increment As Double, cumulative As Double, prevInc As Double
    Dim maxAttempts As Long
    
    maxAttempts = 10  ' Max attempts per increment
    cumulative = startingCumulative
    prevInc = startingIncrement
    
    ReDim outputArray(1 To numRows)
    
    For i = 1 To numRows
        attemptCount = 0
        
        Do
            increment = GenerateIncrement(prevInc, state)
            attemptCount = attemptCount + 1
            
            If ValidateIncrement(increment, cumulative, state) Then
                cumulative = cumulative + increment
                outputArray(i) = cumulative
                prevInc = increment
                Exit Do
            End If
            
            If attemptCount >= maxAttempts Then
                ' Failed to generate valid increment; use fallback
                increment = state.Mean + (Rnd - 0.5) * 0.5 * state.StDev
                cumulative = cumulative + increment
                outputArray(i) = cumulative
                prevInc = increment
                Exit Do
            End If
        Loop
    Next i
    
    GenerateForecastBatch = True
    
End Function

' ============================================================================
' RANDOM SEED CONTROL
' ============================================================================

Public Sub SetRandomSeed(seed As Long)
    ' Use a fixed seed for reproducible forecasts
    useFixedSeed = True
    fixedSeed = seed
    Randomize seed
End Sub

Public Sub ClearRandomSeed()
    ' Return to variable random behavior
    useFixedSeed = False
    Randomize
End Sub
