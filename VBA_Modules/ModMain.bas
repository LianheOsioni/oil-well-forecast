Attribute VB_Name = "ModMain"
' ============================================================================
' MODULE: ModMain
' Main orchestration and workflow control
' ============================================================================

Option Explicit

' ============================================================================
' MAIN ENTRY POINT
' ============================================================================

Public Sub GenerateOilWellForecast()
    ' Main orchestration routine - called by user
    
    Dim config As ConfigSettings
    Dim ws As Worksheet
    Dim excelState As ExcelApplicationState
    
    On Error GoTo ErrorHandler
    
    ' Save and disable Excel settings for performance
    SaveExcelState excelState
    DisableExcelUpdates
    
    ' Get active worksheet
    Set ws = ActiveSheet
    
    ' Load saved configuration
    InitializeDefaults
    LoadConfiguration config
    
    ' ---- PHASE 1: USER INPUT AND CONFIGURATION ----
    If Not PhaseUserInput(ws, config) Then
        GoTo ExitProc
    End If
    
    ' ---- PHASE 2: DATA EXTRACTION AND VALIDATION ----
    If Not PhaseDataExtraction(ws, config) Then
        GoTo ExitProc
    End If
    
    ' ---- PHASE 3: STATISTICAL ANALYSIS ----
    If Not PhaseStatisticalAnalysis(ws, config) Then
        GoTo ExitProc
    End If
    
    ' ---- PHASE 4: TARGET CONFIGURATION ----
    If Not PhaseTargetConfiguration(ws, config) Then
        GoTo ExitProc
    End If
    
    ' ---- PHASE 5: FORECAST GENERATION ----
    If Not PhaseForecastGeneration(ws, config) Then
        GoTo ExitProc
    End If
    
    ' ---- PHASE 6: OUTPUT AND CHARTING ----
    If Not PhaseOutputAndChart(ws, config) Then
        GoTo ExitProc
    End If
    
    ' Save configuration for next run
    SaveConfiguration config
    
    ' Success message
    MsgBox "Forecast generated successfully!" & vbCrLf & vbCrLf & _
           GetConfigSummary(config), vbInformation
    
ExitProc:
    RestoreExcelState excelState
    Exit Sub
    
ErrorHandler:
    RestoreExcelState excelState
    MsgBox "Error in main workflow: " & Err.Description, vbCritical
    
End Sub

' ============================================================================
' PHASE 1: USER INPUT
' ============================================================================

Private Function PhaseUserInput(ws As Worksheet, ByRef config As ConfigSettings) As Boolean
    ' Request and validate all user inputs
    
    Dim colInput As Long
    Dim rowInput As Long
    
    ' Request learning range configuration
    MsgBox "Step 1 of 6: Learning Range Configuration" & vbCrLf & vbCrLf & _
           "Please specify the historical data range used to learn forecasting behavior.", vbInformation
    
    ' Lookup column
    colInput = GetColumnInput("Column containing historical 15-minute increments:", config.LookupCol)
    If colInput = -1 Then
        PhaseUserInput = False
        Exit Function
    End If
    config.LookupCol = colInput
    
    ' Learning start row
    rowInput = GetRowInput("Learning start row:", config.LearnStart)
    If rowInput = -1 Then
        PhaseUserInput = False
        Exit Function
    End If
    config.LearnStart = rowInput
    
    ' Learning end row
    rowInput = GetRowInput("Learning end row:", config.LearnEnd)
    If rowInput = -1 Then
        PhaseUserInput = False
        Exit Function
    End If
    config.LearnEnd = rowInput
    
    ' Validate learning range
    Dim valResult As ValidationResult
    valResult = ValidateLearningRange(config.LearnStart, config.LearnEnd, 10)
    If Not valResult.IsValid Then
        MsgBox valResult.ErrorMessage, vbExclamation
        PhaseUserInput = False
        Exit Function
    End If
    
    ' Request output configuration
    MsgBox "Step 2 of 6: Output Configuration" & vbCrLf & vbCrLf & _
           "Please specify where to write the generated forecast.", vbInformation
    
    ' Output column
    colInput = GetColumnInput("Output column:", config.OutputCol)
    If colInput = -1 Then
        PhaseUserInput = False
        Exit Function
    End If
    config.OutputCol = colInput
    
    ' Output start row
    rowInput = GetRowInput("Output start row:", config.LearnEnd + 1)
    If rowInput = -1 Then
        PhaseUserInput = False
        Exit Function
    End If
    config.OutputStart = rowInput
    
    ' Number of rows
    Dim numInput As Variant
    numInput = GetIntegerInput("Number of 15-minute rows to generate:", config.NumNewRows, 1, 10000)
    If VarType(numInput) = vbBoolean And Not numInput Then
        PhaseUserInput = False
        Exit Function
    End If
    config.NumNewRows = CLng(numInput)
    
    ' Decimal places
    numInput = GetIntegerInput("Decimal places for output:", config.DecimalPlaces, 0, 10)
    If VarType(numInput) = vbBoolean And Not numInput Then
        PhaseUserInput = False
        Exit Function
    End If
    config.DecimalPlaces = CLng(numInput)
    
    PhaseUserInput = True
    
End Function

' ============================================================================
' PHASE 2: DATA EXTRACTION
' ============================================================================

Private Function PhaseDataExtraction(ws As Worksheet, ByRef config As ConfigSettings) As Boolean
    ' Extract and validate historical data
    
    MsgBox "Step 3 of 6: Loading historical data...", vbInformation
    
    Dim extractResult As DataExtractionResult
    extractResult = ExtractDataRange(ws, config.LookupCol, config.LearnStart, config.LearnEnd)
    
    If Not extractResult.Success Then
        MsgBox "Error loading data: " & extractResult.ErrorMessage, vbExclamation
        PhaseDataExtraction = False
        Exit Function
    End If
    
    ' Validate data
    Dim validationResult As DataExtractionResult
    validationResult = ValidateDataArray(extractResult.DataArray)
    
    If Not validationResult.Success Then
        If MsgBox(validationResult.ErrorMessage & vbCrLf & vbCrLf & "Continue anyway?", vbYesNo + vbExclamation) = vbNo Then
            PhaseDataExtraction = False
            Exit Function
        End If
    End If
    
    PhaseDataExtraction = True
    
End Function

' ============================================================================
' PHASE 3: STATISTICAL ANALYSIS
' ============================================================================

Private Function PhaseStatisticalAnalysis(ws As Worksheet, ByRef config As ConfigSettings) As Boolean
    ' Analyze historical data and learn behavior patterns
    
    MsgBox "Step 4 of 6: Analyzing historical patterns...", vbInformation
    
    Dim extractResult As DataExtractionResult
    Dim stats As StatisticsResult
    Dim trendResult As TrendResult
    Dim autoCorr As AutocorrelationResult
    Dim plungeCount As Long, spikeCount As Long
    Dim plungeDepth As Double, spikeHeight As Double
    
    ' Load data
    extractResult = ExtractDataRange(ws, config.LookupCol, config.LearnStart, config.LearnEnd)
    
    ' Calculate statistics
    stats = CalculateStatistics(extractResult.DataArray)
    trendResult = AnalyzeTrend(extractResult.DataArray)
    autoCorr = AnalyzeAutocorrelation(extractResult.DataArray)
    
    ' Detect events
    plungeCount = DetectPlunges(extractResult.DataArray, stats.Mean, stats.StDev)
    spikeCount = DetectSpikes(extractResult.DataArray, stats.Mean, stats.StDev)
    plungeDepth = CalculateAverageEventDepth(extractResult.DataArray, stats.Mean, stats.StDev, True)
    spikeHeight = CalculateAverageEventDepth(extractResult.DataArray, stats.Mean, stats.StDev, False)
    
    ' Display analysis
    Dim analysisMsg As String
    analysisMsg = "Historical Data Analysis:" & vbCrLf & vbCrLf
    analysisMsg = analysisMsg & "Observations: " & stats.Count & vbCrLf
    analysisMsg = analysisMsg & "Mean: " & Format(stats.Mean, "0.00") & vbCrLf
    analysisMsg = analysisMsg & "Std Dev: " & Format(stats.StDev, "0.00") & vbCrLf
    analysisMsg = analysisMsg & "Min: " & Format(stats.MinValue, "0.00") & vbCrLf
    analysisMsg = analysisMsg & "Max: " & Format(stats.MaxValue, "0.00") & vbCrLf & vbCrLf
    analysisMsg = analysisMsg & "Autocorrelation (lag-1): " & Format(autoCorr.Lag1, "0.000") & vbCrLf
    analysisMsg = analysisMsg & "Plunges detected: " & plungeCount & vbCrLf
    analysisMsg = analysisMsg & "Spikes detected: " & spikeCount & vbCrLf
    
    MsgBox analysisMsg, vbInformation
    
    PhaseStatisticalAnalysis = True
    
End Function

' ============================================================================
' PHASE 4: TARGET CONFIGURATION
' ============================================================================

Private Function PhaseTargetConfiguration(ws As Worksheet, ByRef config As ConfigSettings) As Boolean
    ' Request optional target value
    
    MsgBox "Step 5 of 6: Target Configuration (Optional)" & vbCrLf & vbCrLf & _
           "You can specify a final cumulative value target. Leave blank for free simulation.", vbInformation
    
    ' For now, skip target dialog
    config.TargetMode = 0
    
    PhaseTargetConfiguration = True
    
End Function

' ============================================================================
' PHASE 5: FORECAST GENERATION
' ============================================================================

Private Function PhaseForecastGeneration(ws As Worksheet, ByRef config As ConfigSettings) As Boolean
    ' Generate forecast data
    
    MsgBox "Step 6 of 6: Generating forecast...", vbInformation
    
    Dim extractResult As DataExtractionResult
    Dim stats As StatisticsResult
    Dim autoCorr As AutocorrelationResult
    Dim forecastArray() As Double
    Dim session As ForecastSession
    
    ' Load historical data
    extractResult = ExtractDataRange(ws, config.LookupCol, config.LearnStart, config.LearnEnd)
    stats = CalculateStatistics(extractResult.DataArray)
    autoCorr = AnalyzeAutocorrelation(extractResult.DataArray)
    
    ' Initialize forecast session
    Dim plungeCount As Long, spikeCount As Long
    Dim plungeDepth As Double, spikeHeight As Double
    plungeCount = DetectPlunges(extractResult.DataArray, stats.Mean, stats.StDev)
    spikeCount = DetectSpikes(extractResult.DataArray, stats.Mean, stats.StDev)
    plungeDepth = CalculateAverageEventDepth(extractResult.DataArray, stats.Mean, stats.StDev, True)
    spikeHeight = CalculateAverageEventDepth(extractResult.DataArray, stats.Mean, stats.StDev, False)
    
    ' Generate forecast
    Dim startingCumulative As Double
    Dim startingIncrement As Double
    startingCumulative = ws.Cells(config.LearnEnd, config.OutputCol).Value2
    startingIncrement = ws.Cells(config.LearnEnd, config.LookupCol).Value2
    
    If Not GenerateForecastBatch(config.NumNewRows, startingCumulative, startingIncrement, session, forecastArray) Then
        MsgBox "Error generating forecast.", vbExclamation
        PhaseForecastGeneration = False
        Exit Function
    End If
    
    ' Validate output range before writing
    Dim outputEnd As Long
    outputEnd = config.OutputStart + config.NumNewRows - 1
    Dim checkResult As ValidationResult
    checkResult = ValidateOutputRange(ws, config.OutputCol, config.OutputStart, outputEnd, config.LearnEnd)
    
    If Not checkResult.IsValid Then
        MsgBox "Output validation failed: " & checkResult.ErrorMessage, vbExclamation
        PhaseForecastGeneration = False
        Exit Function
    End If
    
    ' Ask for confirmation if overwriting
    If checkResult.ErrorCode = 3002 Then
        If MsgBox(checkResult.ErrorMessage & vbCrLf & vbCrLf & "Overwrite?", vbYesNo + vbExclamation) = vbNo Then
            PhaseForecastGeneration = False
            Exit Function
        End If
    End If
    
    ' Write data
    If Not WriteOutputData(ws, config.OutputCol, config.OutputStart, forecastArray, config.DecimalPlaces) Then
        MsgBox "Error writing output data.", vbCritical
        PhaseForecastGeneration = False
        Exit Function
    End If
    
    PhaseForecastGeneration = True
    
End Function

' ============================================================================
' PHASE 6: OUTPUT AND CHARTING
' ============================================================================

Private Function PhaseOutputAndChart(ws As Worksheet, ByRef config As ConfigSettings) As Boolean
    ' Create chart and final output
    
    Dim chartResponse As Variant
    chartResponse = GetYesNoResponse("Create a chart of the forecast?", True)
    
    If Not IsNull(chartResponse) And chartResponse Then
        Dim chartTitle As String
        chartTitle = "15-Minute Flow Rate Forecast"
        
        If Not CreateOrUpdateChart(ws, config.OutputCol, config.OutputStart, _
                                   config.OutputStart + config.NumNewRows - 1, chartTitle, "Forecast") Then
            MsgBox "Warning: Could not create chart.", vbExclamation
        End If
    End If
    
    PhaseOutputAndChart = True
    
End Function

' ============================================================================
' EXCEL STATE MANAGEMENT
' ============================================================================

Private Type ExcelApplicationState
    ScreenUpdating As Boolean
    Calculation As XlCalculationMode
    EnableEvents As Boolean
End Type

Private Sub SaveExcelState(ByRef state As ExcelApplicationState)
    With state
        .ScreenUpdating = Application.ScreenUpdating
        .Calculation = Application.Calculation
        .EnableEvents = Application.EnableEvents
    End With
End Sub

Private Sub DisableExcelUpdates()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub RestoreExcelState(state As ExcelApplicationState)
    With state
        Application.ScreenUpdating = .ScreenUpdating
        Application.Calculation = .Calculation
        Application.EnableEvents = .EnableEvents
    End With
End Sub
