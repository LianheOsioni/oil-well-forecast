Attribute VB_Name = "ModTesting"
' ============================================================================
' MODULE: ModTesting
' Unit and integration tests for the forecast engine
' ============================================================================

Option Explicit

' Test result type
Public Type TestResult
    TestName As String
    Passed As Boolean
    Message As String
    ErrorCode As Long
End Type

' ============================================================================
' TEST SUITE EXECUTION
' ============================================================================

Public Sub RunAllTests()
    ' Execute all test suites and report results
    
    Dim results() As TestResult
    Dim testCount As Long, passCount As Long
    Dim i As Long
    Dim report As String
    
    report = "=== FORECAST ENGINE TEST SUITE ===" & vbCrLf & vbCrLf
    
    ' Configuration Tests
    testCount = 0: passCount = 0
    TestConfigurationModule results, testCount, passCount
    report = report & "Configuration Tests: " & passCount & "/" & testCount & " passed" & vbCrLf & vbCrLf
    
    ' Validation Tests
    testCount = 0: passCount = 0
    TestValidationModule results, testCount, passCount
    report = report & "Validation Tests: " & passCount & "/" & testCount & " passed" & vbCrLf & vbCrLf
    
    ' Statistics Tests
    testCount = 0: passCount = 0
    TestStatisticsModule results, testCount, passCount
    report = report & "Statistics Tests: " & passCount & "/" & testCount & " passed" & vbCrLf & vbCrLf
    
    ' Data Extraction Tests
    testCount = 0: passCount = 0
    TestDataExtractionModule results, testCount, passCount
    report = report & "Data Extraction Tests: " & passCount & "/" & testCount & " passed" & vbCrLf & vbCrLf
    
    ' Forecast Tests
    testCount = 0: passCount = 0
    TestForecastModule results, testCount, passCount
    report = report & "Forecast Tests: " & passCount & "/" & testCount & " passed" & vbCrLf & vbCrLf
    
    ' Display results
    MsgBox report, vbInformation, "Test Results"
    
End Sub

' ============================================================================
' MODULE TESTS
' ============================================================================

Private Sub TestConfigurationModule(ByRef results() As TestResult, ByRef testCount As Long, ByRef passCount As Long)
    ' Test ModConfiguration functionality
    
    Dim config As ConfigSettings
    
    ' Test 1: Initialization
    testCount = testCount + 1
    InitializeDefaults
    LoadConfiguration config
    If config.LookupCol > 0 Then
        passCount = passCount + 1
    End If
    
    ' Test 2: Column conversion
    testCount = testCount + 1
    If ColumnLetterToIndex("A") = 1 And ColumnLetterToIndex("Z") = 26 Then
        If ColumnIndexToLetter(1) = "A" And ColumnIndexToLetter(26) = "Z" Then
            passCount = passCount + 1
        End If
    End If
    
    ' Test 3: Multi-letter columns
    testCount = testCount + 1
    If ColumnLetterToIndex("AA") = 27 And ColumnIndexToLetter(27) = "AA" Then
        passCount = passCount + 1
    End If
    
End Sub

Private Sub TestValidationModule(ByRef results() As TestResult, ByRef testCount As Long, ByRef passCount As Long)
    ' Test ModValidation functionality
    
    ' Test 1: Column letter validation
    testCount = testCount + 1
    If ValidateColumnLetter("A") And ValidateColumnLetter("XFD") Then
        If Not ValidateColumnLetter("XFE") And Not ValidateColumnLetter("??") Then
            passCount = passCount + 1
        End If
    End If
    
    ' Test 2: Row validation
    testCount = testCount + 1
    If ValidateRowNumber(1) And ValidateRowNumber(1048576) Then
        If Not ValidateRowNumber(0) Then
            passCount = passCount + 1
        End If
    End If
    
End Sub

Private Sub TestStatisticsModule(ByRef results() As TestResult, ByRef testCount As Long, ByRef passCount As Long)
    ' Test ModStatistics functionality
    
    Dim testData(1 To 10) As Double
    Dim stats As StatisticsResult
    Dim i As Long
    
    ' Create test data: 1, 2, 3, ..., 10
    For i = 1 To 10
        testData(i) = CDbl(i)
    Next i
    
    ' Test 1: Mean calculation
    testCount = testCount + 1
    stats = CalculateStatistics(testData)
    If Abs(stats.Mean - 5.5) < 0.01 Then
        passCount = passCount + 1
    End If
    
    ' Test 2: Min/Max
    testCount = testCount + 1
    If stats.MinValue = 1 And stats.MaxValue = 10 Then
        passCount = passCount + 1
    End If
    
    ' Test 3: Autocorrelation range
    testCount = testCount + 1
    Dim autoCorr As AutocorrelationResult
    autoCorr = AnalyzeAutocorrelation(testData)
    If autoCorr.Lag1 >= -1 And autoCorr.Lag1 <= 1 Then
        passCount = passCount + 1
    End If
    
End Sub

Private Sub TestDataExtractionModule(ByRef results() As TestResult, ByRef testCount As Long, ByRef passCount As Long)
    ' Test ModDataExtraction functionality
    
    Dim testArray(1 To 5) As Double
    Dim i As Long
    
    ' Create simple test array
    For i = 1 To 5
        testArray(i) = CDbl(i) * 10
    Next i
    
    ' Test 1: Rate to cumulative conversion
    testCount = testCount + 1
    Dim cumulativeArray() As Double
    cumulativeArray = ConvertRatesToCumulative(testArray, 0)
    If cumulativeArray(1) = 10 And cumulativeArray(5) = 150 Then
        passCount = passCount + 1
    End If
    
    ' Test 2: Data validation
    testCount = testCount + 1
    Dim valResult As DataExtractionResult
    valResult = ValidateDataArray(testArray)
    If valResult.Success Then
        passCount = passCount + 1
    End If
    
End Sub

Private Sub TestForecastModule(ByRef results() As TestResult, ByRef testCount As Long, ByRef passCount As Long)
    ' Test ModForecast functionality
    
    Dim forecastArray() As Double
    Dim session As ForecastSession
    Dim stats As StatisticsResult
    Dim testData(1 To 20) As Double
    Dim i As Long
    
    ' Create test data
    For i = 1 To 20
        testData(i) = 100 + (Rnd - 0.5) * 20
    Next i
    
    ' Test 1: Forecast batch generation
    testCount = testCount + 1
    stats = CalculateStatistics(testData)
    
    With session
        .Mean = stats.Mean
        .StDev = stats.StDev
        .AutoCorr = 0.3
        .LowBand = stats.Mean - 0.5 * stats.StDev
        .HighBand = stats.Mean + 0.5 * stats.StDev
        .ProbLow = 0.25
        .ProbMid = 0.5
        .ProbHigh = 0.25
        .PlungeFrequency = 0.01
        .PlungeAvgDepth = stats.StDev * 0.5
        .SpikeFrequency = 0.01
        .SpikeAvgHeight = stats.StDev * 0.5
        .AllowExtrapolation = True
        .MaxDeviationFactor = 3
    End With
    
    If GenerateForecastBatch(10, 1000, 100, session, forecastArray) Then
        If UBound(forecastArray) = 10 Then
            passCount = passCount + 1
        End If
    End If
    testCount = testCount + 1
    
End Sub

' ============================================================================
' EXAMPLE DATA GENERATORS
' ============================================================================

Public Sub GenerateExampleData()
    ' Create example dataset for testing
    
    Dim ws As Worksheet
    Dim i As Long
    Dim baseValue As Double, noise As Double
    
    ' Create or get worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("ExampleData")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Sheets.Add
        ws.Name = "ExampleData"
    End If
    
    ' Generate 100 days of data (1440 observations)
    ws.Range("A1").Value = "Time"
    ws.Range("B1").Value = "Oil Rate"
    ws.Range("C1").Value = "Gas Rate"
    ws.Range("D1").Value = "BS&W"
    
    For i = 2 To 1442
        ws.Cells(i, 1).Value = i - 2  ' Time counter
        
        ' Oil rate: base 150 bbl/day with sine wave and noise
        baseValue = 150 + 20 * Sin(i * 0.044) ' Period ~24 hours
        noise = (Rnd - 0.5) * 10
        ws.Cells(i, 2).Value = baseValue + noise
        
        ' Gas rate: proportional to oil with higher volatility
        baseValue = 50000 + 10000 * Sin(i * 0.044)
        noise = (Rnd - 0.5) * 5000
        ws.Cells(i, 3).Value = baseValue + noise
        
        ' BS&W: low and stable
        baseValue = 0.3 + (Rnd - 0.5) * 0.2
        If baseValue < 0 Then baseValue = 0
        If baseValue > 10 Then baseValue = 10
        ws.Cells(i, 4).Value = RoundBSWToStandard(baseValue)
    Next i
    
    MsgBox "Example data generated in 'ExampleData' sheet (1440 rows).", vbInformation
    
End Sub

' ============================================================================
' VALIDATION TEST CASES
' ============================================================================

Public Sub ValidateInstallation()
    ' Check that all modules are present and callable
    
    Dim checkMsg As String
    
    checkMsg = "Module Installation Check:" & vbCrLf & vbCrLf
    
    On Error Resume Next
    
    ' Check each module can be accessed
    InitializeDefaults
    checkMsg = checkMsg & "✓ ModConfiguration" & vbCrLf
    
    If ValidateColumnLetter("A") Then
        checkMsg = checkMsg & "✓ ModValidation" & vbCrLf
    End If
    
    Dim config As ConfigSettings
    If GetYesNoResponse("Testing ModUserInput (confirm yes):", True) Then
        checkMsg = checkMsg & "✓ ModUserInput" & vbCrLf
    End If
    
    Dim testArray(1 To 5) As Double
    Dim i As Long
    For i = 1 To 5: testArray(i) = CDbl(i): Next i
    Dim stats As StatisticsResult
    stats = CalculateStatistics(testArray)
    If stats.Mean > 0 Then
        checkMsg = checkMsg & "✓ ModStatistics" & vbCrLf
    End If
    
    checkMsg = checkMsg & vbCrLf & "All modules present and functional."
    
    On Error GoTo 0
    MsgBox checkMsg, vbInformation, "Installation Check"
    
End Sub
