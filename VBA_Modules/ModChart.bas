Attribute VB_Name = "ModChart"
' ============================================================================
' MODULE: ModChart
' Chart creation, management, and display
' ============================================================================

Option Explicit

Private Const CHART_NAME As String = "ForecastChart"
Private Const CHART_WIDTH As Double = 400
Private Const CHART_HEIGHT As Double = 250

' ============================================================================
' CHART CREATION AND MANAGEMENT
' ============================================================================

Public Function CreateOrUpdateChart(ws As Worksheet, dataCol As Long, dataStart As Long, _
                                     dataEnd As Long, chartTitle As String, _
                                     seriesName As String) As Boolean
    ' Create or update a chart showing forecast data
    
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim dataRange As String
    Dim existingChart As ChartObject
    
    On Error GoTo ErrorHandler
    
    ' Check if chart already exists
    Set existingChart = FindChart(ws)
    
    If Not existingChart Is Nothing Then
        ' Ask user whether to replace or create new
        Dim userChoice As VbMsgBoxResult
        userChoice = MsgBox("A chart already exists. Replace it?", vbYesNoCancel + vbQuestion)
        
        Select Case userChoice
            Case vbYes
                existingChart.Delete
            Case vbNo
                ' Create new chart with different name
                Set chartObj = ws.ChartObjects.Add(100, 100, CHART_WIDTH, CHART_HEIGHT)
            Case vbCancel
                CreateOrUpdateChart = False
                Exit Function
        End Select
    End If
    
    ' Create chart if not already done
    If chartObj Is Nothing Then
        Set chartObj = ws.ChartObjects.Add(100, 100, CHART_WIDTH, CHART_HEIGHT)
    End If
    
    Set chart = chartObj.Chart
    
    ' Build data range reference
    dataRange = ColumnIndexToLetter(dataCol) & dataStart & ":" & ColumnIndexToLetter(dataCol) & dataEnd
    
    ' Configure chart
    With chart
        .ChartType = xlLineMarkers
        .HasLegend = True
        .Legend.Position = xlRight
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
    End With
    
    ' Add series
    With chart.SeriesCollection
        If .Count > 0 Then .Item(1).Delete
        .NewSeries
        With .Item(1)
            .Name = seriesName
            .Values = ws.Range(dataRange)
            .XValues = ws.Range(ColumnIndexToLetter(dataCol - 1) & dataStart & ":" & _
                               ColumnIndexToLetter(dataCol - 1) & dataEnd)
        End With
    End With
    
    ' Format axes
    With chart.Axes(xlValue, xlPrimary)
        .HasMajorGridlines = True
        .HasMinorGridlines = False
    End With
    
    CreateOrUpdateChart = True
    Exit Function
    
ErrorHandler:
    CreateOrUpdateChart = False
    
End Function

Public Function CreateMultiSeriesChart(ws As Worksheet, seriesData As Variant, _
                                       chartTitle As String, yAxisLabel As String) As Boolean
    ' Create chart with multiple series (oil, gas, etc.)
    
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim existingChart As ChartObject
    
    On Error GoTo ErrorHandler
    
    ' Look for existing chart
    Set existingChart = FindChart(ws)
    
    If Not existingChart Is Nothing Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Replace existing chart?", vbYesNoCancel + vbQuestion)
        If response = vbYes Then
            existingChart.Delete
        ElseIf response = vbCancel Then
            CreateMultiSeriesChart = False
            Exit Function
        End If
    End If
    
    ' Create new chart
    Set chartObj = ws.ChartObjects.Add(100, 100, CHART_WIDTH * 1.2, CHART_HEIGHT)
    Set chart = chartObj.Chart
    
    With chart
        .ChartType = xlLineMarkers
        .HasLegend = True
        .Legend.Position = xlRight
        .HasTitle = True
        .ChartTitle.Text = chartTitle
    End With
    
    ' Add series (implementation depends on seriesData structure)
    CreateMultiSeriesChart = True
    Exit Function
    
ErrorHandler:
    CreateMultiSeriesChart = False
    
End Function

' ============================================================================
' CHART FINDING AND MANIPULATION
' ============================================================================

Private Function FindChart(ws As Worksheet) As ChartObject
    ' Find existing forecast chart on the worksheet
    
    Dim chartObj As ChartObject
    
    On Error Resume Next
    Set FindChart = ws.ChartObjects(CHART_NAME)
    On Error GoTo 0
    
    If FindChart Is Nothing Then
        ' Try to find any chart
        If ws.ChartObjects.Count > 0 Then
            Set FindChart = ws.ChartObjects(1)
        End If
    End If
    
End Function

Public Function DeleteChart(ws As Worksheet) As Boolean
    ' Remove the forecast chart
    
    Dim chartObj As ChartObject
    
    On Error Resume Next
    Set chartObj = ws.ChartObjects(CHART_NAME)
    On Error GoTo 0
    
    If Not chartObj Is Nothing Then
        chartObj.Delete
        DeleteChart = True
    Else
        DeleteChart = False
    End If
    
End Function

Public Sub PositionChart(ws As Worksheet, anchorColumn As Long, anchorRow As Long)
    ' Position chart at specified location
    
    Dim chartObj As ChartObject
    Dim topCell As Range
    
    On Error Resume Next
    Set chartObj = FindChart(ws)
    On Error GoTo 0
    
    If Not chartObj Is Nothing Then
        Set topCell = ws.Cells(anchorRow, anchorColumn)
        With chartObj
            .Left = topCell.Left
            .Top = topCell.Top
        End With
    End If
    
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

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
