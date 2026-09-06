Attribute VB_Name = "ModConfiguration"
' ============================================================================
' MODULE: ModConfiguration
' Manages saved settings, configuration persistence, and parameter defaults
' ============================================================================

Option Explicit

' Configuration storage locations
Private Const CONFIG_SHEET_NAME As String = "_Config"
Private Const LEGACY_CONFIG_CELLS As String = "Z1:Z20"

' Configuration keys (for future expansion to named ranges or config sheet)
Enum ConfigKey
    CfgLookupCol = 1
    CfgLearnStart = 2
    CfgLearnEnd = 3
    CfgOutputCol = 4
    CfgOutputStart = 5
    CfgNumNewRows = 6
    CfgDecimalPlaces = 7
    CfgPlungeFrequency = 8
    CfgOilRateCol = 9
    CfgOilCumulCol = 10
    CfgGasRateCol = 11
    CfgGasCumulCol = 12
    CfgBSWCol = 13
    CfgTargetMode = 14
    CfgTarget1 = 15
    CfgTarget2 = 16
    CfgAllowExtrapolation = 17
    CfgUseRandomSeed = 18
    CfgRandomSeed = 19
End Enum

' Configuration storage type
Public Type ConfigSettings
    ' Learning range
    LookupCol As Long               ' Primary lookup column (numeric index)
    LearnStart As Long              ' First row of learning data
    LearnEnd As Long                ' Last row of learning data
    
    ' Output configuration
    OutputCol As Long               ' Output column (numeric index)
    OutputStart As Long             ' First output row
    NumNewRows As Long              ' Number of rows to generate
    DecimalPlaces As Long           ' Decimal places for rounding
    
    ' Multi-variable support
    OilRateCol As Long              ' Oil 15-minute rate column
    OilCumulCol As Long             ' Oil cumulative column (0 = not used)
    GasRateCol As Long              ' Gas 15-minute rate column
    GasCumulCol As Long             ' Gas cumulative column (0 = not used)
    BSWCol As Long                  ' BS&W direct measurement column
    
    ' Learning and statistical parameters
    PlungeFrequency As Double       ' User-overridable plunge frequency
    
    ' Target and forecasting modes
    TargetMode As Long              ' 0=None, 1=One target, 2=Two targets
    Target1 As Double               ' Final target value
    Target2 As Double               ' Secondary target value
    
    ' Generation options
    AllowExtrapolation As Boolean   ' Allow values beyond historical range
    UseRandomSeed As Boolean        ' Use fixed seed for reproducibility
    RandomSeed As Long              ' Seed value if UseRandomSeed = True
    
End Type

Private defaultConfig As ConfigSettings

' ============================================================================
' INITIALIZATION
' ============================================================================

Public Sub InitializeDefaults()
    ' Set reasonable default values for configuration
    
    With defaultConfig
        .LookupCol = 2                  ' Column B
        .LearnStart = 2                 ' Row 2 (after header)
        .LearnEnd = 51                  ' 50 rows of learning data
        .OutputCol = 3                  ' Column C (or user's choice)
        .OutputStart = 52               ' Immediately after learning range
        .NumNewRows = 96                ' 96 intervals = 24 hours
        .DecimalPlaces = 2              ' Two decimal places
        
        .OilRateCol = 0                 ' Not used by default
        .OilCumulCol = 0                ' Not used by default
        .GasRateCol = 0                 ' Not used by default
        .GasCumulCol = 0                ' Not used by default
        .BSWCol = 0                     ' Not used by default
        
        .PlungeFrequency = 0.03         ' Approximately once every 5 hours
        .TargetMode = 0                 ' No target by default
        .Target1 = 0
        .Target2 = 0
        
        .AllowExtrapolation = True      ' Allow extrapolation by default
        .UseRandomSeed = False          ' Free randomness by default
        .RandomSeed = 0
    End With
    
End Sub

' ============================================================================
' LOAD CONFIGURATION
' ============================================================================

Public Sub LoadConfiguration(ByRef config As ConfigSettings)
    ' Load saved configuration from worksheet
    ' If configuration does not exist or is invalid, load defaults
    
    Dim ws As Worksheet
    Dim configSheet As Worksheet
    
    ' Try to get config sheet; if it doesn't exist, use active sheet
    On Error Resume Next
    Set configSheet = ActiveWorkbook.Sheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    
    If configSheet Is Nothing Then
        Set configSheet = ActiveSheet
    End If
    
    ' Load values from hidden cells Z1:Z20
    With configSheet
        config.LookupCol = CLng(GetConfigCell(.Range("Z1"), defaultConfig.LookupCol))
        config.LearnStart = CLng(GetConfigCell(.Range("Z2"), defaultConfig.LearnStart))
        config.LearnEnd = CLng(GetConfigCell(.Range("Z3"), defaultConfig.LearnEnd))
        config.OutputCol = CLng(GetConfigCell(.Range("Z4"), defaultConfig.OutputCol))
        config.OutputStart = CLng(GetConfigCell(.Range("Z5"), defaultConfig.OutputStart))
        config.NumNewRows = CLng(GetConfigCell(.Range("Z6"), defaultConfig.NumNewRows))
        config.DecimalPlaces = CLng(GetConfigCell(.Range("Z7"), defaultConfig.DecimalPlaces))
        config.PlungeFrequency = CDbl(GetConfigCell(.Range("Z8"), defaultConfig.PlungeFrequency))
        
        config.OilRateCol = CLng(GetConfigCell(.Range("Z9"), defaultConfig.OilRateCol))
        config.OilCumulCol = CLng(GetConfigCell(.Range("Z10"), defaultConfig.OilCumulCol))
        config.GasRateCol = CLng(GetConfigCell(.Range("Z11"), defaultConfig.GasRateCol))
        config.GasCumulCol = CLng(GetConfigCell(.Range("Z12"), defaultConfig.GasCumulCol))
        config.BSWCol = CLng(GetConfigCell(.Range("Z13"), defaultConfig.BSWCol))
        
        config.TargetMode = CLng(GetConfigCell(.Range("Z14"), defaultConfig.TargetMode))
        config.Target1 = CDbl(GetConfigCell(.Range("Z15"), defaultConfig.Target1))
        config.Target2 = CDbl(GetConfigCell(.Range("Z16"), defaultConfig.Target2))
        
        config.AllowExtrapolation = CBool(GetConfigCell(.Range("Z17"), IIf(defaultConfig.AllowExtrapolation, 1, 0)))
        config.UseRandomSeed = CBool(GetConfigCell(.Range("Z18"), IIf(defaultConfig.UseRandomSeed, 1, 0)))
        config.RandomSeed = CLng(GetConfigCell(.Range("Z19"), defaultConfig.RandomSeed))
    End With
    
    ' Validate loaded configuration
    ValidateConfiguration config
    
End Sub

' ============================================================================
' SAVE CONFIGURATION
' ============================================================================

Public Sub SaveConfiguration(ByRef config As ConfigSettings)
    ' Save current configuration to worksheet
    
    Dim configSheet As Worksheet
    
    ' Try to get or create config sheet
    On Error Resume Next
    Set configSheet = ActiveWorkbook.Sheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    
    If configSheet Is Nothing Then
        ' Config sheet doesn't exist; use active sheet
        Set configSheet = ActiveSheet
    End If
    
    ' Write values to hidden cells Z1:Z20
    With configSheet
        .Range("Z1").Value = config.LookupCol
        .Range("Z2").Value = config.LearnStart
        .Range("Z3").Value = config.LearnEnd
        .Range("Z4").Value = config.OutputCol
        .Range("Z5").Value = config.OutputStart
        .Range("Z6").Value = config.NumNewRows
        .Range("Z7").Value = config.DecimalPlaces
        .Range("Z8").Value = config.PlungeFrequency
        .Range("Z9").Value = config.OilRateCol
        .Range("Z10").Value = config.OilCumulCol
        .Range("Z11").Value = config.GasRateCol
        .Range("Z12").Value = config.GasCumulCol
        .Range("Z13").Value = config.BSWCol
        .Range("Z14").Value = config.TargetMode
        .Range("Z15").Value = config.Target1
        .Range("Z16").Value = config.Target2
        .Range("Z17").Value = IIf(config.AllowExtrapolation, 1, 0)
        .Range("Z18").Value = IIf(config.UseRandomSeed, 1, 0)
        .Range("Z19").Value = config.RandomSeed
        
        ' Hide the config cells
        .Range(LEGACY_CONFIG_CELLS).EntireRow.Hidden = True
    End With
    
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function GetConfigCell(cell As Range, defaultValue As Variant) As Variant
    ' Safely retrieve a configuration value from a cell
    ' If cell is empty or contains error, return default
    
    If IsEmpty(cell) Or IsError(cell) Then
        GetConfigCell = defaultValue
    Else
        GetConfigCell = cell.Value2
    End If
    
End Function

Private Sub ValidateConfiguration(ByRef config As ConfigSettings)
    ' Ensure all configuration values are within acceptable ranges
    ' Apply defaults where necessary
    
    ' Column validation
    If config.LookupCol < 1 Then config.LookupCol = defaultConfig.LookupCol
    If config.OutputCol < 1 Then config.OutputCol = defaultConfig.OutputCol
    If config.OilRateCol < 0 Then config.OilRateCol = 0
    If config.GasRateCol < 0 Then config.GasRateCol = 0
    If config.BSWCol < 0 Then config.BSWCol = 0
    
    ' Row validation
    If config.LearnStart < 1 Then config.LearnStart = defaultConfig.LearnStart
    If config.LearnEnd < config.LearnStart + 1 Then _
        config.LearnEnd = config.LearnStart + defaultConfig.LearnEnd - defaultConfig.LearnStart
    If config.OutputStart < 1 Then config.OutputStart = config.LearnEnd + 1
    
    ' Generation parameters
    If config.NumNewRows < 1 Then config.NumNewRows = defaultConfig.NumNewRows
    If config.DecimalPlaces < 0 Then config.DecimalPlaces = defaultConfig.DecimalPlaces
    If config.DecimalPlaces > 10 Then config.DecimalPlaces = 10
    
    ' Plunge frequency
    If config.PlungeFrequency < 0.005 Then config.PlungeFrequency = 0.005
    If config.PlungeFrequency > 0.15 Then config.PlungeFrequency = 0.15
    
    ' Target validation
    If config.TargetMode < 0 Or config.TargetMode > 2 Then config.TargetMode = 0
    
    ' Seed validation
    If config.RandomSeed < 0 Then config.RandomSeed = 0
    
End Sub

' ============================================================================
' CONFIGURATION INSPECTION
' ============================================================================

Public Function GetConfigSummary(config As ConfigSettings) As String
    ' Return a human-readable summary of current configuration
    
    Dim summary As String
    
    summary = "=== Forecast Configuration ===" & vbCrLf & vbCrLf
    summary = summary & "Learning Range:" & vbCrLf
    summary = summary & "  Lookup Column: " & ColumnIndexToLetter(config.LookupCol) & vbCrLf
    summary = summary & "  From Row: " & config.LearnStart & vbCrLf
    summary = summary & "  To Row: " & config.LearnEnd & vbCrLf
    summary = summary & "  Observations: " & (config.LearnEnd - config.LearnStart + 1) & vbCrLf & vbCrLf
    
    summary = summary & "Output Configuration:" & vbCrLf
    summary = summary & "  Output Column: " & ColumnIndexToLetter(config.OutputCol) & vbCrLf
    summary = summary & "  Start Row: " & config.OutputStart & vbCrLf
    summary = summary & "  Rows to Generate: " & config.NumNewRows & vbCrLf
    summary = summary & "  Decimal Places: " & config.DecimalPlaces & vbCrLf & vbCrLf
    
    summary = summary & "Multi-Variable:" & vbCrLf
    If config.OilRateCol > 0 Then
        summary = summary & "  Oil Rate: " & ColumnIndexToLetter(config.OilRateCol)
        If config.OilCumulCol > 0 Then _
            summary = summary & " (Cumulative: " & ColumnIndexToLetter(config.OilCumulCol) & ")"
        summary = summary & vbCrLf
    Else
        summary = summary & "  Oil Rate: Not used" & vbCrLf
    End If
    
    If config.GasRateCol > 0 Then
        summary = summary & "  Gas Rate: " & ColumnIndexToLetter(config.GasRateCol)
        If config.GasCumulCol > 0 Then _
            summary = summary & " (Cumulative: " & ColumnIndexToLetter(config.GasCumulCol) & ")"
        summary = summary & vbCrLf
    Else
        summary = summary & "  Gas Rate: Not used" & vbCrLf
    End If
    
    If config.BSWCol > 0 Then
        summary = summary & "  BS&W: " & ColumnIndexToLetter(config.BSWCol) & vbCrLf
    Else
        summary = summary & "  BS&W: Not used" & vbCrLf
    End If
    
    summary = summary & vbCrLf & "Target Mode: "
    Select Case config.TargetMode
        Case 0: summary = summary & "None (free simulation)" & vbCrLf
        Case 1: summary = summary & "Target 1 = " & Format(config.Target1, "0.00") & vbCrLf
        Case 2: summary = summary & "Target 1 = " & Format(config.Target1, "0.00") & _
                           ", Target 2 = " & Format(config.Target2, "0.00") & vbCrLf
    End Select
    
    summary = summary & "Plunge Frequency: " & Format(config.PlungeFrequency, "0.000") & vbCrLf
    summary = summary & "Extrapolation Allowed: " & (IIf(config.AllowExtrapolation, "Yes", "No")) & vbCrLf
    
    If config.UseRandomSeed Then
        summary = summary & "Random Seed: " & config.RandomSeed & " (Reproducible)" & vbCrLf
    Else
        summary = summary & "Random Seed: None (Variable)" & vbCrLf
    End If
    
    GetConfigSummary = summary
    
End Function

' ============================================================================
' UTILITY FUNCTION (moved from ModCore)
' ============================================================================

Public Function ColumnIndexToLetter(col As Long) As String
    ' Convert numeric column index to letter(s)
    Dim result As String
    Dim originalCol As Long
    
    originalCol = col
    
    ' Handle invalid input
    If col < 1 Then
        ColumnIndexToLetter = "?"
        Exit Function
    End If
    
    ' Convert to base-26
    result = ""
    Do While col > 0
        result = Chr(64 + ((col - 1) Mod 26) + 1) & result
        col = (col - 1) \ 26
    Loop
    
    ColumnIndexToLetter = result
    
End Function

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
