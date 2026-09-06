Attribute VB_Name = "ModUserInput"
' ============================================================================
' MODULE: ModUserInput
' Request and validate user inputs with proper cancellation handling
' ============================================================================

Option Explicit

' ============================================================================
' COLUMN INPUT
' ============================================================================

Public Function GetColumnInput(prompt As String, defaultCol As Long) As Long
    ' Request a column letter from the user
    ' Returns numeric column index, or -1 if cancelled
    
    Dim colLetter As String
    Dim colIndex As Long
    Dim defaultLetter As String
    
    ' Convert default numeric index to letter for display
    defaultLetter = ColumnIndexToLetter(defaultCol)
    
    Do
        ' Use InputBox to get column letter
        colLetter = Application.InputBox(prompt, "Column Selection", _
                                        defaultLetter, Type:=2)
        
        ' Check for cancellation (empty string from user pressing Cancel)
        If colLetter = "" Then
            GetColumnInput = -1
            Exit Function
        End If
        
        ' Validate the letter
        If Not ValidateColumnLetter(colLetter) Then
            If MsgBox("'" & colLetter & "' is not a valid column letter." & vbCrLf & _
                      "Please enter a column from A to XFD.", vbRetryCancel) = vbCancel Then
                GetColumnInput = -1
                Exit Function
            End If
        Else
            ' Convert and return
            colIndex = ColumnLetterToIndex(colLetter)
            GetColumnInput = colIndex
            Exit Do
        End If
    Loop
    
End Function

Public Function GetRowInput(prompt As String, defaultRow As Long) As Long
    ' Request a row number from the user
    ' Returns row number, or -1 if cancelled
    
    Dim rowNum As Variant
    Dim defaultStr As String
    
    defaultStr = CStr(defaultRow)
    
    Do
        rowNum = Application.InputBox(prompt, "Row Selection", defaultStr, Type:=1)
        
        ' Check for cancellation (False is returned when user cancels numeric input)
        If VarType(rowNum) = vbBoolean And Not rowNum Then
            GetRowInput = -1
            Exit Function
        End If
        
        ' Validate
        If Not ValidateRowNumber(rowNum) Then
            If MsgBox("'" & CStr(rowNum) & "' is not a valid row number." & vbCrLf & _
                      "Please enter a number from 1 to " & 1048576, _
                      vbRetryCancel) = vbCancel Then
                GetRowInput = -1
                Exit Function
            End If
        Else
            GetRowInput = CLng(rowNum)
            Exit Do
        End If
    Loop
    
End Function

Public Function GetNumericInput(prompt As String, defaultVal As Double) As Variant
    ' Request a numeric value from the user
    ' Returns the value, or a False Boolean if cancelled
    
    Dim val As Variant
    Dim defaultStr As String
    
    defaultStr = Format(defaultVal, "0.00")
    
    Do
        val = Application.InputBox(prompt, "Numeric Input", defaultStr, Type:=1)
        
        ' Check for cancellation (False is returned when user cancels)
        If VarType(val) = vbBoolean And Not val Then
            GetNumericInput = False
            Exit Function
        End If
        
        ' Validate
        If Not IsNumericValue(val) Then
            If MsgBox("'" & CStr(val) & "' is not a valid number.", _
                      vbRetryCancel) = vbCancel Then
                GetNumericInput = False
                Exit Function
            End If
        Else
            GetNumericInput = CDbl(val)
            Exit Do
        End If
    Loop
    
End Function

Public Function GetIntegerInput(prompt As String, defaultVal As Long, _
                                Optional minVal As Long = 0, _
                                Optional maxVal As Long = 1000) As Variant
    ' Request an integer value with optional range checking
    ' Returns the value, or False if cancelled
    
    Dim val As Variant
    Dim intVal As Long
    Dim defaultStr As String
    
    defaultStr = CStr(defaultVal)
    
    Do
        val = Application.InputBox(prompt & " (" & minVal & " to " & maxVal & ")", _
                                   "Integer Input", defaultStr, Type:=1)
        
        ' Check for cancellation
        If VarType(val) = vbBoolean And Not val Then
            GetIntegerInput = False
            Exit Function
        End If
        
        ' Validate
        If Not IsNumericValue(val) Then
            If MsgBox("'" & CStr(val) & "' is not a valid number.", vbRetryCancel) = vbCancel Then
                GetIntegerInput = False
                Exit Function
            End If
        Else
            intVal = CLng(val)
            If intVal < minVal Or intVal > maxVal Then
                If MsgBox("Value must be between " & minVal & " and " & maxVal & ".", _
                         vbRetryCancel) = vbCancel Then
                    GetIntegerInput = False
                    Exit Function
                End If
            Else
                GetIntegerInput = intVal
                Exit Do
            End If
        End If
    Loop
    
End Function

Public Function GetDecimalInput(prompt As String, defaultVal As Double, _
                                Optional minVal As Double = 0, _
                                Optional maxVal As Double = 100) As Variant
    ' Request a decimal value with optional range checking
    ' Returns the value, or False if cancelled
    
    Dim val As Variant
    Dim dblVal As Double
    Dim defaultStr As String
    
    defaultStr = Format(defaultVal, "0.00")
    
    Do
        val = Application.InputBox(prompt & " (" & Format(minVal, "0.00") & " to " & Format(maxVal, "0.00") & ")", _
                                   "Decimal Input", defaultStr, Type:=1)
        
        ' Check for cancellation
        If VarType(val) = vbBoolean And Not val Then
            GetDecimalInput = False
            Exit Function
        End If
        
        ' Validate
        If Not IsNumericValue(val) Then
            If MsgBox("'" & CStr(val) & "' is not a valid number.", vbRetryCancel) = vbCancel Then
                GetDecimalInput = False
                Exit Function
            End If
        Else
            dblVal = CDbl(val)
            If dblVal < minVal Or dblVal > maxVal Then
                If MsgBox("Value must be between " & Format(minVal, "0.00") & " and " & Format(maxVal, "0.00") & ".", _
                         vbRetryCancel) = vbCancel Then
                    GetDecimalInput = False
                    Exit Function
                End If
            Else
                GetDecimalInput = dblVal
                Exit Do
            End If
        End If
    Loop
    
End Function

Public Function GetYesNoResponse(prompt As String, Optional defaultYes As Boolean = True) As Variant
    ' Request a Yes/No response from the user
    ' Returns True for Yes, False for No, or Null if cancelled
    
    Dim buttonClicked As VbMsgBoxResult
    Dim defaultButton As VbMsgBoxStyle
    
    If defaultYes Then
        defaultButton = vbDefaultButton1
    Else
        defaultButton = vbDefaultButton2
    End If
    
    buttonClicked = MsgBox(prompt, vbYesNoCancel + defaultButton + vbQuestion)
    
    Select Case buttonClicked
        Case vbYes
            GetYesNoResponse = True
        Case vbNo
            GetYesNoResponse = False
        Case vbCancel
            GetYesNoResponse = Null
    End Select
    
End Function

' ============================================================================
' MULTI-VARIABLE CONFIGURATION DIALOG
' ============================================================================

Public Sub RequestMultiVariableConfiguration(ByRef config As ConfigSettings)
    ' Guide user through selecting which variables to forecast
    
    Dim response As Variant
    Dim selectedOil As Boolean, selectedGas As Boolean, selectedBSW As Boolean
    
    ' Ask about each variable
    response = GetYesNoResponse("Do you want to generate oil production data?", True)
    If IsNull(response) Then Exit Sub
    selectedOil = response
    
    If selectedOil Then
        response = GetYesNoResponse("Do you have an oil cumulative column (in addition to rates)?", False)
        If Not IsNull(response) Then
            If response Then
                Dim oilCumulCol As Long
                oilCumulCol = GetColumnInput("Enter oil cumulative column:", 4)
                If oilCumulCol <> -1 Then
                    config.OilCumulCol = oilCumulCol
                End If
            End If
        End If
    End If
    
    response = GetYesNoResponse("Do you want to generate gas production data?", False)
    If IsNull(response) Then Exit Sub
    selectedGas = response
    
    If selectedGas Then
        response = GetYesNoResponse("Do you have a gas cumulative column (in addition to rates)?", False)
        If Not IsNull(response) Then
            If response Then
                Dim gasCumulCol As Long
                gasCumulCol = GetColumnInput("Enter gas cumulative column:", 5)
                If gasCumulCol <> -1 Then
                    config.GasCumulCol = gasCumulCol
                End If
            End If
        End If
    End If
    
    response = GetYesNoResponse("Do you want to generate BS&W data?", False)
    If IsNull(response) Then Exit Sub
    selectedBSW = response
    
    If Not (selectedOil Or selectedGas Or selectedBSW) Then
        MsgBox "At least one variable must be selected.", vbExclamation
        RequestMultiVariableConfiguration config
    End If
    
End Sub

' ============================================================================
' TARGET CONFIGURATION DIALOG
' ============================================================================

Public Sub RequestTargetConfiguration(ByRef config As ConfigSettings, currentValue As Double, _
                                      mean As Double, stdev As Double, numRows As Long)
    ' Guide user through setting target values
    
    Dim response As Variant
    Dim targetVal As Variant
    Dim validationResult As ValidationResult
    
    config.TargetMode = 0  ' Default: no target
    
    ' Ask if user wants to use a target
    response = GetYesNoResponse("Do you want to reach a specific final cumulative value?", False)
    
    If IsNull(response) Then
        Exit Sub
    ElseIf Not response Then
        Exit Sub  ' No target
    End If
    
    ' Get Target 1
    Do
        targetVal = GetNumericInput("Enter Target 1 (final cumulative value):", currentValue + numRows * mean)
        
        If VarType(targetVal) = vbBoolean And Not targetVal Then
            Exit Sub  ' Cancelled
        End If
        
        validationResult = ValidateTarget(targetVal, currentValue, mean, stdev, numRows)
        
        If Not validationResult.IsValid Then
            If MsgBox(validationResult.ErrorMessage & vbCrLf & vbCrLf & "Continue anyway?", vbYesNo + vbExclamation) = vbNo Then
                ' Retry
            Else
                ' Accept target
                config.TargetMode = 1
                config.Target1 = CDbl(targetVal)
                Exit Do
            End If
        Else
            config.TargetMode = 1
            config.Target1 = CDbl(targetVal)
            Exit Do
        End If
    Loop
    
    ' Ask if user wants to use a second target
    If config.TargetMode = 1 Then
        response = GetYesNoResponse("Do you want to specify a second intermediate target?", False)
        
        If Not IsNull(response) And response Then
            Dim target2Val As Variant
            target2Val = GetNumericInput("Enter Target 2 (intermediate cumulative value):", currentValue + numRows * mean * 0.5)
            
            If Not (VarType(target2Val) = vbBoolean And Not target2Val) Then
                config.TargetMode = 2
                config.Target2 = CDbl(target2Val)
            End If
        End If
    End If
    
End Sub

' ============================================================================
' HELPER FUNCTIONS
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
