Attribute VB_Name = "ModErrorHandling"
' ============================================================================
' MODULE: ModErrorHandling
' Global error handling and logging
' ============================================================================

Option Explicit

Private Const LOG_SHEET_NAME As String = "_ErrorLog"
Private logEnabled As Boolean
Private logSheetCreated As Boolean

' Error logging type
Public Type ErrorLogEntry
    Timestamp As Date
    ErrorNumber As Long
    ErrorDescription As String
    ErrorSource As String
    UserAction As String
End Type

' ============================================================================
' ERROR HANDLING
' ============================================================================

Public Sub EnableErrorLogging()
    logEnabled = True
End Sub

Public Sub DisableErrorLogging()
    logEnabled = False
End Sub

Public Sub LogError(errNumber As Long, errDescription As String, errSource As String, _
                    Optional userAction As String = "")
    ' Log an error to the error log sheet or message box
    
    Dim errorMsg As String
    
    errorMsg = "Error #" & errNumber & ": " & errDescription & vbCrLf
    errorMsg = errorMsg & "Source: " & errSource & vbCrLf
    
    If userAction <> "" Then
        errorMsg = errorMsg & "During: " & userAction & vbCrLf
    End If
    
    If logEnabled Then
        ' Try to write to error log
        On Error Resume Next
        WriteErrorLog errNumber, errDescription, errSource, userAction
        On Error GoTo 0
    End If
    
End Sub

Private Sub WriteErrorLog(errNumber As Long, errDescription As String, errSource As String, _
                         userAction As String)
    ' Write error to a hidden log sheet
    
    Dim logSheet As Worksheet
    Dim logRow As Long
    
    ' Try to get or create log sheet
    On Error Resume Next
    Set logSheet = ActiveWorkbook.Sheets(LOG_SHEET_NAME)
    On Error GoTo 0
    
    If logSheet Is Nothing Then
        Set logSheet = ActiveWorkbook.Sheets.Add
        logSheet.Name = LOG_SHEET_NAME
        logSheet.Visible = xlSheetHidden
        logSheetCreated = True
    End If
    
    ' Find last row
    logRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Write log entry
    With logSheet
        .Cells(logRow, 1).Value = Now
        .Cells(logRow, 2).Value = errNumber
        .Cells(logRow, 3).Value = errDescription
        .Cells(logRow, 4).Value = errSource
        .Cells(logRow, 5).Value = userAction
    End With
    
End Sub

' ============================================================================
' VALIDATION ERROR REPORTING
' ============================================================================

Public Sub ReportValidationError(result As ValidationResult)
    ' Display validation error to user
    
    Dim message As String
    message = result.ErrorMessage
    
    If result.ErrorCode > 0 Then
        message = message & vbCrLf & vbCrLf & "(Error code: " & result.ErrorCode & ")"
    End If
    
    MsgBox message, vbExclamation
    
End Sub

' ============================================================================
' USER-FRIENDLY ERROR MESSAGES
' ============================================================================

Public Function GetUserFriendlyErrorMessage(errorCode As Long) As String
    ' Convert error code to user-friendly message
    
    Select Case errorCode
        Case 1001
            GetUserFriendlyErrorMessage = "Learning start row must be at least row 1."
        Case 1002
            GetUserFriendlyErrorMessage = "Learning end row must be after the start row."
        Case 1003
            GetUserFriendlyErrorMessage = "The learning range does not contain enough data. Please select a larger range."
        Case 2001
            GetUserFriendlyErrorMessage = "The selected range contains empty cells. Please remove empty rows or select a different range."
        Case 2002
            GetUserFriendlyErrorMessage = "The selected range contains error values. Please fix errors or select a different range."
        Case 2003
            GetUserFriendlyErrorMessage = "The selected range contains non-numeric data. Please select a range with only numbers."
        Case 3001
            GetUserFriendlyErrorMessage = "The output range overlaps with the learning data. Please choose a different output location."
        Case 3002
            GetUserFriendlyErrorMessage = "The output location already contains data. You will be asked to confirm overwriting."
        Case 4001
            GetUserFriendlyErrorMessage = "Target value must be numeric."
        Case 4002, 4003
            GetUserFriendlyErrorMessage = "Target value may be unrealistic given the historical data. Consider reviewing your target."
        Case 5001
            GetUserFriendlyErrorMessage = "Invalid output range specification."
        Case 5002
            GetUserFriendlyErrorMessage = "Output range would overwrite historical learning data. Please choose a different location."
        Case Else
            GetUserFriendlyErrorMessage = "An unknown error occurred (code: " & errorCode & "). Please try again."
    End Select
    
End Function

' ============================================================================
' GRACEFUL FAILURE HANDLING
' ============================================================================

Public Sub CancelOperation(Optional reason As String = "")
    ' Safely cancel the current operation
    
    If reason <> "" Then
        MsgBox "Operation cancelled: " & reason, vbInformation
    Else
        MsgBox "Operation cancelled by user.", vbInformation
    End If
    
End Sub

Public Function PromptToRetry(failureMessage As String) As Boolean
    ' Ask user whether to retry after a failure
    
    Dim response As VbMsgBoxResult
    response = MsgBox(failureMessage & vbCrLf & vbCrLf & "Retry?", vbYesNo + vbExclamation)
    PromptToRetry = (response = vbYes)
    
End Function
