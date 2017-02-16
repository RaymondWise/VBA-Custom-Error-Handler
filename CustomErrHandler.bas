Attribute VB_Name = "CustomErrHandler"
Option Explicit
Public Enum CustomError
    CustomError1 = vbObjectError + 42
    CustomError2 = vbObjectError + 43
    CustomError3 = vbObjectError + 44
    CustomError4 = vbObjectError + 45
    CustomError5 = vbObjectError + 46
End Enum
Public Sub CustomErrorHandler(Err As Object)
    Select Case Err.Number
        Case CustomError.CustomError1
            MsgBox "Custom Error Message 1", vbExclamation
        
        Case CustomError.CustomError2
            MsgBox "Custom Error Message 2", vbExclamation
        
        Case CustomError.CustomError3
            MsgBox "Custom Error Message 3", vbExclamation
        
        Case CustomError.CustomError4
            MsgBox "Custom Error Message 4", vbExclamation
        
        Case CustomError.CustomError5
            MsgBox "Custom Error Message 5", vbExclamation
        
        Case Else
            MsgBox "Unexpected Error: " & Err.Number & "- " & Err.Description, vbCritical
    End Select
End Sub

Public Sub Void()
    'Example of how to use the custom handler
    Application.ScreenUpdating = False
    On Error GoTo CleanFail
    
    Dim x As Long
    x = 1
    Dim y As Long
    y = 2
    
    If x = y Then Err.Raise CustomError.CustomError1
    
    
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    CustomErrorHandler Err
    Resume CleanExit
End Sub
