' Enforce explicit declaration of all variables (prevents typos from creating new variables)
Option Explicit

' Declare variables for user inputs, numeric values, and threshold boundaries
Dim strInput, strBaseInput, dblNum, MinNum, MaxNum, Base

' Set the lower bound of valid scores to 0
MinNum = 0

' Set the upper bound of valid scores to 100
MaxNum = 100

' Start a loop that repeats until a valid numeric score is entered
Do
    ' Prompt the user for a score and store the text response in strInput
    strInput = InputBox("Enter a score between 0 and 100:", "Score Check")
    
    ' Check if the user pressed Cancel or left the input box empty
    If strInput = "" Then
        ' Inform the user that the operation was cancelled
        MsgBox "No input provided. Exiting.", vbInformation, "Cancelled"
        
        ' Terminate the script immediately before processing further
        WScript.Quit
    End If
    
' Repeat the Do block as long as strInput is NOT numeric
Loop Until IsNumeric(strInput)

' Convert the verified score string into a double-precision number
dblNum = CDbl(strInput)

' Check if the score is outside the allowed 0-100 range before asking for a passing threshold
If dblNum < MinNum Or dblNum > MaxNum Then
    ' Display an out-of-range warning with the score and valid bounds
    MsgBox "Score " & dblNum & " is out of range (" & MinNum & "-" & MaxNum & ").", vbExclamation, "Out of Range"
    
    ' Stop the script since there is no valid score to evaluate against a base
    WScript.Quit
End If

' Start a loop that repeats until a valid numeric passing base is entered
Do
    ' Prompt the user in English to define the passing threshold (base)
    strBaseInput = InputBox("Enter the passing threshold (base) score:", "Passing Threshold")
    
    ' Check if the user pressed Cancel or left the base input box empty
    If strBaseInput = "" Then
        ' Inform the user that no passing threshold was provided
        MsgBox "No passing threshold provided. Exiting.", vbInformation, "Cancelled"
        
        ' Terminate the script immediately
        WScript.Quit
    End If
    
' Repeat the Do block as long as strBaseInput is NOT numeric
Loop Until IsNumeric(strBaseInput)

' Convert the verified threshold string into a double-precision number
Base = CDbl(strBaseInput)

' Compare the score against the user-defined passing threshold
If dblNum >= Base Then
    ' Display a passing result message including the chosen base
    MsgBox "Score: " & dblNum & " - PASS (Base: " & Base & ")", vbInformation, "Result"
Else
    ' Display a failing result message including the chosen base
    MsgBox "Score: " & dblNum & " - FAIL (Base: " & Base & ")", vbExclamation, "Result"
End If