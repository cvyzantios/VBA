' Enforce explicit declaration of all variables (prevents typos from creating new variables)
Option Explicit

' Declare variables for raw user input, numeric score, and threshold boundaries
Dim strInput, dblNum, MinNum, MaxNum, Base

' Set the lower bound of valid scores to 0
MinNum = 0

' Set the upper bound of valid scores to 100
MaxNum = 100

' Set the passing threshold score to 50
Base = 50

' Start a loop that repeats until a valid numeric value is entered
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

' Convert the verified string input into a double-precision number
dblNum = CDbl(strInput)

' Check if the score is strictly below MinNum OR strictly above MaxNum
If dblNum < MinNum Or dblNum > MaxNum Then
    ' Display an out-of-range warning with the score and valid bounds
    MsgBox "Score " & dblNum & " is out of range (" & MinNum & "-" & MaxNum & ").", vbExclamation, "Out of Range"
    
' Otherwise, if the score is greater than or equal to the passing threshold
ElseIf dblNum >= Base Then
    ' Display a passing result message
    MsgBox "Score: " & dblNum & " - PASS", vbInformation, "Result"
    
' For any remaining valid score below Base (0 up to 49.99...)
Else
    ' Display a failing result message
    MsgBox "Score: " & dblNum & " - FAIL", vbExclamation, "Result"
    
' End the conditional block
End If