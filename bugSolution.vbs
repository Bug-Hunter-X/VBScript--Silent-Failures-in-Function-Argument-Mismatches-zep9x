Option Explicit

' Function to demonstrate the problem
Function CalculateSum(num1, num2)
  If IsMissing(num1) Or IsMissing(num2) Then
    Err.Raise 1, , "Both arguments are required for CalculateSum function."
  ElseIf IsNumeric(num1) = False Or IsNumeric(num2) = False Then
    Err.Raise 2, , "Both arguments must be numbers for CalculateSum function."
  Else
    CalculateSum = num1 + num2
  End If
End Function

'Example usage:
On Error GoTo ErrHandler
Dim result
result = CalculateSum(5, 10) ' Correct
WScript.Echo "Correct Result: " & result
result = CalculateSum(5) 'Incorrect Number of Arguments
WScript.Echo "Incorrect Number of Arguments: " & result
result = CalculateSum("5", 10) 'Incorrect Argument Type
WScript.Echo "Incorrect Argument Type: " & result
result = CalculateSum(5, 10, 15) ' Incorrect number of arguments
WScript.Echo "Incorrect Number of Arguments: " & result
WScript.Quit

ErrHandler:
Select Case Err.Number
  Case 1: WScript.Echo "Error: " & Err.Description
  Case 2: WScript.Echo "Error: " & Err.Description
  Case Else: WScript.Echo "An unexpected error occurred."
End Select
WScript.Quit