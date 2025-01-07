Function f(a, b)
  If IsNumeric(a) And IsNumeric(b) Then
    If CDbl(a) > CDbl(b) Then
      MsgBox "a is greater than b"
    ElseIf CDbl(a) < CDbl(b) Then
      MsgBox "a is less than b"
    Else
      MsgBox "a is equal to b"
    End If
  Else
    MsgBox "Error: Invalid input types.  Please provide numeric values."
  End If
end function

'Call the function
f 10, 5
f "10", 5
f 10, "5"