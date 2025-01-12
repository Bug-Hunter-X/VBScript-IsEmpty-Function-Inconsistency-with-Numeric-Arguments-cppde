Function f(a,b)
  If IsEmpty(a) Then
    WScript.Echo "a is empty"
    Exit Function
  End If
  If IsEmpty(b) Then
    WScript.Echo "b is empty"
    Exit Function
  End If 
  f = a+b
end function

WScript.Echo f(1,2)