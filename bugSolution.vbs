Function f(a, b)
  If a = " "" Then 
    WScript.Echo "a is empty"
    Exit Function
  End If
  If b = " "" Then
    WScript.Echo "b is empty"
    Exit Function
  End If
  f = CDbl(a) + CDbl(b) 
End Function

WScript.Echo f(1,2)
WScript.Echo f(1, "")
WScript.Echo f("",2)
WScript.Echo f("","") 