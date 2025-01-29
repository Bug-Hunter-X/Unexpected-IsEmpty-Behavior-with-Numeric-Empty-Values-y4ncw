Function f(a, b)
  If VarType(a) = vbEmpty Then
    a = 0
  End If
  If VarType(b) = vbEmpty Then
    b = 0
  End If
  result = a + b
End Function

MsgBox f(1,Empty)