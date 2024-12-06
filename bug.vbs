Function GetObject(path)
  On Error Resume Next
  Set obj = GetObject(path)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = Nothing
  End If
  Set GetObject = obj
End Function

' Example usage
Set myObj = GetObject("C:\\MyFile.txt")
if myObj is Nothing then
  WScript.Echo "Could not get object!"
else
  WScript.Echo "Got object!"
end if