Function GetObjectSafe(path)
  Dim obj, errNum
  On Error Resume Next
  Set obj = GetObject(path)
  errNum = Err.Number
  On Error GoTo 0
  If errNum <> 0 Then
    WScript.Echo "Error getting object: " & Err.Description
    Set obj = Nothing
  End If
  Set GetObjectSafe = obj
End Function

'Example Usage
Set myObj = GetObjectSafe("C:\\MyFile.txt")
if myObj is Nothing then
  WScript.Echo "Could not get object!"
else
  WScript.Echo "Got object!"
  ' Remember to release the object when finished
  Set myObj = Nothing
end if