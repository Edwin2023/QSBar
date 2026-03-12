On Error Resume Next
Set obj = CreateObject("QSBar.WpsAddIn")
If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "Success! COM object created."
End If
