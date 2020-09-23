Attribute VB_Name = "Module1"
Declare Function mciSendString Lib "WINMM.DLL" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Sub PlaySound(StrPath As String)
On Error Resume Next
Call mciSendString("play " & StrPath, 0&, 0, 0)
End Sub
Sub StopSound(StrPath As String)
On Error Resume Next
Call mciSendString("stop " & StrPath, 0&, 0, 0)
End Sub

