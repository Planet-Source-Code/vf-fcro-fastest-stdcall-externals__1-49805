Attribute VB_Name = "Test"



'Redirected Copy Memory & MessageBoxW
Public Sub CopyMemory(ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
End Sub

'For string operation use Word type of Export function (UNICODE type)
Public Function MessageBox_W(ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
End Function




Sub Main()
Dim HModule As Long
Dim Proc As Long
'REDIRECT COPYMEMORY
HModule = GetModuleHandle("kernel32.dll")
Proc = GetProcAddress(HModule, "RtlMoveMemory")
Redirect AddressOf CopyMemory, Proc

'REDIRECT MESSAGEBOX_W (UNICODE)
HModule = GetModuleHandle("user32.dll")
Proc = GetProcAddress(HModule, "MessageBoxW")
Redirect AddressOf MessageBox_W, Proc

Form1.Show
End Sub
