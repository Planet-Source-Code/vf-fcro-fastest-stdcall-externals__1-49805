Attribute VB_Name = "PatchIt"
'REDIRECT CALLING by Vanja Fuckar
'most effective & faster way to call the EXTERNALS!
'It require & work only with compiled code @ 2003



Public Type PatchMem
Align As Integer
Data As Integer
Address As Long
HoldAddress As Long
End Type

Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal HModule As Long, ByVal lpProcName As String) As Long



Sub Redirect(ByVal FunctionAddress As Long, ByVal RedirectTo As Long)
Dim Patch As PatchMem
Patch.Data = &H25FF
Patch.Address = FunctionAddress + 6
Patch.HoldAddress = RedirectTo
WriteProcessMemory -1, ByVal FunctionAddress, Patch.Data, 10&, ByVal 0&
End Sub




