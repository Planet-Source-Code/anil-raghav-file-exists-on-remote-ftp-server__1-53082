Attribute VB_Name = "modShell"
Public Type SHELLEXECUTEINFO
   cbSize As Long
   fMask As Long
   hwnd As Long
   lpVerb As String
   lpFile As String
   lpParameters As String
   lpDirectory As String
   nShow As Long
   hInstApp As Long
   lpIDList As Long
   lpClass As String
   hkeyClass As Long
   dwHotKey As Long
   hIcon As Long
   hProcess As Long
End Type

Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As _
   SHELLEXECUTEINFO) As Long
Public Const SE_ERR_FNF = 2
Public Const SE_ERR_NOASSOC = 31
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal _
   dwMilliseconds As Long) As Long
Public Const INFINITE = &HFFFF
Public Const WAIT_TIMEOUT = &H102

Public Function SuperShell(parm As String) As Boolean
On Error Resume Next
Dim sei As SHELLEXECUTEINFO
Dim retval As Long

   With sei
       .cbSize = Len(sei)
       .fMask = SEE_MASK_NOCLOSEPROCESS
       .hwnd = frmTransliterate.hwnd
       .lpVerb = "open"
       .lpFile = parm
       .lpDirectory = directorywherethefilesare
       .nShow = SW_HIDE 'this makes the action to be executed in hide mode
   End With
   retval = ShellExecuteEx(sei)
   If retval = 0 Then
        MsgBox "Some unexpected error ocurred"
   Else
       Do
           DoEvents
           retval = WaitForSingleObject(sei.hProcess, 0)
       Loop While retval = WAIT_TIMEOUT
    End If
End Function



