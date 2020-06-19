option explicit

Public Const CF_TEXT = 1
Public Const GHND = &H42


Declare Function lstrcpy Lib "kernel32" ( _
         ByVal lpString1 As Any, _
         ByVal lpString2 As Any) As Long


Declare Function GlobalAlloc Lib "kernel32" ( _
         ByVal wFlags As Long, _
         ByVal dwBytes As Long) As Long

 ' GlobalLock {
  '      Compare with GlobalUnlock
    Declare Function GlobalLock Lib "kernel32" ( _
         ByVal hMem As Long) As Long
  ' }

  ' GlobalUnlock {
  '     Compare with GlobalLock
    Declare Function GlobalUnlock Lib "kernel32" ( _
         ByVal hMem As Long) As Long
  ' }


 Declare Function OpenClipboard Lib "User32" ( _
         ByVal hwnd As Long) As Long
         
 Declare Function EmptyClipboard Lib "User32" () As Long



 Declare Function SetClipboardData Lib "User32" ( _
         ByVal wFormat As Long, _
         ByVal hMem As Long) As Long



 Declare Function CloseClipboard Lib "User32" () As Long


