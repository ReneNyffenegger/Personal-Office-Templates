option explicit

public const CF_TEXT =    1
public const GHND    = &H42

declare function lstrcpy Lib "kernel32" ( _
         byVal lpString1 as any, _
         byVal lpString2 as any) as long


declare function GlobalAlloc Lib "kernel32" ( _
         byVal wFlags  as long, _
         byVal dwBytes as long) as long

' GlobalLock {
'      Compare with GlobalUnlock
   declare function GlobalLock Lib "kernel32" ( _
        byVal hMem as long) as long
 ' }

 ' GlobalUnlock {
 '     Compare with GlobalLock
   declare function GlobalUnlock Lib "kernel32" ( _
        byVal hMem as long) as long
 ' }


declare function OpenClipboard Lib "User32" ( _
        byVal hwnd as long) as long
        

declare function EmptyClipboard Lib "User32" () as long


declare function SetClipboardData Lib "User32" ( _
        byval wformat  as long, _
        byVal hMem     as long) as long


declare function CloseClipboard Lib "User32" () as long
