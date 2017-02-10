Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Const WH_MOUSE_LL = 14
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const WM_MOUSEWHEEL = &H20A
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As MSLLHOOKSTRUCT, ByVal Source As Long, ByVal Length As Long)

Public hookId As Long
Public Direction As Boolean
Public EventRaised As Boolean
Public Type MSLLHOOKSTRUCT
    ptx As Long
    pty As Long
    deltax As Long
    deltay As Long
    time As Long
    extinfo As Long
End Type

Public Function MouseProc(ByVal ncode As Long, ByVal wp As Long, ByVal lp As Long) As Long
    Dim ll As MSLLHOOKSTRUCT
    If wp = WM_MOUSEWHEEL Then
        CopyMemory ll, lp, Len(ll)
        If ll.deltax < 0 Then
            Direction = False
        Else
            Direction = True
        End If
        EventRaised = True
    End If
    MouseProc = CallNextHookEx(hookId, ncode, wp, lp)
End Function
