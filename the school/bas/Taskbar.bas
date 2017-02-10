Attribute VB_Name = "Taskbar"
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48


Public Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Public Function GetTaskbarHeight() As Integer
Dim lRes As Long
Dim rectVal As RECT

lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

