Attribute VB_Name = "PublicDim"
Public SearchNum As Integer
Public SetStr

Public Sub LoadBrowserSet()
Dim cw As String
Dim t3 As String
Open App.path & "\set.dat" For Input As #1
  Do While Not EOF(1)
    Line Input #1, cw$
    t3 = cw
    Loop
Close #1
SetStr = Split(t3, ",")
End Sub
