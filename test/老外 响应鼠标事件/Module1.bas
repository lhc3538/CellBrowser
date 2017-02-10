Attribute VB_Name = "Module1"
Public Sub sub_MouseWheel(ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single '控制每次移动几行
    Dim iA
    On Error Resume Next
    iA = 0

    With Screen.ActiveControl
        Lstep = .Height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 10 Then
            Lstep = 10
        End If
        Lstep = 1
        If Rotation > 0 Then
            NewValue = .TopRow - Lstep
            If NewValue < 1 Then
                NewValue = 1
            End If
        Else
            NewValue = .TopRow + Lstep
            If NewValue > .Rows - 1 Then
                NewValue = .Rows - 1
            End If
        End If
        
        If .Rows > .FixedRows Then
           iA = IIf(.FixedRows >= NewValue, .FixedRows, NewValue)
           If iA > .Rows Then iA = .Rows - 1
           .TopRow = iA
        End If
        
    End With
End Sub

