Attribute VB_Name = "Module1"

Public Sub Timeout(duration)

    On Error Resume Next

    DoEvents
    StartTime = Timer

    Do While Timer - StartTime < duration

        DoEvents

    Loop

End Sub

Public Sub Center(frmform As Form)

    frmform.Left = (Screen.Width - frmform.Width) / 2
    frmform.Top = (Screen.Height - frmform.Height) / 2

End Sub

