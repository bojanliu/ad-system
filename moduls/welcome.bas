Attribute VB_Name = "welcome"
Option Explicit

Sub b()
    Application.OnTime Now + TimeValue("00:00:01"), "a"
End Sub

Sub a()
    »¶Σ­΄°Με.time.Caption = »¶Σ­΄°Με.time.Caption - 1
    If »¶Σ­΄°Με.time.Caption - 1 >= 0 Then
        Call b
    Else
        »¶Σ­΄°Με.Hide
        Φχ΄°Με.Show
    End If
End Sub
