' https://social.msdn.microsoft.com/forums/en-US/9f6891f2-d0c4-47a6-b63f-48405aae4022/powerpoint-run-macro-on-timer

Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As LongPtr) As LongPtr

Dim lngTimerID As LongPtr
Dim blnTimer As Boolean

Sub StartOnTime(milliSecs As Long, ByVal lpTimerFunc As LongPtr)
    'MsgBox "StartOnTime" + Str(blnTimer)

    If blnTimer Then
        lngTimerID = KillTimer(0, lngTimerID)
        If lngTimerID = 0 Then
            MsgBox "Error : Timer Not Stopped"
            Exit Sub
        End If
        blnTimer = False
    Else
        lngTimerID = SetTimer(0, 0, milliSecs, lpTimerFunc)
        If lngTimerID = 0 Then
            MsgBox "Error : Timer Not Generated"
            Exit Sub
        End If
        blnTimer = True
    End If
End Sub

Sub RestartOnTime(milliSecs As Long, ByVal lpTimerFunc As LongPtr)
    'MsgBox "RestartOnTime" + Str(blnTimer)

    If blnTimer Then
        lngTimerID = KillTimer(0, lngTimerID)
        If lngTimerID = 0 Then
            MsgBox "Error : Timer Not Stopped"
            Exit Sub
        End If
    End If
    
    lngTimerID = SetTimer(0, 0, milliSecs, lpTimerFunc)
    If lngTimerID = 0 Then
        MsgBox "Error : Timer Not Generated"
        Exit Sub
    End If
    blnTimer = True
End Sub

Sub KillOnTime()
    lngTimerID = KillTimer(0, lngTimerID)
    blnTimer = False
End Sub

