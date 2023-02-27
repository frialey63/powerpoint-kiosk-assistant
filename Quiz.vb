Private Const filePath As String = "c:\stage\slide_changed.txt"

Private Const MILLI_SECS As Long = 10000

Sub OnSlideShowPageChange(ByVal SSW As SlideShowWindow)
    WriteSlideChangeFile SSW
         
    TimeUtils.RestartOnTime MILLI_SECS, AddressOf GotoSlideOne
End Sub

Sub OnSlideShowTerminate(ByVal oWindow As SlideShowWindow)
    TimeUtils.KillOnTime
End Sub

' TODO could be used for stats?
Private Sub WriteSlideChangeFile(ByVal SSW As SlideShowWindow)
    Open filePath For Output As 1
        Print #1, Format(Now(), "HH:mm:ss") + " " + Application.ActivePresentation.name + Str(SSW.View.CurrentShowPosition) + " EOL"
    Close #1
End Sub

Private Sub GotoSlideOne()
    For Each pres In Presentations
        If IsMerged(pres.name) Then
            pres.SlideShowSettings.Run
        End If
    Next
    
    For Each ssWin In SlideShowWindows
        If IsMerged(ssWin.Presentation.name) Then
            With ssWin
               .View.GotoSlide 1
               .Activate
            End With
        End If
    Next
End Sub

Function IsMerged(name As String) As Boolean
    If name = "Merged.pptm" Or name = "Merged.ppsm" Then
        IsMerged = True
    Else
        IsMerged = False
    End If
End Function

