' Triggered on Value change and on OnStartRunning to keep SimulationModeEnabled
' coherent if the InternalTag is reset by a Studio change.
Sub SimulationMode_OnChange()
    ScanFolder Application.GetObject("XATM_Data.Substation"), (Value = True)
End Sub

' Returns True if at least one position tag has good quality — breaker does NOT need sim mode.
Function AnyBreakerTagHealthy(breaker)
    If TagHealthy(breaker.PositionOpen)      Then AnyBreakerTagHealthy = True : Exit Function
    If TagHealthy(breaker.PositionClosed)    Then AnyBreakerTagHealthy = True : Exit Function
    If TagHealthy(breaker.PositionOpenAlt)   Then AnyBreakerTagHealthy = True : Exit Function
    If TagHealthy(breaker.PositionClosedAlt) Then AnyBreakerTagHealthy = True : Exit Function
    AnyBreakerTagHealthy = False
End Function

Function TagHealthy(tag)
    On Error Resume Next
    Dim q
    q = tag.Quality
    TagHealthy = (Err.Number = 0) And (q >= 192)
    On Error GoTo 0
End Function

Sub ScanFolder(folder, simEnabled)
    Dim item
    For Each item In folder
        Select Case TypeName(item)
            Case "XATM_Breaker"
                Dim shouldSim
                If simEnabled Then
                    shouldSim = Not AnyBreakerTagHealthy(item)
                Else
                    shouldSim = False
                End If
                item.Item("Data").Item("SimulationModeEnabled").WriteEx shouldSim
            Case Else
                On Error Resume Next
                ScanFolder item, simEnabled
                On Error GoTo 0
        End Select
    Next
End Sub
