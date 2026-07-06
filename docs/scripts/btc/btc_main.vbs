' XATM_BTC — Bus Transfer Control
' Data-driven FSM: GetTMSteps / GetNMSteps / GetTASteps return the ordered
' command array; this dispatcher iterates it with per-step confirmation and
' timeout guarding. Step content and path selection live entirely in the
' Get*Steps functions (layout file) — not here.

' ===========================================================
'  DISPATCHER
' ===========================================================

Sub Main_Main()

    Const TIMEOUT = 30  ' seconds per step (applies to every wait state)

    If Not xatm_BTC.Enabled Then
        WriteLog "[BTC] Automation not enabled."
        Exit Sub
    End If

    If xatm_BTC.Running Then Exit Sub

    Select Case Value
        Case 0  : Step01
        Case 99 : Complete
        Case Else
            If StepTimer() < TIMEOUT Then
                StepAdvance
            Else
                Timeout "Step " & (Value - 1)
            End If
    End Select

    IncrementTimer

End Sub

' ===========================================================
'  STEP 01 — Initialization
' ===========================================================

Sub Step01()
    ' TM: read impedimentos; call GetTMSteps(trigger, imp) → build step array;
    '     execute steps(0) — always a CLOSE command (make-before-break: close
    '     the backup TIE before touching the source breaker).
    '     Path step counts: TR1/TR2 Path A = 6 steps, Path B/C = 4 steps;
    '                       TR3/TR4 = 2 steps.
    '
    ' NM: call GetNMSteps(trigger, imp) → build step array;
    '     execute steps(0) — inverse sequence of TM (first command is an OPEN).
    '
    ' TA: read impedimentos; call GetTASteps(trigger, imp) → build step array;
    '     execute steps(0) — always OPEN CB[trigger] to confirm isolation;
    '     CB may already be open (tripped by protection relay); command double-checks.
    '     Path step counts: same shape as TM (2–6 steps), starts with OPENs.
    '
    ' Layout — 2T+1TIE: always 2 steps; TM = CLOSE TIE → OPEN CB;
    '                                    TA = OPEN CB → CLOSE TIE.

    Select Case xatm_BTC.AutomationType
        Case "TM"
        Case "NM"
        Case "TA"
    End Select

    ResetTimer
    Value = 1

End Sub

' ===========================================================
'  STEP ADVANCE — Generic confirm + execute next
' ===========================================================

Sub StepAdvance()
    ' Confirm the previously issued command steps(Value−1) by checking the
    ' target device's position:
    '   CLOSE command → device.Position = CLOSED
    '   OPEN  command → device.Position = OPEN
    ' If not yet confirmed: Exit Sub — wait for the next polling cycle.
    '
    ' If confirmed and more steps remain:
    '   ExecuteStep steps(Value) → ResetTimer → Value = Value + 1
    ' If confirmed and no more steps:
    '   Value = 99
    '
    ' All three modes (TM / NM / TA) follow the same confirmation rule —
    ' the difference is in the command array built by Step01, not here.
    '
    ' Layout — shorter paths (TR3, TR4, Path B/C) reach Value = 99 earlier;
    '          the TotalSteps count stored in Step01 drives the exit condition.

    ' [resolve steps(Value−1) device; check position; Exit Sub if not confirmed]

    ' [if Value < TotalSteps]:
    '   ExecuteStep steps(Value)
    ResetTimer
    Value = Value + 1

    ' [else]:
    '   Value = 99

End Sub

' ===========================================================
'  Helpers
' ===========================================================

Function StepTimer()
    StepTimer = Parent.Item("StepTimer").Value
End Function

Sub ResetTimer()
    Parent.Item("StepTimer").Value = 0
End Sub

Sub IncrementTimer()
    Parent.Item("StepTimer").Value = Parent.Item("StepTimer").Value + 1
End Sub

Sub Timeout(label)
    WriteLog "[BTC] " & label & " – Timeout exceeded. Automation blocked."
    xatm_BTC.Running  = False
    xatm_BTC.StepFailed = True
End Sub

Sub Complete()
    WriteLog "[BTC] [" & xatm_BTC.AutomationType & "] Completed successfully."
    xatm_BTC.Running = False
    xatm_BTC.CompletedSuccessfully = True
End Sub
