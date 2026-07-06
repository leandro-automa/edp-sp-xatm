' ==================================================================
'  BTC — Step 1: first commanded action of the sequence
'
'  TM Step 1 is ALWAYS a CLOSE (make-before-break). WHICH breaker to
'  close depends on the layout:
'    - Common layouts            -> the single bus-tie (700). No options.
'    - 4TR4LV_6BB6TIERING (GUL)  -> depends on the trigger transformer
'      AND which transformers are locked out (impedimento): Path A/B/C.
'
'  The full ordered plan is resolved ONCE here and committed, so every
'  later step reads the SAME path and never re-derives it from a lockout
'  state that could change mid-sequence.
' ==================================================================

' --- GUL ring breaker IDs (agreed scheme; swappable) --------------
'  600-series: ring-closure tie      700-series: bus-section ties
Const DJ10 = 600   ' TIE[1A,4A]  ring closure (normally open)
Const DJ20 = 700   ' TIE[1B,1A]
Const DJ30 = 710   ' TIE[2A,1B]
Const DJ40 = 720   ' TIE[2B,2A]
Const DJ50 = 730   ' TIE[3A,2B]
Const DJ60 = 740   ' TIE[4A,3A]

Const TIE_MAIN = 700   ' common-layout single bus-tie


Sub Main_Step01()

    Dim triggerId
    triggerId = xatm_BTC.Item("Commands").Item("TriggerTransformerId").Value

    Dim plan
    Select Case Parent.Item("AutomationType").Value
        Case "TM" : plan = GetTMPlan(triggerId)
        ' Case "NM" : plan = GetNMPlan(triggerId)   ' same shape, inverse order
        ' Case "TA" : plan = GetTAPlan(triggerId)   ' same shape, OPEN CB first
    End Select

    If Not IsArray(plan) Then
        WriteLog "Step 1: No valid transfer path for the current lockout state - aborting."
        Abort
        Exit Sub
    End If

    ' Commit the whole plan; later steps index into it, never re-derive.
    StorePlan plan

    ExecuteStep plan(0)
    WriteLog "Step 1: " & plan(0) & " commanded - proceeding to confirmation."

    ResetTimer
    Value = 2

End Sub


' Returns the full ordered TM step array, or Empty if no valid path.
Function GetTMPlan(triggerId)

    Select Case LayoutType()

        Case "4TR4LV_6BB6TIERING"
            GetTMPlan = GetTMPlan_Ring(triggerId)

        Case Else
            ' Common layout: one tie, one path, no contingency handling.
            GetTMPlan = Array("CLOSE:" & TIE_MAIN, _
                              "OPEN:"  & SecondaryBreakerId(triggerId))

    End Select

End Function


' GUL 4-transformer ring. Path chosen by trigger + lockouts,
' mirroring docs/layouts/4T_ring_guarulhos.md.
Function GetTMPlan_Ring(triggerId)

    Dim imp, cb
    imp = ReadLockouts()                    ' imp(1..4) = True if that TR is locked out
    cb  = SecondaryBreakerId(triggerId)     ' this transformer's secondary breaker

    Select Case triggerId

        Case 1
            If imp(2) Then                                  ' Path B — TR2 already out
                GetTMPlan_Ring = Array("CLOSE:"&DJ20, "OPEN:"&cb,   "CLOSE:"&DJ40, "OPEN:"&DJ30)
            ElseIf imp(3) Or imp(4) Then                    ' Path C — TR3/TR4 already out
                GetTMPlan_Ring = Array("CLOSE:"&DJ10, "OPEN:"&DJ20, "CLOSE:"&DJ30, "OPEN:"&cb)
            Else                                            ' Path A — no contingency
                GetTMPlan_Ring = Array("CLOSE:"&DJ10, "OPEN:"&DJ20, "CLOSE:"&DJ50, "OPEN:"&DJ40, "CLOSE:"&DJ30, "OPEN:"&cb)
            End If

        Case 2
            If imp(1) Then                                  ' Path B — TR1 already out
                GetTMPlan_Ring = Array("CLOSE:"&DJ20, "OPEN:"&DJ30, "CLOSE:"&DJ40, "OPEN:"&cb)
            ElseIf imp(3) Or imp(4) Then                    ' Path C — TR3/TR4 already out
                GetTMPlan_Ring = Array("CLOSE:"&DJ10, "OPEN:"&DJ20, "CLOSE:"&DJ30, "OPEN:"&cb)
            Else                                            ' Path A — no contingency
                GetTMPlan_Ring = Array("CLOSE:"&DJ10, "OPEN:"&DJ20, "CLOSE:"&DJ50, "OPEN:"&DJ40, "CLOSE:"&DJ30, "OPEN:"&cb)
            End If

        Case 3, 4                                           ' single path, no contingency
            GetTMPlan_Ring = Array("CLOSE:"&DJ60, "OPEN:"&cb)

    End Select

End Function


' ==================================================================
'  Helpers
' ==================================================================

Function LayoutType()
    LayoutType = xatm_BTC.Item("Data").Item("LayoutType").Value
End Function

' Low-side breaker of a transformer bay: T1=120, T2=220, T3=320, T4=420.
Function SecondaryBreakerId(triggerId)
    SecondaryBreakerId = triggerId * 100 + 20
End Function

' imp(1..4) — True if the transformer is locked out (impedimento).
Function ReadLockouts()
    Dim imp(4), i, tr, exists
    For i = 1 To 4
        Set tr = GetDeviceById(i * 100, exists)
        If exists Then
            imp(i) = tr.Item("Data").Item("LockedOut").Value
        Else
            imp(i) = False
        End If
    Next
    ReadLockouts = imp
End Function

' Persist the resolved plan so Steps 2..N read the committed path.
Sub StorePlan(plan)
    Parent.Item("StepPlan").Value = Join(plan, ",")
End Sub
