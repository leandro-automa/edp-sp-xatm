' ============================================================================
'  BTC — TM breaker plan (ETD Guarulhos 4TR4LV_6BB6TIERING)
'
'  Single source of truth for "which breaker, which command, each step,
'  each transformer". Verified line-by-line against the specification
'  (Descritivo Automatismos ETD Guarulhos), sections 1.4.1.x / 1.4.2.x.
'
'  Replaces the per-step Select Case triggerId logic scattered across
'  S1TM..S6TM: build the plan ONCE in Step00, then a generic executor
'  runs each entry.
'
'  Entry format "id:cmd" where cmd 2 = close (LIGA), 1 = open (DESLIGA).
'  Note: cmd value == confirmed Position value (closed=2, open=1).
' ============================================================================

' --- Ring breaker IDs (verified against the existing TR1 code) --------------
Const DJ10 = 900   ' TIE[1A,4A]  ring closure (normally open)
Const DJ20 = 700   ' TIE[1B,1A]
Const DJ30 = 710   ' TIE[2A,1B]
Const DJ40 = 720   ' TIE[2B,2A]
Const DJ50 = 730   ' TIE[3A,2B]
Const DJ60 = 740   ' TIE[4A,3A]
Const DJ03 = 120   ' CB TR1 (secondary)
Const DJ04 = 220   ' CB TR2
Const DJ05 = 320   ' CB TR3
Const DJ06 = 420   ' CB TR4


' ---------------------------------------------------------------------------
'  TM plan table
'
'  TR1 (100)                                          spec
'    none        : C DJ10 O DJ20 C DJ50 O DJ40 C DJ30 O DJ03   1.4.1.1
'    TR2 imp     : C DJ20 O DJ03 C DJ40 O DJ30               1.4.2.1
'    TR3/TR4 imp : C DJ10 O DJ20 C DJ30 O DJ03               1.4.2.3/1.4.2.5
'
'  TR2 (200)
'    none        : C DJ10 O DJ20 C DJ50 O DJ40 C DJ30 O DJ04   1.4.1.3
'    TR1 imp     : C DJ20 O DJ30 C DJ40 O DJ04               1.4.2.7
'    TR3/TR4 imp : C DJ50 O DJ40 C DJ30 O DJ04               1.4.2.9/1.4.2.11
'
'  TR3 (300)   -- contingencies are all DISTINCT (no 3||4 grouping)
'    none        : C DJ60 O DJ05                             1.4.1.5
'    TR1 imp     : C DJ40 O DJ50 C DJ60 O DJ05               1.4.2.13
'    TR2 imp     : C DJ20 O DJ10 C DJ60 O DJ05               1.4.2.15
'    TR4 imp     : C DJ10 O DJ60 C DJ50 O DJ05               1.4.2.17
'
'  TR4 (400)   -- contingencies are all DISTINCT
'    none        : C DJ60 O DJ06                             1.4.1.7
'    TR1 imp     : C DJ40 O DJ50 C DJ60 O DJ06               1.4.2.19
'    TR2 imp     : C DJ20 O DJ10 C DJ60 O DJ06               1.4.2.21
'    TR3 imp     : C DJ50 O DJ60 C DJ10 O DJ05  (*)          1.4.2.23
'
'  (*) TR4/TR3 imp ends on DJ05 (TR3 breaker), NOT DJ06, per spec 1.4.2.23.
'      Counterintuitive — flagged for confirmation with the utility.
' ---------------------------------------------------------------------------
Function GetTMPlan_Ring(triggerId, lockouts)

	Select Case triggerId

		Case 100
			If lockouts(2) Then
				GetTMPlan_Ring = Array(DJ20&":2", DJ03&":1", DJ40&":2", DJ30&":1")
			ElseIf lockouts(3) Or lockouts(4) Then
				GetTMPlan_Ring = Array(DJ10&":2", DJ20&":1", DJ30&":2", DJ03&":1")
			Else
				GetTMPlan_Ring = Array(DJ10&":2", DJ20&":1", DJ50&":2", DJ40&":1", DJ30&":2", DJ03&":1")
			End If

		Case 200
			If lockouts(1) Then
				GetTMPlan_Ring = Array(DJ20&":2", DJ30&":1", DJ40&":2", DJ04&":1")
			ElseIf lockouts(3) Or lockouts(4) Then
				GetTMPlan_Ring = Array(DJ50&":2", DJ40&":1", DJ30&":2", DJ04&":1")
			Else
				GetTMPlan_Ring = Array(DJ10&":2", DJ20&":1", DJ50&":2", DJ40&":1", DJ30&":2", DJ04&":1")
			End If

		Case 300
			If lockouts(1) Then
				GetTMPlan_Ring = Array(DJ40&":2", DJ50&":1", DJ60&":2", DJ05&":1")
			ElseIf lockouts(2) Then
				GetTMPlan_Ring = Array(DJ20&":2", DJ10&":1", DJ60&":2", DJ05&":1")
			ElseIf lockouts(4) Then
				GetTMPlan_Ring = Array(DJ10&":2", DJ60&":1", DJ50&":2", DJ05&":1")
			Else
				GetTMPlan_Ring = Array(DJ60&":2", DJ05&":1")
			End If

		Case 400
			If lockouts(1) Then
				GetTMPlan_Ring = Array(DJ40&":2", DJ50&":1", DJ60&":2", DJ06&":1")
			ElseIf lockouts(2) Then
				GetTMPlan_Ring = Array(DJ20&":2", DJ10&":1", DJ60&":2", DJ06&":1")
			ElseIf lockouts(3) Then
				GetTMPlan_Ring = Array(DJ50&":2", DJ60&":1", DJ10&":2", DJ05&":1")   ' (*) ends on DJ05
			Else
				GetTMPlan_Ring = Array(DJ60&":2", DJ06&":1")
			End If

	End Select

End Function


' ---------------------------------------------------------------------------
'  Layout dispatch + Step00 build. Common layouts: single tie, single path.
' ---------------------------------------------------------------------------
Function GetTMPlan(triggerId, lockouts)

	Select Case GetLayoutType()
		Case "4TR4LV_6BB6TIERING"
			GetTMPlan = GetTMPlan_Ring(triggerId, lockouts)
		Case Else
			' Common: close the tie, open the trigger's secondary breaker.
			GetTMPlan = Array("700:2", (triggerId + 20) & ":1")
	End Select

End Function


' Build the plan once, store it, arm the first step.
Sub Main_Step00_Build()

	Dim plan
	Select Case Parent.Item("AutomationType").Value
		Case "TM" : plan = GetTMPlan(Parent.Item("TriggerTransformerId").Value, ReadLockouts())
		' Case "NM" : plan = GetNMPlan(...)
		' Case "TA" : plan = GetTAPlan(...)
	End Select

	Parent.Item("StepPlan").WriteEx Join(plan, ",")
	Value = 1

End Sub


' ---------------------------------------------------------------------------
'  Generic step executor — replaces S1TM..S6TM entirely.
'  Confirms the current entry's breaker position; if not yet there, drives
'  the command state machine; on confirmation advances (or finishes at 99).
' ---------------------------------------------------------------------------
Sub ExecuteStep()

	Dim plan
	plan = Split(Parent.Item("StepPlan").Value, ",")

	Dim idx
	idx = Value - 1                         ' Value 1..N -> plan(0..N-1)

	Dim parts, breakerId, cmd
	parts     = Split(plan(idx), ":")
	breakerId = CLng(parts(0))
	cmd       = CInt(parts(1))              ' 2 = close, 1 = open

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then
		WriteLog "Step " & Value & ": Circuit breaker ID=" & breakerId & " not found - Global lockout."
		Main_GlobalLockout()
		Exit Sub
	End If

	' Confirmed?
	If breaker.Item("Data").Item("Position").Value = cmd Then
		WriteLog "Step " & Value & ": Circuit breaker operated - Proceeding to the next step."
		ResetTimer()
		If idx >= UBound(plan) Then
			Value = 99
		Else
			Value = Value + 1
		End If
		Exit Sub
	End If

	' Not yet — drive the command.
	Select Case breaker.Item("Data").Item("CommandInProgress").Value
		Case 0, 3
			breaker.Item("Data").Item("CommandOpenClose").WriteEx cmd
		Case 1
			WriteLog "Step " & Value & ": Command execution failed - Global lockout."
			Main_GlobalLockout()
		Case 2
			' Command in progress - wait next cycle.
	End Select

End Sub
