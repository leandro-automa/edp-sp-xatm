-----------------------
Documentação de Scripts
-----------------------
Start
Mon Jul  6 15:44:45 2026
-----------------------

<xatm_BTC.Commands.Start:Start_OnChangedValue()>
Sub Start_OnChangedValue()

	If Trim(Value) = "" Then Exit Sub      ' self-cleared write, ignore

	Dim ts
	ts = TimeStamp	' preserve for the silent clear

	Dim parts
	parts = Split(Me.Value, ":")

	If UBound(parts) <> 1 Then
		Reject "Invalid Start format '" & Me.Value & "' (expected MODE:trigger).", ts
		Exit Sub
	End If

	Dim mode
	mode = UCase(Trim(parts(0)))

	Select Case mode
		
		Case "TM", "NM", "TA"
		
		Case Else
			
			Reject "Unknown automation type '" & mode & "'.", ts
			Exit Sub
			
	End Select

	If Not IsNumeric(parts(1)) Then
		
		Reject "Invalid trigger transformer '" & parts(1) & "'.", ts
		Exit Sub
		
	End If

	Dim triggerId
	triggerId = CInt(parts(1))

	' ================
	' GATE CHECKS
	' ================

	If Not xatm_BTC.Enabled Then
		
		Reject "Automation disabled.", ts
		Exit Sub

	End If

	If xatm_BTC.Running Then
		
		Reject "Already running.", ts
		Exit Sub

	End If

	' If xatm_BTC.OperatorBlock Then
		
	' 	Reject "Blocked by operator.", ts
	' 	Exit Sub

	' End If

	' If xatm_BTC.GeneralBlock Then
		
	' 	Reject "General interlock active.", ts
	' 	Exit Sub

	' End If

	If AnyOtherAutomationRunning() Then
		
		Reject "Another automation is in progress.", ts
		Exit Sub

	End If


	' TODO: operator, general, preconditions and other automation interlocks


	' ================
	' START
	' ================
	
	xatm_BTC.Item("FSM").Item("AutomationType").WriteEx mode
	xatm_BTC.Item("FSM").Item("TriggerTransformerId").WriteEx triggerId
	xatm_BTC.Item("FSM").Item("StepTimer").WriteEx 0
	xatm_BTC.Item("FSM").Item("Main").WriteEx 0
	xatm_BTC.Running = True

	WriteLog "[BTC] Start - " & mode & " TR" & triggerId

	WriteEx "", ts ' clear without re-firing

End Sub


' True if any OTHER automation object is currently running (mutual exclusion).
' Relies on a common Running property instead of enumerating each type.
Function AnyOtherAutomationRunning()

	AnyOtherAutomationRunning = False
	
	Dim obj
	For Each obj In Application.GetObject("XATM_Data.Automation")

		If Not (obj Is xatm_BTC) Then
		
			Dim running
			running = False
			On Error Resume Next
			running = obj.Running
			On Error GoTo 0

			If running Then
				AnyOtherAutomationRunning = True
				Exit Function
			End If

		End If

	Next

End Function

' Logs the rejection and clears the trigger without re-firing.
' Uses the explicit tag path (not Me) so it is safe to call from a helper.
Sub Reject(reason, ts)

	WriteLog "[BTC] Not started - " & reason
	WriteEx "", ts
	
End Sub

Sub WriteLog(message)
	
	Dim consoleLogEngine
	Set consoleLogEngine = Nothing
	
	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Parent.Parent.Name & "] - " & message
	On Error Goto 0
	
	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Parent.Parent.Name & "] - " & message
	End If
	
End Sub

