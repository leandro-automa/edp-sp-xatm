-----------------------
Documentação de Scripts
-----------------------
xatm_Breaker
Wed Jul  8 18:04:07 2026
-----------------------

<xatm_Breaker.Data.CommandInProgress:CommandInProgress_OnChangedValue()>
Sub CommandInProgress_OnChangedValue()

	Select Case Value

		Case 2
			' Command in progress - start the countdown from the configured timeout.
			Parent.Item("Timers").Item("CommandTimer").WriteEx xatm_Breaker.CommandTimeout

		Case 1
			' Command execution failed.
			WriteLog "Command execution failed."

		Case 3
			' Command completed successfully - back to idle.

	End Select
	
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

<xatm_Breaker.Data.CommandOpenClose:CommandOpenClose_OnChangedValue()>
Sub CommandOpenClose_OnChangedValue()
	
	Dim command
	command = Value      ' 1 = open, 2 = close

	' --- Position interlock: don't command a breaker already at target ---
	Dim position
	position = Parent.Item("Position").Value

	Select Case command

		Case 1   ' Open: only when currently closed
			If position <> 2 Then
				WriteLog "Open command not executed - position interlock (breaker not closed)."
				Exit Sub
			End If

		Case 2   ' Close: only when currently open
			If position <> 1 Then
				WriteLog "Close command not executed - position interlock (breaker not open)."
				Exit Sub
			End If

		Case Else
			WriteLog "Command not executed - invalid command value: " & command
			Exit Sub

	End Select

	' --- Issue the command. Handle with position only (no provisional position) ---
	If CBool(Parent.Item("SimulationModeEnabled").Value) Then

		WriteLog IIf(command = 1, "Open", "Close") & " command sent (Simulation Mode)."

		' The breaker "responds" by moving to the commanded position - unless we are
		' simulating a failure, in which case Position is left unchanged so the
		' command is never confirmed and times out.
		If Not CBool(Parent.Item("SimulateCommandFailure").Value) Then
			Parent.Item("Position").WriteEx command      ' 1 -> open, 2 -> closed
		End If

	Else

		' TODO: write the real command output IOTag(s) here
		WriteLog IIf(command = 1, "Open", "Close") & " command sent."

	End If

	' --- Flag command in progress (drives the command timer / BTC) ---
	If Parent.Item("CommandInProgress").Value <> 2 Then
		Parent.Item("CommandInProgress").WriteEx 2
	End If

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

<xatm_Breaker.Data.CommunicationFailure:CommunicationFailure_E_OnChangedOpenPositionQuality()>
Sub CommunicationFailure_E_OnChangedOpenPositionQuality()
	
	SetCommunicationFailure
		
End Sub

<xatm_Breaker.Data.CommunicationFailure:CommunicationFailure_Functions()>
Sub CommunicationFailure_Functions()
End Sub


Function IsCommunicationHealthy()
	
	IsCommunicationHealthy = False

	If xatm_Breaker.UseDoublePoints Then

		' ======================
		' Double Point
		' ======================
		
		Dim tags
		tags = Array(xatm_Breaker.PositionClosed, _
					xatm_Breaker.PositionOpen, _
					xatm_Breaker.PositionClosedAlt, _
					xatm_Breaker.PositionOpenAlt)
		
		Dim i
		
		For i = 0 To UBound(tags)
        	
			Dim tag
			Set tag = Nothing
			
			On Error Resume Next
			Set tag = tags(i)
			On Error GoTo 0
        	
			If Not tag Is Nothing Then
				If tag.Quality >= 192 Then
				
					IsCommunicationHealthy = True
					Exit For
					
				End If
        	End If

    	Next			
	
	Else
	
		' ======================
		' Single Point
		' ======================
		
		Dim mainHealthy
		mainHealthy = (xatm_Breaker.PositionOpen.Quality >= 192) And (xatm_Breaker.PositionClosed.Quality >= 192)
		
		Dim altHealthy
		altHealthy  = (xatm_Breaker.PositionOpenAlt.Quality >= 192) And (xatm_Breaker.PositionClosedAlt.Quality >= 192)
		
		IsCommunicationHealthy = mainHealthy Or altHealthy

	End If
	
End Function


Sub SetCommunicationFailure()
	
	Dim communicationFailure
	communicationFailure = Not IsCommunicationHealthy()
	
	WriteEx communicationFailure 
	
End Sub

<xatm_Breaker.Data.CommunicationFailure:CommunicationFailure_OnChangedSimMode()>
Sub CommunicationFailure_OnChangedSimMode()
	
	If CBool(Parent.Item("SimulationModeEnabled").Value) Then
	
		WriteEx False
	
	Else 
	
		SetCommunicationFailure
	
	End If
		
End Sub

<xatm_Breaker.Data.CommunicationFailure:CommunicationFailure_OnStartRunning()>
Sub CommunicationFailure_OnStartRunning()
	
	SetCommunicationFailure
		
End Sub

<xatm_Breaker.Data.MemorizedPosition:MemorizedPosition_OnChangedPosition()>
Sub MemorizedPosition_OnChangedPosition()

	If Value = 0 Then
		
		Value = Parent.Item("Position").Value
		
	Else
		
		' Init timer
		DocString = 120
		
	End If
	
End Sub

<xatm_Breaker.Data.MemorizedPosition:MemorizedPosition_Timer()>
Sub MemorizedPosition_Timer()
	
	Dim currentPosition
	currentPosition = Parent.Item("Position").Value
	
	If DocString = "0" Then
		
		If currentPosition = 1 Or currentPosition = 2 Then
		
			Value = currentPosition
			DocString = "-1"
			
		End If
		
		Exit Sub
		
	End If
	
	If Value = currentPosition Then
		
		DocString = "-1"
		Exit Sub
		
	End If	
	
	DocString = CStr(CInt(DocString) - 1)
	
End Sub

<xatm_Breaker.Data.Position:Position_E_OnChangedClosedPosition()>
Sub Position_E_OnChangedClosedPosition()

	CalculatePosition
	
End Sub

<xatm_Breaker.Data.Position:Position_E_OnChangedClosedQuality()>
Sub Position_E_OnChangedClosedQuality()
	
	CalculatePosition
	
End Sub

<xatm_Breaker.Data.Position:Position_E_OnChangedOpenPosition()>
Sub Position_E_OnChangedOpenPosition()
	
	CalculatePosition
	
End Sub

<xatm_Breaker.Data.Position:Position_E_OnChangedOpenPositionQuality()>
Sub Position_E_OnChangedOpenPositionQuality()
	
	CalculatePosition
	
End Sub

<xatm_Breaker.Data.Position:Position_Functions()>
Sub Position_Functions()

End Sub

Sub CalculatePosition()

	Dim resolved
	resolved = ResolvePosition()
	
	If Value <> resolved Then

		WriteEx resolved
		
		On Error Resume Next
		Application.Trace "[" & xatm_Breaker.Name & "] - Position changed to: " & resolved
		On Error Goto 0
		
	End If

End Sub

Function ResolvePosition()

	ResolvePosition = 0

	If xatm_Breaker.UseDoublePoints Then

		' ======================
		' Double Point
		' ======================
	
		tags = Array(xatm_Breaker.PositionClosed, _
					xatm_Breaker.PositionOpen, _
					xatm_Breaker.PositionClosedAlt, _
					xatm_Breaker.PositionOpenAlt)
    
		For i = LBound(tags) To UBound(tags)

			If IsObjectExists(tags(i)) Then
			
				If tags(i).Quality >= 192 Then
					ResolvePosition = CalculateDoublePointState(tags(i).Value)
					Exit Function
				End If

			End If

		Next
		
		'On Error Resume Next
		'Application.Trace Me.PathName & ": No valid positions found with sufficient quality."
		'On Error Goto 0

	Else

		' ======================
		' Single Point
		' ======================

		If IsObjectExists(xatm_Breaker.PositionClosed) And IsObjectExists(xatm_Breaker.PositionOpen) Then
			
			If xatm_Breaker.PositionClosed.Quality >= 192 And xatm_Breaker.PositionOpen.Quality >= 192 Then
				
				ResolvePosition = CalculateSinglePointState(xatm_Breaker.PositionOpen.Value, _
						xatm_Breaker.PositionClosed.Value)
				Exit Function

			End If

		End If
		
		If IsObjectExists(xatm_Breaker.PositionClosedAlt) And IsObjectExists(xatm_Breaker.PositionOpenAlt) Then
			
			If xatm_Breaker.PositionClosedAlt.Quality >= 192 And xatm_Breaker.PositionOpenAlt.Quality >= 192 Then
				
				ResolvePosition = CalculateSinglePointState(xatm_Breaker.PositionOpenAlt.Value, _
						xatm_Breaker.PositionClosedAlt.Value)
				Exit Function

			End If

		Else

			On Error Resume Next
			Application.Trace Me.PathName & ": No valid pair of positions found with sufficient quality."
			On Error Goto 0

		End If
	End If

End Function

Function IsObjectExists(obj)

	Dim value
	Dim quality
	On Error Resume Next
	value = obj.Value
	quality = obj.Quality
	IsObjectExists = (Err = 0)
	On Error Goto 0

End Function

Function CalculateDoublePointState(currentRawValue)
	
	Select Case currentRawValue
		
		Case xatm_Breaker.RawValueOpen
			CalculateDoublePointState = 1
		
		Case xatm_Breaker.RawValueClosed
			CalculateDoublePointState = 2
		
		Case Else
			CalculateDoublePointState = 0

	End Select

End Function

Function CalculateSinglePointState(rawValueOpen, rawValueClosed)
	
	If rawValueOpen = xatm_Breaker.RawValueOpen And rawValueClosed <> xatm_Breaker.RawValueClosed Then
		
		CalculateSinglePointState = 1
	
	ElseIf rawValueOpen <> xatm_Breaker.RawValueOpen And rawValueClosed = xatm_Breaker.RawValueClosed Then
		
		CalculateSinglePointState = 2
	
	Else
	
		CalculateSinglePointState = 0
	
	End If

End Function

Sub Foo()	
End Sub

<xatm_Breaker.Data.Position:Position_OnChangedSimMode()>
Sub Position_OnChangedSimMode()

	If CBool(Parent.Item("SimulationModeEnabled").Value) Then
		
		WriteEx (xatm_Breaker.NormalState + 1)
		
	Else
		
		WriteEx 0
		
		CalculatePosition
	
	End If
	
End Sub

<xatm_Breaker.Data.Position:Position_OnStartRunning()>
Sub Position_OnStartRunning()

	CalculatePosition
	
End Sub

<xatm_Breaker.Data.Reset:Reset_OnChangedTimeStamp()>
Sub Reset_OnChangedTimeStamp()

	If Not CBool(Value) Then
		Exit Sub
	End If
	
	Parent.Item("MemorizedPosition").Value = Parent.Item("Position").Value
	Parent.Item("MemorizedPosition").DocString = "-1"
	
	Parent.Item("CommandInProgress").WriteEx Empty, 0
	
	' ================
	' RESET TIMERS
	' ================
	Dim child
	For Each child In Parent.Item("Timers")
		
		If TypeName(child) = "InternalTag" Then
			Child.WriteEx -1, 0
		End If
		
	Next
End Sub

<xatm_Breaker.Data.Timers.CommandTimer:CommandTimer_Counter()>
Sub CommandTimer_Counter()

	' Counts DOWN from CommandTimeout (seeded when CommandInProgress = 2).
	' Runs every second while Value >= 0.

	Dim command
	command = Parent.Parent.Item("CommandOpenClose").Value     ' 1 = open, 2 = close

	Dim position
	position = Parent.Parent.Item("Position").Value            ' 1 = open, 2 = closed

	' --- Completed: position reached the commanded target ---
	If position = command Then
		Parent.Parent.Item("CommandInProgress").WriteEx 3
		Value = -1
		Exit Sub
	End If

	' --- Timeout: target not reached in time -> failed ---
	If Value <= 0 Then
		Parent.Parent.Item("CommandInProgress").WriteEx 1
		Value = -1
		Exit Sub
	End If

	' --- Retry once, at half the timeout ---
	If Value = Int(xatm_Breaker.CommandTimeout / 2) Then
		
		WriteLog "Command not confirmed, resending (retry)."
		Parent.Parent.Item("CommandOpenClose").WriteEx command
		
	End If

	Value = Value - 1

End Sub

Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Parent.Parent.Parent.Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Parent.Parent.Parent.Name & "] - " & message
	End If

End Sub

