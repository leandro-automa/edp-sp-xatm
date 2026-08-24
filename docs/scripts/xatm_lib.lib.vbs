-----------------------
Documentação de Scripts
-----------------------
XATM_LIB (C:\ProjDev\edp_sp\xatm_lib.lib)
Mon Aug 24 14:05:47 2026
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
	Dim commandSent
	commandSent = False

	If CBool(Parent.Item("SimulationModeEnabled").Value) Then

		WriteLog IIf(command = 1, "Open", "Close") & " command sent (Simulation Mode)."

		' The breaker "responds" by moving to the commanded position - unless we are
		' simulating a failure, in which case Position is left unchanged so the
		' command is never confirmed and times out.
		If Not CBool(Parent.Item("SimulateCommandFailure").Value) Then
			Parent.Item("Position").WriteEx command      ' 1 -> open, 2 -> closed
		End If

		commandSent = True

	Else

		' Nothing real leaves a demo build.
		'
		' Asked here and not before the simulation branch on purpose: a demo
		' build drives simulated equipment exactly as the runtime one does,
		' so a whole sequence can still be watched end to end. It is only the
		' switchyard it is kept away from.
		If IsDemoBuild() Then

			WriteLog IIf(command = 1, "Open", "Close") & _
			         " command not executed - DEMO build, only simulated equipment is operated."
			Exit Sub

		End If

		' Route the command to a relay that is actually communicating: try the
		' configured priority source first, then fall back to the redundant one.
		Dim useAlt
		'useAlt = CBool(xatm_Breaker.PreferAlternateSource)
		useAlt = False

		If IsSourceHealthy(useAlt) Then
			commandSent = SendCommand(useAlt, command)
		End If

		If Not commandSent Then
			If IsSourceHealthy(Not useAlt) Then
				commandSent = SendCommand(Not useAlt, command)
			End If
		End If

		If Not commandSent Then
			WriteLog IIf(command = 1, "Open", "Close") & " command not executed - no relay available."
			Exit Sub
		End If

	End If

	' --- Flag command in progress (drives the command timer / TMTNM) ---
	If Parent.Item("CommandInProgress").Value <> 2 Then
		Parent.Item("CommandInProgress").WriteEx 2
	End If

End Sub

' Whether the library driving this breaker is the demo build.
'
' Read off the xatm_Build the breaker carries, so the answer comes from
' the same library file as this code - not from a constant copied into
' this scope, which would be a second thing to flip, and not from a
' singleton in another project, which a site might not have deployed.
'
' Fails closed. A breaker with no Build inside it is one driven by a
' library too old to have the class at all, and that is not a library to
' let near a switchyard.
Function IsDemoBuild()

	IsDemoBuild = True

	On Error Resume Next
	IsDemoBuild = CBool(xatm_Breaker.Item("Build").Demo)
	On Error Goto 0

End Function


' Writes the command output of one relay - main (useAlt = False) or redundant
' (useAlt = True). Returns True only when the output tag was actually written.
Function SendCommand(useAlt, command)

	SendCommand = False

	Dim outputTag
	Set outputTag = Nothing
	
	Dim selectTag
	Set selectTag = Nothing
	
	Dim sourceName
	If useAlt Then
		sourceName = "relay 2"
	Else
		sourceName = "relay 1"
	End If
	
	Dim rawValue
		
	On Error Resume Next
	
	If command = 1 Then

		' ======================
		' Open
		' ======================
		rawValue = xatm_Breaker.RawValueCommandOpen

		If useAlt Then
			Set outputTag = xatm_Breaker.CommandOpenAlt
			Set selectTag = xatm_Breaker.CommandSBOOpenAlt
		Else
			Set outputTag = xatm_Breaker.CommandOpen
			Set selectTag = xatm_Breaker.CommandSBOOpen
		End If

	Else

		' ======================
		' Close
		' ======================
		rawValue = xatm_Breaker.RawValueCommandClose

		If useAlt Then
			Set outputTag = xatm_Breaker.CommandCloseAlt
			Set selectTag = xatm_Breaker.CommandSBOCloseAlt
		Else
			Set outputTag = xatm_Breaker.CommandClose
			Set selectTag = xatm_Breaker.CommandSBOClose
		End If

	End If

	Err.Clear
	On Error Goto 0

	If Not IsTagLinked(outputTag) Then
		WriteLog IIf(command = 1, "Open", "Close") & " command not executed - output not linked (" & sourceName & ")."
		Exit Function
	End If

	' Select-Before-Operate: only performed when a select tag is configured.
	If IsTagLinked(selectTag) Then
		selectTag.WriteEx rawValue
	End If

	outputTag.WriteEx rawValue

	WriteLog IIf(command = 1, "Open", "Close") & " command sent (" & sourceName & ") - value = " & rawValue & "."

	SendCommand = True

End Function

' Communication check for a single relay, same criteria used by CommunicationFailure:
' double point needs one readable tag, single point needs the whole contact pair.
Function IsSourceHealthy(useAlt)

	Dim tagOpen
	Dim tagClosed

	Set tagOpen = Nothing
	Set tagClosed = Nothing

	On Error Resume Next

	If useAlt Then
		Set tagOpen   = xatm_Breaker.PositionOpenAlt
		Set tagClosed = xatm_Breaker.PositionClosedAlt
	Else
		Set tagOpen   = xatm_Breaker.PositionOpen
		Set tagClosed = xatm_Breaker.PositionClosed
	End If

	Err.Clear
	On Error Goto 0

	If xatm_Breaker.UseDoublePoints Then
		IsSourceHealthy = IsTagHealthy(tagOpen) Or IsTagHealthy(tagClosed)
	Else
		IsSourceHealthy = IsTagHealthy(tagOpen) And IsTagHealthy(tagClosed)
	End If

End Function

Function IsTagHealthy(tag)

	IsTagHealthy = False

	If tag Is Nothing Then
		Exit Function
	End If

	Dim quality
	On Error Resume Next
	quality = tag.Quality
	IsTagHealthy = (Err.Number = 0) And (quality >= 192)
	Err.Clear
	On Error Goto 0

End Function

' An unassociated XObject property yields Nothing (or an object that cannot be read),
' which is how an unused redundant relay / missing SBO tag is detected.
Function IsTagLinked(tag)

	IsTagLinked = False

	If tag Is Nothing Then
		Exit Function
	End If

	Dim probe
	On Error Resume Next
	probe = tag.Value
	IsTagLinked = (Err.Number = 0)
	Err.Clear
	On Error Goto 0

End Function

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
	
	' ================
	' RESET COMMAND FAILURES
	' ================
	' The two latches CommandTimer sets on a timeout. Here is the only
	' place they are cleared - a failure that cleared itself on the next
	' command would be gone before anybody saw it.
	xatm_Breaker.CommandOpenFailed  = False
	xatm_Breaker.CommandCloseFailed = False
	
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

		' Which of the two it was, latched for whoever looks later - the
		' automation goes to global lockout on CommandInProgress alone,
		' and by the time anybody reads a screen the timer has stopped
		' counting and CommandOpenClose says nothing about what failed.
		'
		' Set after the stop above and not before it, so a breaker whose
		' class has not been given these properties yet still stops its
		' timer and still tells the automation the command failed.
		If command = 1 Then
			xatm_Breaker.CommandOpenFailed = True
		ElseIf command = 2 Then
			xatm_Breaker.CommandCloseFailed = True
		End If

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

<xatm_Build:xatm_Build_OnStartRunning()>
Sub xatm_Build_OnStartRunning()

	Const DEMO_BUILD = True

	' What this build of the library is, published where anything can read
	' it without having to know where it came from.
	'
	' One of these sits inside every breaker and every disconnector, so the
	' command gate can ask the device it is already holding rather than
	' reach into another project for the answer.
	'
	' Nothing here is configured. The class comes from whichever
	' xatm_lib.lib the domain has loaded, so swapping the file swaps what
	' every instance says at once - which is the whole mechanism.
	Demo = DEMO_BUILD

	' The same fact as a word, for a label to bind to. Derived rather than
	' declared, so it cannot come to disagree with the flag - and the gate
	' reads the flag and never this, because a typo or a stray capital in
	' a string comparison is not a thing to put in front of a switchyard.
	Edition = IIf(DEMO_BUILD, "DEMO", "RUNTIME")
		
End Sub

<xatm_ConsoleLogEngine.WriteLine:xatm_ConsoleLogEngine_OnWriteLineChanged()>
Sub xatm_ConsoleLogEngine_OnWriteLineChanged()

    If IsEmpty(Me.WriteLine) Then 
        Exit Sub
    End If

    Dim MAX_LINES
    MAX_LINES = MaxLines

    ' ====================================
    ' Format message with timestamp
    ' ====================================
    Dim formattedTimeStamp
    On Error Resume Next
    formattedTimeStamp = E3Format(Now, DateTimeFormat)
    If Err Then
    	formattedTimeStamp = E3Format(Now, "dd/MM/yyyy HH:mm:ss")
    End If
    On Error Goto 0
    
    Dim message
    message = formattedTimeStamp & vbTab & WriteLine

    ' ====================================
    ' Append to content array
    ' ====================================
    Dim contentTag
    Set contentTag = Item("Data").Item("Content")

    Dim contentArr
    If IsArray(contentTag.Value) Then
        contentArr = contentTag.Value
    Else
        contentArr = Array()
    End If

    Dim currentSize
    currentSize = UBound(contentArr) + 1
	
    ' Determine the index for the new message
    Dim index
    
    If currentSize < MAX_LINES Then
        
        ' Array still growing — just append
        index = currentSize
        ReDim Preserve contentArr(currentSize)
        
    Else
        
        index = MAX_LINES - 1

        ' Array at capacity — shift and append
        Dim j
        For j = 0 To MAX_LINES - 2
            contentArr(j) = contentArr(j + 1)
        Next
        
    End If

    ' Add new message at the end
    contentArr(index) = message

    ' Trace the message for debugging
    If TraceEnabled Then
		On Error Resume Next
		Application.Trace message, False
		On Error Goto 0
	End If
	
    ' ================================
	' Update the content tag with the new array
    ' ================================
	contentTag.Value = contentArr
    
    
    ' ================================
    ' Clear the WriteLine property to prepare for the next message
    ' ================================
    Me.WriteLine = Empty
	
End Sub

<xatm_Disconnector.Data.CommandInProgress:CommandInProgress_OnChangedValue()>
Sub CommandInProgress_OnChangedValue()

	Select Case Value

		Case 2
			' Command in progress - start the countdown from the configured timeout.
			Parent.Item("Timers").Item("CommandTimer").WriteEx xatm_Disconnector.CommandTimeout

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

<xatm_Disconnector.Data.CommandOpenClose:CommandOpenClose_OnChangedValue()>
Sub CommandOpenClose_OnChangedValue()
	
	Dim command
	command = Value      ' 1 = open, 2 = close

	' --- Position interlock: don't command a disconnector already at target ---
	Dim position
	position = Parent.Item("Position").Value

	Select Case command

		Case 1   ' Open: only when currently closed
			If position <> 2 Then
				WriteLog "Open command not executed - position interlock (disconnector not closed)."
				Exit Sub
			End If

		Case 2   ' Close: only when currently open
			If position <> 1 Then
				WriteLog "Close command not executed - position interlock (disconnector not open)."
				Exit Sub
			End If

		Case Else
			WriteLog "Command not executed - invalid command value: " & command
			Exit Sub

	End Select

	' --- Issue the command. Handle with position only (no provisional position) ---
	Dim commandSent
	commandSent = False

	If CBool(Parent.Item("SimulationModeEnabled").Value) Then

		WriteLog IIf(command = 1, "Open", "Close") & " command sent (Simulation Mode)."

		' The disconnector "responds" by moving to the commanded position - unless we are
		' simulating a failure, in which case Position is left unchanged so the
		' command is never confirmed and times out.
		If Not CBool(Parent.Item("SimulateCommandFailure").Value) Then
			Parent.Item("Position").WriteEx command      ' 1 -> open, 2 -> closed
		End If

		commandSent = True

	Else
		
		' Nothing real leaves a demo build.
		'
		' Asked here and not before the simulation branch on purpose: a demo
		' build drives simulated equipment exactly as the runtime one does,
		' so a whole sequence can still be watched end to end. It is only the
		' switchyard it is kept away from.
		If IsDemoBuild() Then

			WriteLog IIf(command = 1, "Open", "Close") & _
			         " command not executed - DEMO build, only simulated equipment is operated."
			Exit Sub

		End If

		' Route the command to a relay that is actually communicating: try the
		' configured priority source first, then fall back to the redundant one.
		Dim useAlt
		'useAlt = CBool(xatm_Disconnector.PreferAlternateSource)
		useAlt = False

		If IsSourceHealthy(useAlt) Then
			commandSent = SendCommand(useAlt, command)
		End If

		If Not commandSent Then
			If IsSourceHealthy(Not useAlt) Then
				commandSent = SendCommand(Not useAlt, command)
			End If
		End If

		If Not commandSent Then
			WriteLog IIf(command = 1, "Open", "Close") & " command not executed - no relay available."
			Exit Sub
		End If

	End If

	' --- Flag command in progress (drives the command timer / TMTNM) ---
	If Parent.Item("CommandInProgress").Value <> 2 Then
		Parent.Item("CommandInProgress").WriteEx 2
	End If

End Sub

' Whether the library driving this breaker is the demo build.
'
' Read off the xatm_Build the breaker carries, so the answer comes from
' the same library file as this code - not from a constant copied into
' this scope, which would be a second thing to flip, and not from a
' singleton in another project, which a site might not have deployed.
'
' Fails closed. A breaker with no Build inside it is one driven by a
' library too old to have the class at all, and that is not a library to
' let near a switchyard.
Function IsDemoBuild()

	IsDemoBuild = True

	On Error Resume Next
	IsDemoBuild = CBool(xatm_Disconnector.Item("Build").Demo)
	On Error Goto 0

End Function


' Writes the command output of one relay - main (useAlt = False) or redundant
' (useAlt = True). Returns True only when the output tag was actually written.
Function SendCommand(useAlt, command)

	SendCommand = False

	Dim outputTag
	Set outputTag = Nothing
	
	Dim selectTag
	Set selectTag = Nothing
	
	Dim sourceName
	If useAlt Then
		sourceName = "relay 2"
	Else
		sourceName = "relay 1"
	End If
	
	Dim rawValue
		
	On Error Resume Next
	
	If command = 1 Then

		' ======================
		' Open
		' ======================
		rawValue = xatm_Disconnector.RawValueCommandOpen

		If useAlt Then
			Set outputTag = xatm_Disconnector.CommandOpenAlt
			Set selectTag = xatm_Disconnector.CommandSBOOpenAlt
		Else
			Set outputTag = xatm_Disconnector.CommandOpen
			Set selectTag = xatm_Disconnector.CommandSBOOpen
		End If

	Else

		' ======================
		' Close
		' ======================
		rawValue = xatm_Disconnector.RawValueCommandClose

		If useAlt Then
			Set outputTag = xatm_Disconnector.CommandCloseAlt
			Set selectTag = xatm_Disconnector.CommandSBOCloseAlt
		Else
			Set outputTag = xatm_Disconnector.CommandClose
			Set selectTag = xatm_Disconnector.CommandSBOClose
		End If

	End If

	Err.Clear
	On Error Goto 0

	If Not IsTagLinked(outputTag) Then
		WriteLog IIf(command = 1, "Open", "Close") & " command not executed - output not linked (" & sourceName & ")."
		Exit Function
	End If

	' Select-Before-Operate: only performed when a select tag is configured.
	If IsTagLinked(selectTag) Then
		selectTag.WriteEx rawValue
	End If

	outputTag.WriteEx rawValue

	WriteLog IIf(command = 1, "Open", "Close") & " command sent (" & sourceName & ") - value = " & rawValue & "."

	SendCommand = True

End Function

' Communication check for a single relay, same criteria used by CommunicationFailure:
' double point needs one readable tag, single point needs the whole contact pair.
Function IsSourceHealthy(useAlt)

	Dim tagOpen
	Dim tagClosed

	Set tagOpen = Nothing
	Set tagClosed = Nothing

	On Error Resume Next

	If useAlt Then
		Set tagOpen   = xatm_Disconnector.PositionOpenAlt
		Set tagClosed = xatm_Disconnector.PositionClosedAlt
	Else
		Set tagOpen   = xatm_Disconnector.PositionOpen
		Set tagClosed = xatm_Disconnector.PositionClosed
	End If

	Err.Clear
	On Error Goto 0

	If xatm_Disconnector.UseDoublePoints Then
		IsSourceHealthy = IsTagHealthy(tagOpen) Or IsTagHealthy(tagClosed)
	Else
		IsSourceHealthy = IsTagHealthy(tagOpen) And IsTagHealthy(tagClosed)
	End If

End Function

Function IsTagHealthy(tag)

	IsTagHealthy = False

	If tag Is Nothing Then
		Exit Function
	End If

	Dim quality
	On Error Resume Next
	quality = tag.Quality
	IsTagHealthy = (Err.Number = 0) And (quality >= 192)
	Err.Clear
	On Error Goto 0

End Function

' An unassociated XObject property yields Nothing (or an object that cannot be read),
' which is how an unused redundant relay / missing SBO tag is detected.
Function IsTagLinked(tag)

	IsTagLinked = False

	If tag Is Nothing Then
		Exit Function
	End If

	Dim probe
	On Error Resume Next
	probe = tag.Value
	IsTagLinked = (Err.Number = 0)
	Err.Clear
	On Error Goto 0

End Function

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

<xatm_Disconnector.Data.CommunicationFailure:CommunicationFailure_E_OnChangedOpenPositionQuality()>
Sub CommunicationFailure_E_OnChangedOpenPositionQuality()
	
	SetCommunicationFailure
		
End Sub

<xatm_Disconnector.Data.CommunicationFailure:CommunicationFailure_Functions()>
Sub CommunicationFailure_Functions()
End Sub


Function IsCommunicationHealthy()
	
	IsCommunicationHealthy = False

	If xatm_Disconnector.UseDoublePoints Then

		' ======================
		' Double Point
		' ======================
		
		Dim tags
		tags = Array(xatm_Disconnector.PositionClosed, _
					xatm_Disconnector.PositionOpen, _
					xatm_Disconnector.PositionClosedAlt, _
					xatm_Disconnector.PositionOpenAlt)
		
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
		mainHealthy = (xatm_Disconnector.PositionOpen.Quality >= 192) And (xatm_Disconnector.PositionClosed.Quality >= 192)
		
		Dim altHealthy
		altHealthy  = (xatm_Disconnector.PositionOpenAlt.Quality >= 192) And (xatm_Disconnector.PositionClosedAlt.Quality >= 192)
		
		IsCommunicationHealthy = mainHealthy Or altHealthy

	End If
	
End Function


Sub SetCommunicationFailure()
	
	Dim communicationFailure
	communicationFailure = Not IsCommunicationHealthy()
	
	WriteEx communicationFailure 
	
End Sub

<xatm_Disconnector.Data.CommunicationFailure:CommunicationFailure_OnChangedSimMode()>
Sub CommunicationFailure_OnChangedSimMode()
	
	If CBool(Parent.Item("SimulationModeEnabled").Value) Then
	
		WriteEx False
	
	Else 
	
		SetCommunicationFailure
	
	End If
		
End Sub

<xatm_Disconnector.Data.CommunicationFailure:CommunicationFailure_OnStartRunning()>
Sub CommunicationFailure_OnStartRunning()
	
	SetCommunicationFailure
		
End Sub

<xatm_Disconnector.Data.MemorizedPosition:MemorizedPosition_OnChangedPosition()>
Sub MemorizedPosition_OnChangedPosition()

	If Value = 0 Then
		
		Value = Parent.Item("Position").Value
		
	Else
		
		' Init timer
		DocString = 120
		
	End If
	
End Sub

<xatm_Disconnector.Data.MemorizedPosition:MemorizedPosition_Timer()>
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

<xatm_Disconnector.Data.Position:Position_E_OnChangedClosedPosition()>
Sub Position_E_OnChangedClosedPosition()

	CalculatePosition
	
End Sub

<xatm_Disconnector.Data.Position:Position_E_OnChangedClosedQuality()>
Sub Position_E_OnChangedClosedQuality()
	
	CalculatePosition
	
End Sub

<xatm_Disconnector.Data.Position:Position_E_OnChangedOpenPosition()>
Sub Position_E_OnChangedOpenPosition()
	
	CalculatePosition
	
End Sub

<xatm_Disconnector.Data.Position:Position_E_OnChangedOpenPositionQuality()>
Sub Position_E_OnChangedOpenPositionQuality()
	
	CalculatePosition
	
End Sub

<xatm_Disconnector.Data.Position:Position_Functions()>
Sub Position_Functions()

End Sub

Sub CalculatePosition()

	Dim resolved
	resolved = ResolvePosition()
	
	If Value <> resolved Then

		WriteEx resolved
		
		On Error Resume Next
		Application.Trace "[" & xatm_Disconnector.Name & "] - Position changed to: " & resolved
		On Error Goto 0
		
	End If

End Sub

Function ResolvePosition()

	ResolvePosition = 0

	If xatm_Disconnector.UseDoublePoints Then

		' ======================
		' Double Point
		' ======================
	
		tags = Array(xatm_Disconnector.PositionClosed, _
					xatm_Disconnector.PositionOpen, _
					xatm_Disconnector.PositionClosedAlt, _
					xatm_Disconnector.PositionOpenAlt)
    
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

		If IsObjectExists(xatm_Disconnector.PositionClosed) And IsObjectExists(xatm_Disconnector.PositionOpen) Then
			
			If xatm_Disconnector.PositionClosed.Quality >= 192 And xatm_Disconnector.PositionOpen.Quality >= 192 Then
				
				ResolvePosition = CalculateSinglePointState(xatm_Disconnector.PositionOpen.Value, _
						xatm_Disconnector.PositionClosed.Value)
				Exit Function

			End If

		End If
		
		If IsObjectExists(xatm_Disconnector.PositionClosedAlt) And IsObjectExists(xatm_Disconnector.PositionOpenAlt) Then
			
			If xatm_Disconnector.PositionClosedAlt.Quality >= 192 And xatm_Disconnector.PositionOpenAlt.Quality >= 192 Then
				
				ResolvePosition = CalculateSinglePointState(xatm_Disconnector.PositionOpenAlt.Value, _
						xatm_Disconnector.PositionClosedAlt.Value)
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
		
		Case xatm_Disconnector.RawValueOpen
			CalculateDoublePointState = 1
		
		Case xatm_Disconnector.RawValueClosed
			CalculateDoublePointState = 2
		
		Case Else
			CalculateDoublePointState = 0

	End Select

End Function

Function CalculateSinglePointState(rawValueOpen, rawValueClosed)
	
	If rawValueOpen = xatm_Disconnector.RawValueOpen And rawValueClosed <> xatm_Disconnector.RawValueClosed Then
		
		CalculateSinglePointState = 1
	
	ElseIf rawValueOpen <> xatm_Disconnector.RawValueOpen And rawValueClosed = xatm_Disconnector.RawValueClosed Then
		
		CalculateSinglePointState = 2
	
	Else
	
		CalculateSinglePointState = 0
	
	End If

End Function

Sub Foo()	
End Sub

<xatm_Disconnector.Data.Position:Position_OnChangedSimMode()>
Sub Position_OnChangedSimMode()

	If CBool(Parent.Item("SimulationModeEnabled").Value) Then
		
		WriteEx (xatm_Disconnector.NormalState + 1)
		
	Else
		
		WriteEx 0
		
		CalculatePosition
	
	End If
	
End Sub

<xatm_Disconnector.Data.Position:Position_OnStartRunning()>
Sub Position_OnStartRunning()

	CalculatePosition
	
End Sub

<xatm_Disconnector.Data.Reset:Reset_OnChangedTimeStamp()>
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
	
	' ================
	' RESET COMMAND FAILURES
	' ================
	' The two latches CommandTimer sets on a timeout. Here is the only
	' place they are cleared - a failure that cleared itself on the next
	' command would be gone before anybody saw it.
	xatm_Disconnector.CommandOpenFailed  = False
	xatm_Disconnector.CommandCloseFailed = False
	
End Sub

<xatm_Disconnector.Data.Timers.CommandTimer:CommandTimer_Counter()>
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

		' Which of the two it was, latched for whoever looks later - the
		' automation goes to global lockout on CommandInProgress alone,
		' and by the time anybody reads a screen the timer has stopped
		' counting and CommandOpenClose says nothing about what failed.
		'
		' Set after the stop above and not before it, so a disconnector whose
		' class has not been given these properties yet still stops its
		' timer and still tells the automation the command failed.
		If command = 1 Then
			xatm_Disconnector.CommandOpenFailed = True
		ElseIf command = 2 Then
			xatm_Disconnector.CommandCloseFailed = True
		End If

		Exit Sub

	End If

	' --- Retry once, at half the timeout ---
	If Value = Int(xatm_Disconnector.CommandTimeout / 2) Then
		
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

<xatm_RASEAT.Commands.OperatorBlock:OperatorBlock_CommandOperatorBlock()>
Sub OperatorBlock_CommandOperatorBlock()

	Dim v
	v = False

	On Error Resume Next
	v = CBool(xatm_RASEAT.CommandOperatorBlock.Value)
	On Error Goto 0

	xatm_RASEAT.OperatorBlock = v

End Sub

<xatm_RASEAT.Commands.OperatorBlock:OperatorBlock_OnChangedValue()>
Sub OperatorBlock_OnChangedValue()

	If xatm_RASEAT.OperatorBlock Then
		WriteLog "Blocked by operator."
	Else
		WriteLog "Released by operator."
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

<xatm_RASEAT.Commands.Reset:Reset_OnChangedValue()>
Sub Reset_OnChangedValue()

	If CBool(Value) Then Reset()

End Sub

Sub Reset()

	xatm_RASEAT.Running      = False
	xatm_RASEAT.GeneralBlock = False
	xatm_RASEAT.Successful   = False
	xatm_RASEAT.Unsuccessful = False

	Dim i
	For i = 1 To 6

		Dim propName
		propName = "StepExecutionFailed" & i

		On Error Resume Next
		Execute("xatm_RASEAT." & propName & " = False")
		On Error Goto 0

	Next

	' The state the run was carrying, cleared with it. A step number left
	' behind would be read by the next start as a maneuver already under
	' way.
	Dim tag
	For Each tag In xatm_RASEAT.Item("FSM")

		If TypeName(tag) = "InternalTag" Then
			tag.WriteEx Empty, tag.TimeStamp
		End If

	Next
	
	' ================
	' RESET ALL DEVICES
	' ================
	ResetDevices Application.GetObject("XATM_Data.Substation")

	WriteLog "Reset"

End Sub


Sub ResetDevices(folder)
	
	For Each obj In folder
	
		If TypeName(obj) = "xatm_Breaker" Then
		
			On Error Resume Next
			obj.Item("Data").Item("Reset").WriteEx True
			On Error Goto 0
			
		Else
		
			ResetDevices obj
		
		End If
	
	Next

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

<xatm_RASEAT.Commands.Reset:Reset_Reset()>
Sub Reset_Reset()

	If xatm_RASEAT.CommandReset.Value = 0 Then Exit Sub

	WriteEx True
	
End Sub

<xatm_RASEAT.Commands.Start:Start_OnChangedValue()>
Sub Start_OnChangedValue()

	' Asked for by a transformer whose lockout relay operated, and by nothing
	' There is no operator command for this: the entry is already dark by
	' the time anybody could press one.
	'
	' The value carries the Id of the transformer that asked, which is for
	' the log alone - Step 0 reads the relay of every transformer for itself,
	' because more than one can have operated and the sequence has to wait
	' for all of them.

	If Trim(Value) = "" Then Exit Sub      ' self-cleared write, ignore

	Dim ts
	ts = TimeStamp	' preserve for the silent clear

	Dim askedBy
	askedBy = 0
	If IsNumeric(Value) Then askedBy = CInt(Value)

	' ================
	' GATE CHECKS
	' ================

	If Not xatm_RASEAT.Enabled Then

		Reject "Automation disabled.", ts
		Exit Sub

	End If

	' Refused and logged, like any other. Two transformers losing their relay
	' together, or one shortly after the other, is unusual enough to be
	' worth a line - and Step 0 already took whichever were carrying it
	' when the sequence began, so the second request has nothing to add.
	If xatm_RASEAT.Running Then

		Reject "Already running.", ts
		Exit Sub

	End If

	If xatm_RASEAT.OperatorBlock Then

		Reject "Blocked by operator.", ts
		Exit Sub

	End If

	If xatm_RASEAT.GeneralBlock Then

		Reject "General interlock active.", ts
		Exit Sub

	End If

	If xatm_RASEAT.AutomaticBlock Then

		Reject "Blocked by field conditions.", ts
		Exit Sub

	End If

	' Reads the other way round - the expression says when the maneuver may
	' go ahead, so it is the absence of it that rejects.
	If Not xatm_RASEAT.Preconditions Then

		Reject "Preconditions are not met.", ts
		Exit Sub

	End If

	If AnyOtherAutomationRunning() Then

		Reject "Another automation is in progress.", ts
		Exit Sub

	End If

	' ================
	' START
	' ================
	'
	' The results of the last run are cleared here rather than at the end
	' of it, so the control room keeps seeing how the previous reclosing
	' went until a new one actually begins.
	xatm_RASEAT.Successful   = False
	xatm_RASEAT.Unsuccessful = False

	xatm_RASEAT.Item("FSM").Item("StepTimer").WriteEx 0
	xatm_RASEAT.Item("FSM").Item("Main").WriteEx 0
	xatm_RASEAT.Running = True

	WriteLog "Start - asked by " & TransformerName(askedBy)

	WriteEx "", ts ' clear without re-firing

End Sub


' True if any OTHER automation object is currently running.
'
' The transfer that answers the same trip is not an "other". One lockout
' relay starts both now: the transfer moves the medium voltage off the
' transformer that tripped, and this recloses the entry above it. Different
' equipment, the same event, and neither waits for the other - which is the
' whole point of the change, and the opposite of what this check used to do.
'
' A manual transfer still blocks, in both directions. That one is somebody
' operating, and a reclosing should not cut across it any more than it
' should be cut across.
Function AnyOtherAutomationRunning()

	AnyOtherAutomationRunning = False

	Dim obj
	For Each obj In Application.GetObject("XATM_Data.Automation")

		If Not (obj Is xatm_RASEAT) And TypeName(obj) <> "xatm_TA" Then

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


' The name of the transformer that asked, for the log. Its own scope's
' copy: E3 gives no way to call the one the state machine keeps.
Function TransformerName(id)

	TransformerName = "ID " & id

	Dim obj
	For Each obj In Application.GetObject("XATM_Data.Substation")

		Dim found
		Set found = FindTransformer(obj, id)

		If Not found Is Nothing Then
			TransformerName = found.Name
			Exit Function
		End If

	Next

End Function

Function FindTransformer(folder, id)

	Set FindTransformer = Nothing

	If UCase(TypeName(folder)) = "XATM_TRANSFORMER" Then

		On Error Resume Next
		If folder.Id = id Then Set FindTransformer = folder
		On Error Goto 0

		Exit Function

	End If

	Dim obj
	On Error Resume Next
	For Each obj In folder

		Dim found
		Set found = FindTransformer(obj, id)

		If Not found Is Nothing Then
			Set FindTransformer = found
			Exit Function
		End If

	Next
	On Error Goto 0

End Function


' Logs the rejection and clears the trigger without re-firing.
Sub Reject(reason, ts)

	WriteLog "Not started - " & reason
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

<xatm_RASEAT.FSM.Main:Main_Completed()>
Sub Main_Completed()

	Parent.Item("PrimaryBreakerId").WriteEx    Empty, 0
	Parent.Item("TriggerTransformers").WriteEx Empty, 0
	Parent.Item("StepTimer").WriteEx           Empty, 0
	WriteEx Empty, 0

	xatm_RASEAT.Running = False
	WriteLog "Reclosing completed."
		
End Sub

<xatm_RASEAT.FSM.Main:Main_Functions()>
Sub Main_Functions()
End Sub


Function GetDeviceById(id, ByRef exists)

	exists = False
	Set GetDeviceById = FindInFolder(Application.GetObject("XATM_Data.Substation"), id, exists)

End Function

Function FindInFolder(folder, id, ByRef exists)

	Set FindInFolder = Nothing

	Dim obj
	For Each obj In folder

		Select Case UCase(TypeName(obj))

			Case "XATM_BREAKER", "XATM_TRANSFORMER", "XATM_DISCONNECTOR"
				If obj.Id = id Then
					Set FindInFolder = obj
					exists = True
					Exit Function
				End If

			Case Else
				Dim found
				On Error Resume Next
				Set found = FindInFolder(obj, id, exists)
				On Error GoTo 0

				If exists Then
					Set FindInFolder = found
					Exit Function
				End If

		End Select

	Next

End Function


' The incomer axis on its own, which is the one this automation answers to.
' The transfer concatenates Transformer and Busbar because its step tables
' turn on both; no step here operates a busbar tie.
Function GetIncomerLayout()

	GetIncomerLayout = ""

	On Error Resume Next
	GetIncomerLayout = Application.GetObject("XATM_Data.Automation.Layout.Incomer").Value
	On Error Goto 0

End Function


' The incomer breakers, in the order the layout declares them.
'
' Ids and not names: the tenth and twentieth of the reserved 10-99 band are
' the two entry bays, and the layout document is what says so.
Function IncomerIds()

	Select Case UCase(GetIncomerLayout() & "")
		Case "2BR2BB" : IncomerIds = Array(10, 20)
		Case Else     : IncomerIds = Array()
	End Select

End Function


' Every transformer the substation has, as a comma separated list of Ids.
' Walked rather than assumed, so a layout with two transformers answers
' with two.
Function TransformerIds()

	Dim acc
	acc = ""

	On Error Resume Next
	CollectTransformers Application.GetObject("XATM_Data.Substation"), acc
	On Error Goto 0

	TransformerIds = acc

End Function

Sub CollectTransformers(folder, ByRef acc)

	Dim obj
	For Each obj In folder

		Select Case UCase(TypeName(obj))

			Case "XATM_TRANSFORMER"

				Dim id
				id = Empty

				On Error Resume Next
				id = obj.Id
				On Error Goto 0

				If IsNumeric(id) Then
					If acc <> "" Then acc = acc & ","
					acc = acc & CInt(id)
				End If

			Case "XATM_BREAKER", "XATM_DISCONNECTOR"

				' a device, and not one this is looking for

			Case Else

				On Error Resume Next
				CollectTransformers obj, acc
				On Error Goto 0

		End Select

	Next

End Sub


' Which incomer was carrying the station when the relay operated, and the
' other one. Both are snapshotted at Step 0 and held for the run.
Function ReadPrimaryId()

	ReadPrimaryId = 0

	Dim id
	id = Empty

	On Error Resume Next
	id = Parent.Item("PrimaryBreakerId").Value
	On Error Goto 0

	If IsNumeric(id) Then ReadPrimaryId = CInt(id)

End Function

Function ReadBackupId()

	ReadBackupId = 0

	Dim ids, i, primary
	ids = IncomerIds()
	primary = ReadPrimaryId()

	For i = 0 To UBound(ids)
		If ids(i) <> primary Then
			ReadBackupId = ids(i)
			Exit Function
		End If
	Next

End Function


' The transformers whose lockout relay operated, as the run treats them.
'
' Snapshotted rather than re-read, because the relay resets on its own and
' sequence would otherwise stop waiting for a transformer half way through
' isolating - the same reason the reference logic memorises it.
Function ReadTriggerTransformers()

	Dim s
	s = ""

	On Error Resume Next
	s = Parent.Item("TriggerTransformers").Value
	On Error Goto 0

	If IsEmpty(s) Or IsNull(s) Then s = ""

	If s = "" Then
		ReadTriggerTransformers = Array()
	Else
		ReadTriggerTransformers = Split(s, ",")
	End If

End Function


Sub ResetTimer()

	Parent.Item("StepTimer").Value = 0

End Sub

Sub IncrementTimer()

	Parent.Item("StepTimer").Value = Parent.Item("StepTimer").Value + 1

End Sub

Function Elapsed()

	Elapsed = Parent.Item("StepTimer").Value

End Function


' True when the step is done with, so the caller can advance. Written as a
' question rather than a command because three of the steps answer it by
' expiring and three by being answered.
Function TimedOut(seconds)

	TimedOut = (Elapsed() >= seconds)

End Function


Function OperationName(operation)

	If operation = 2 Then
		OperationName = "closed"
	ElseIf operation = 1 Then
		OperationName = "opened"
	Else
		OperationName = "operated"
	End If

End Function


' Moves to the next step and starts its clock. Every step ends this way, so
' no step can advance and leave the previous step's elapsed time behind it.
Sub Advance(nextStep)

	ResetTimer()
	Value = nextStep

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

<xatm_RASEAT.FSM.Main:Main_GlobalLockout()>
Sub Main_GlobalLockout()

	xatm_RASEAT.Running        = False
	xatm_RASEAT.GeneralBlock   = True
	xatm_RASEAT.Unsuccessful   = True

	Select Case Value
		Case 1 : xatm_RASEAT.StepExecutionFailed1 = True
		Case 2 : xatm_RASEAT.StepExecutionFailed2 = True
		Case 3 : xatm_RASEAT.StepExecutionFailed3 = True
		Case 4 : xatm_RASEAT.StepExecutionFailed4 = True
		Case 5 : xatm_RASEAT.StepExecutionFailed5 = True
		Case 6 : xatm_RASEAT.StepExecutionFailed6 = True
	End Select

	WriteLog "General block - the reclosing did not complete."	
	
End Sub

<xatm_RASEAT.FSM.Main:Main_Main()>
Sub Main_Main()

	If Not xatm_RASEAT.Enabled Then

		WriteLog "Automation not enabled."
		xatm_RASEAT.Running = False
		Exit Sub

	End If

	If Not xatm_RASEAT.Running Then Exit Sub

	' Dispatch and nothing else.
	'
	' The transfer checks one timeout here before calling any step, because
	' every one of its steps fails when its timer runs out. Half of these do
	' not: step 3 is a dwell that succeeds by expiring, and steps 4 and 5
	' fall through to the next way of confirming the same close. Only 1, 2
	' and 6 end the maneuver, so what expiry means is the step's to say.
	Select Case Value

		Case 0  : Main_Step00()
		Case 1  : Main_Step01()
		Case 2  : Main_Step02()
		Case 3  : Main_Step03()
		Case 4  : Main_Step04()
		Case 5  : Main_Step05()
		Case 6  : Main_Step06()
		Case 99 : Main_Completed()

	End Select

	IncrementTimer()

End Sub


' The step timings, in seconds.
'
' Declared here rather than configured: they are numbers this sequence is
' written around, not dials a panel offers. ISOLATION_TIMEOUT is the one to
' watch - it has to cover the slowest tripped transformer finishing its
' isolation, two motorised disconnectors travelling included, and 25 is the
' figure the reference logic used for a different station.
Const BREAKER_TIMEOUT   = 5     ' a breaker confirming a position
Const ISOLATION_TIMEOUT = 25    ' every tripped transformer reporting Isolated
Const CURRENT_TIMEOUT   = 5     ' load current vouching for a close
Const LATCH_DELAY       = 2     ' the bistable mechanism settling
Const RESET_DELAY       = 5     ' the results standing before they are cleared

Sub Foo()	
End Sub

<xatm_RASEAT.FSM.Main:Main_Step00()>
Sub Main_Step00()

	' Everything the run is decided by, read once.
	'
	' Both of these change under a running sequence if they are re-read: a
	' The relay resets on its own, and MemorizedPosition adopts the new position
	' once its own countdown expires. Snapshotting is what keeps step 6
	' operating the breaker step 4 was talking about.
	StoreTriggerTransformers()
	StorePrimaryBreaker()

	If ReadPrimaryId() = 0 Then

		WriteLog "Step 0: neither incomer was carrying the station - nothing to reclose."
		Main_GlobalLockout()
		Exit Sub

	End If

	WriteLog "Starting " & DescribeReclosing()

	Advance 1

End Sub


' Which transformers had their lockout relay operate, written down as a list.
'
' Taken from the live signal at trigger time. A transformer whose relay has
' already reset by the time step 2 runs is still waited on, because this is
' what step 2 reads - which is the whole of what the reference logic's
' 28 second memory was for.
Sub StoreTriggerTransformers()

	Dim ids, i, id, transformer, exists, acc
	ids = Split(TransformerIds(), ",")
	acc = ""

	For i = 0 To UBound(ids)

		If ids(i) <> "" Then

			id = CInt(ids(i))
			Set transformer = GetDeviceById(id, exists)

			If exists Then

				Dim operated
				operated = False

				On Error Resume Next
				operated = CBool(transformer.LockingOutRelay)
				On Error Goto 0

				If operated Then
					If acc <> "" Then acc = acc & ","
					acc = acc & id
				End If

			End If

		End If

	Next

	Parent.Item("TriggerTransformers").WriteEx acc

End Sub


' Which incomer was carrying the station.
'
' MemorizedPosition and not Position: by the time this runs the breaker may
' already have tripped, and the memorised value is the one that still says
' what the topology was. It holds the pre-change position until its own
' countdown expires, which is longer than this sequence takes - but only
' if nobody re-reads it later, which is why this is written down.
Sub StorePrimaryBreaker()

	Dim ids, i, breaker, exists, primary
	ids = IncomerIds()
	primary = 0

	For i = 0 To UBound(ids)

		Set breaker = GetDeviceById(ids(i), exists)

		If exists Then

			Dim memorized
			memorized = 0

			On Error Resume Next
			memorized = breaker.Item("Data").Item("MemorizedPosition").Value
			On Error Goto 0

			' 2 is closed. The first one found carrying is taken, which is
			' the only answer when the station is running normally - and a
			' defined one for the abnormal case where both were closed.
			If memorized = 2 And primary = 0 Then primary = ids(i)

		End If

	Next

	Parent.Item("PrimaryBreakerId").WriteEx primary

End Sub


' What the log says a run is, in the words the control room uses.
Function DescribeReclosing()

	Dim trs, i, list
	trs = ReadTriggerTransformers()
	list = ""

	For i = 0 To UBound(trs)
		If list <> "" Then list = list & ", "
		list = list & DeviceName(CInt(trs(i)))
	Next

	If list = "" Then list = "no transformer"

	DescribeReclosing = "reclosing on " & DeviceName(ReadPrimaryId()) & _
	                    ", tripped by " & list

End Function


' A device's name for the log, falling back to its Id when the project has
' no such device.
Function DeviceName(id)

	Dim device, exists
	Set device = GetDeviceById(id, exists)

	If exists Then
		DeviceName = device.Name
	Else
		DeviceName = "ID " & id
	End If

End Function

' A scope may not end on a Function, and E3 says nothing when one does.
Sub Step00Functions()
	
End Sub

<xatm_RASEAT.FSM.Main:Main_Step01()>
Sub Main_Step01()

	' The whole entry has to be dead before anything is reclosed.
	'
	' Both incomers, not just the one that was carrying: the relay trips
	' them together, and a station with one still closed is not a station
	' waiting to be restored - it is one nobody understands yet.
	Dim ids, i, breaker, exists, allOpen
	ids = IncomerIds()

	If UBound(ids) < 0 Then

		WriteLog "Step 1: the incomer layout declares no entry bays - nothing to reclose."
		Main_GlobalLockout()
		Exit Sub

	End If

	allOpen = True

	For i = 0 To UBound(ids)

		Set breaker = GetDeviceById(ids(i), exists)

		If Not exists Then

			WriteLog "Step 1: incomer with ID=" & ids(i) & " was not found. Please check the configuration - general block"
			Main_GlobalLockout()
			Exit Sub

		End If

		If breaker.Item("Data").Item("Position").Value <> 1 Then allOpen = False

	Next

	If allOpen Then

		WriteLog "Step 1: both incomers open - proceeding to the next step."
		Advance 2
		Exit Sub

	End If

	If TimedOut(BREAKER_TIMEOUT) Then

		WriteLog "Step 1: an incomer did not open inside " & BREAKER_TIMEOUT & "s - general block"
		Main_GlobalLockout()

	End If
		
End Sub

<xatm_RASEAT.FSM.Main:Main_Step02()>
Sub Main_Step02()
	
	' Every transformer whose lockout relay operated, not one of them.
	'
	' Several can trip together, and that is not an ambiguity to resolve:
	' the entry may not be re-energised while any of them is still tied to
	' it, so the step waits for the slowest.
	Dim trs, i, id, transformer, exists, allIsolated
	trs = ReadTriggerTransformers()

	' No relay anywhere means the trigger was withdrawn before this ran. There
	' is nothing to isolate, and nothing to wait for.
	If UBound(trs) < 0 Then

		WriteLog "Step 2: no transformer is carrying its lockout relay - proceeding to the next step."
		Advance 3
		Exit Sub

	End If

	allIsolated = True

	For i = 0 To UBound(trs)

		id = CInt(trs(i))
		Set transformer = GetDeviceById(id, exists)

		If Not exists Then

			WriteLog "Step 2: transformer with ID=" & id & " was not found. Please check the configuration - general block"
			Main_GlobalLockout()
			Exit Sub

		End If

		Dim isolated
		isolated = False

		On Error Resume Next
		isolated = CBool(transformer.Isolated)
		On Error Goto 0

		If Not isolated Then allIsolated = False

	Next

	If allIsolated Then

		WriteLog "Step 2: every tripped transformer is isolated - proceeding to the next step."
		Advance 3
		Exit Sub

	End If
	
	' Indicates what the step is waiting for after 3 seconds
	If Elapsed() = 3 Then

		WriteLog "Step 2: waiting for " & NotIsolated() & " to isolate."

	End If

	If TimedOut(ISOLATION_TIMEOUT) Then

		WriteLog "Step 2: " & NotIsolated() & " did not isolate inside " & ISOLATION_TIMEOUT & "s - general block"
		Main_GlobalLockout()

	End If

End Sub


' The transformers still holding step 2 up, named for the log. Worth the
' walk: the step waits on all of them, and being told which one refused is
' the difference between a fault to chase and a message to shrug at.
Function NotIsolated()

	Dim trs, i, id, transformer, exists, list
	trs = ReadTriggerTransformers()
	list = ""

	For i = 0 To UBound(trs)

		id = CInt(trs(i))
		Set transformer = GetDeviceById(id, exists)

		If exists Then

			Dim isolated
			isolated = False

			On Error Resume Next
			isolated = CBool(transformer.Isolated)
			On Error Goto 0

			If Not isolated Then
				If list <> "" Then list = list & ", "
				list = list & transformer.Name
			End If

		End If

	Next

	If list = "" Then list = "a transformer"

	NotIsolated = list

End Function

' A scope may not end on a Function, and E3 says nothing when one does.
Sub Step02Functions()	
End Sub

<xatm_RASEAT.FSM.Main:Main_Step03()>
Sub Main_Step03()
	
	' A dwell, and the only step that succeeds by running out of time.
	'
	' The breaker's mechanism has to settle before it is asked to close
	' again; closing into an unlatched bistable is how a close is lost
	' without anything reporting a failure.
	If TimedOut(LATCH_DELAY) Then

		WriteLog "Step 3: waited " & LATCH_DELAY & "s for the mechanism - proceeding to the next step."
		Advance 4

	End If
		
End Sub

<xatm_RASEAT.FSM.Main:Main_Step04()>
Sub Main_Step04()
	
	' Close the incomer that was carrying the station.
	'
	' Expiring here is not a failure. It means the position contacts have
	' not vouched for the close, and step 5 asks the current instead - the
	' breaker may well be closed and only saying so slowly.
	Dim breaker, exists
	Set breaker = GetDeviceById(ReadPrimaryId(), exists)

	If Not exists Then

		WriteLog "Step 4: incomer with ID=" & ReadPrimaryId() & " was not found. Please check the configuration - general block"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = 2 Then

		WriteLog "Step 4: " & breaker.Name & " closed - the reclosing succeeded."
		Succeed()
		Exit Sub

	End If

	If TimedOut(BREAKER_TIMEOUT) Then
		
		' The breaker is told to drop the command before this step lets go
		' of it.
		'
		' Its own timer is still running: a command it has not confirmed is
		' resent once, at half of CommandTimeout, which is far longer than
		' this step waited. Left alone it would close this incomer long
		' after step 6 had closed the other, and the station would find
		' itself with both.
		breaker.Item("Data").Item("Reset").WriteEx True

		WriteLog "Step 4: " & breaker.Name & " has not confirmed closed - asking the current."
		Advance 5
		Exit Sub

	End If

	IssueClose breaker
		
End Sub

<xatm_RASEAT.FSM.Main:Main_Step05()>
Sub Main_Step05()

	' The same close, vouched for by the current instead of the contacts.
	'
	' A breaker carrying load is closed whatever its auxiliary contacts
	' say, and saying so here is what keeps a sticking contact from
	' sending the station to its second incomer for no reason.
	Dim breaker, exists
	Set breaker = GetDeviceById(ReadPrimaryId(), exists)

	If Not exists Then

		WriteLog "Step 5: incomer with ID=" & ReadPrimaryId() & " was not found. Please check the configuration - general block"
		Main_GlobalLockout()
		Exit Sub

	End If

	Dim carrying
	carrying = False

	On Error Resume Next
	carrying = CBool(breaker.HasLoadCurrent)
	On Error Goto 0

	If carrying Then

		WriteLog "Step 5: " & breaker.Name & " is carrying load - the reclosing succeeded."
		Succeed()
		Exit Sub

	End If

	If TimedOut(CURRENT_TIMEOUT) Then

		WriteLog "Step 5: " & breaker.Name & " is not carrying load - trying the other incomer."
		Advance 6

	End If
	
End Sub

<xatm_RASEAT.FSM.Main:Main_Step06()>
Sub Main_Step06()

	' The other incomer, and the last thing tried.
	'
	' Expiring here does end the maneuver: there is no third entry to fall
	' back to.
	Dim breaker, exists
	Set breaker = GetDeviceById(ReadBackupId(), exists)

	If Not exists Then

		WriteLog "Step 6: there is no second incomer to fall back to - general block"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = 2 Then

		WriteLog "Step 6: " & breaker.Name & " closed - the reclosing succeeded."
		Succeed()
		Exit Sub

	End If

	' Asked before the command and not once at the start: the backup sat
	' unused through five steps, and whether it may be closed is a question
	' about now rather than about then.
	Dim defective
	defective = False

	On Error Resume Next
	defective = CBool(breaker.Defective)
	On Error Goto 0

	If defective Then

		WriteLog "Step 6: " & breaker.Name & " is defective and must not be operated - general block"
		Main_GlobalLockout()
		Exit Sub

	End If

	If TimedOut(BREAKER_TIMEOUT) Then

		WriteLog "Step 6: " & breaker.Name & " did not close - general block"
		Main_GlobalLockout()
		Exit Sub

	End If

	IssueClose breaker

End Sub


' Sends the close, once, and reads what the breaker says about the one it
' is already carrying out. The same three answers the transfer reads, and
' for the same reasons.
Sub IssueClose(breaker)

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command
			breaker.Item("Data").Item("CommandOpenClose").WriteEx 2

		Case 1

			' The breaker reported the command failed. Nothing is retried
			' here: the step's own clock decides what happens next, and on
			' step 4 that is to ask the current rather than to give up.

		Case 2

			' In progress

	End Select

End Sub


' What every successful ending does. Held rather than cleared, so the
' control room sees the result before Main_Completed takes it away.
Sub Succeed()

	xatm_RASEAT.Successful = True
	Advance 99
	
End Sub

<xatm_TA.Commands.OperatorBlock:OperatorBlock_CommandOperatorBlock()>
Sub OperatorBlock_CommandOperatorBlock()
	
	Dim v
	v = False
	
	On Error Resume Next
	v = CBool(xatm_TA.CommandOperatorBlock.Value)
	On Error Goto 0
	
	xatm_TA.OperatorBlock = v

End Sub

<xatm_TA.Commands.OperatorBlock:OperatorBlock_OnChangedValue()>
Sub OperatorBlock_OnChangedValue()
	
	If xatm_TA.OperatorBlock Then
		WriteLog "Blocked by operator."
	Else
		WriteLog "Released by operator."
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

<xatm_TA.Commands.Reset:Reset_OnChangedTimeStamp()>
Sub Reset_OnChangedTimeStamp()
	
	If CBool(Value) Then
		
		Reset()
		
	End If
	
End Sub

Sub Reset()
	
	xatm_TA.Running 		= False
	xatm_TA.GeneralBlock	= False
	
	' ===============
	' RESET ALL STEP FAIL POINTS
	' ===============
	Dim i
	For i = 1 To 6
		Dim propName
		propName = "StepExecutionFailed" & i
		
		On Error Resume Next
		Execute("xatm_TA." & propName & " = False")
		On Error Goto 0
	Next
	
	Dim tag
	For Each tag In xatm_TA.Item("FSM")
		
		If TypeName(tag) = "InternalTag" Then

			tag.WriteEx Empty, tag.TimeStamp
		
		End If
		
	Next
	
	' ================
	' RESET ALL DEVICES
	' ================
	ResetDevices Application.GetObject("XATM_Data.Substation")
	
	WriteLog "Reset"

End Sub

Sub ResetDevices(folder)
	
	For Each obj In folder
	
		If TypeName(obj) = "xatm_Breaker" Then
		
			On Error Resume Next
			obj.Item("Data").Item("Reset").WriteEx True
			On Error Goto 0
			
		Else
		
			ResetDevices obj
		
		End If
	
	Next

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

<xatm_TA.Commands.Reset:Reset_Reset()>
Sub Reset_Reset()
	
	If xatm_TA.CommandReset.Value = 0 Then Exit Sub
	
	WriteEx True
		
End Sub

<xatm_TA.Commands.Start:Start_OnChangedValue()>
Sub Start_OnChangedValue()

	If Trim(Value) = "" Then Exit Sub      ' self-cleared write, ignore

	Dim ts
	ts = TimeStamp	' preserve for the silent clear

	Dim parts
	parts = Split(Me.Value, ":")

	' Two fields, and only two.
	'
	' Nobody commands an automatic transfer - it answers a trip - so there
	' is no third field naming a contingency, because there was no operator
	' to pick one. Which transformer is out is read off the field instead,
	' below. That is the whole of what separates this class from the manual
	' one, and it is why the two were pulled apart.
	If UBound(parts) <> 1 Then
		Reject "Invalid Start format '" & Me.Value & "' (expected TA:trigger).", ts
		Exit Sub
	End If

	Dim mode
	mode = UCase(Trim(parts(0)))

	If mode <> "TA" Then
		Reject "This automation runs TA and nothing else - '" & mode & "' was asked for.", ts
		Exit Sub
	End If

	If Not IsNumeric(parts(1)) Then
		Reject "Invalid trigger transformer '" & parts(1) & "'.", ts
		Exit Sub
	End If

	Dim triggerId
	triggerId = CInt(parts(1))

	' ================
	' WHICH TRANSFORMER IS OUT
	' ================
	'
	' Read off the field, because the switchyard is what asked and the
	' switchyard says nothing.
	Dim impedeId, outCount
	impedeId = FieldImpediment(triggerId, outCount)

	' Tabela 1 situations 6-11 are two transformers out, and no sequence
	' leaves them. Refusing is the honest answer; picking one of the two
	' would run a sequence whose starting state is false.
	If outCount > 1 Then

		Reject "More than one transformer is out of service - no sequence is defined from there.", ts
		Exit Sub

	End If

	' ================
	' IS THE REQUEST ITSELF COHERENT
	' ================
	'
	' Ahead of the gates below, because these say the request is malformed
	' rather than that the switchyard is not ready for it.

	If Not IsConfiguredTransformer(triggerId) Then

		Reject "Trigger transformer " & triggerId & " is not configured in this substation.", ts
		Exit Sub

	End If

	' A transfer needs somewhere to put the load. Anything that is neither
	' the trigger nor out of service can take it; with none left there is
	' nothing to transfer to and no sequence to run.
	If CountDestinations(triggerId, impedeId) = 0 Then

		Reject "No transformer is left to take the load.", ts
		Exit Sub

	End If

	' ================
	' GATE CHECKS
	' ================

	If Not xatm_TA.Enabled Then

		Reject "Automation disabled.", ts
		Exit Sub

	End If

	If xatm_TA.Running Then

		Reject "Already running.", ts
		Exit Sub

	End If

	If xatm_TA.OperatorBlock Then

	 	Reject "Blocked by operator.", ts
		Exit Sub

	End If

	If xatm_TA.GeneralBlock Then

	 	Reject "General interlock active.", ts
	 	Exit Sub

	End If

	' One pair of gates, and named outright rather than built from the mode.
	'
	' The manual class keeps a pair per contingency because an operator
	' picks which maneuver to run and each one closes a different path. This
	' one is picked by a trip; there is nothing for a gate name to vary
	' with, and a scheme that restores a dead busbar should not have five
	' ways to be barred.
	'
	' Both are read before either is judged, so a manifest that has not been
	' rebuilt is reported as the configuration fault it is rather than
	' surfacing as a maneuver that will not start for no stated reason.
	Dim missing, blocked, permitted
	missing = ""

	blocked   = ReadGate("AutomaticBlock", False, missing)
	permitted = ReadGate("Preconditions",  True,  missing)

	If missing <> "" Then

		Reject "this automation has no " & missing & " - the manifest has not been rebuilt.", ts
		Exit Sub

	End If

	If blocked Then

		Reject "Blocked by field conditions.", ts
		Exit Sub

	End If

	' This one reads the other way round - the expression on it says when
	' the maneuver may go ahead, so it is the absence of it that rejects.
	If Not permitted Then

		Reject "Preconditions are not met.", ts
		Exit Sub

	End If

	If AnyOtherAutomationRunning() Then

		Reject "Another automation is in progress.", ts
		Exit Sub

	End If

	' ================
	' START
	' ================

	xatm_TA.Item("FSM").Item("AutomationType").WriteEx mode
	xatm_TA.Item("FSM").Item("TriggerTransformerId").WriteEx triggerId
	xatm_TA.Item("FSM").Item("ImpededTransformerId").WriteEx impedeId
	xatm_TA.Item("FSM").Item("StepTimer").WriteEx 0
	xatm_TA.Item("FSM").Item("Main").WriteEx 0
	xatm_TA.Running = True

	If impedeId = 0 Then
		WriteLog "Start - TA TR" & triggerId
	Else
		WriteLog "Start - TA TR" & triggerId & " with TR" & impedeId & " out of service"
	End If

	WriteEx "", ts ' clear without re-firing

End Sub

' True if any OTHER automation object is currently running (mutual exclusion).
' Relies on a common Running property instead of enumerating each type.
'
' The reclosing is not an "other" any more. One lockout relay starts both:
' this one moves the medium voltage off the transformer that tripped, that
' one recloses the entry above it. Different equipment, the same event, and
' neither has anything to wait for in the other.
'
' A manual transfer still blocks, in both directions. That one is somebody
' operating, and two sets of hands on the same busbars is what this check
' was put here for.
Function AnyOtherAutomationRunning()

	AnyOtherAutomationRunning = False

	Dim obj
	For Each obj In Application.GetObject("XATM_Data.Automation")

		If Not (obj Is xatm_TA) And TypeName(obj) <> "xatm_RASEAT" Then

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


' Every transformer the substation has, as "<id>:<0|1>" pairs - the flag is
' its out-of-service reading at the moment of asking.
'
' One walk, because the three questions below - is this Id configured, how
' many are out, which one is out - are all answered off the same list.
'
' FSM.Main has GetDeviceById for the same job. E3 gives no way to call a
' procedure in another object's scope, so this scope keeps its own, narrower
' version: it only ever looks for transformers.
Function TransformerStates()

	Dim acc
	acc = ""

	On Error Resume Next
	CollectTransformers Application.GetObject("XATM_Data.Substation"), acc
	On Error Goto 0

	TransformerStates = acc

End Function

Sub CollectTransformers(folder, ByRef acc)

	Dim obj
	For Each obj In folder

		Select Case UCase(TypeName(obj))

			Case "XATM_TRANSFORMER"

				Dim id
				id = Empty

				On Error Resume Next
				id = obj.Id
				On Error Goto 0

				If IsNumeric(id) Then

					Dim flag
					flag = 0

					On Error Resume Next
					If CBool(obj.OutOfService) Then flag = 1
					On Error Goto 0

					If acc <> "" Then acc = acc & ","
					acc = acc & CInt(id) & ":" & flag

				End If

			Case "XATM_BREAKER", "XATM_DISCONNECTOR"

				' a device, and not one this is looking for

			Case Else

				' Not a device -> treat as a subfolder and recurse.
				On Error Resume Next
				CollectTransformers obj, acc
				On Error Goto 0

		End Select

	Next

End Sub

' True when id names a transformer this substation actually has.
'
' Without it a typo'd screen binding, or a command tag left unconfigured,
' falls through every ElseIf in the step logic and lands on the Else that
' means "nothing is out" - running a real maneuver on a live switchyard
' under an assumption nobody made.
Function IsConfiguredTransformer(id)

	IsConfiguredTransformer = False

	Dim entries, i, pair
	entries = Split(TransformerStates(), ",")

	For i = 0 To UBound(entries)

		If entries(i) <> "" Then

			pair = Split(entries(i), ":")

			If CInt(pair(0)) = id Then
				IsConfiguredTransformer = True
				Exit Function
			End If

		End If

	Next

End Function

' How many transformers could take the load: everything that is neither the
' trigger, nor the one declared out, nor out of service in the field.
Function CountDestinations(triggerId, impedeId)

	CountDestinations = 0

	Dim entries, i, pair, id
	entries = Split(TransformerStates(), ",")

	For i = 0 To UBound(entries)

		If entries(i) <> "" Then

			pair = Split(entries(i), ":")
			id   = CInt(pair(0))

			If id <> triggerId And id <> impedeId And pair(1) = "0" Then
				CountDestinations = CountDestinations + 1
			End If

		End If

	Next

End Function

' What StoreImpediments used to snapshot, reduced to the one Id the
' sequences can take an argument for. This is the TA path: no operator said
' anything, so the field is asked.
'
' The trigger is skipped. A transformer that has just tripped may carry the
' out-of-service flag as well - stale, or set by whoever was about to work
' on it - and counting it would make the maneuver its own impediment, which
' Start then rejects. A protection-driven restoration refusing to run is the
' worst way to be wrong: the busbars stay dead. StoreImpediments had the
' same immunity, though by accident rather than on purpose - the step logic
' simply never tested imp(trigger).
'
' outCount is how many OTHERS were found out, so the caller can tell "none
' out" from "more out than any sequence covers" - 0 is returned for both.
Function FieldImpediment(triggerId, ByRef outCount)

	FieldImpediment = 0
	outCount = 0

	Dim entries, i, pair, id
	entries = Split(TransformerStates(), ",")

	For i = 0 To UBound(entries)

		If entries(i) <> "" Then

			pair = Split(entries(i), ":")
			id   = CInt(pair(0))

			If id <> triggerId And pair(1) = "1" Then

				outCount = outCount + 1
				If outCount = 1 Then FieldImpediment = id

			End If

		End If

	Next

End Function

' One gate, read by the name built for it.
'
' E3 gives no way to index an XObject properties, so the read is late
' bound - the same Execute and scratch cell the configuration import uses
' to write one.
'
' A property that is not there leaves the default and adds its name to
' missing. Reported rather than assumed either way round: defaulting a
' precondition to True would let a maneuver run on a gate nobody wrote,
' and defaulting it to False would bar every maneuver with nothing said
' about why.
Function ReadGate(propertyName, defaultValue, ByRef missing)

	gGateValue = defaultValue

	Dim failed
	failed = False

	On Error Resume Next
	Err.Clear
	Execute "gGateValue = xatm_TA." & propertyName
	If Err.Number <> 0 Then failed = True
	Err.Clear
	On Error Goto 0

	If failed Then

		If missing <> "" Then missing = missing & ", "
		missing = missing & propertyName
		gGateValue = defaultValue

	End If

	ReadGate = CBool(gGateValue)

End Function

' Scratch cell for the late-bound read, the way the configuration import
' keeps one for its writes.
Dim gGateValue


' Logs the rejection and clears the trigger without re-firing.
' Uses the explicit tag path (not Me) so it is safe to call from a helper.
Sub Reject(reason, ts)

	WriteLog "Not started - " & reason
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

<xatm_TA.FSM.Main:Main_Completed()>
Sub Main_Completed()
		
	Parent.Item("TriggerTransformerId").WriteEx  Empty, 0
	Parent.Item("AutomationType").WriteEx Empty, 0
	Parent.Item("StepTimer").WriteEx  Empty, 0
	WriteEx  Empty, 0
	
    xatm_TA.Running = False
    WriteLog "Automation completed successfully."
            	
End Sub

<xatm_TA.FSM.Main:Main_Functions()>
Sub Main_Functions()
End Sub


Function GetDeviceById(id, ByRef exists)
    
    exists = False
    Set GetDeviceById = FindInFolder(Application.GetObject("XATM_Data.Substation"), id, exists)
    
End Function

Function FindInFolder(folder, id, ByRef exists)
    
    Set FindInFolder = Nothing
        
    Dim obj

    For Each obj In folder
        
        Select Case UCase(TypeName(obj))

            Case "XATM_BREAKER", "XATM_TRANSFORMER", "XATM_DISCONNECTOR"
                If obj.Id = id Then
                    Set FindInFolder = obj
                    exists = True
                    Exit Function
                End If

            Case Else
                ' Not a device -> treat as a subfolder and recurse.
                Dim found
                On Error Resume Next
                Set found = FindInFolder(obj, id, exists)
                On Error GoTo 0

                If exists Then
                    Set FindInFolder = found
                    Exit Function
                End If

        End Select
        
    Next

End Function

Function GetLayoutType()

	GetLayoutType = Application.GetObject("XATM_Data.Automation.Layout.Transformer").Value & "_" & _
					Application.GetObject("XATM_Data.Automation.Layout.Busbar").Value
	
End Function

Sub ResetTimer()

    Parent.Item("StepTimer").Value = 0
    
End Sub

Sub IncrementTimer()

    Parent.Item("StepTimer").Value = Parent.Item("StepTimer").Value + 1
    
End Sub

Function OperationName(operation)
	If operation = 2 Then
		OperationName = "closed"
	ElseIf operation = 1 Then
		OperationName = "opened"
	Else
		OperationName = "operated"
	End If
End Function


Function ReadImpededId()

	Dim id
	id = 0

	On Error Resume Next
	id = Parent.Item("ImpededTransformerId").Value
	On Error Goto 0

	If IsEmpty(id) Or IsNull(id) Then id = 0
	If Not IsNumeric(id) Then id = 0

	ReadImpededId = CInt(id)

End Function


' Whether step 1 sends the open command, or only confirms the breaker.
'
' True for a class that has not been given the property, which is what
' every station did before it existed - and the forgiving direction
' either way. Step 1 leaves early when the breaker is already open, so a
' redundant order only ever reaches one still closed; a wait nobody
' satisfies runs the step to its timeout instead.
Function SendsOpenCommand()

	SendsOpenCommand = True

	On Error Resume Next
	SendsOpenCommand = CBool(xatm_TA.SendLowVoltageBreakerOpenCommand)
	On Error Goto 0

End Function


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

<xatm_TA.FSM.Main:Main_GlobalLockout()>
Sub Main_GlobalLockout()
	
	xatm_TA.Running 		= False
	xatm_TA.GeneralBlock 	= True

	Select Case Value
		Case 1 : xatm_TA.StepExecutionFailed1 = True
		Case 2 : xatm_TA.StepExecutionFailed2 = True
		Case 3 : xatm_TA.StepExecutionFailed3 = True
		Case 4 : xatm_TA.StepExecutionFailed4 = True
		Case 5 : xatm_TA.StepExecutionFailed5 = True
		Case 6 : xatm_TA.StepExecutionFailed6 = True
	End Select
	
	WriteLog "Global lockout activated due to automation failure."
	
End Sub

<xatm_TA.FSM.Main:Main_Main()>
Sub Main_Main()
	
	Const STEP_1_TIMER = 30

	If Not xatm_TA.Enabled Then
		
		WriteLog "Automation not enabled."
		xatm_TA.Running = False
		Exit Sub
	
	End If
	
	If Not xatm_TA.Running Then
		Exit Sub
	End If
	
	Select Case Value
		
		Case 0
			
			Main_Step00()
		
		Case 1
			
			If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
				
				Main_Step01()
				
			Else
				
				WriteLog "Step 1: Execution failed - Timeout exceeded."
				Main_GlobalLockout
				Exit Sub
				
			End If

        Case 2

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step02()
                
            Else
                
                WriteLog "Step 2: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
				
            End If
        
        Case 3

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step03()
                
            Else
                
                WriteLog "Step 3: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
                
            End If

        Case 4
            
            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step04()
                
            Else
                
                WriteLog "Step 4: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
				
            End If
        
        Case 5
            
            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step05()
                
            Else
                
                WriteLog "Step 5: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
				
            End If

        Case 6

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step06()
                
            Else
                
                WriteLog "Step 6: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
				
            End If

        Case 99
			
			Main_Completed()
			
	End Select
	
	IncrementTimer()
	
End Sub

<xatm_TA.FSM.Main:Main_Step00()>
Sub Main_Step00()

	WriteLog "Starting " & DescribeAutomation()

	Value = 1

End Sub

' Human-readable summary of the running automation for the step log,
' using transformer names, e.g. "TA TR-01 - TR-02 out of service".
'
' The mode is not read off the FSM the way the manual class reads it. There
' is only one thing this automation runs, and a line that fetched the
' answer it already knows would only be able to disagree with itself.
Function DescribeAutomation()

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value

	Dim impedeId
	impedeId  = ReadImpededId()

	If impedeId = 0 Then
		DescribeAutomation = "TA " & TransformerName(triggerId) & " - no transformer out of service"
	Else
		DescribeAutomation = "TA " & TransformerName(triggerId) & " - " & TransformerName(impedeId) & " out of service"
	End If

End Function

' Transformer name for an id, falling back to "ID <n>" when not found.
Function TransformerName(id)

	Dim transformer, exists
	Set transformer = GetDeviceById(id, exists)

	If exists Then
		TransformerName = transformer.Name
	Else
		TransformerName = "ID " & id
	End If

End Function

Sub Foo()
	
End Sub

<xatm_TA.FSM.Main:Main_Step01()>
Sub Main_Step01()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()

	S1TA triggerId, impedeId

End Sub

Sub S1TA(triggerId, impedeId)

	Dim breakerId, action
	action = 1

	' TA always starts by opening the tripped transformer's own secondary
	' breaker (triggerId + 20) to confirm isolation - same for every layout.
	' TODO: (Spec 1.4.3.16 mistakenly writes CB3 for the TR4/TR3-out case; 
	' opening the trigger's own CB is correct and makes that a non-issue.)
	breakerId = triggerId + 20

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 1: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 1: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 2
		Exit Sub

	End If
	
	' Opened here, or only confirmed here.
	'
	' Where the lockout relay trips the secondary breaker itself there is
	' nothing for this step to command - only something to confirm, and an
	' order sent into a breaker the protection is already opening is a
	' second order for one movement.
	'
	' Where it is not wired that way, this step is what opens it. The step
	' timer bounds the wait either way: a breaker that never opens is a
	' failed step and not a sequence that sits there.
	If Not SendsOpenCommand() Then

		' Said once, and only once the wait is long enough to look like a
		' stall. The protection normally has the breaker open before the
		' first tick, so there is usually nothing to say at all - and a line
		' every second until the step times out would bury the one that
		' matters. Three, the same figure the reclosing waits before it says
		' what it is waiting for.
		If Parent.Item("StepTimer").Value = 3 Then
			WriteLog "Step 1: waiting for " & breaker.Name & " to be opened by the protection."
		End If

		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select
	
End Sub

<xatm_TA.FSM.Main:Main_Step02()>
Sub Main_Step02()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()

	S2TA triggerId, impedeId

End Sub

Sub S2TA(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action, nextStep
	nextStep = 3

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1                      ' default: open a tie to isolate the dead busbar

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					breakerId = 710
				Else
					breakerId = 700
				End If

			Case 200
				If impedeId = 100 Then
					breakerId = 710
				Else
					breakerId = 720
				End If

			Case 300
				If impedeId = 100 Then
					breakerId = 730
				ElseIf impedeId = 200 Then
					breakerId = 900
				ElseIf impedeId = 400 Then
					breakerId = 740
				Else
					breakerId = 740        ' no contingency: 2-step restore
					action = 2
					nextStep = 99
				End If

			Case 400
				If impedeId = 100 Then
					breakerId = 730
				ElseIf impedeId = 200 Then
					breakerId = 900
				ElseIf impedeId = 300 Then
					breakerId = 740
				Else
					breakerId = 740        ' no contingency: 2-step restore
					action = 2
					nextStep = 99
				End If

		End Select

	Else

		' ==============
		' DEFAULT
		' ==============
		action = 2
		breakerId = 700
		nextStep = 99

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 2: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 2: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = nextStep
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select
		
End Sub

<xatm_TA.FSM.Main:Main_Step03()>
Sub Main_Step03()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()

	S3TA triggerId, impedeId

End Sub

Sub S3TA(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action
	action = 2                          ' restore step: close a tie

	If layoutType = "4TR4LV_6BB6TIERING" Then

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					breakerId = 700
				Else
					breakerId = 900
				End If

			Case 200
				If impedeId = 100 Then
					breakerId = 700
				Else
					breakerId = 730
				End If

			Case 300
				If impedeId = 100 Then
					breakerId = 740
				ElseIf impedeId = 200 Then
					breakerId = 700
				ElseIf impedeId = 400 Then
					breakerId = 900
				End If

			Case 400
				If impedeId = 300 Then
					breakerId = 900
				Else
					breakerId = 740
				End If

		End Select

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 3: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 3: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 4
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select
	
End Sub

<xatm_TA.FSM.Main:Main_Step04()>
Sub Main_Step04()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()

	S4TA triggerId, impedeId

End Sub

Sub S4TA(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action, nextStep
	nextStep = 99

	If layoutType = "4TR4LV_6BB6TIERING" Then

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					action = 2
					breakerId = 720
				ElseIf impedeId = 300 Or impedeId = 400 Then
					action = 2
					breakerId = 710
				Else
					action = 1             ' no contingency: open, sequence continues
					breakerId = 720
					nextStep = 5
				End If

			Case 200
				If impedeId = 100 Then
					action = 2
					breakerId = 720
				ElseIf impedeId = 300 Or impedeId = 400 Then
					action = 2
					breakerId = 710
				Else
					action = 1             ' no contingency: open, sequence continues
					breakerId = 700
					nextStep = 5
				End If

			Case 300
				action = 2
				If impedeId = 100 Then
					breakerId = 720
				ElseIf impedeId = 200 Then
					breakerId = 740
				ElseIf impedeId = 400 Then
					breakerId = 730
				End If

			Case 400
				action = 2
				If impedeId = 100 Then
					breakerId = 720
				ElseIf impedeId = 200 Then
					breakerId = 700
				ElseIf impedeId = 300 Then
					breakerId = 730
				End If

		End Select

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 4: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 4: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = nextStep
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

<xatm_TA.FSM.Main:Main_Step05()>
Sub Main_Step05()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()

	S5TA triggerId, impedeId

End Sub

Sub S5TA(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action
	action = 2

	If layoutType = "4TR4LV_6BB6TIERING" Then

		Select Case triggerId

			Case 100
				breakerId = 730

			Case 200
				breakerId = 900

		End Select

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 5: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 5: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 6
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

<xatm_TA.FSM.Main:Main_Step06()>
Sub Main_Step06()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()

	S6TA triggerId, impedeId

End Sub

Sub S6TA(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action
	action = 2

	If layoutType = "4TR4LV_6BB6TIERING" Then

		Select Case triggerId

			Case 100, 200
				breakerId = 710

		End Select

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 6: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 6: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 99
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select
		
End Sub

<xatm_TMTNM.Commands.OperatorBlock:OperatorBlock_CommandOperatorBlock()>
Sub OperatorBlock_CommandOperatorBlock()
	
	Dim v
	v = False
	
	On Error Resume Next
	v = CBool(xatm_TMTNM.CommandOperatorBlock.Value)
	On Error Goto 0
	
	xatm_TMTNM.OperatorBlock = v

End Sub

<xatm_TMTNM.Commands.OperatorBlock:OperatorBlock_OnChangedValue()>
Sub OperatorBlock_OnChangedValue()
	
	If xatm_TMTNM.OperatorBlock Then
		WriteLog "Blocked by operator."
	Else
		WriteLog "Released by operator."
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

<xatm_TMTNM.Commands.Reset:Reset_OnChangedTimeStamp()>
Sub Reset_OnChangedTimeStamp()
	
	If CBool(Value) Then
		
		Reset()
		
	End If
	
End Sub

Sub Reset()
	
	xatm_TMTNM.Running 		= False
	xatm_TMTNM.GeneralBlock	= False
	xatm_TMTNM.Reverting	= False
	xatm_TMTNM.Successful	= False
	xatm_TMTNM.Unsuccessful	= False
	
	' ===============
	' RESET ALL STEP FAIL POINTS
	' ===============
	Dim i
	For i = 1 To 6
		Dim propName
		propName = "StepExecutionFailed" & i
		
		On Error Resume Next
		Execute("xatm_TMTNM." & propName & " = False")
		On Error Goto 0
	Next
	
	Dim tag
	For Each tag In xatm_TMTNM.Item("FSM")
		
		If TypeName(tag) = "InternalTag" Then

			tag.WriteEx Empty, tag.TimeStamp
		
		End If
		
	Next
	
	' ================
	' RESET ALL DEVICES
	' ================
	ResetDevices Application.GetObject("XATM_Data.Substation")
	
	WriteLog "Reset"

End Sub

Sub ResetDevices(folder)
	
	For Each obj In folder
	
		If TypeName(obj) = "xatm_Breaker" Then
		
			On Error Resume Next
			obj.Item("Data").Item("Reset").WriteEx True
			On Error Goto 0
			
		Else
		
			ResetDevices obj
		
		End If
	
	Next

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

<xatm_TMTNM.Commands.Reset:Reset_Reset()>
Sub Reset_Reset()
	
	If xatm_TMTNM.CommandReset.Value = 0 Then Exit Sub
	
	WriteEx True
		
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartNM()>
Sub Start_CommandStartNM()
	
	If xatm_TMTNM.CommandStartNM.Value = 0 Then Exit Sub
	
	StartMode "NM", 0
	
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartNM100()>
Sub Start_CommandStartNM100()

	If xatm_TMTNM.CommandStartNM100.Value = 0 Then Exit Sub

	StartMode "NM", 100
	
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartNM200()>
Sub Start_CommandStartNM200()

	If xatm_TMTNM.CommandStartNM200.Value = 0 Then Exit Sub

	StartMode "NM", 200
	
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartNM300()>
Sub Start_CommandStartNM300()

	If xatm_TMTNM.CommandStartNM300.Value = 0 Then Exit Sub

	StartMode "NM", 300
	
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartNM400()>
Sub Start_CommandStartNM400()

	If xatm_TMTNM.CommandStartNM400.Value = 0 Then Exit Sub

	StartMode "NM", 400
	
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartTM()>
Sub Start_CommandStartTM()

	If xatm_TMTNM.CommandStartTM.Value = 0 Then Exit Sub

	StartMode "TM", 0
	
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartTM100()>
Sub Start_CommandStartTM100()

	If xatm_TMTNM.CommandStartTM100.Value = 0 Then Exit Sub

	StartMode "TM", 100
	
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartTM200()>
Sub Start_CommandStartTM200()

	If xatm_TMTNM.CommandStartTM200.Value = 0 Then Exit Sub

	StartMode "TM", 200
	
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartTM300()>
Sub Start_CommandStartTM300()

	If xatm_TMTNM.CommandStartTM300.Value = 0 Then Exit Sub

	StartMode "TM", 300
	
End Sub

<xatm_TMTNM.Commands.Start:Start_CommandStartTM400()>
Sub Start_CommandStartTM400()

	If xatm_TMTNM.CommandStartTM400.Value = 0 Then Exit Sub

	StartMode "TM", 400
	
End Sub

<xatm_TMTNM.Commands.Start:Start_OnChangedValue()>
Sub Start_OnChangedValue()

	If Trim(Value) = "" Then Exit Sub      ' self-cleared write, ignore

	Dim ts
	ts = TimeStamp	' preserve for the silent clear

	Dim parts
	parts = Split(Me.Value, ":")

	' Two fields or three. The third names the transformer the maneuver is
	' to assume out of service, and is what the operator now picks; a
	' request without it is a request with nothing out.
	If UBound(parts) < 1 Or UBound(parts) > 2 Then
		Reject "Invalid Start format '" & Me.Value & "' (expected MODE:trigger[:impeded]).", ts
		Exit Sub
	End If

	Dim mode
	mode = UCase(Trim(parts(0)))

	Select Case mode

		Case "TM", "NM"

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
	' WHICH TRANSFORMER IS OUT
	' ================
	'
	' The person asking says which one, and that is the whole of the
	' triggering philosophy here. An automatic transfer reads the field
	' for it instead, because the switchyard is what asked and the
	' switchyard says nothing - which is xatm_TA's business now.
	Dim impedeId
	impedeId = 0

	If UBound(parts) = 2 Then

		If Not IsNumeric(parts(2)) Then

			Reject "Invalid impeded transformer '" & parts(2) & "'.", ts
			Exit Sub

		End If

		impedeId = CInt(parts(2))

	End If

	' ================
	' IS THE REQUEST ITSELF COHERENT
	' ================
	'
	' Ahead of the gates below, because these say the request is malformed
	' rather than that the switchyard is not ready for it. Each one is a
	' case the step logic cannot express: its Else branch would take a
	' nonsense argument for "nothing is out" and run a real maneuver.
	'
	' Two kinds of automation asking for two different things, so they are
	' checked apart. An instance carrying a BusbarPair moves that pair of
	' busbars and is bound to no transformer at all - there is no busbar
	' XObject to bind to, which is why the pair is carried as text. Every
	' other instance moves its own transformer's load, and answers the
	' questions that were always asked here.
	Dim runMode
	runMode = mode

	If BusbarPair() <> "" Then

		If impedeId <> 0 Then

			Reject "a busbar transfer takes no impediment - it has one path and it does not vary.", ts
			Exit Sub

		End If

		If GetLayoutType() <> "4TR4LV_6BB6TIERING" Then

			Reject "a busbar transfer is defined only on the 4TR4LV_6BB6TIERING layout.", ts
			Exit Sub

		End If

		runMode = BusbarMode(mode, BusbarPair())

		If runMode = "" Then

			Reject "'" & BusbarPair() & "' is not a pair of busbars that can be transferred between.", ts
			Exit Sub

		End If

	Else

		If Not IsConfiguredTransformer(triggerId) Then

			Reject "Trigger transformer " & triggerId & " is not configured in this substation.", ts
			Exit Sub

		End If

		If impedeId <> 0 Then

			If impedeId = triggerId Then

				Reject "Transformer " & triggerId & " cannot be both the trigger and the one out of service.", ts
				Exit Sub

			End If

			If Not IsConfiguredTransformer(impedeId) Then

				Reject "Impeded transformer " & impedeId & " is not configured in this substation.", ts
				Exit Sub

			End If

		End If

		' A transfer needs somewhere to put the load. Anything that is
		' neither the trigger nor out of service can take it; with none
		' left there is nothing to transfer to and no sequence to run.
		'
		' On the four transformer ring one impediment still leaves two, so
		' this only bites on the small layouts - two transformers with the
		' adjacent one impeded, which is the case the client asked to have
		' refused.
		If CountDestinations(triggerId, impedeId) = 0 Then

			Reject "No transformer is left to take the load.", ts
			Exit Sub

		End If

	End If

	' ================
	' GATE CHECKS
	' ================

	If Not xatm_TMTNM.Enabled Then

		Reject "Automation disabled.", ts
		Exit Sub

	End If

	If xatm_TMTNM.Running Then

		Reject "Already running.", ts
		Exit Sub

	End If

	If xatm_TMTNM.OperatorBlock Then

	 	Reject "Blocked by operator.", ts
		Exit Sub

	End If

	If xatm_TMTNM.GeneralBlock Then

	 	Reject "General interlock active.", ts
	 	Exit Sub

	End If

	' The switchyard withholding it and the switchyard permitting it - but
	' asked of this maneuver own pair of gates, not of one pair standing
	' for every maneuver the automation can run.
	'
	' Both are read before either is judged, so a manifest that has not
	' been rebuilt is reported as the configuration fault it is rather than
	' surfacing as a maneuver that will not start for no stated reason.
	Dim gate, missing, blocked, permitted
	gate    = GateSuffix(mode, impedeId)
	missing = ""

	blocked   = ReadGate("AutomaticBlock" & gate, False, missing)
	permitted = ReadGate("Preconditions"  & gate, True,  missing)

	If missing <> "" Then

		Reject "this automation has no " & missing & " - the manifest has not been rebuilt.", ts
		Exit Sub

	End If

	If blocked Then

		Reject "Blocked by field conditions (AutomaticBlock" & gate & ").", ts
		Exit Sub

	End If

	' This one reads the other way round - the expression on it says when
	' the maneuver may go ahead, so it is the absence of it that rejects.
	If Not permitted Then

		Reject "Preconditions are not met (Preconditions" & gate & ").", ts
		Exit Sub

	End If

	If AnyOtherAutomationRunning() Then

		Reject "Another automation is in progress.", ts
		Exit Sub

	End If

	' ================
	' START
	' ================

	' The resolved maneuver and not the command: a busbar instance was
	' asked for TM and what it runs is TMB1A, which is what every step
	' from here on reads.
	' The last run's result stands until this one starts, so the control
	' room keeps seeing how the previous attempt went.
	xatm_TMTNM.Successful   = False
	xatm_TMTNM.Unsuccessful = False
	xatm_TMTNM.Reverting    = False
	
	xatm_TMTNM.Item("FSM").Item("AutomationType").WriteEx runMode
	xatm_TMTNM.Item("FSM").Item("TriggerTransformerId").WriteEx triggerId
	xatm_TMTNM.Item("FSM").Item("ImpededTransformerId").WriteEx impedeId

	' Emptied here and not only at the end of a run. A run that stopped in
	' lockout never reached Main_Completed, and its record appended to would
	' make this run's revert undo that one's work as well.
	xatm_TMTNM.Item("FSM").Item("PerformedSteps").WriteEx ""

	xatm_TMTNM.Item("FSM").Item("StepTimer").WriteEx 0
	xatm_TMTNM.Item("FSM").Item("Main").WriteEx 0
	xatm_TMTNM.Running = True

	If runMode <> mode Then
		WriteLog "Start - " & BusbarLabel(runMode)
	ElseIf impedeId = 0 Then
		WriteLog "Start - " & mode & " TR" & triggerId
	Else
		WriteLog "Start - " & mode & " TR" & triggerId & " with TR" & impedeId & " out of service"
	End If

	WriteEx "", ts ' clear without re-firing

End Sub

' True if any OTHER automation object is currently running (mutual exclusion).
' Relies on a common Running property instead of enumerating each type.
Function AnyOtherAutomationRunning()

	AnyOtherAutomationRunning = False

	Dim obj
	For Each obj In Application.GetObject("XATM_Data.Automation")

		If Not (obj Is xatm_TMTNM) Then

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


' Asks for a maneuver by writing the very tag this runs on.
'
' Start_OnChangedValue is what reads it, so every gate it keeps - enabled,
' not already running, the three blocks, the preconditions, no other
' automation in progress - is kept for a command off a screen too, and is
' kept in the one place that has them all. A command that is refused is
' logged there and says why, exactly as a written trigger is.
'
' impedeId is which transformer the maneuver is to treat as out of service,
' or 0 for none. It is not read off the command tag: there is one tag per
' maneuver, and which tag was written is what says which maneuver was asked
' for. The value on it means nothing beyond "asked for".
'
' That is the shape Level 3 needs. Over IEC 60870-5-104 a command is a point
' address and carries no argument, so an address has to mean one maneuver by
' itself; the E3 operation screen works the same way round, a button writes
' a tag rather than composing one. A single tag carrying which transformer
' is out would have to be a setpoint, and a setpoint followed by an execute
' is two telegrams with nothing binding them together - not something to put
' under a switching sequence.
'
' There is no TA among the commands. That one is asked for by the
' switchyard, through the transformer's own triggers, and never by a
' person - which is also why TA takes no impediment argument and reads the
' field for it instead.
Sub StartMode(mode, impedeId)

	' A busbar automation is bound to nothing, so there is nothing to
	' look up and no trigger to name: which pair it moves is a property
	' of its own, and that pair is the whole of what the maneuver needs.
	If BusbarPair() <> "" Then

		WriteEx mode & ":0:0"
		Exit Sub

	End If

	Dim found, id
	id = BoundTransformerId(found)

	' Not something the operator did wrong: the automation has been left
	' unbound, and Start could only reject a maneuver on transformer 0.
	If Not found Then

		WriteLog "Not started - no transformer is bound to this automation."
		Exit Sub

	End If

	WriteEx mode & ":" & id & ":" & impedeId

End Sub

' The Id of the transformer this instance drives, which is the trigger of
' anything asked for from a screen. found stays False where nothing is
' bound, or where what is bound carries no Id yet - a half-configured
' automation rather than a maneuver to reject.
Function BoundTransformerId(ByRef found)

	found = False
	BoundTransformerId = 0

	Dim transformer
	Set transformer = Nothing

	On Error Resume Next
	Set transformer = xatm_TMTNM.Transformer
	On Error Goto 0

	If transformer Is Nothing Then Exit Function

	Dim id
	id = Empty

	On Error Resume Next
	id = transformer.Id
	On Error Goto 0

	If IsEmpty(id) Or IsNull(id) Then Exit Function
	If Not IsNumeric(id) Then Exit Function

	BoundTransformerId = CInt(id)
	found = True

End Function

' Every transformer the substation has, as "<id>:<0|1>" pairs - the flag is
' its out-of-service reading at the moment of asking.
'
' One walk, because the three questions below - is this Id configured, how
' many are out, which one is out - are all answered off the same list.
'
' FSM.Main has GetDeviceById for the same job. E3 gives no way to call a
' procedure in another object's scope, so this scope keeps its own, narrower
' version: it only ever looks for transformers.
Function TransformerStates()

	Dim acc
	acc = ""

	On Error Resume Next
	CollectTransformers Application.GetObject("XATM_Data.Substation"), acc
	On Error Goto 0

	TransformerStates = acc

End Function

Sub CollectTransformers(folder, ByRef acc)

	Dim obj
	For Each obj In folder

		Select Case UCase(TypeName(obj))

			Case "XATM_TRANSFORMER"

				Dim id
				id = Empty

				On Error Resume Next
				id = obj.Id
				On Error Goto 0

				If IsNumeric(id) Then

					Dim flag
					flag = 0

					On Error Resume Next
					If CBool(obj.OutOfService) Then flag = 1
					On Error Goto 0

					If acc <> "" Then acc = acc & ","
					acc = acc & CInt(id) & ":" & flag

				End If

			Case "XATM_BREAKER", "XATM_DISCONNECTOR"

				' a device, and not one this is looking for

			Case Else

				' Not a device -> treat as a subfolder and recurse.
				On Error Resume Next
				CollectTransformers obj, acc
				On Error Goto 0

		End Select

	Next

End Sub

' True when id names a transformer this substation actually has.
'
' Without it a typo'd screen binding, or a command tag left unconfigured,
' falls through every ElseIf in the step logic and lands on the Else that
' means "nothing is out" - running a real maneuver on a live switchyard
' under an assumption nobody made.
Function IsConfiguredTransformer(id)

	IsConfiguredTransformer = False

	Dim entries, i, pair
	entries = Split(TransformerStates(), ",")

	For i = 0 To UBound(entries)

		If entries(i) <> "" Then

			pair = Split(entries(i), ":")

			If CInt(pair(0)) = id Then
				IsConfiguredTransformer = True
				Exit Function
			End If

		End If

	Next

End Function

' How many transformers could take the load: everything that is neither the
' trigger, nor the one declared out, nor out of service in the field.
Function CountDestinations(triggerId, impedeId)

	CountDestinations = 0

	Dim entries, i, pair, id
	entries = Split(TransformerStates(), ",")

	For i = 0 To UBound(entries)

		If entries(i) <> "" Then

			pair = Split(entries(i), ":")
			id   = CInt(pair(0))

			If id <> triggerId And id <> impedeId And pair(1) = "0" Then
				CountDestinations = CountDestinations + 1
			End If

		End If

	Next

End Function

' The pair of busbars this instance moves, or "" when it moves none.
'
' Text and not a link, because a busbar is not an XObject there is anything
' to bind to - the four other properties that name equipment are links, and
' this one could not be.
'
' One property and not two. A flag reading "this is a busbar automation"
' beside a name saying which busbar would be two ways of writing one fact,
' and two ways of writing one fact can be made to disagree; a name that is
' either there or is not answers both questions and cannot.
'
' Empty is the ordinary case and the safe one: an instance nobody
' configured, or one predating the property, is a transformer automation.
Function BusbarPair()

	BusbarPair = ""

	On Error Resume Next
	BusbarPair = Trim(xatm_TMTNM.BusbarPair & "")
	On Error Goto 0

End Function


' Which of the four busbar sequences a TM or an NM on this instance is, and
' "" for a pair that names none of them.
'
' The mode says which way round - out or back - and the pair says which
' busbars, so the two together name one sequence: TM on B1A-B4A is TMB1A
' and NM on B2B-B3A is NMB2B.
Function BusbarMode(mode, pair)

	BusbarMode = ""

	Select Case UCase(pair)

		Case "B1A-B4A"
			BusbarMode = mode & "B1A"

		Case "B2B-B3A"
			BusbarMode = mode & "B2B"

	End Select

End Function


' The maneuver as the spec writes it, for anything a person reads.
Function BusbarLabel(mode)

	Select Case mode

		Case "TMB1A"
			BusbarLabel = "TM B1A-B4A"

		Case "NMB1A"
			BusbarLabel = "NM B1A-B4A"

		Case "TMB2B"
			BusbarLabel = "TM B2B-B3A"

		Case "NMB2B"
			BusbarLabel = "NM B2B-B3A"

		Case Else
			BusbarLabel = mode

	End Select

End Function


' The layout this substation is configured as, read the way the step logic
' reads it. Written out again here because one E3 object cannot call a
' procedure in another's scope.
Function GetLayoutType()

	GetLayoutType = Application.GetObject("XATM_Data.Automation.Layout.Transformer").Value & "_" & _
					Application.GetObject("XATM_Data.Automation.Layout.Busbar").Value

End Function


' Which pair of gates a maneuver answers to.
'
' The mode, and the transformer it assumes is out where there is one: TM,
' TM200, NM400. The same names the command tags carry, so the command and
' the gates that let it through are one string apart.
'
Function GateSuffix(mode, impedeId)

	If impedeId = 0 Then
		GateSuffix = mode
	Else
		GateSuffix = mode & impedeId
	End If

End Function


' One gate, read by the name built for it.
'
' E3 gives no way to index an XObject properties, so the read is late
' bound - the same Execute and scratch cell the configuration import uses
' to write one.
'
' A property that is not there leaves the default and adds its name to
' missing. Reported rather than assumed either way round: defaulting a
' precondition to True would let a maneuver run on a gate nobody wrote,
' and defaulting it to False would bar every maneuver with nothing said
' about why.
Function ReadGate(propertyName, defaultValue, ByRef missing)

	gGateValue = defaultValue

	Dim failed
	failed = False

	On Error Resume Next
	Err.Clear
	Execute "gGateValue = xatm_TMTNM." & propertyName
	If Err.Number <> 0 Then failed = True
	Err.Clear
	On Error Goto 0

	If failed Then

		If missing <> "" Then missing = missing & ", "
		missing = missing & propertyName
		gGateValue = defaultValue

	End If

	ReadGate = CBool(gGateValue)

End Function

' Scratch cell for the late-bound read, the way the configuration import
' keeps one for its writes.
Dim gGateValue


' Logs the rejection and clears the trigger without re-firing.
' Uses the explicit tag path (not Me) so it is safe to call from a helper.
Sub Reject(reason, ts)

	WriteLog "Not started - " & reason
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

<xatm_TMTNM.FSM.Main:Main_Completed()>
Sub Main_Completed()
	
	' Read before the state is torn down, because tearing it down is what
	' this does.
	Dim reverted
	reverted = IsReverting()

	Parent.Item("TriggerTransformerId").WriteEx  Empty, 0
	Parent.Item("AutomationType").WriteEx Empty, 0
	Parent.Item("PerformedSteps").WriteEx Empty, 0
	Parent.Item("StepTimer").WriteEx  Empty, 0
	WriteEx  Empty, 0

	xatm_TMTNM.Running   = False
	xatm_TMTNM.Reverting = False

	' A revert reaches the end of a sequence without the maneuver having
	' happened. Nothing is blocked and the substation is where it started,
	' so there is nothing to clear and nothing to stop - but what was asked
	' for did not take place, and that is the part the control room needs.
	If reverted Then

		xatm_TMTNM.Unsuccessful = True
		WriteLog "Reverted - the substation is back the way it was, and the maneuver did not complete."

	Else

		xatm_TMTNM.Successful = True
		WriteLog "Automation completed successfully."

	End If
            	
End Sub

<xatm_TMTNM.FSM.Main:Main_Functions()>
Sub Main_Functions()
End Sub


Function GetDeviceById(id, ByRef exists)
    
    exists = False
    Set GetDeviceById = FindInFolder(Application.GetObject("XATM_Data.Substation"), id, exists)
    
End Function

Function FindInFolder(folder, id, ByRef exists)
    
    Set FindInFolder = Nothing
        
    Dim obj

    For Each obj In folder
        
        Select Case UCase(TypeName(obj))

            Case "XATM_BREAKER", "XATM_TRANSFORMER", "XATM_DISCONNECTOR"
                If obj.Id = id Then
                    Set FindInFolder = obj
                    exists = True
                    Exit Function
                End If

            Case Else
                ' Not a device -> treat as a subfolder and recurse.
                Dim found
                On Error Resume Next
                Set found = FindInFolder(obj, id, exists)
                On Error GoTo 0

                If exists Then
                    Set FindInFolder = found
                    Exit Function
                End If

        End Select
        
    Next

End Function

Function GetLayoutType()

	GetLayoutType = Application.GetObject("XATM_Data.Automation.Layout.Transformer").Value & "_" & _
					Application.GetObject("XATM_Data.Automation.Layout.Busbar").Value
	
End Function

Sub ResetTimer()

    Parent.Item("StepTimer").Value = 0
    
End Sub

Sub IncrementTimer()

    Parent.Item("StepTimer").Value = Parent.Item("StepTimer").Value + 1
    
End Sub

Function OperationName(operation)
	If operation = 2 Then
		OperationName = "closed"
	ElseIf operation = 1 Then
		OperationName = "opened"
	Else
		OperationName = "operated"
	End If
End Function
' Which transformer the running maneuver is treating as out of service, or 0
' for none. Written once by Start_OnChangedValue - declared by the operator
' for TM and NM, read off the field for TA - and constant for the whole run,
' so no step can be handed a different contingency than the step before it.
'
' The spec pairs every maneuver with at most one impediment: Tabela 1
' situations 2-5 are the states a sequence starts from, and 6-11 - two
' transformers out - are the states they end in, with no sequence leaving
' them. So this is one Id and not a set.
' What this run has actually operated, in the order it operated it.
'
' "900:2,700:1" - the device and the action that was sent to it, oldest
' first. A string and not an array because it holds at most one entry per
' step, both fields are integers so no delimiter can collide with them, and
' a WatchWindow shows it as the record it is. The reclosing keeps its
' transformer list the same way.
Function ReadPerformed()

	ReadPerformed = ""

	On Error Resume Next
	ReadPerformed = Trim(Parent.Item("PerformedSteps").Value & "")
	On Error Goto 0

End Function


' Writes a command down as it is issued, so that a revert can put back
' exactly what was moved and nothing else.
'
' A step that finds its device already in position writes nothing, and that
' is the whole point: a device that was already where the step wanted it is
' one this run did not touch, whoever put it there and for whatever reason.
' Driving it anywhere on the way back would be undoing somebody else's work.
'
' Recorded on issue rather than on confirmation. A breaker that moved while
' its position feedback failed still has to be put back, and the undo is
' position-guarded like every other step, so an entry for a device that
' never actually moved costs one tick and nothing else.
Sub RecordPerformed(id, action)

	Dim list
	list = ReadPerformed()

	' Not the same device twice running. A step issues on CommandInProgress
	' 0 or 3, so a command that completed without the position check seeing
	' it would come back through here. Undoing twice is harmless, but the
	' record should say what happened rather than what happened twice.
	If LastPerformedId(list) = CLng(id) Then Exit Sub

	If list <> "" Then list = list & ","
	list = list & id & ":" & action

	Parent.Item("PerformedSteps").WriteEx list

End Sub


' The device of the last entry, and 0 for an empty record.
Function LastPerformedId(list)

	LastPerformedId = 0

	If list = "" Then Exit Function

	Dim entries, last
	entries = Split(list, ",")
	last    = Split(entries(UBound(entries)), ":")

	If IsNumeric(last(0)) Then LastPerformedId = CLng(last(0))

End Function


' Drops the entry the revert has just finished with.
Sub DropLastPerformed(entries)

	Dim keep, i
	keep = ""

	For i = 0 To UBound(entries) - 1
		If keep <> "" Then keep = keep & ","
		keep = keep & entries(i)
	Next

	Parent.Item("PerformedSteps").WriteEx keep

End Sub


' One device per pass, walking back through what the run actually did.
'
' Stays at 98 until the record is empty, so the FSM value itself says a
' revert is under way - which is what keeps a failure here from starting
' another one, without a flag to read or trust.
'
' Nothing about the maneuver is consulted: not the mode, not the trigger,
' not the contingency. The record is the whole of what there is to undo,
' which is why this one procedure serves a transfer, a normalisation, a
' busbar pair and anything added later.
Sub Main_Revert()

	Dim list
	list = ReadPerformed()

	If list = "" Then

		WriteLog "Revert: everything this run operated is back where it was."
		Value = 99
		Exit Sub

	End If

	Dim entries, entry
	entries = Split(list, ",")
	entry   = Split(entries(UBound(entries)), ":")

	Dim id, performed, target
	id        = CLng(entry(0))
	performed = CLng(entry(1))
	target    = IIf(performed = 2, 1, 2)

	Dim device, exists
	Set device = GetDeviceById(id, exists)

	If Not exists Then

		WriteLog "Revert: device with ID=" & id & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If device.Item("Data").Item("Position").Value = target Then

		WriteLog "Revert: " & device.Name & " " & OperationName(target) & " - " & _
		         UBound(entries) & " left to put back."

		DropLastPerformed entries
		ResetTimer()
		Exit Sub

	End If

	Select Case device.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command
			device.Item("Data").Item("CommandOpenClose").WriteEx target

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

Function ReadImpededId()

	Dim id
	id = 0

	On Error Resume Next
	id = Parent.Item("ImpededTransformerId").Value
	On Error Goto 0

	If IsEmpty(id) Or IsNull(id) Then id = 0
	If Not IsNumeric(id) Then id = 0

	ReadImpededId = CInt(id)

End Function

' The maneuver as the spec writes it, for anything a person reads. Its own
' copy, because one E3 object cannot call a procedure in another's scope.
Function BusbarLabel(mode)

	Select Case mode

		Case "TMB1A"
			BusbarLabel = "TM B1A-B4A"

		Case "NMB1A"
			BusbarLabel = "NM B1A-B4A"

		Case "TMB2B"
			BusbarLabel = "TM B2B-B3A"

		Case "NMB2B"
			BusbarLabel = "NM B2B-B3A"

		Case Else
			BusbarLabel = mode

	End Select

End Function


' Whether the running maneuver is one of the two busbar ones.
Function IsBusbarMode(mode)

	Select Case mode

		Case "TMB1A", "NMB1A", "TMB2B", "NMB2B"
			IsBusbarMode = True

		Case Else
			IsBusbarMode = False

	End Select

End Function


' The two ties a busbar maneuver works, in the order it works them: which
' = 1 is the one to close, which = 2 the one to open.
'
'   TM B1A-B4A   close DJ10 (900), open DJ20 (700)   B1A leaves TR1 for TR4
'   NM B1A-B4A   close DJ20 (700), open DJ10 (900)   and comes back
'   TM B2B-B3A   close DJ50 (730), open DJ40 (720)   B2B leaves TR2 for TR3
'   NM B2B-B3A   close DJ40 (720), open DJ50 (730)   and comes back
'
' Make before break both ways round, and that is worth saying because the
' spec's normalisation tables read the other way at a glance: they carry an
' upward arrow and are read from the bottom row up, so what looks like an
' open followed by a close is a close followed by an open. The busbar is
' live throughout either maneuver and stays live - it is the pair of ties
' that swaps, and all that separates a transfer from its normalisation is
' which of the two closes first.
'
' 0 for a mode that names no tie, which the step reports rather than acting
' on: a busbar maneuver that reached a step with nothing to operate is a
' maneuver this table does not know, not a breaker that failed.
Function BusbarTie(mode, which)

	BusbarTie = 0

	Select Case mode

		Case "TMB1A"
			If which = 1 Then BusbarTie = 900 Else BusbarTie = 700

		Case "NMB1A"
			If which = 1 Then BusbarTie = 700 Else BusbarTie = 900

		Case "TMB2B"
			If which = 1 Then BusbarTie = 730 Else BusbarTie = 720

		Case "NMB2B"
			If which = 1 Then BusbarTie = 720 Else BusbarTie = 730

	End Select

End Function


' One step of a busbar maneuver: operate the tie this step is for, and go
' on once it has moved.
'
' Both steps in one procedure because both are the same procedure - one
' breaker taken to one position - where each transformer maneuver picks its
' path from the trigger and the contingency and so needs a step of its own.
' Step 1 closes and hands over to step 2; step 2 opens and the sequence is
' done, which is why nothing beyond step 2 ever sees these modes.
Sub SBB(stepNumber, mode)

	Dim action, nextStep

	If stepNumber = 1 Then
		action   = 2
		nextStep = 2
	Else
		action   = 1
		nextStep = 99
	End If

	Dim breakerId
	breakerId = BusbarTie(mode, stepNumber)

	If breakerId = 0 Then

		WriteLog "Step " & stepNumber & ": " & mode & " names no tie breaker for this step. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step " & stepNumber & ": Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step " & stepNumber & ": " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = nextStep
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can  put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

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

<xatm_TMTNM.FSM.Main:Main_GlobalLockout()>
Sub Main_GlobalLockout()

	' Where every failure lands, and so where the choice between unwinding
	' and stopping is made. A step body that cannot go on calls this and
	' says nothing about what should happen next - the timeout in Main_Main
	' does the same, and so does a breaker the configuration does not have.

	' Latched first, and whichever way this goes. Which step failed is the
	' same fact either way, and once a revert has driven the substation back
	' to normal it is the only trace left of what went wrong.
	Select Case Value
		Case 1 : xatm_TMTNM.StepExecutionFailed1 = True
		Case 2 : xatm_TMTNM.StepExecutionFailed2 = True
		Case 3 : xatm_TMTNM.StepExecutionFailed3 = True
		Case 4 : xatm_TMTNM.StepExecutionFailed4 = True
		Case 5 : xatm_TMTNM.StepExecutionFailed5 = True
		Case 6 : xatm_TMTNM.StepExecutionFailed6 = True
	End Select

	If ShouldRevert() Then

		StartRevert
		Exit Sub

	End If

	xatm_TMTNM.Running      = False
	xatm_TMTNM.GeneralBlock = True
	xatm_TMTNM.Unsuccessful = True

	WriteLog "Global lockout activated due to automation failure."

End Sub


' Whether this failure unwinds the maneuver instead of stopping on it.
'
' Four things have to hold, and the last three are what keep it from being
' a loop, a contradiction or a gesture:
'
'   - the instance is configured for it;
'   - the run is not already a revert, which the FSM value says by itself,
'     or a revert that failed would start another one;
'   - something was actually operated. A run that skipped every step it
'     reached has nothing to put back, and saying so is a better account
'     than a sequence of no-ops.
Function ShouldRevert()

	ShouldRevert = False

	If Not RevertsOnFailure() Then Exit Function
	If Value = 98 Then Exit Function
	If ReadPerformed() = "" Then Exit Function

	ShouldRevert = True

End Function


' Hands the run to 98, which walks the record of what it did back to front.
'
' Running stays True and GeneralBlock stays clear: the maneuver is not over
' and nothing is barred yet.
'
' Nothing about the failed maneuver is disturbed. The mode, the trigger and
' the contingency stay where they are - the revert has no use for them, and
' the log still wants to be able to name what it was that failed.
Sub StartRevert()

	WriteLog "Step " & Value & " failed - reverting, putting back what this run operated."

	xatm_TMTNM.Reverting = True

	' Not a maneuver of its own: 98 runs the record back and consults
	' nothing about the one that failed. The mode, the trigger and the
	' contingency are all left where they are, for the log to name.
	Parent.Item("StepTimer").WriteEx 0

	Value = 98

End Sub


' The two questions asked of the class, both failing to the answer that
' changes nothing: an instance whose class predates these properties keeps
' stopping on a failure, the way it always did.
Function RevertsOnFailure()

	RevertsOnFailure = False

	On Error Resume Next
	RevertsOnFailure = CBool(xatm_TMTNM.RevertOnFailure)
	On Error Goto 0

End Function

Function IsReverting()

	IsReverting = False

	On Error Resume Next
	IsReverting = CBool(xatm_TMTNM.Reverting)
	On Error Goto 0

End Function

Sub EndOfScope()

End Sub

<xatm_TMTNM.FSM.Main:Main_Main()>
Sub Main_Main()
	
	Const STEP_1_TIMER = 30

	If Not xatm_TMTNM.Enabled Then
		
		WriteLog "Automation not enabled."
		xatm_TMTNM.Running = False
		Exit Sub
	
	End If
	
	If Not xatm_TMTNM.Running Then
		Exit Sub
	End If
	
	Select Case Value
		
		Case 0
			
			Main_Step00()
		
		Case 1
			
			If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
				
				Main_Step01()
				
			Else
				
				WriteLog "Step 1: Execution failed - Timeout exceeded."
				Main_GlobalLockout
				Exit Sub
				
			End If

        Case 2

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step02()
                
            Else
                
                WriteLog "Step 2: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
				
            End If
        
        Case 3

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step03()
                
            Else
                
                WriteLog "Step 3: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
                
            End If

        Case 4
            
            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step04()
                
            Else
                
                WriteLog "Step 4: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
				
            End If
        
        Case 5
            
            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step05()
                
            Else
                
                WriteLog "Step 5: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
				
            End If

        Case 6

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Step06()
                
            Else
                
                WriteLog "Step 6: Execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
				
            End If
		
		Case 98

            ' The revert stays here, one device per pass, until the
            ' record is empty. Each device gets its own window: the
            ' timer is reset as each one is put back.
            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then
                
                Main_Revert()
                
            Else
                
                WriteLog "Revert: execution failed - Timeout exceeded."
                Main_GlobalLockout
				Exit Sub
				
            End If

        Case 99
			
			Main_Completed()
			
	End Select
	
	IncrementTimer()

End Sub

<xatm_TMTNM.FSM.Main:Main_Step00()>
Sub Main_Step00()

	WriteLog "Starting " & DescribeAutomation()

	Value = 1

End Sub

' Human-readable summary of the running automation for the step log,
' using transformer names, e.g. "TM TR-01 - TR-02 out of service".
Function DescribeAutomation()

	Dim autoType
	autoType  = Parent.Item("AutomationType").Value

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value

	Dim impedeId
	impedeId  = ReadImpededId()

	' A busbar maneuver is not about the transformer and has no
	' contingency, so it names itself and stops there.
	If IsBusbarMode(autoType) Then
		DescribeAutomation = BusbarLabel(autoType)
	ElseIf impedeId = 0 Then
		DescribeAutomation = autoType & " " & TransformerName(triggerId) & " - no transformer out of service"
	Else
		DescribeAutomation = autoType & " " & TransformerName(triggerId) & " - " & TransformerName(impedeId) & " out of service"
	End If

End Function

' Transformer name for an id, falling back to "ID <n>" when not found.
Function TransformerName(id)

	Dim transformer, exists
	Set transformer = GetDeviceById(id, exists)

	If exists Then
		TransformerName = transformer.Name
	Else
		TransformerName = "ID " & id
	End If

End Function

' E3 will not take a Function as the last thing in a scope.
Sub EndOfScope()

End Sub

<xatm_TMTNM.FSM.Main:Main_Step01()>
Sub Main_Step01()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S1TM triggerId, impedeId
		
		Case "NM"

			S1NM triggerId, impedeId
		
		' The busbar maneuvers take neither argument: which busbar and
		' which way round is the whole of what they need, and the mode
		' carries both.
		Case "TMB1A", "NMB1A", "TMB2B", "NMB2B"

			SBB 1, Parent.Item("AutomationType").Value
			
	End Select
	
End Sub

Sub S1NM(triggerId, impedeId)

	Dim breakerId, action
	action = 2

	' NM always starts by re-closing the trigger transformer's own breaker.
	breakerId = triggerId + 20

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 1: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 1: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 2
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

Sub S1TM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					breakerId = 700
				Else
					breakerId = 900
				End If

			Case 200
				If impedeId = 100 Then
					breakerId = 700
				ElseIf impedeId = 300 Or impedeId = 400 Then
					breakerId = 730
				Else
					breakerId = 900
				End If

			Case 300
				If impedeId = 100 Then
					breakerId = 720
				ElseIf impedeId = 200 Then
					breakerId = 700
				ElseIf impedeId = 400 Then
					breakerId = 900
				Else
					breakerId = 740
				End If

			Case 400
				If impedeId = 100 Then
					breakerId = 720
				ElseIf impedeId = 200 Then
					breakerId = 700
				ElseIf impedeId = 300 Then
					breakerId = 730
				Else
					breakerId = 740
				End If

		End Select

	Else

		' ================
		' DEFAULT
		' ================
		action = 2
	 	breakerId = 700

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 1: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 1: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 2
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert canput 
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select
	
End Sub

<xatm_TMTNM.FSM.Main:Main_Step02()>
Sub Main_Step02()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S2TM triggerId, impedeId
		
		Case "NM"

			S2NM triggerId, impedeId
		
		' The busbar maneuvers take neither argument: which busbar and
		' which way round is the whole of what they need, and the mode
		' carries both.
		Case "TMB1A", "NMB1A", "TMB2B", "NMB2B"

			SBB 2, Parent.Item("AutomationType").Value
			
	End Select
	
End Sub

Sub S2NM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action, nextStep
	nextStep = 3

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					breakerId = 700
				Else
					breakerId = 710
				End If

			Case 200
				If impedeId = 100 Then
					breakerId = 720
				Else
					breakerId = 710
				End If

			Case 300
				If impedeId = 400 Then
					breakerId = 730
				ElseIf impedeId = 100 Or impedeId = 200 Then
					breakerId = 740
				Else
					breakerId = 740
					nextStep = 99    ' no contingency: sequence ends here
				End If

			Case 400
				If impedeId = 300 Then
					breakerId = 900
				ElseIf impedeId = 100 Or impedeId = 200 Then
					breakerId = 740
				Else
					breakerId = 740
					nextStep = 99    ' no contingency: sequence ends here
				End If

		End Select

	Else

		' ==============
		' DEFAULT
		' ==============
		action = 1
		breakerId = 700
		nextStep = 99

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 2: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 2: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = nextStep
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert canput 
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

Sub S2TM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		nextStep = 3

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					breakerId = 120
				Else
					breakerId = 700
				End If

			Case 200
				If impedeId = 100 Then
					breakerId = 710
				ElseIf impedeId = 300 Or impedeId = 400 Then
					breakerId = 720
				Else
					breakerId = 700
				End If

			Case 300
				If impedeId = 100 Then
					breakerId = 730
				ElseIf impedeId = 200 Then
					breakerId = 900
				ElseIf impedeId = 400 Then
					breakerId = 740
				Else
					breakerId = 320
					nextStep = 99    ' no contingency: sequence ends here
				End If

			Case 400
				If impedeId = 100 Then
					breakerId = 730
				ElseIf impedeId = 200 Then
					breakerId = 900
				ElseIf impedeId = 300 Then
					breakerId = 740
				Else
					breakerId = 420
					nextStep = 99    ' no contingency: sequence ends here
				End If

		End Select

	Else

		' ==============
		' DEFAULT
		' ==============
		action = 1
		breakerId = triggerId + 20
		nextStep = 99

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 2: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 2: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = nextStep
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

<xatm_TMTNM.FSM.Main:Main_Step03()>
Sub Main_Step03()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S3TM triggerId, impedeId
		
		Case "NM"

			S3NM triggerId, impedeId
			
	End Select
	
End Sub

Sub S3NM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					breakerId = 710
				ElseIf impedeId = 300 Or impedeId = 400 Then
					breakerId = 700
				Else
					breakerId = 720
				End If

			Case 200
				If impedeId = 100 Then
					breakerId = 710
				Else
					breakerId = 720
				End If

			Case 300
				If impedeId = 100 Then
					breakerId = 730
				ElseIf impedeId = 200 Then
					breakerId = 900
				ElseIf impedeId = 400 Then
					breakerId = 740
				End If

			Case 400
				If impedeId = 100 Then
					breakerId = 730
				ElseIf impedeId = 200 Then
					breakerId = 900
				ElseIf impedeId = 300 Then
					breakerId = 740
				End If

		End Select

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 3: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 3: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 4
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

Sub S3TM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					breakerId = 720
				ElseIf impedeId = 300 Or impedeId = 400 Then
					breakerId = 710
				Else
					breakerId = 730
				End If

			Case 200
				If impedeId = 100 Then
					breakerId = 720
				ElseIf impedeId = 300 Or impedeId = 400 Then
					breakerId = 710
				Else
					breakerId = 730
				End If

			Case 300
				If impedeId = 100 Or impedeId = 200 Then
					breakerId = 740
				ElseIf impedeId = 400 Then
					breakerId = 730
				End If

			Case 400
				If impedeId = 100 Or impedeId = 200 Then
					breakerId = 740
				ElseIf impedeId = 300 Then
					breakerId = 900
				End If

		End Select

	Else

		action = 2

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 3: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 3: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 4
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select
	
End Sub

<xatm_TMTNM.FSM.Main:Main_Step04()>
Sub Main_Step04()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S4TM triggerId, impedeId
		
		Case "NM"

			S4NM triggerId, impedeId
			
	End Select
	
End Sub

Sub S4NM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action, nextStep
	nextStep = 99

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					breakerId = 720
				ElseIf impedeId = 300 Or impedeId = 400 Then
					breakerId = 900
				Else
					breakerId = 730
					nextStep = 5    ' sequence continues
				End If

			Case 200
				If impedeId = 100 Then
					breakerId = 700
				ElseIf impedeId = 300 Or impedeId = 400 Then
					breakerId = 730
				Else
					breakerId = 730
					nextStep = 5    ' sequence continues
				End If

			Case 300
				If impedeId = 100 Then
					breakerId = 720
				ElseIf impedeId = 200 Then
					breakerId = 700
				ElseIf impedeId = 400 Then
					breakerId = 900
				End If

			Case 400
				If impedeId = 100 Then
					breakerId = 720
				ElseIf impedeId = 200 Then
					breakerId = 700
				ElseIf impedeId = 300 Then
					breakerId = 730
				End If

		End Select

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 4: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 4: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = nextStep
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

Sub S4TM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	Dim nextStep
	nextStep = 99

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Select Case triggerId

			Case 100
				If impedeId = 200 Then
					breakerId = 710
				ElseIf impedeId = 300 Or impedeId = 400 Then
					breakerId = 120
				Else
					breakerId = 720
					nextStep = 5    ' sequence continues
				End If

			Case 200
				If impedeId = 100 Or impedeId = 300 Or impedeId = 400 Then
					breakerId = 220
				Else
					breakerId = 720
					nextStep = 5    ' sequence continues
				End If

			Case 300
				breakerId = 320

			Case 400
				breakerId = 420

		End Select

	Else

		action = 1

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 4: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 4: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = nextStep
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

<xatm_TMTNM.FSM.Main:Main_Step05()>
Sub Main_Step05()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S5TM triggerId, impedeId
		
		Case "NM"

			S5NM triggerId, impedeId
			
	End Select
	
End Sub

Sub S5NM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Select Case triggerId

			Case 100, 200
				breakerId = 700

		End Select

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 5: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 5: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 6
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

Sub S5TM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Select Case triggerId

			Case 100, 200

				breakerId = 710

		End Select

	Else

		action = 2

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 5: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 5: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 6
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

<xatm_TMTNM.FSM.Main:Main_Step06()>
Sub Main_Step06()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S6TM triggerId, impedeId
		
		Case "NM"

			S6NM triggerId, impedeId
			
	End Select
	
End Sub

Sub S6NM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Select Case triggerId

			Case 100, 200
				breakerId = 900

		End Select

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 6: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 6: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 99
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

Sub S6TM(triggerId, impedeId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Select Case triggerId

			Case 100
				breakerId = 120

			Case 200
				breakerId = 220

		End Select

	Else

		action = 1

	End If

	Dim breaker, breakerExists
	Set breaker = GetDeviceById(breakerId, breakerExists)

	If Not breakerExists Then

		WriteLog "Step 6: Circuit breaker with ID=" & breakerId & " was not found. Please check the configuration - Global lockout"
		Main_GlobalLockout()
		Exit Sub

	End If

	If breaker.Item("Data").Item("Position").Value = action Then

		WriteLog "Step 6: " & breaker.Name & " " & OperationName(action) & " - Proceeding to the next step."
		ResetTimer()
		Value = 99
		Exit Sub

	End If

	Select Case breaker.Item("Data").Item("CommandInProgress").Value

		Case 0, 3

			' Idle - issue the command, and write it down so a revert can put
			' back what this run moved and nothing else.
			RecordPerformed breaker.Id, action
			breaker.Item("Data").Item("CommandOpenClose").WriteEx action

		Case 1

			' Command execution failed
			Main_GlobalLockout()
			Exit Sub

		Case 2

			' Command in progress
			Exit Sub

	End Select

End Sub

<xatm_Transformer.Data.Timers.UndervoltageRelay:UndervoltageRelay_Counter()>
Sub UndervoltageRelay_Counter()
	
	If Value > 0 Then
		
		If Value <= 3 Then
			
			WriteLog "Undervoltage trigger is about to be activated in " & Value & "..."
			
		End If
			
		Value = Value - 1
		Exit Sub
	
	End If
	
	
	
	' ================
	' TRIGGER TMTNM
	' ================
	
	WriteLog "Undervoltage relay counter has reached zero - requesting TA."
	
	Parent.Parent.Item("UndervoltageTrip").WriteEx True
	
	Value = -1

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

<xatm_Transformer.Data.Timers.UndervoltageRelay:UndervoltageRelay_OnChangedValue()>
Sub UndervoltageRelay_OnChangedValue()

	If xatm_Transformer.UndervoltageRelay Then
		
		Dim delay
		
		If xatm_Transformer.UndervoltageDelay > 0 Then
			delay = xatm_Transformer.UndervoltageDelay
		Else
			delay = 25 ' default value
		End If
		
		WriteLog "Undervoltage timer initialized."
		Value = delay
	
	Else
		
		If Value <> -1 Then
			WriteLog "Undervoltage timer reset."
			Value = -1
		End If
		
	End If
	
End Sub

<xatm_Transformer.Data.Triggers.RASEAT:RASEAT_Functions()>
Sub RASEAT_Functions()
End Sub


' Asks the high-voltage reclosing to run, because this transformer's busbar
' relay operated.
'
' Found rather than configured. There is one reclosing automation in a
' station and it holds no references to equipment, so the transformer looks
' for it by class - the same way RequestTMTNM looks for the transfer bound to
' this transformer, and for the same reason: E3 gives no way to call across
' objects, so the only way to ask for work is to write to a tag on the
' object that does it.
'
' A station whose incomer layout declares no entry bays has no reclosing to
' find, and says so once rather than failing quietly.
Sub RequestRASEAT()

	Dim obj, bound
	Set bound = Nothing

	For Each obj In Application.GetObject("XATM_Data.Automation")

		If TypeName(obj) = "xatm_RASEAT" Then
			Set bound = obj
			Exit For
		End If

	Next

	If bound Is Nothing Then
		WriteLog "No RASEAT in this substation - reclosing request ignored."
		Exit Sub
	End If

	WriteLog "CR operated - reclosing requested via " & bound.Name & "."

	' The Id travels so the reclosing can name who asked. It reads the CR
	' of every transformer for itself: more than one can operate together,
	' and the sequence has to wait for all of them to isolate.
	bound.Item("Commands").Item("Start").WriteEx xatm_Transformer.Id

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

<xatm_Transformer.Data.Triggers.RASEAT:RASEAT_OnCRTrip()>
Sub RASEAT_OnCRTrip()
	
	RequestRASEAT()
	
End Sub

<xatm_Transformer.Data.Triggers.TA:TA_Functions()>
Sub TA_Functions()
End Sub

Sub RequestTA(mode)

	Dim obj, bound, tr
	Set bound = Nothing

	' The automatic transfer bound to this transformer, and nothing else.
	'
	' Only xatm_TA is looked at now. The manual automations share this
	' folder and some of them are bound to a transformer too - a busbar
	' one is bound to whichever transformer feeds the busbar it moves -
	' but none of them has a TA in it, and the class is what says so.
	' That used to be a flag this loop had to read.
	For Each obj In Application.GetObject("XATM_Data.Automation")
		If TypeName(obj) = "xatm_TA" Then

			Set tr = Nothing
			On Error Resume Next
			Set tr = obj.Transformer          ' may be unbound
			On Error GoTo 0

			If Not tr Is Nothing Then
				If tr.Id = xatm_Transformer.Id Then
					Set bound = obj
					Exit For
				End If
			End If

		End If
	Next

	If bound Is Nothing Then
		WriteLog "No TA automation is bound to this transformer - " & mode & " request ignored."
		Exit Sub
	End If

	WriteLog mode & " requested via " & bound.Name & "."
	bound.Item("Commands").Item("Start").WriteEx mode & ":" & xatm_Transformer.Id

End Sub

' Asks the high-voltage reclosing to run, because this transformer's lockout
' relay operated.
'
' Found rather than configured. There is one reclosing automation in a
' station and it holds no references to equipment, so the transformer looks
' for it by class - the same way RequestTA looks for the transfer bound to
' this transformer, and for the same reason: E3 gives no way to call across
' objects, so the only way to ask for work is to write to a tag on the
' object that does it.
'
' A station whose incomer layout declares no entry bays has no reclosing to
' find, and says so once rather than failing quietly.
Sub RequestRASEAT()

	Dim obj, bound
	Set bound = Nothing

	For Each obj In Application.GetObject("XATM_Data.Automation")

		If TypeName(obj) = "xatm_RASEAT" Then
			Set bound = obj
			Exit For
		End If

	Next

	If bound Is Nothing Then
		WriteLog "No RASEAT in this substation - reclosing request ignored."
		Exit Sub
	End If

	WriteLog "Reclosing requested via " & bound.Name & "."

	' The Id travels so the reclosing can name who asked. It reads the
	' lockout relay of every transformer for itself: more than one can
	' operate together, and the sequence has to wait for all of them to
	' isolate.
	bound.Item("Commands").Item("Start").WriteEx xatm_Transformer.Id

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

<xatm_Transformer.Data.Triggers.TA:TA_OnLockingoutRelayTrip()>
Sub TA_OnLockingoutRelayTrip()
	
	' Both, off the one relay.
	'
	' The lockout takes the transformer off the medium-voltage busbars
	' and the entry above it at the same time, so both answers are wanted
	' and neither is the other's business: the transfer moves the load,
	' the reclosing brings the entry back. Each refuses itself if it is
	' not in a state to run, and each says so on the console.
	'
	' Undervoltage below asks for the transfer alone. Nothing has opened
	' at the entry there, so there is nothing to reclose.
	RequestTA "TA"
	RequestRASEAT
		
End Sub

<xatm_Transformer.Data.Triggers.TA:TA_OnUndervoltageTrip()>
Sub TA_OnUndervoltageTrip()

	RequestTA "TA"
	
End Sub

<xatm_Version.Data.Version:Version_Functions()>
Sub Version_Functions()
End Sub


' The console the automation logs to, and the E3 trace either way.
'
' Parent.Parent and not the bare Name: this runs in a tag inside the
' object, so Name is the tag's own and would put the same word on every
' line no matter which instance said it.
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

<xatm_Version.Data.Version:Version_OnChangedValue()>
Sub Version_OnChangedValue()
	
	' The report is asked for by writing this tag, which is the shape every
	' command in the library takes: one E3 object cannot call a procedure in
	' another's scope, so a screen writes a tag and the object acts on it.
	'
	' The value means nothing beyond "asked for" - which tag was written is
	' what says what was wanted.
	If Trim(Value & "") = "" Then Exit Sub

	Dim ts
	ts = TimeStamp

	ShowReport

	' Cleared with the timestamp it arrived with, so emptying the tag is not
	' itself a second command. The guard above covers the other case, where
	' the event is driven by the value rather than by the timestamp.
	WriteEx "", ts
		
End Sub

<xatm_Version.Data.Version:Version_OnStartRunning()>
Sub Version_OnStartRunning()

	' Which library is running, published for anything that wants to know.
	'
	' This one stands alone in the data project rather than sitting inside
	' every device, because what it has to offer - three tags an operation
	' centre can read, and a report somebody opens - is worth having once.
	' Whether the build is a demo lives on xatm_Build instead, one per
	' device, where the command gate needs it.
	'
	' Every write is qualified. This runs in a tag inside the object and not
	' in the object's own scope, so a bare Version = ... would quietly make a
	' script variable of that name and leave the property empty.
	xatm_Version.Version      = LIBRARY_VERSION
	xatm_Version.ReleaseNotes = RELEASE_NOTES

	' The version again as three numbers, split from the one string so that
	' cutting a release stays a single edit and the parts cannot come to
	' disagree with the whole.
	Dim parts
	parts = Split(LIBRARY_VERSION & "", ".")

	xatm_Version.Major = VersionPart(parts, 0)
	xatm_Version.Minor = VersionPart(parts, 1)
	xatm_Version.Patch = VersionPart(parts, 2)

	WriteLog Announcement()

End Sub


' What library this is, and what it says it changed.
'
' The version lives here and the demo flag lives on xatm_Build. Two facts
' about the same file, kept apart because they are wanted in different
' places: the flag has to be inside every device for the command gate, and
' this has to be somewhere the interface and the distribution can walk to.
Const LIBRARY_VERSION = "1.0.0"
Const RELEASE_NOTES   = "Automatic reclosing (RASEAT); per-maneuver preconditions and blocks; one command per contingency."


' One number out of the version string, and 0 for anything that is not
' there or is not a number - a version somebody wrote as "1.0" still
' publishes a patch, and it is 0.
Function VersionPart(parts, i)

	VersionPart = 0

	If i > UBound(parts) Then Exit Function
	If IsNumeric(parts(i)) Then VersionPart = CLng(parts(i))

End Function


' The line the console carries at start.
'
' The edition comes from the plant rather than from a constant here: the
' devices carry it, and asking them means this cannot announce one build
' while the gate enforces the other.
'
' Three shapes and not two. Nothing orders the start of one object against
' another, so this can run before a single xatm_Build has written its flag,
' and a start line reading "runtime build" on a demo build is the one wrong
' answer the whole mechanism exists to prevent. With nobody yet answering,
' it gives the version and stops there.
Function Announcement()

	Select Case LibraryEdition()

		Case "DEMO"

			Announcement = "xatm_lib " & LIBRARY_VERSION & " - DEMO build - " & _
			               "commands reach simulated equipment only; nothing is sent to the switchyard."

		Case "RUNTIME"

			Announcement = "xatm_lib " & LIBRARY_VERSION & " - runtime build."

		Case Else

			Announcement = "xatm_lib " & LIBRARY_VERSION & "."

	End Select

End Function


' Which build is driving this station, as a word - and "" when nothing has
' answered.
'
' Asked of the equipment: every breaker and disconnector carries a Build,
' all of them come from the same library file, so the first one found
' answers for the rest. Nothing answering means a station not yet built, a
' library too old to have the class, or simply a start that has not reached
' the switchyard objects - none of which is an edition, and none of which
' is worth guessing at.
Function LibraryEdition()

	LibraryEdition = ""

	Dim substation
	Set substation = Nothing

	On Error Resume Next
	Set substation = Application.GetObject("XATM_Data.Substation")
	On Error Goto 0

	If substation Is Nothing Then Exit Function

	Dim found
	found = Empty

	FindBuild substation, found

	If IsEmpty(found) Then Exit Function

	LibraryEdition = IIf(CBool(found), "DEMO", "RUNTIME")

End Function

Sub FindBuild(folder, ByRef found)

	If Not IsEmpty(found) Then Exit Sub

	Dim obj
	For Each obj In folder

		If Not IsEmpty(found) Then Exit Sub

		Dim build
		Set build = Nothing

		On Error Resume Next
		Set build = obj.Item("Build")
		On Error Goto 0

		If Not build Is Nothing Then

			On Error Resume Next
			found = CBool(build.Demo)
			On Error Goto 0

			If Not IsEmpty(found) Then Exit Sub

		End If

		On Error Resume Next
		FindBuild obj, found
		On Error Goto 0

	Next

End Sub

