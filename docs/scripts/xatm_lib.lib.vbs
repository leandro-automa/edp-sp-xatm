-----------------------
Documentação de Scripts
-----------------------
XATM_LIB (C:\ProjDev\edp_sp\xatm_lib.lib)
Wed Aug 12 16:59:28 2026
-----------------------

<xatm_BTC.Commands.OperatorBlock:OperatorBlock_CommandOperatorBlock()>
Sub OperatorBlock_CommandOperatorBlock()
	
	Dim v
	v = False
	
	On Error Resume Next
	v = CBool(xatm_BTC.CommandOperatorBlock.Value)
	On Error Goto 0
	
	xatm_BTC.OperatorBlock = v

End Sub

<xatm_BTC.Commands.OperatorBlock:OperatorBlock_OnChangedValue()>
Sub OperatorBlock_OnChangedValue()
	
	If xatm_BTC.OperatorBlock Then
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

<xatm_BTC.Commands.Reset:Reset_OnChangedTimeStamp()>
Sub Reset_OnChangedTimeStamp()
	
	If CBool(Value) Then
		
		Reset()
		
	End If
	
End Sub

Sub Reset()
	
	xatm_BTC.Running 		= False
	xatm_BTC.GeneralBlock	= False
	
	' ===============
	' RESET ALL STEP FAIL POINTS
	' ===============
	Dim i
	For i = 1 To 6
		Dim propName
		propName = "StepExecutionFailed" & i
		
		On Error Resume Next
		Execute("xatm_BTC." & propName & " = False")
		On Error Goto 0
	Next
	
	Dim tag
	For Each tag In xatm_BTC.Item("FSM")
		
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

<xatm_BTC.Commands.Reset:Reset_Reset()>
Sub Reset_Reset()
	
	If xatm_BTC.CommandReset.Value = 0 Then Exit Sub
	
	WriteEx True
		
End Sub

<xatm_BTC.Commands.Start:Start_CommandStartNM()>
Sub Start_CommandStartNM()

	If xatm_BTC.CommandStartNM.Value = 0 Then Exit Sub

	StartMode "NM", 0

End Sub

<xatm_BTC.Commands.Start:Start_CommandStartNM100()>
Sub Start_CommandStartNM100()

	If xatm_BTC.CommandStartNM100.Value = 0 Then Exit Sub

	StartMode "NM", 100

End Sub

<xatm_BTC.Commands.Start:Start_CommandStartNM200()>
Sub Start_CommandStartNM200()

	If xatm_BTC.CommandStartNM200.Value = 0 Then Exit Sub

	StartMode "NM", 200

End Sub

<xatm_BTC.Commands.Start:Start_CommandStartNM300()>
Sub Start_CommandStartNM300()

	If xatm_BTC.CommandStartNM300.Value = 0 Then Exit Sub

	StartMode "NM", 300

End Sub

<xatm_BTC.Commands.Start:Start_CommandStartNM400()>
Sub Start_CommandStartNM400()

	If xatm_BTC.CommandStartNM400.Value = 0 Then Exit Sub

	StartMode "NM", 400

End Sub

<xatm_BTC.Commands.Start:Start_CommandStartTM()>
Sub Start_CommandStartTM()

	If xatm_BTC.CommandStartTM.Value = 0 Then Exit Sub

	StartMode "TM", 0

End Sub

<xatm_BTC.Commands.Start:Start_CommandStartTM100()>
Sub Start_CommandStartTM100()

	If xatm_BTC.CommandStartTM100.Value = 0 Then Exit Sub

	StartMode "TM", 100

End Sub

<xatm_BTC.Commands.Start:Start_CommandStartTM200()>
Sub Start_CommandStartTM200()

	If xatm_BTC.CommandStartTM200.Value = 0 Then Exit Sub

	StartMode "TM", 200

End Sub

<xatm_BTC.Commands.Start:Start_CommandStartTM300()>
Sub Start_CommandStartTM300()

	If xatm_BTC.CommandStartTM300.Value = 0 Then Exit Sub

	StartMode "TM", 300

End Sub

<xatm_BTC.Commands.Start:Start_CommandStartTM400()>
Sub Start_CommandStartTM400()

	If xatm_BTC.CommandStartTM400.Value = 0 Then Exit Sub

	StartMode "TM", 400

End Sub

<xatm_BTC.Commands.Start:Start_OnChangedValue()>
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
	' WHICH TRANSFORMER IS OUT
	' ================
	'
	' TM and NM are asked for by a person, and the person says which one -
	' that is the whole of the new triggering philosophy. TA is asked for
	' by the switchyard, which says nothing, so the field is read instead.
	Dim impedeId
	impedeId = 0

	If mode = "TA" Then

		Dim outCount
		impedeId = FieldImpediment(triggerId, outCount)

		' Tabela 1 situations 6-11 are two transformers out, and no
		' sequence leaves them. Refusing is the honest answer; picking one
		' of the two would run a sequence whose starting state is false.
		If outCount > 1 Then

			Reject "More than one transformer is out of service - no sequence is defined from there.", ts
			Exit Sub

		End If

	ElseIf UBound(parts) = 2 Then

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

	' A transfer needs somewhere to put the load. Anything that is neither
	' the trigger nor out of service can take it; with none left there is
	' nothing to transfer to and no sequence to run.
	'
	' On the four transformer ring one impediment still leaves two, so this
	' only bites on the small layouts - two transformers with the adjacent
	' one impeded, which is the case the client asked to have refused.
	If CountDestinations(triggerId, impedeId) = 0 Then

		Reject "No transformer is left to take the load.", ts
		Exit Sub

	End If

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

	If xatm_BTC.OperatorBlock Then

	 	Reject "Blocked by operator.", ts
		Exit Sub

	End If

	If xatm_BTC.GeneralBlock Then

	 	Reject "General interlock active.", ts
	 	Exit Sub

	End If

	' The switchyard withholding it and the switchyard permitting it - but
	' asked of this maneuver own pair of gates, not of one pair standing
	' for every maneuver the automation can run.
	'
	' Both are read before either is judged, so a gate the XObject has not
	' been given is reported as the configuration fault it is, rather than
	' surfacing as a maneuver that will not start for no stated reason.
	'
	' The manifest only describes a property - it drives the config screen,
	' the interface, the alarms and the distribution. The XObject has to
	' carry the property itself, or there is nothing for any of them to
	' describe.
	Dim gate, missing, blocked, permitted
	gate    = GateSuffix(mode, impedeId)
	missing = ""

	blocked   = ReadGate("AutomaticBlock" & gate, False, missing)
	permitted = ReadGate("Preconditions"  & gate, True,  missing)

	If missing <> "" Then

		Reject "this automation has no " & missing & " - add the property to the xatm_BTC XObject.", ts
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

	xatm_BTC.Item("FSM").Item("AutomationType").WriteEx mode
	xatm_BTC.Item("FSM").Item("TriggerTransformerId").WriteEx triggerId
	xatm_BTC.Item("FSM").Item("ImpededTransformerId").WriteEx impedeId
	xatm_BTC.Item("FSM").Item("StepTimer").WriteEx 0
	xatm_BTC.Item("FSM").Item("Main").WriteEx 0
	xatm_BTC.Running = True

	If impedeId = 0 Then
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
	Set transformer = xatm_BTC.Transformer
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

' Which pair of gates a maneuver answers to.
'
' The mode, and the transformer it assumes is out where there is one: TM,
' TM200, NM400. The same names the command tags carry, so the command and
' the gates that let it through are one string apart.
'
' TA is always plain TA. It has one pair and not five - nobody commands
' it, so there is no command for a gate to pair with.
Function GateSuffix(mode, impedeId)

	If mode = "TA" Then
		GateSuffix = "TA"
	ElseIf impedeId = 0 Then
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
	Execute "gGateValue = xatm_BTC." & propertyName
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
<xatm_BTC.FSM.Main:Main_Completed()>
Sub Main_Completed()
	
	Parent.Item("TriggerTransformerId").WriteEx  Empty, 0
	Parent.Item("AutomationType").WriteEx Empty, 0
	Parent.Item("StepTimer").WriteEx  Empty, 0
	WriteEx  Empty, 0
	
    xatm_BTC.Running = False
    WriteLog "Automation completed successfully."
            	
End Sub

<xatm_BTC.FSM.Main:Main_Functions()>
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

<xatm_BTC.FSM.Main:Main_GlobalLockout()>
Sub Main_GlobalLockout()
	
	xatm_BTC.Running 		= False
	xatm_BTC.GeneralBlock 	= True

	Select Case Value
		Case 1 : xatm_BTC.StepExecutionFailed1 = True
		Case 2 : xatm_BTC.StepExecutionFailed2 = True
		Case 3 : xatm_BTC.StepExecutionFailed3 = True
		Case 4 : xatm_BTC.StepExecutionFailed4 = True
		Case 5 : xatm_BTC.StepExecutionFailed5 = True
		Case 6 : xatm_BTC.StepExecutionFailed6 = True
	End Select
	
	WriteLog "Global lockout activated due to automation failure."
	
End Sub

<xatm_BTC.FSM.Main:Main_Main()>
Sub Main_Main()
	
	Const STEP_1_TIMER = 30

	If Not xatm_BTC.Enabled Then
		
		WriteLog "Automation not enabled."
		xatm_BTC.Running = False
		Exit Sub
	
	End If
	
	If Not xatm_BTC.Running Then
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

<xatm_BTC.FSM.Main:Main_Step00()>
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

	If impedeId = 0 Then
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

' A scope may not end on a Function - E3 takes the script without complaint
' and then behaves as though the last one were not there. StoreImpediments
' used to close this section; with it gone, this stub does. A guard, not
' debris.
Sub Step00Functions()
End Sub

<xatm_BTC.FSM.Main:Main_Step01()>
Sub Main_Step01()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S1TM triggerId, impedeId
		
		Case "NM"

			S1NM triggerId, impedeId
		
		Case "TA"

			S1TA triggerId, impedeId
			
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

<xatm_BTC.FSM.Main:Main_Step02()>
Sub Main_Step02()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S2TM triggerId, impedeId
		
		Case "NM"

			S2NM triggerId, impedeId

		Case "TA"

			S2TA triggerId, impedeId
			
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

<xatm_BTC.FSM.Main:Main_Step03()>
Sub Main_Step03()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S3TM triggerId, impedeId
		
		Case "NM"

			S3NM triggerId, impedeId
		
		Case "TA"

			S3TA triggerId, impedeId
			
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

<xatm_BTC.FSM.Main:Main_Step04()>
Sub Main_Step04()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S4TM triggerId, impedeId
		
		Case "NM"

			S4NM triggerId, impedeId
		
		Case "TA"

			S4TA triggerId, impedeId
			
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
				' Ends on DJ06/420. Spec 1.4.2.23's step table writes
				' DESLIGA DJ05, but its own prose says "Ligar DJ-10 e
				' desligar DJ-06", its end state puts B4A back on TR1, and
				' DJ05 is already open whenever TR3 is impeded - so the
				' table's step is a no-op that would leave TR4 connected
				' and report success. Same typo as 1.4.3.16 for TA.
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

<xatm_BTC.FSM.Main:Main_Step05()>
Sub Main_Step05()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S5TM triggerId, impedeId
		
		Case "NM"

			S5NM triggerId, impedeId
		
		Case "TA"

			S5TA triggerId, impedeId
			
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

<xatm_BTC.FSM.Main:Main_Step06()>
Sub Main_Step06()

	Dim triggerId, impedeId
	triggerId = Parent.Item("TriggerTransformerId").Value
	impedeId  = ReadImpededId()
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S6TM triggerId, impedeId
		
		Case "NM"

			S6NM triggerId, impedeId
		
		Case "TA"

			S6TA triggerId, impedeId
			
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

	' --- Flag command in progress (drives the command timer / BTC) ---
	If Parent.Item("CommandInProgress").Value <> 2 Then
		Parent.Item("CommandInProgress").WriteEx 2
	End If

End Sub

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
    
    ' ====================================
    ' Append to content array
    ' ====================================
    '
    ' One entry per physical line, not one per call.
    '
    ' A caller may hand over a whole report in a single write - BuildAlarms
    ' builds its complaint with a vbCrLf before every item - and two things
    ' break when that lands in one entry. The ring below counts entries, so
    ' MaxLines stops meaning lines and a twenty line report evicts nineteen
    ' lines that should have stayed. And the list box the Footer draws with
    ' refuses an item carrying a line break outright, which takes the whole
    ' redraw down with it rather than just that row.
    '
    ' The timestamp goes on the first line and the rest are indented under
    ' it, so a report still reads as one event instead of twenty.
    Dim contentTag
    Set contentTag = Item("Data").Item("Content")

    Dim contentArr
    If IsArray(contentTag.Value) Then
        contentArr = contentTag.Value
    Else
        contentArr = Array()
    End If

    Dim msgBody
    msgBody = Replace(Replace(WriteLine & "", vbCrLf, vbLf), vbCr, vbLf)

    Dim msgLines
    msgLines = Split(msgBody, vbLf)

    Dim continuationPad
    continuationPad = Space(Len(formattedTimeStamp))

    Dim n, lineText, currentSize, index, j

    For n = 0 To UBound(msgLines)

        ' A blank continuation line carries nothing and would only push a
        ' real one out of the ring. The first line is kept either way: an
        ' empty message is still an event worth timestamping.
        If n = 0 Or Trim(msgLines(n)) <> "" Then

            If n = 0 Then
                lineText = formattedTimeStamp & vbTab & RTrim(msgLines(n))
            Else
                lineText = continuationPad & vbTab & RTrim(msgLines(n))
            End If

            currentSize = UBound(contentArr) + 1

            If currentSize < MAX_LINES Then

                ' Array still growing - just append
                index = currentSize
                ReDim Preserve contentArr(currentSize)

            Else

                index = MAX_LINES - 1

                ' Array at capacity - shift and append
                For j = 0 To MAX_LINES - 2
                    contentArr(j) = contentArr(j + 1)
                Next

            End If

            contentArr(index) = lineText

            ' Traced a line at a time, so the E3 trace and the console
            ' agree on what one line is.
            If TraceEnabled Then
                On Error Resume Next
                Application.Trace lineText, False
                On Error Goto 0
            End If

        End If

    Next

    ' ================================
    ' Update the content tag with the new array
    ' ================================
    '
    ' Written once, after every line. The tag drives the Footer's redraw,
    ' so writing it per line would redraw the screen once per line of a
    ' report.
    contentTag.Value = contentArr

    ' ================================
    ' Clear the WriteLine property to prepare for the next message
    ' ================================    Me.WriteLine = Empty
	
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
	' TRIGGER BTC
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

<xatm_Transformer.Data.Triggers.BTC:BTC_Functions()>
Sub BTC_Functions()
End Sub

Sub RequestBTC(mode)

	Dim obj, bound, tr
	Set bound = Nothing

	For Each obj In Application.GetObject("XATM_Data.Automation")
		If TypeName(obj) = "xatm_BTC" Then

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
		WriteLog "No BTC bound to this transformer - " & mode & " request ignored."
		Exit Sub
	End If

	WriteLog mode & " requested via " & bound.Name & "."
	bound.Item("Commands").Item("Start").WriteEx mode & ":" & xatm_Transformer.Id

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

<xatm_Transformer.Data.Triggers.BTC:BTC_OnLockingoutRelayTrip()>
Sub BTC_OnLockingoutRelayTrip()
	
	RequestBTC "TA"
		
End Sub

<xatm_Transformer.Data.Triggers.BTC:BTC_OnUndervoltageTrip()>
Sub BTC_OnUndervoltageTrip()

	RequestBTC "TA"
	
End Sub

