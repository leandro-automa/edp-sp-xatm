-----------------------
Documentação de Scripts
-----------------------
XATM_LIB (C:\ProjDev\edp_sp\xatm_lib.lib)
Tue Aug  4 14:08:27 2026
-----------------------

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

	If xatm_BTC.OperatorBlock Then
		
	 	Reject "Blocked by operator.", ts
		Exit Sub

	End If

	If xatm_BTC.GeneralBlock Then
		
	 	Reject "General interlock active.", ts
	 	Exit Sub

	End If

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

	WriteLog "Start - " & mode & " TR" & triggerId

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

' The layout combinations the step sequences below are written for. Any
' other combination has no sequence and must not be run.
Function IsSupportedLayout(layoutType)

	Select Case layoutType

		Case "4TR4LV_6BB6TIERING", "2TR2LV_2BB1TIE"
			IsSupportedLayout = True

		Case Else
			IsSupportedLayout = False

	End Select

End Function

' Ends the run when a step is reached that the active layout has no
' sequence for, so a command is never sent to a breaker picked by default.
Sub UnsupportedStep(stepNumber, layoutType)

	WriteLog "Step " & stepNumber & ": layout '" & layoutType & _
	         "' has no step " & stepNumber & " - Global lockout"
	Main_GlobalLockout()

End Sub

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

	Dim layoutType
	layoutType = GetLayoutType()

	If Not IsSupportedLayout(layoutType) Then

		Main_Abort "layout '" & layoutType & "' has no step sequence."
		Exit Sub

	End If

	' ===============
	' Snapshot the substation transformers lockout state at trigger time.
	' ===============
	StoreImpediments()

	Dim refusal
	refusal = ContingencyRefusal(layoutType, Parent.Item("TriggerTransformerId").Value)

	If refusal <> "" Then

		Main_Abort refusal
		Exit Sub

	End If

	WriteLog "Starting " & DescribeAutomation()

	Value = 1

End Sub

' Ends the run before any command has been sent. A refusal is not a step
' failure, so it clears the run the way a completion does instead of
' raising the global lockout.
Sub Main_Abort(reason)

	Parent.Item("TriggerTransformerId").WriteEx Empty, 0
	Parent.Item("AutomationType").WriteEx Empty, 0
	Parent.Item("StepTimer").WriteEx Empty, 0
	WriteEx Empty, 0

	xatm_BTC.Running = False
	WriteLog "Not executed - " & reason

End Sub

' Why the substation state rules the automation out, or "" when it does
' not. The ring always keeps an alternative path open; the two-transformer
' layout has a single tie, so every sequence leans on the other
' transformer and none of them can run while it is out of service.
Function ContingencyRefusal(layoutType, triggerId)

	ContingencyRefusal = ""

	If layoutType <> "2TR2LV_2BB1TIE" Then Exit Function

	Dim otherIndex

	Select Case triggerId

		Case 100 : otherIndex = 2
		Case 200 : otherIndex = 1

		Case Else
			ContingencyRefusal = TransformerName(triggerId) & " is not part of the " & _
			                     layoutType & " layout."
			Exit Function

	End Select

	Dim impediments
	impediments = ReadImpediments()

	If impediments(otherIndex) Then
		ContingencyRefusal = TransformerName(otherIndex * 100) & " is out of service - " & _
		                     "it is the only source left for both busbars."
	End If

End Function

' Human-readable summary of the running automation for the step log,
' using transformer names, e.g. "TM TR-01 - TR-02 out of service".
Function DescribeAutomation()

	Dim autoType
	autoType  = Parent.Item("AutomationType").Value
	
	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value
	
	Dim impediments
	impediments  = ReadImpediments()
	
	Dim lockedList
	lockedList = ""

	Dim i
	For i = 1 To 4
		If impediments(i) Then
			If lockedList <> "" Then lockedList = lockedList & ", "
			lockedList = lockedList & TransformerName(i * 100)
		End If
	Next

	If lockedList = "" Then
		DescribeAutomation = autoType & " " & TransformerName(triggerId) & " - no transformer out of service"
	Else
		DescribeAutomation = autoType & " " & TransformerName(triggerId) & " - " & lockedList & " out of service"
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

Function ReadImpediments()

    ReadImpediments = Parent.Item("TransformerImpediments").Value

End Function

Sub StoreImpediments()

	Dim impediments(4)
	Dim i, transformer, transformerExists

	For i = 1 To 4

		Set transformer = GetDeviceById(i * 100, transformerExists)

		If transformerExists Then
			impediments(i) = transformer.OutOfService
		Else
			impediments(i) = False
		End If

	Next

	Parent.Item("TransformerImpediments").WriteEx impediments

End Sub

<xatm_BTC.FSM.Main:Main_Step01()>
Sub Main_Step01()

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S1TM triggerId
		
		Case "NM"

			S1NM triggerId
		
		Case "TA"

			S1TA triggerId
			
	End Select
	
End Sub

Sub S1NM(triggerId)

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

Sub S1TM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				If impediments(2) Then
					breakerId = 700
				Else
					breakerId = 900
				End If

			Case 200
				If impediments(1) Then
					breakerId = 700
				ElseIf impediments(3) Or impediments(4) Then
					breakerId = 730
				Else
					breakerId = 900
				End If

			Case 300
				If impediments(1) Then
					breakerId = 720
				ElseIf impediments(2) Then
					breakerId = 700
				ElseIf impediments(4) Then
					breakerId = 900
				Else
					breakerId = 740
				End If

			Case 400
				If impediments(1) Then
					breakerId = 720
				ElseIf impediments(2) Then
					breakerId = 700
				ElseIf impediments(3) Then
					breakerId = 730
				Else
					breakerId = 740
				End If

		End Select

	ElseIf layoutType = "2TR2LV_2BB1TIE" Then

		' The transformer is still live, so the tie closes first and both
		' busbars are fed before it is taken out.
		action = 2
		breakerId = 700

	Else

		UnsupportedStep 1, layoutType
		Exit Sub

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

Sub S1TA(triggerId)

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

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S2TM triggerId
		
		Case "NM"

			S2NM triggerId

		Case "TA"

			S2TA triggerId
			
	End Select
	
End Sub

Sub S2NM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action, nextStep
	nextStep = 3

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				If impediments(2) Then
					breakerId = 700
				Else
					breakerId = 710
				End If

			Case 200
				If impediments(1) Then
					breakerId = 720
				Else
					breakerId = 710
				End If

			Case 300
				If impediments(4) Then
					breakerId = 730
				ElseIf impediments(1) Or impediments(2) Then
					breakerId = 740
				Else
					breakerId = 740
					nextStep = 99    ' no contingency: sequence ends here
				End If

			Case 400
				If impediments(3) Then
					breakerId = 900
				ElseIf impediments(1) Or impediments(2) Then
					breakerId = 740
				Else
					breakerId = 740
					nextStep = 99    ' no contingency: sequence ends here
				End If

		End Select

	ElseIf layoutType = "2TR2LV_2BB1TIE" Then

		' The transformer is back on its own busbar, so the tie that was
		' carrying it opens. Last step of the sequence.
		action = 1
		breakerId = 700
		nextStep = 99

	Else

		UnsupportedStep 2, layoutType
		Exit Sub

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

Sub S2TM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Dim impediments
		impediments = ReadImpediments()

		nextStep = 3

		Select Case triggerId

			Case 100
				If impediments(2) Then
					breakerId = 120
				Else
					breakerId = 700
				End If

			Case 200
				If impediments(1) Then
					breakerId = 710
				ElseIf impediments(3) Or impediments(4) Then
					breakerId = 720
				Else
					breakerId = 700
				End If

			Case 300
				If impediments(1) Then
					breakerId = 730
				ElseIf impediments(2) Then
					breakerId = 900
				ElseIf impediments(4) Then
					breakerId = 740
				Else
					breakerId = 320
					nextStep = 99    ' no contingency: sequence ends here
				End If

			Case 400
				If impediments(1) Then
					breakerId = 730
				ElseIf impediments(2) Then
					breakerId = 900
				ElseIf impediments(3) Then
					breakerId = 740
				Else
					breakerId = 420
					nextStep = 99    ' no contingency: sequence ends here
				End If

		End Select

	ElseIf layoutType = "2TR2LV_2BB1TIE" Then

		' The tie is closed, so opening the trigger's own breaker hands its
		' busbar to the other transformer. Last step of the sequence.
		action = 1
		breakerId = triggerId + 20
		nextStep = 99

	Else

		UnsupportedStep 2, layoutType
		Exit Sub

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

Sub S2TA(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action, nextStep
	nextStep = 3

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1                      ' default: open a tie to isolate the dead busbar

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				If impediments(2) Then
					breakerId = 710
				Else
					breakerId = 700
				End If

			Case 200
				If impediments(1) Then
					breakerId = 710
				Else
					breakerId = 720
				End If

			Case 300
				If impediments(1) Then
					breakerId = 730
				ElseIf impediments(2) Then
					breakerId = 900
				ElseIf impediments(4) Then
					breakerId = 740
				Else
					breakerId = 740        ' no contingency: 2-step restore
					action = 2
					nextStep = 99
				End If

			Case 400
				If impediments(1) Then
					breakerId = 730
				ElseIf impediments(2) Then
					breakerId = 900
				ElseIf impediments(3) Then
					breakerId = 740
				Else
					breakerId = 740        ' no contingency: 2-step restore
					action = 2
					nextStep = 99
				End If

		End Select

	ElseIf layoutType = "2TR2LV_2BB1TIE" Then

		' The trigger's breaker is confirmed open, so closing the tie restores
		' its busbar from the other transformer. Last step of the sequence.
		action = 2
		breakerId = 700
		nextStep = 99

	Else

		UnsupportedStep 2, layoutType
		Exit Sub

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

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S3TM triggerId
		
		Case "NM"

			S3NM triggerId
		
		Case "TA"

			S3TA triggerId
			
	End Select
	
End Sub

Sub S3NM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				If impediments(2) Then
					breakerId = 710
				ElseIf impediments(3) Or impediments(4) Then
					breakerId = 700
				Else
					breakerId = 720
				End If

			Case 200
				If impediments(1) Then
					breakerId = 710
				Else
					breakerId = 720
				End If

			Case 300
				If impediments(1) Then
					breakerId = 730
				ElseIf impediments(2) Then
					breakerId = 900
				ElseIf impediments(4) Then
					breakerId = 740
				End If

			Case 400
				If impediments(1) Then
					breakerId = 730
				ElseIf impediments(2) Then
					breakerId = 900
				ElseIf impediments(3) Then
					breakerId = 740
				End If

		End Select

	Else

		UnsupportedStep 3, layoutType
		Exit Sub

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

Sub S3TM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				If impediments(2) Then
					breakerId = 720
				ElseIf impediments(3) Or impediments(4) Then
					breakerId = 710
				Else
					breakerId = 730
				End If

			Case 200
				If impediments(1) Then
					breakerId = 720
				ElseIf impediments(3) Or impediments(4) Then
					breakerId = 710
				Else
					breakerId = 730
				End If

			Case 300
				If impediments(1) Or impediments(2) Then
					breakerId = 740
				ElseIf impediments(4) Then
					breakerId = 730
				End If

			Case 400
				If impediments(1) Or impediments(2) Then
					breakerId = 740
				ElseIf impediments(3) Then
					breakerId = 900
				End If

		End Select

	Else

		UnsupportedStep 3, layoutType
		Exit Sub

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

Sub S3TA(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action
	action = 2                          ' restore step: close a tie

	If layoutType = "4TR4LV_6BB6TIERING" Then

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				If impediments(2) Then
					breakerId = 700
				Else
					breakerId = 900
				End If

			Case 200
				If impediments(1) Then
					breakerId = 700
				Else
					breakerId = 730
				End If

			Case 300
				If impediments(1) Then
					breakerId = 740
				ElseIf impediments(2) Then
					breakerId = 700
				ElseIf impediments(4) Then
					breakerId = 900
				End If

			Case 400
				If impediments(3) Then
					breakerId = 900
				Else
					breakerId = 740
				End If

		End Select

	Else

		UnsupportedStep 3, layoutType
		Exit Sub

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

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S4TM triggerId
		
		Case "NM"

			S4NM triggerId
		
		Case "TA"

			S4TA triggerId
			
	End Select
	
End Sub

Sub S4NM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action, nextStep
	nextStep = 99

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				If impediments(2) Then
					breakerId = 720
				ElseIf impediments(3) Or impediments(4) Then
					breakerId = 900
				Else
					breakerId = 730
					nextStep = 5    ' sequence continues
				End If

			Case 200
				If impediments(1) Then
					breakerId = 700
				ElseIf impediments(3) Or impediments(4) Then
					breakerId = 730
				Else
					breakerId = 730
					nextStep = 5    ' sequence continues
				End If

			Case 300
				If impediments(1) Then
					breakerId = 720
				ElseIf impediments(2) Then
					breakerId = 700
				ElseIf impediments(4) Then
					breakerId = 900
				End If

			Case 400
				If impediments(1) Then
					breakerId = 720
				ElseIf impediments(2) Then
					breakerId = 700
				ElseIf impediments(3) Then
					breakerId = 730
				End If

		End Select

	Else

		UnsupportedStep 4, layoutType
		Exit Sub

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

Sub S4TM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	Dim nextStep
	nextStep = 99

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				If impediments(2) Then
					breakerId = 710
				ElseIf impediments(3) Or impediments(4) Then
					breakerId = 120
				Else
					breakerId = 720
					nextStep = 5    ' sequence continues
				End If

			Case 200
				If impediments(1) Or impediments(3) Or impediments(4) Then
					breakerId = 220
				Else
					breakerId = 720
					nextStep = 5    ' sequence continues
				End If

			Case 300
				breakerId = 320

			Case 400
				If impediments(3) Then
					breakerId = 320
				Else
					breakerId = 420
				End If

		End Select

	Else

		UnsupportedStep 4, layoutType
		Exit Sub

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

Sub S4TA(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action, nextStep
	nextStep = 99

	If layoutType = "4TR4LV_6BB6TIERING" Then

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				If impediments(2) Then
					action = 2
					breakerId = 720
				ElseIf impediments(3) Or impediments(4) Then
					action = 2
					breakerId = 710
				Else
					action = 1             ' no contingency: open, sequence continues
					breakerId = 720
					nextStep = 5
				End If

			Case 200
				If impediments(1) Then
					action = 2
					breakerId = 720
				ElseIf impediments(3) Or impediments(4) Then
					action = 2
					breakerId = 710
				Else
					action = 1             ' no contingency: open, sequence continues
					breakerId = 700
					nextStep = 5
				End If

			Case 300
				action = 2
				If impediments(1) Then
					breakerId = 720
				ElseIf impediments(2) Then
					breakerId = 740
				ElseIf impediments(4) Then
					breakerId = 730
				End If

			Case 400
				action = 2
				If impediments(1) Then
					breakerId = 720
				ElseIf impediments(2) Then
					breakerId = 700
				ElseIf impediments(3) Then
					breakerId = 730
				End If

		End Select

	Else

		UnsupportedStep 4, layoutType
		Exit Sub

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

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S5TM triggerId
		
		Case "NM"

			S5NM triggerId
		
		Case "TA"

			S5TA triggerId
			
	End Select
	
End Sub

Sub S5NM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Select Case triggerId

			Case 100, 200
				breakerId = 700

		End Select

	Else

		UnsupportedStep 5, layoutType
		Exit Sub

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

Sub S5TM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 2

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100, 200

				breakerId = 710

		End Select

	Else

		UnsupportedStep 5, layoutType
		Exit Sub

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

Sub S5TA(triggerId)

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

	Else

		UnsupportedStep 5, layoutType
		Exit Sub

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

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value
		
	Select Case Parent.Item("AutomationType").Value
		
		Case "TM"
			
			S6TM triggerId
		
		Case "NM"

			S6NM triggerId
		
		Case "TA"

			S6TA triggerId
			
	End Select
	
End Sub

Sub S6NM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Select Case triggerId

			Case 100, 200
				breakerId = 900

		End Select

	Else

		UnsupportedStep 6, layoutType
		Exit Sub

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

Sub S6TM(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action

	If layoutType = "4TR4LV_6BB6TIERING" Then

		action = 1

		Dim impediments
		impediments = ReadImpediments()

		Select Case triggerId

			Case 100
				breakerId = 120

			Case 200
				breakerId = 220

		End Select

	Else

		UnsupportedStep 6, layoutType
		Exit Sub

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

Sub S6TA(triggerId)

	Dim layoutType
	layoutType = GetLayoutType()

	Dim breakerId, action
	action = 2

	If layoutType = "4TR4LV_6BB6TIERING" Then

		Select Case triggerId

			Case 100, 200
				breakerId = 710

		End Select

	Else

		UnsupportedStep 6, layoutType
		Exit Sub

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

