-----------------------
Documentação de Scripts
-----------------------
Main
Mon Jul  6 15:12:44 2026
-----------------------

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


' Past-tense word for an operation, used in step logs.
Function OperationName(operation)
	If operation = 2 Then
		OperationName = "closed"
	Else
		OperationName = "opened"
	End If
End Function


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

	xatm_BTC.Running = False

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

            End If

        Case 3

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then

                Main_Step03()

            Else

                WriteLog "Step 3: Execution failed - Timeout exceeded."

            End If

        Case 4

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then

                Main_Step04()

            Else

                WriteLog "Step 4: Execution failed - Timeout exceeded."

            End If

        Case 5

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then

                Main_Step05()

            Else

                WriteLog "Step 5: Execution failed - Timeout exceeded."

            End If

        Case 6

            If Parent.Item("StepTimer").Value < STEP_1_TIMER Then

                Main_Step06()

            Else

                WriteLog "Step 6: Execution failed - Timeout exceeded."

            End If

        Case 99

			Main_Completed()

	End Select

	IncrementTimer()

End Sub

<xatm_BTC.FSM.Main:Main_Step00()>
Sub Main_Step00()

	' ===============
	' Snapshot the substation transformers out-of-service state at trigger time.
	' ===============
	StoreImpediments()

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
	
	Dim impediments
	impediments  = ReadImpediments()
	
	Dim outList
	outList = ""

	Dim i
	For i = 1 To 4
		If impediments(i) Then
			If outList <> "" Then outList = outList & ", "
			outList = outList & TransformerName(i * 100)
		End If
	Next

	If outList = "" Then
		DescribeAutomation = autoType & " " & TransformerName(triggerId) & " - no transformer out of service"
	Else
		DescribeAutomation = autoType & " " & TransformerName(triggerId) & " - " & outList & " out of service"
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

	End Select

End Sub

' ================================
' TM - Manual transfer
' ================================
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

<xatm_BTC.FSM.Main:Main_Step02()>
Sub Main_Step02()

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value

	Select Case Parent.Item("AutomationType").Value

		Case "TM"

			S2TM triggerId

		Case "NM"

			S2NM triggerId

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

<xatm_BTC.FSM.Main:Main_Step03()>
Sub Main_Step03()

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value

	Select Case Parent.Item("AutomationType").Value

		Case "TM"

			S3TM triggerId

		Case "NM"

			S3NM triggerId

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

<xatm_BTC.FSM.Main:Main_Step04()>
Sub Main_Step04()

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value

	Select Case Parent.Item("AutomationType").Value

		Case "TM"

			S4TM triggerId

		Case "NM"

			S4NM triggerId

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

<xatm_BTC.FSM.Main:Main_Step05()>
Sub Main_Step05()

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value

	Select Case Parent.Item("AutomationType").Value

		Case "TM"

			S5TM triggerId

		Case "NM"

			S5NM triggerId

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

<xatm_BTC.FSM.Main:Main_Step06()>
Sub Main_Step06()

	Dim triggerId
	triggerId = Parent.Item("TriggerTransformerId").Value

	Select Case Parent.Item("AutomationType").Value

		Case "TM"

			S6TM triggerId

		Case "NM"

			S6NM triggerId

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


' ================================
' NM - Manual restore (inverse of TM: restores the substation to normal)
' Read from the spec bottom-to-top. Step actions alternate close/open
' (odd close, even open), same as TM; the breaker selection differs.
' ================================
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
