-----------------------
Documentação de Scripts
-----------------------
XATM_LIBCONFIG (C:\ProjDev\edp_sp\xatm_libconfig.lib)
Thu Jul 30 17:04:40 2026
-----------------------

<xatm_BTCStatus.Source:xatm_BTCStatus_OnSourceChanged()>
Sub xatm_BTCStatus_OnSourceChanged()
	
	Links.CreateLink "CurrentStep", Source.Item("FSM").Item("Main").PathName & ".Value"
	Links.CreateLink "StepTimer", Source.Item("FSM").Item("StepTimer").PathName & ".Value"
	
End Sub

<xatm_BreakerSymbol.Source:xatm_BreakerSymbol_OnSourceChanged()>
Sub xatm_BreakerSymbol_OnSourceChanged()
	
	Item("lblName").Value = Source.Name
	
	Links.CreateLink "Position", Source.Item("Data").Item("Position").PathName & ".Value"
	
	Links.CreateLink "MemorizedPosition", Source.Item("Data").Item("MemorizedPosition").PathName & ".Value"
	Links.CreateLink "MemorizedPositionTimer", Source.Item("Data").Item("MemorizedPosition").PathName & ".DocString"
	
	Links.CreateLink "CommunicationFailure", Source.Item("Data").Item("CommunicationFailure").PathName & ".Value"
	
	Links.CreateLink "SimulationModeEnabled", Source.Item("Data").Item("SimulationModeEnabled").PathName & ".Value"
	
	Links.CreateLink "SimulateCommandFailure", Source.Item("Data").Item("SimulateCommandFailure").PathName & ".Value"
	
	Links.CreateLink "CommandInProgress", Source.Item("Data").Item("CommandInProgress").PathName & ".Value"
	
	Links.CreateLink "CommandTimer", Source.Item("Data").Item("Timers").Item("CommandTimer").PathName & ".Value"
		
End Sub

<xatm_BreakerSymbol.objButton:objButton_Click()>
Sub objButton_Click()
	
	If Not xatm_BreakerSymbol.SimulationModeEnabled Then
		
		MsgBox "Simulation mode is not enabled for this equipment.", vbExclamation, "Simulation Mode Disabled"
		Exit Sub
	
	End If
	
	Dim source
	Set source = xatm_BreakerSymbol.Source
	
	Dim openCmd, closeCmd
	openCmd  = "Open"
	closeCmd = "Close"

	Select Case xatm_BreakerSymbol.Position
		Case 1 : openCmd  = "!Open"   ' already open- disable Open
		Case 2 : closeCmd = "!Close"  ' already closed- disable Close
	End Select

	Dim failTag
	Set failTag = source.Item("Data").Item("SimulateCommandFailure")

	Dim options
	options = "Command{" & openCmd & "|" & closeCmd & "||" & _
	          IIf(CBool(failTag.Value), "*", "") & "Command Failure}|" & _
	          "Reset||Cancel"

	Dim userOption
	userOption = Application.SelectMenu(options)

	Select Case userOption

		Case 1, 2

			source.Item("Data").Item("CommandOpenClose").WriteEx userOption

		Case 3

			failTag.WriteEx Not CBool(failTag.Value)

		Case 4

			source.Item("Data").Item("Reset").WriteEx True

	End Select	
	
End Sub

<xatm_SelectLayout.Index:xatm_SelectLayout_OnIndexChanged()>
Sub xatm_SelectLayout_OnIndexChanged()

	Item("Image").Filename = ParentFolderPath & Index & FileExtension
		
End Sub

<xatm_SelectLayout.btnNext:btnNext_Click()>
Sub btnNext_Click()

	Dim imageList
	imageList = xatm_SelectLayout.ImageList

	Dim index
	index = xatm_SelectLayout.Index

	If Trim(imageList) = "" Then Exit Sub

	Dim parts
	parts = Split(imageList, ",")

	Dim layouts()
	ReDim layouts(UBound(parts))

	Dim count
	count = 0

	Dim current
	current = -1

	Dim i
	For i = 0 To UBound(parts)
		If Trim(parts(i)) <> "" Then
			layouts(count) = Trim(parts(i))
			If StrComp(layouts(count), Trim(index), vbTextCompare) = 0 Then current = count
			count = count + 1
		End If
	Next

	If count = 0 Then Exit Sub

	If current < 0 Then
		xatm_SelectLayout.Index = layouts(0)
		Exit Sub
	End If

	If current = count - 1 Then Exit Sub

	xatm_SelectLayout.Index = layouts(current + 1)
	
End Sub

<xatm_SelectLayout.btnPrevious:btnPrevious_Click()>
Sub btnPrevious_Click()

	Dim imageList
	imageList = xatm_SelectLayout.ImageList

	Dim index
	index = xatm_SelectLayout.Index

	If Trim(imageList) = "" Then Exit Sub

	Dim parts
	parts = Split(imageList, ",")

	Dim layouts()
	ReDim layouts(UBound(parts))

	Dim count
	count = 0

	Dim current
	current = -1

	Dim i
	For i = 0 To UBound(parts)
		If Trim(parts(i)) <> "" Then
			layouts(count) = Trim(parts(i))
			If StrComp(layouts(count), Trim(index), vbTextCompare) = 0 Then current = count
			count = count + 1
		End If
	Next

	If count = 0 Then Exit Sub

	If current < 0 Then
		xatm_SelectLayout.Index = layouts(0)
		Exit Sub
	End If

	If current = 0 Then Exit Sub

	xatm_SelectLayout.Index = layouts(current - 1)
		
End Sub

<xatm_TransformerSymbol.Source:xatm_TransformerSymbol_OnSourceChanged()>
Sub xatm_TransformerSymbol_OnSourceChanged()
	
	Item("lblName").Value = Source.Name
	
	Links.CreateLink "SimulationModeEnabled", Source.Item("Data").Item("SimulationModeEnabled").PathName & ".Value"
		
End Sub

<xatm_TransformerSymbol.objButton:objButton_Click()>
Sub objButton_Click()
	
	If Not xatm_TransformerSymboL.SimulationModeEnabled Then
		
		MsgBox "Simulation mode is not enabled for this equipment.", vbExclamation, "Simulation Mode Disabled"
		Exit Sub
	
	End If
	
	Dim source
	Set source = xatm_TransformerSymboL.Source
	
	Dim options
	options = IIf(source.OutOfService,      "*", "") & "Out Of Service|" & _
	          IIf(source.UndervoltageRelay, "*", "") & "Undervoltage (27)|" & _
	          IIf(source.LockingOutRelay,   "*", "") & "Locking Out Relay (86)|" & _
	          "|Reset|Cancel"

	Dim userOption
	userOption = Application.SelectMenu(options)
	
	Select Case userOption

		Case 1
			source.OutOfService = Not source.OutOfService

		Case 2
			source.UndervoltageRelay = Not source.UndervoltageRelay

		Case 3
			source.LockingOutRelay = Not source.LockingOutRelay

		Case 4
			' TODO: Reset

	End Select
	
End Sub

