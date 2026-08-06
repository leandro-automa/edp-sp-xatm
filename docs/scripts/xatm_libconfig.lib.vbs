-----------------------
Documentação de Scripts
-----------------------
XATM_LIBCONFIG (C:\ProjDev\edp_sp\xatm_libconfig.lib)
Thu Aug  6 14:29:56 2026
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

<xatm_PropertyRow.Texto1:Texto1_Click()>
Sub Texto1_Click()
	Dim d
	Set d = CreateObject("Scripting.Dictionary")
	d.Add "Exposure" & vbTab & ": " & xatm_PropertyRow.Exposure, Empty
	d.Add "ObjectType" & vbTab & ": " & xatm_PropertyRow.ObjectType, Empty
	d.Add "PropertyName" & vbTab & ": " & xatm_PropertyRow.PropertyName, Empty
	d.Add "PropertySource" & vbTab & ": " & xatm_PropertyRow.PropertySource, Empty
	d.Add "PropertyType" & vbTab & ": " & xatm_PropertyRow.PropertyType, Empty
	d.Add "Source" & vbTab & ": " & xatm_PropertyRow.Source, Empty
	d.Add "Value" & vbTab & ": " & xatm_PropertyRow.Value, Empty
	d.Add "Kind" & vbTab & ": " & xatm_PropertyRow.Kind, Empty

	MsgBox Join(d.Keys, vbCrLf), vbInformation, "Property Information"
End Sub

<xatm_PropertyRow.txtConfiguredValue:txtConfiguredValue_Validate(Cancel, NewValue)>
Sub txtConfiguredValue_Validate(Cancel, NewValue)

	' Nothing to send when the field comes back the way it was found.
	If CStr(NewValue) = CStr(xatm_PropertyRow.Value) Then Exit Sub

	Dim command
	Set command = Nothing

	On Error Resume Next
	Set command = Application.GetObject(SET_PROPERTY)
	On Error Goto 0

	Dim valueTag
	Set valueTag = Nothing

	On Error Resume Next
	Set valueTag = Application.GetObject(PROPERTY_VALUE)
	On Error Goto 0
		
	If command Is Nothing Or valueTag Is Nothing Then
		WriteLog "Not sent - " & SET_PROPERTY & " and " & PROPERTY_VALUE & " are what carry an edit."
		Cancel = True
		Exit Sub
	End If

	' The value first, so it is standing by when the command fires, and on
	' a tag of its own so nothing the operator typed has to be escaped.
	valueTag.WriteEx NewValue
	
	' kind|path|name - tokens of ours, none of which can carry a bar.
	command.WriteEx xatm_PropertyRow.Kind & "|" & _
	                xatm_PropertyRow.Source & "|" & _
	                NewValue

	If command.DocString <> EXIT_SUCCESS Then

		' The document would not take it and has said why on the console.
		' Cancelling leaves the operator in the field with what they typed,
		' rather than swallowing the edit.
		Cancel = True
		Exit Sub

	End If

	' What the field shows is what the document holds.
	xatm_PropertyRow.Value = NewValue
	
End Sub


' The command tag in xatm_config that owns the document. Reached by path
' because a library has no other way up to the project it runs under -
' and it is the project that knows what XML is, not this control.
Const SET_PROPERTY = "xatm_config_data.XML.SetProperty"
Const PROPERTY_VALUE = "xatm_config_data.XML.PropertyValue"
Const EXIT_SUCCESS = "EXIT_SUCCESS"


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & xatm_PropertyRow.Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & xatm_PropertyRow.Name & "] - " & message
	End If
	
End Sub

<xatm_PropertyRow.txtCurrentValue:txtCurrentValue_Validate(Cancel, NewValue)>
Sub txtCurrentValue_Validate(Cancel, NewValue)

	' Nothing to send when the field comes back the way it was found.
	If CStr(NewValue) = CStr(xatm_PropertyRow.Value) Then Exit Sub

	Dim command
	Set command = Nothing

	On Error Resume Next
	Set command = Application.GetObject(SET_PROPERTY)
	On Error Goto 0

	Dim valueTag
	Set valueTag = Nothing

	On Error Resume Next
	Set valueTag = Application.GetObject(PROPERTY_VALUE)
	On Error Goto 0
		
	If command Is Nothing Or valueTag Is Nothing Then
		WriteLog "Not sent - " & SET_PROPERTY & " and " & PROPERTY_VALUE & " are what carry an edit."
		Cancel = True
		Exit Sub
	End If

	' The value first, so it is standing by when the command fires, and on
	' a tag of its own so nothing the operator typed has to be escaped.
	valueTag.WriteEx NewValue
	
	' kind|path|name - tokens of ours, none of which can carry a bar.
	command.WriteEx xatm_PropertyRow.Kind & "|" & _
	                xatm_PropertyRow.Source & "|" & _
	                NewValue

	If command.DocString <> EXIT_SUCCESS Then

		' The document would not take it and has said why on the console.
		' Cancelling leaves the operator in the field with what they typed,
		' rather than swallowing the edit.
		Cancel = True
		Exit Sub

	End If

	' What the field shows is what the document holds.
	xatm_PropertyRow.Value = NewValue
	
End Sub


' The command tag in xatm_config that owns the document. Reached by path
' because a library has no other way up to the project it runs under -
' and it is the project that knows what XML is, not this control.
Const SET_PROPERTY = "xatm_config_data.XML.SetProperty"
Const PROPERTY_VALUE = "xatm_config_data.XML.PropertyValue"
Const EXIT_SUCCESS = "EXIT_SUCCESS"


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & xatm_PropertyRow.Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & xatm_PropertyRow.Name & "] - " & message
	End If
	
End Sub

<xatm_PropertyRow:xatm_PropertyRow_OnStartRunning()>
Sub xatm_PropertyRow_OnStartRunning()

	' The current column shows what E3 holds this moment, and the configured
	' column what the document says. Seeing the two differ is how an edit
	' that has not been saved yet makes itself visible.
	Dim linkTo
	linkTo = LiveSource()

	If linkTo = "" Then
		Item("txtCurrentValue").Value = ""
		Exit Sub
	End If

	Item("txtCurrentValue").Links.CreateLink "Value", linkTo

End Sub


' Where the current value is read from, "" when there is nothing to read.
'
' An IOTag is read through the tag it is associated with. Everything else
' is read off the object itself - which is what ObjectPath is for. Source
' is no use here: it is a key, id:700 for anything with an Id, so that a
' rename does not strand the row that was built before it.
'
' A property that points at another object has no value of its own to
' show, so it gets no link.
Function LiveSource()

	LiveSource = ""

	If LCase(PropertyType & "") = "iotag" Then

		If Trim(PropertySource & "") <> "" Then LiveSource = PropertySource & ".Value"
		Exit Function

	End If

	If LCase(Left(PropertyType & "", 5)) = "xatm_" Then Exit Function

	If Trim(ObjectPath & "") = "" Then Exit Function
	If Trim(PropertyName & "") = "" Then Exit Function

	LiveSource = ObjectPath & "." & PropertyName

End Function

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

<xatm_SelectLayout.objArea:objArea_OnStartRunning()>
Sub objArea_OnStartRunning()
	
	Visible = False
		
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

