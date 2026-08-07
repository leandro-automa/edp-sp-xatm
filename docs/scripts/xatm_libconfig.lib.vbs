-----------------------
Documentação de Scripts
-----------------------
XATM_LIBCONFIG (C:\ProjDev\edp_sp\xatm_libconfig.lib)
Fri Aug  7 14:18:57 2026
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

<xatm_PropertyRow.txtConfiguredValue:txtConfiguredValue_Validate(Cancel, NewValue)>
Sub txtConfiguredValue_Validate(Cancel, NewValue)

	' What was typed has to fit the type the manifest declares, or the
	' document would end up holding abcd where a breaker wants a timeout -
	' and the import would put that into the project on the next Save.
	Dim problem
	problem = TypeProblem(NewValue)

	If problem <> "" Then

		WriteLog "Not sent - " & problem
		Cancel = True
		Exit Sub

	End If

	' Spelled the way the export spells it, so the document does not end up
	' carrying true, TRUE and -1 for the same thing.
	Dim toSend
	toSend = Normalised(NewValue)

	' Nothing to send when the field comes back the way it was found -
	' compared after normalising, so retyping true over True is not an edit.
	If toSend = CStr(xatm_PropertyRow.Value & "") Then Exit Sub

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
	valueTag.WriteEx toSend

	' kind|key|name - tokens of ours, none of which can carry a bar. The
	' value is not among them: it went to PropertyValue above, whole and
	' unparsed.
	'
	' key rather than path: Source is id:700 for anything with an Id, so
	' that renaming an object does not strand the row that was built
	' before it.
	command.WriteEx xatm_PropertyRow.Kind & "|" & _
	                xatm_PropertyRow.Source & "|" & _
	                xatm_PropertyRow.PropertyName

	If command.DocString <> EXIT_SUCCESS Then

		' The document would not take it and has said why on the console.
		' Cancelling leaves the operator in the field with what they typed,
		' rather than swallowing the edit.
		Cancel = True
		Exit Sub

	End If

	' What the field shows is what the document holds.
	xatm_PropertyRow.Value = toSend

End Sub


' Why what was typed does not fit the declared type, "" when it does.
'
' An empty field always fits: it clears the property, which the document
' records by dropping the value attribute rather than writing an empty
' one. A type nothing is said about here - String, and the name row -
' takes whatever is typed, and SetName is what refuses a blank name or a
' duplicate.
Function TypeProblem(text)

	TypeProblem = ""

	Dim s
	s = Trim(CStr(text) & "")
	If s = "" Then Exit Function

	Select Case LCase(xatm_PropertyRow.PropertyType & "")

		Case "boolean"

			If Not IsBooleanText(s) Then
				TypeProblem = "'" & s & "' is not True or False."
			End If

		Case "integer", "edbswitchstate"

			If Not IsWholeNumber(s) Then
				TypeProblem = "'" & s & "' is not a whole number."
			End If

	End Select

End Function


' True and False, and the numbers that stand for them - zero is False and
' anything else is True, which is how a protocol carries a boolean.
Function IsBooleanText(s)

	IsBooleanText = True

	Select Case LCase(s)
		Case "true", "false" : Exit Function
	End Select

	If IsWholeNumber(s) Then Exit Function

	IsBooleanText = False

End Function


' Digits, with a sign in front if you like, and nothing else.
'
' Deliberately not IsNumeric, which says yes to 1e3 and to 12.5 - and to
' 12,5 or not, depending on which decimal separator the machine running
' E3 happens to use. None of these properties is ever fractional, so the
' plain test is both stricter and the same in every locale.
Function IsWholeNumber(s)

	IsWholeNumber = False

	Dim body
	body = s

	If Left(body, 1) = "-" Or Left(body, 1) = "+" Then body = Mid(body, 2)
	If body = "" Then Exit Function

	Dim i
	For i = 1 To Len(body)
		If InStr("0123456789", Mid(body, i, 1)) = 0 Then Exit Function
	Next

	IsWholeNumber = True

End Function


' What is actually sent. CLng and CStr are safe here because the text has
' already been checked digit by digit.
Function Normalised(text)

	Dim s
	s = Trim(CStr(text) & "")

	Normalised = s
	If s = "" Then Exit Function

	Select Case LCase(xatm_PropertyRow.PropertyType & "")

		Case "boolean"

			If IsWholeNumber(s) Then
				Normalised = CStr(CLng(s) <> 0)
			Else
				Normalised = CStr(LCase(s) = "true")
			End If

		Case "integer", "edbswitchstate"

			Normalised = CStr(CLng(s))

	End Select

End Function


' The command tag in xatm_config that owns the document. Reached by path
' because a library has no other way up to the project it runs under -
' and it is the project that knows what XML is, not this control.
Const SET_PROPERTY = "xatm_config_data.Config.SetProperty"
Const PROPERTY_VALUE = "xatm_config_data.Config.PropertyValue"
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

	' Two columns that mean different things. The current one shows what the
	' project holds this moment, read live; the configured one shows what
	' the document says it should hold. The two disagreeing is an edit that
	' has been staged and not saved.

	' --- current: linked, because it has to follow the plant ------------
	Dim linkTo
	linkTo = LiveSource()

	If linkTo = "" Then
		Item("txtCurrentValue").Value = ""
	Else
		Item("txtCurrentValue").Links.CreateLink "Value", linkTo
	End If

	' Never typed into. The current column is the plant, and the only thing
	' that writes straight to it is a force.
	Item("txtCurrentValue").IsSetPoint = False

	' --- configured: handed over, because it came with the row ---------
	'
	' No link: this is what the export read out of Studio and what the
	' operator has staged since, and the row was given it. Linking it live
	' would put the project's value back over a staged edit the moment it
	' was typed.
	Item("txtConfiguredValue").Value = ConfiguredText()

	' And it takes typing only where the manifest says the value may be
	' typed - Id and a command output are shown and never edited.
	Item("txtConfiguredValue").IsSetPoint = Can(EXPOSE_EDIT)

	' --- configured source: what it is wired to, rather than what it holds
	'
	' A path or an expression is neither of the other two columns. It is not
	' the plant's reading and it is not a value that was typed - it is where
	' the value comes from, which a breaker's twelve tags and a BTC's
	' Transformer both have and a CommandTimeout does not.
	Item("txtConfiguredSource").Value = ConfiguredSource()

	' Read only. A tag or an object is picked, not spelled out, and it is
	' the picker that will set this.
	Item("txtConfiguredSource").IsSetPoint = False

End Sub


' The numbers behind the Exposure the row was built with. Written out
' rather than shared: the flags are declared in the manifests over in
' xatm_config, and one E3 object cannot read another's constants - which
' is also why the Exposure property documents them in its own help.
Const EXPOSE_VIEW       = 1
Const EXPOSE_VALUE      = 2
Const EXPOSE_EDIT       = 4
Const EXPOSE_EXPRESSION = 8
Const EXPOSE_FORCE      = 16
Const EXPOSE_SAVED      = 32

Const MAX_SOURCE_LENGTH = 40

' Whether the manifest allowed this. Empty And anything is 0, so a row
' built without an Exposure allows nothing.
Function Can(flag)

	Can = ((Exposure And flag) <> 0)

End Function


' What the configured column shows. A property the document carries no
' value for is unset rather than blank, and either way there is nothing
' to put in the field.
'
' An association is left out of it: it has a column of its own, and the
' path of the transformer a BTC drives reads badly in a field sized for
' a timeout.
Function ConfiguredText()

	ConfiguredText = ""

	If IsObjectProperty() Then Exit Function

	If IsEmpty(Value) Or IsNull(Value) Then Exit Function

	ConfiguredText = CStr(Value)

End Function


' Where this property's value comes from, or what it points at - three
' shapes that all answer the same question, so they share a column.
'
' An IOTag names the tag it is wired to. A bound property carries the
' expression somebody wrote for it. A property typed xatm_ holds the path
' of another object, and that one arrives in Value rather than in
' PropertySource - a wrinkle in how the document is written, not one
' worth putting on the screen.
Function ConfiguredSource()

	ConfiguredSource = ""

	If IsObjectProperty() Then

		If Not (IsEmpty(Value) Or IsNull(Value)) Then ConfiguredSource = CStr(Value)
		Exit Function

	End If

	If IsEmpty(PropertySource) Or IsNull(PropertySource) Then Exit Function
	
	If Len(PropertySource) > MAX_SOURCE_LENGTH Then
		ConfiguredSource = Left(PropertySource, MAX_SOURCE_LENGTH - 3) & "..."
	Else
		ConfiguredSource = CStr(PropertySource)
	End If
	
End Function


' A property that points at another object rather than holding a value of
' its own. The manifest names the class, as xatm_Transformer.
Function IsObjectProperty()

	IsObjectProperty = (LCase(Left(PropertyType & "", 5)) = "xatm_")

End Function


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

Sub Foo()

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

