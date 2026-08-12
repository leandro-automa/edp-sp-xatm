-----------------------
Documentação de Scripts
-----------------------
XATM_LIBCONFIG (C:\ProjDev\edp_sp\xatm_libconfig.lib)
Wed Aug 12 11:18:51 2026
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

<xatm_PropertyRow.btnPickTag:btnPickTag_Click()>
Sub btnPickTag_Click()

	Const BROWSER_SCREEN = "xatm_config_screens.DomainBrowser"
	Const BROWSER_TITLE  = "Select a tag for "
	Const BROWSER_LEFT   = 0
	Const BROWSER_TOP    = 0
	Const BROWSER_WIDTH  = 510
	Const BROWSER_HEIGHT = 558
	Dim FLAGS : FLAGS = 1 + 2 + 16 + 64 + 256 + 2048


	' Where the browser opens: the tag this property is wired to now, and
	' "" where it is wired to nothing. It travels as the Arg, which is
	' what the screen reads in its OnPreShow.
	Dim current
	current = Trim(xatm_PropertyRow.PropertySource & "")

	' Modal, so nothing else on the panel can be touched while a tag is
	' being chosen, and so the choice comes straight back here - Close
	' hands its argument to DoModal as its return.
	
	
	Application.DoModal	BROWSER_SCREEN, _
	                     BROWSER_TITLE & xatm_PropertyRow.PropertyName, _
	                     , , _
	                     BROWSER_WIDTH, BROWSER_HEIGHT, _
	                     current, FLAGS
	
	
	Dim picked
	picked = Application.GetObject("xatm_config_data.Catalog.SelectedPath").Value
	' Nothing was chosen. Both shapes are asked about: Cancel closes with
	' "", and the X on the window closes without a Close at all, which
	' arrives here as Empty.
	If IsEmpty(picked) Or IsNull(picked) Then Exit Sub

	picked = Trim(CStr(picked))
	If picked = "" Then Exit Sub

	' The tag it already holds, chosen again, is not an edit.
	If StrComp(picked, current, vbTextCompare) = 0 Then Exit Sub

	' The object this row belongs to, and the tag that was chosen. Both
	' paths come from PathName - the row's out of the document, the tag's
	' out of the catalog - so both already resolve, bracketed where they
	' had to be, and neither is bracketed again here.
	Dim obj
	Set obj = Nothing

	On Error Resume Next
	Set obj = Application.GetObject(xatm_PropertyRow.ObjectPath)
	On Error Goto 0

	If obj Is Nothing Then
		MsgBox "Not wired - the project has nothing at " & _
		       xatm_PropertyRow.ObjectPath & "."
		Exit Sub
	End If

	Set gPickedTag = Nothing

	On Error Resume Next
	Set gPickedTag = Application.GetObject(picked)
	On Error Goto 0

	If gPickedTag Is Nothing Then
		MsgBox "Not wired - the project has nothing at " & picked & "."
		Exit Sub
	End If

	' An IOTag property holds the tag object itself and not its path, so
	' the tag is associated rather than assigned - a Set, and through
	' Execute because the property is named at runtime. gPickedTag is a
	' cell of its own for the same reason the import keeps one: what
	' Execute runs is a line of text, and an object cannot be spelled out
	' in it.
	Dim failed
	failed = ""

	On Error Resume Next

	Execute "obj." & xatm_PropertyRow.PropertyName & " = gPickedTag.PathName"

	If Err.Number <> 0 Then
		failed = Err.Description
		Err.Clear
	End If

	On Error Goto 0

	If failed <> "" Then
		MsgBox "Not wired - " & xatm_PropertyRow.PropertyName & " would not take " & _
		       picked & " - " & failed
		Exit Sub
	End If

	' The document has to be told as well. The write above went into the
	' project, and the panel is built from the document - so left alone
	' the row comes back wired to the old tag the next time it is built,
	' and the file it is saved to keeps the old one too.
	SendSource picked

	' What the row holds, and what its source column shows, now that the
	' project holds it too.
	xatm_PropertyRow.PropertySource = picked

	On Error Resume Next
	xatm_PropertyRow.Item("txtConfiguredSource").Value = SourceText(picked)
	On Error Goto 0

End Sub


' The command tag in xatm_config that owns the document, and the kind of
' edit this button sends. Written out again here because one E3 object
' cannot call another's, and a library has no other way up to the project
' it runs under.
Const SET_PROPERTY   = "xatm_config_data.Config.SetProperty"
Const PROPERTY_VALUE = "xatm_config_data.Config.PropertyValue"
Const KIND_SOURCE    = "source"
Const EXIT_SUCCESS   = "EXIT_SUCCESS"


' Scratch cell for the association written through Execute. The property
' is named at runtime, so the line is built as text, and an object has to
' be waiting in a variable for it to name.
Dim gPickedTag


' How much of a path the source column has room for - the same cut the
' row makes for itself when it is built.
Const MAX_SOURCE_LENGTH = 128

Function SourceText(path)

	If Len(path) > MAX_SOURCE_LENGTH Then
		SourceText = Left(path, MAX_SOURCE_LENGTH - 3) & "..."
	Else
		SourceText = CStr(path)
	End If

End Function


' Stages the association in the document, on the two tags every other
' edit on this panel travels on: the path on one, and the command that
' reads it on the other.
'
' Said out loud when it does not go through. The project already holds
' the tag by this point, so a document that was not told is the two
' disagreeing - which is worth knowing about rather than finding out at
' the next rebuild.
Sub SendSource(path)

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
		MsgBox "The project holds the tag, but the document was not told - " & _
		       SET_PROPERTY & " and " & PROPERTY_VALUE & " are what carry an edit."
		Exit Sub
	End If

	' The path first, so it is standing by when the command fires, and on
	' a tag of its own so nothing in it has to be escaped or split.
	valueTag.WriteEx path

	' kind|key|name. key rather than path: Source is id:700 for anything
	' with an Id, so renaming an object does not strand the row.
	command.WriteEx KIND_SOURCE & "|" & _
	                xatm_PropertyRow.Source & "|" & _
	                xatm_PropertyRow.PropertyName

	If command.DocString <> EXIT_SUCCESS Then
		MsgBox "The project holds the tag, but the document would not take it. " & _
		       "The console says why."
	End If

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
	
	' --- the two pickers ------------------------------------------------
	'
	' Neither is typed into, which is why each needs a button of its own.
	' An IOTag names a tag in the project, and a tag is picked from what
	' is there rather than spelled out. A property the manifest lets be
	' bound carries an expression, which is written on a screen with room
	' for it.
	'
	' A row that is neither - a timeout, an Id - shows no button, and the
	' two are exclusive: no property is both an IOTag and bindable.
	ShowItem "btnPickTag", PicksTag()
	ShowItem "btnBuildExpression", Can(EXPOSE_EXPRESSION)

End Sub


' The numbers behind the Exposure the row was built with. Written out
' rather than shared: the flags are declared in the manifests over in
' xatm_config, and one E3 object cannot read another's constants - which
' is also why the Exposure property documents them in its own help.
Const EXPOSE_VIEW       = 1     ' a row appears for it
Const EXPOSE_VALUE      = 2     ' its value is shown - never for a write-only command
Const EXPOSE_EDIT       = 4     ' the value can be typed
Const EXPOSE_EXPRESSION = 8     ' it can be bound to an expression
Const EXPOSE_FORCE      = 16    ' it can be forced at runtime, and is never saved
Const EXPOSE_SAVED      = 32    ' its value is configuration, not a reading
Const EXPOSE_INTERFACE  = 64    ' the Elipse application is given a tag for it

Const MAX_SOURCE_LENGTH = 128

' Whether the manifest allowed this. Empty And anything is 0, so a row
' built without an Exposure allows nothing.
Function Can(flag)

	Can = ((Exposure And flag) <> 0)

End Function

' Whether this row wires a tag, and so wants the picker.
'
' EXPOSE_EDIT is not the question, though it reads like it: that flag
' says the value field may be typed into, which is the one thing an
' tag row does not do - what is configured is the tag it is associated
' with, not a value - and no manifest gives it to one. The type is what
' decides, and EXPOSE_SAVED that the association is configuration rather
' than something read off the plant.
Function PicksTag()

	PicksTag = False

	If Not IsTagProperty() Then Exit Function

	PicksTag = Can(EXPOSE_SAVED)

End Function



' Shows or hides one of the row's objects, and says nothing when the
' control has not been given it - so a row still builds while a button
' is only half added to it in Studio.
Sub ShowItem(itemName, shown)

	On Error Resume Next
	Item(itemName).Visible = shown
	On Error Goto 0

End Sub

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
	
	' A tag property has no value of its own to configure - the tag it is
	' wired to is the whole of it, and that belongs in the source column.
	' Asked here rather than left to the document, so a row built from an
	' export written before this was understood shows a blank field and
	' not a path in the wrong column.
	If IsTagProperty() Then Exit Function
	
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


' A property configured by the tag it is wired to rather than by a value
' of its own. An IOTag names a tag out in the project, an InternalTag one
' of the object's own - a breaker's trip output and a BTC's start command
' - and the row shows the two alike, because what there is to show of
' either is which tag it points at.
Function IsTagProperty()

	IsTagProperty = False

	Select Case LCase(PropertyType & "")
		Case "iotag", "internaltag" : IsTagProperty = True
	End Select

End Function


' Where the current value is read from, "" when there is nothing to read.
'
' A command is written and never read. The manifest says so by withholding
' EXPOSE_VALUE - a breaker's trip output and its select tag both - and
' linking one anyway reads back whatever was last written to the tag, which
' the current column would then show as a state the plant is in. Nothing is
' linked for it, and the column stays empty.
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

	If Not Can(EXPOSE_VALUE) Then Exit Function

	If IsTagProperty() Then

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

