-----------------------
Documentação de Scripts
-----------------------
XATM_CONFIG (C:\ProjDev\edp_sp\xatm_config.prj)
Mon Aug 31 17:10:01 2026
-----------------------

<xatm_config_data.Catalog.XMLBuilderAfterDelay:XMLBuilderAfterDelay_Functions()>
Sub XMLBuilderAfterDelay_Functions()
End Sub

' The whole domain, walked and left on one tag.
'
' Nothing is written to disk. It used to name a file in C:\temp and save
' to it, and that path outlived its purpose - the tree reads the tag, the
' file was only ever there to be looked at, and a second copy of the
' catalogue is a second thing that can be out of date.
Sub ExportElipseCatalog()

	Dim doc
	Set doc = CreateObject("MSXML2.DOMDocument.6.0")
	doc.async = False
	doc.preserveWhiteSpace = True

	Dim pi
	Set pi = doc.createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
	doc.appendChild pi

	Dim root
	Set root = doc.createElement("ElipseCatalog")
	root.setAttribute "source", "Application"
	root.setAttribute "generatedAt", IsoNow()
	root.setAttribute "generator", Me.PathName
	doc.appendChild root

	Application.Trace "Elipse catalog export started."

	WalkEnumerable doc, root, Application, "Application", 0

	' Parent is Catalog, so this is xatm_config_data.Catalog.XMLContent -
	' the one tag the browser reads and the only place the catalogue lives.
	Parent.Item("XMLContent").WriteEx doc.xml

	Application.Trace "Elipse catalog export finished."

End Sub

Sub WalkEnumerable(doc, parentNode, container, containerPath, depth)

	Dim sawChild
	sawChild = False

	On Error Resume Next
	Dim child
	For Each child In container
		If Err.Number <> 0 Then
			AddError doc, parentNode, containerPath, "enumerate", Err.Number, Err.Description
			Err.Clear
			Exit For
		End If

		sawChild = True
		AddObjectTree doc, parentNode, child, containerPath, depth + 1

		If Err.Number <> 0 Then
			AddError doc, parentNode, containerPath, "walk-child", Err.Number, Err.Description
			Err.Clear
		End If
	Next

	If Err.Number <> 0 Then
		AddError doc, parentNode, containerPath, "enumerate", Err.Number, Err.Description
		Err.Clear
	End If
	On Error GoTo 0

	If Not sawChild Then
		WalkByItemFallback doc, parentNode, container, containerPath, depth
	End If

End Sub

Sub AddObjectTree(doc, parentNode, obj, parentPath, depth)

	Dim objType
	objType = TypeName(obj)

	Dim objName
	objName = SafeRead(obj, "Name")

	Dim objPath
	objPath = SafeRead(obj, "PathName")

	Dim childCount
	childCount = SafeRead(obj, "Count")

	If Len(objPath) = 0 Then
		If Len(objName) > 0 Then
			objPath = parentPath & "." & objName
		Else
			objPath = parentPath & ".[" & objType & "]"
		End If
	End If

	Dim node
	Set node = doc.createElement("Object")
	node.setAttribute "type", objType
	node.setAttribute "name", objName
	node.setAttribute "path", objPath
	node.setAttribute "depth", CStr(depth)
	node.setAttribute "count", childCount
	parentNode.appendChild node

	AddCommonProperties doc, node, obj
	WalkEnumerable doc, node, obj, objPath, depth

End Sub

Sub AddCommonProperties(doc, objectNode, obj)

	Dim propsNode
	Set propsNode = doc.createElement("Properties")
	objectNode.appendChild propsNode

	AddProperty doc, propsNode, obj, "DocString"

End Sub

Sub AddProperty(doc, propsNode, obj, propertyName)

	Dim value
	value = SafeRead(obj, propertyName)

	Dim propNode
	Set propNode = doc.createElement("Property")
	propNode.setAttribute "name", propertyName
	propNode.setAttribute "value", value
	propsNode.appendChild propNode

End Sub

Sub WalkByItemFallback(doc, parentNode, container, containerPath, depth)

	Dim total
	total = SafeRead(container, "Count")
	If Not IsNumeric(total) Then Exit Sub
	If CLng(total) < 1 Then Exit Sub

	Dim i
	For i = 1 To CLng(total)
		On Error Resume Next
		Dim child
		Set child = container.Item(i)
		If Err.Number <> 0 Then
			AddError doc, parentNode, containerPath, "Item(" & CStr(i) & ")", Err.Number, Err.Description
			Err.Clear
		Else
			AddObjectTree doc, parentNode, child, containerPath, depth + 1
		End If
		Set child = Nothing
		On Error GoTo 0
	Next

End Sub

Function SafeRead(obj, propertyName)

	Dim value
	value = ""

	On Error Resume Next
	Select Case propertyName
		Case "Name"
			value = obj.Name
		Case "PathName"
			value = obj.PathName
		Case "Count"
			value = obj.Count
		Case "DocString"
			value = obj.DocString
		Case Else
			value = ""
	End Select

	If Err.Number <> 0 Then
		value = ""
		Err.Clear
	End If
	On Error GoTo 0

	SafeRead = CStr(value)

End Function

Sub AddError(doc, parentNode, objectPath, operation, number, description)

	Dim errNode
	Set errNode = doc.createElement("Error")
	errNode.setAttribute "path", objectPath
	errNode.setAttribute "operation", operation
	errNode.setAttribute "number", CStr(number)
	errNode.setAttribute "description", description
	parentNode.appendChild errNode

End Sub

Function IsoNow()

	Dim dt
	dt = Now

	IsoNow = Year(dt) & "-" & Pad2(Month(dt)) & "-" & Pad2(Day(dt)) & _
			 "T" & Pad2(Hour(dt)) & ":" & Pad2(Minute(dt)) & ":" & Pad2(Second(dt))

End Function

Function Pad2(value)

	If value < 10 Then
		Pad2 = "0" & CStr(value)
	Else
		Pad2 = CStr(value)
	End If

End Function

Sub EndOfScript()
		
End Sub

<xatm_config_data.Catalog.XMLBuilderAfterDelay:XMLBuilderAfterDelay_OnStartRunning()>
Sub XMLBuilderAfterDelay_OnStartRunning()

	Value = 3	' delay in seconds
		
End Sub

<xatm_config_data.Catalog.XMLBuilderAfterDelay:XMLBuilderAfterDelay_ThickCountdown()>
Sub XMLBuilderAfterDelay_ThickCountdown()
	
	If Value > 0 Then
			
		Value = Value - 1
	
	Else
		
		' stop
		Value = -1
		
		ExportElipseCatalog
	
	End If
		
End Sub

<xatm_config_data.Config.BuildAlarms:BuildAlarms_OnChangedValue()>
Sub BuildAlarms_OnChangedValue()

	If IsEmpty(Value) Or IsNull(Value) Then Exit Sub
	If Trim(CStr(Value)) = "" Then Exit Sub

	Dim ts
	ts = TimeStamp                         ' preserved for the silent clear

	Dim substationPath
	substationPath = Trim(CStr(Value))

	Dim substation
	Set substation = Nothing
	On Error Resume Next
	Set substation = Application.GetObject(substationPath)
	On Error Goto 0

	If substation Is Nothing Then
		Fail "there is nothing at " & substationPath & ".", ts
		Exit Sub
	End If

	' Where the alarms bind. Read from the project rather than assumed,
	' because the interface sits at the root of the data project in some
	' layouts and inside XATM_Data in others - and an alarm bound to a
	' path that is merely plausible is an alarm that never raises.
	Dim interfaceServer
	Set interfaceServer = InterfaceRoot()

	If interfaceServer Is Nothing Then
		Fail "the project has no interface data server, so there is " & _
		     "nothing for the alarms to bind to.", ts
		Exit Sub
	End If

	gInterfaceRoot = interfaceServer.PathName

	' The folder the alarms are built in - emptied and made again rather
	' than brought up to date, the way the interface is. A device can be
	' renamed on the Config screen, and a folder named after the name it
	' used to have cannot be told from one nobody wants any more.
	Dim folder
	Set folder = ChildFolder(substation, AUTOMATION_FOLDER)

	If folder Is Nothing Then
		Fail "the substation would not take a " & AUTOMATION_FOLDER & " folder.", ts
		Exit Sub
	End If

	ClearFolder folder

	gFolders  = 0
	gAlarms   = 0
	gCommands = 0

	Dim problem
	problem = ""

	AlarmFolder folder, Application.GetObject(DATA_ROOT), "", problem

	' Written here rather than left to the Save that persists XATM_Data:
	' these objects are in the substation's container, which is not the
	' container holding the automation.
	'
	' Save answers True or False and raises nothing, so the answer is what
	' decides. An On Error still wraps the call for a container that will
	' not give one at all, and saved starts False so that a call which
	' never returns is a failure rather than a silence.
	Dim saved
	saved = False

	On Error Resume Next
	Err.Clear

	saved = folder.Context("Container").Save()

	If Err.Number <> 0 Then Err.Clear
	On Error Goto 0

	If Not saved Then
		problem = problem & vbCrLf & "  not saved - E3 would not write the " & _
		          "substation's project. A licence that will not carry this " & _
		          "many tags is the usual reason."
	End If

	' Said out loud even when it all went through, for the same reason
	' the interface rebuild says it: a routine that mentions only what it
	' could not build reads the same whether it ran or was never called.
	If gAlarms = 0 And gCommands = 0 Then
		WriteLog "Automation - no alarms and no commands are interfaced."
	Else
		WriteLog "Automation - " & gAlarms & " alarms and " & gCommands & _
		         " commands on " & gFolders & " objects, under " & _
		         substationPath & "." & AUTOMATION_FOLDER & "."
	End If

	If problem <> "" Then
		WriteLog "Automation is incomplete:" & problem
		DocString = EXIT_FAILURE
		WriteEx Empty, ts
		Exit Sub
	End If

	DocString = EXIT_SUCCESS
	WriteEx Empty, ts                      ' clear without re-firing

End Sub


' The folder built under the chosen PowerSubstation. Named for what it
' holds rather than for the alarms alone: the substation side both reads
' the alarms in it and writes the commands beside them.
Const AUTOMATION_FOLDER  = "Automatismos"

' What the folder is made of. The first two carry a reading and a
' command alike; the last two are the alarm's own, and a command tag
' gets neither.
Const AUTOMATION_FOLDER_CLASS = "DataFolder"
Const AUTOMATION_TAG_CLASS    = "InternalTag"
Const ALARM_SOURCE_CLASS      = "DigitalAlarmSource"
Const ALARM_SOURCE_NAME       = "Alarm"

' What marks a tag as this configuration's to remake - the same mark the
' distribution puts on its IOTags and the interface rebuild puts on its
' own. The folder is emptied and built again on every run, so nothing here
' reads the mark yet; it is written so that a tag under the substation can
' be told from one an engineer put there by hand.
Const SELF_MARK               = "$XATM$"

' How a link carries the value - the BindType argument of CreateLink,
' which is a Simple Link when it is left off.
'
'   simple    source -> property   the tag follows the interface
'   reverse   property -> source   the interface follows the tag
'
' Both are one-way, and both live on the substation tag, so which one is
' asked for is the whole of the difference between a reading and a
' command. The link is not moved to the other end for a command - it
' would be wiped by the next interface rebuild if it were.
Const LINK_SIMPLE  = 0
Const LINK_REVERSE = 5

Const DATA_ROOT          = "XATM_Data"
Const CONFIG_DATA        = "xatm_config_data"
Const HELPER_FOLDER      = "PropertiesHelper"

Const INTERFACE_ROOT     = "XATM_Interface"
Const INTERFACE_ROOT_ALT = "XATM_Data.XATM_Interface"

Const EXIT_SUCCESS       = "EXIT_SUCCESS"
Const EXIT_FAILURE       = "EXIT_FAILURE"

' Anything this severe or worse has to be acknowledged. A status the
' control room is merely told about - a sequence being in progress - is
' not something to make an operator clear off the list by hand.
'
' Tested with <=, because Power's severity scale runs backwards:
' critical is -2 and low is 2, so "this severe or worse" is the smaller
' number. Written out as a literal because the SEV_ constants belong to
' the manifests, and this scope cannot see them.
Const ACK_FROM_SEVERITY  = 1     ' medium

' The area an alarm is filed under, written onto the source rather than
' left to the folder it sits in. Power gives an alarm the area of its
' place in the tree, and this tree is shaped for the data project, not
' for the way a control room reads its list.
'
'   a device       its own name - the name of the folder holding it
'   an automation  XATM, the whole of the automation as one area
'
' The automation classes are the ones named, and everything else falls
' to its own name, because a layout gains devices - another bay,
' another transformer - and a class nobody thought of when this was
' written should read as the equipment it is rather than be filed
' under the automation.
Const AUTOMATION_AREA    = "XATM"

' Where the interface was found this run, as the head of every path the
' alarms bind to.
Dim gInterfaceRoot

' What the last build came out at, for the line it writes to the log.
Dim gFolders
Dim gAlarms
Dim gCommands


' One folder of the data project: an alarm for every object in it a
' manifest speaks for, and then the same for whatever is under it.
'
' The walk the interface rebuild does, over the same tree and reading
' the same manifests - so the folders come out shaped alike and an
' alarm is found where the tree says its device is.
Sub AlarmFolder(root, folder, path, problem)

	Dim item, bag

	For Each item In folder

		Set bag = ManifestOf(TypeName(item))

		If Not bag Is Nothing Then

			AlarmObject root, path, item, bag, problem

		Else

			On Error Resume Next
			AlarmFolder root, item, ChildPath(path, item.Name), problem
			On Error Goto 0

		End If

	Next

End Sub


' One name onto a path, and the first name is the whole of it.
Function ChildPath(path, objectName)

	If path = "" Then
		ChildPath = objectName
	Else
		ChildPath = path & "." & objectName
	End If

End Function


' The folder at a path under the alarm root, each level made where it is
' not already there. An object with nothing on the alarm table gets no
' folder, rather than an empty one to wonder about.
Function EnsureFolder(root, path, problem)

	Set EnsureFolder = root
	If path = "" Then Exit Function

	Dim names
	names = Split(path, ".")

	Dim parent, child, i
	Set parent = root

	For i = 0 To UBound(names)

		Set child = Nothing

		On Error Resume Next
		Set child = parent.Item(names(i))
		On Error Goto 0

		If child Is Nothing Then Set child = NewChild(parent, AUTOMATION_FOLDER_CLASS, names(i))

		If child Is Nothing Then
			problem = problem & vbCrLf & "  no folder for " & path
			Set EnsureFolder = Nothing
			Exit Function
		End If

		Set parent = child

	Next

	Set EnsureFolder = parent

End Function


' One object's folder, and in it a tag for each property worth carrying:
' an alarm for a reading the manifest puts on the alarm table, and a
' plain tag for each command the interface holds. An object with neither
' gets no folder at all.
Sub AlarmObject(root, path, obj, bag, problem)

	Dim key, p

	Dim wanted
	wanted = False

	For Each key In bag.Keys
		If bag(key).IsAlarmed() Then wanted = True
		If IsCommand(bag(key)) Then wanted = True
	Next

	If Not wanted Then Exit Sub

	Dim parent
	Set parent = EnsureFolder(root, path, problem)
	If parent Is Nothing Then Exit Sub

	Dim folder
	Set folder = NewChild(parent, AUTOMATION_FOLDER_CLASS, obj.Name)

	If folder Is Nothing Then
		problem = problem & vbCrLf & "  no folder for " & obj.Name
		Exit Sub
	End If

	gFolders = gFolders + 1

	' Where this object's readings are on the interface. The alarms bind
	' there and never to the XObject: the interface is where the Elipse
	' application meets the automation, and an alarm is the application
	' asking to be told about a reading.
	Dim interfacePath
	interfacePath = gInterfaceRoot & "." & ChildPath(path, obj.Name)

	For Each key In bag.Keys

		Set p = bag(key)

		If p.IsAlarmed() Then
			AlarmTag folder, interfacePath, obj, p, problem
		ElseIf IsCommand(p) Then
			CommandTag folder, interfacePath, obj, p, problem
		End If

	Next

End Sub


' A command the interface carries: an InternalTag property a manifest
' interfaces. RebuildInterface leaves these unlinked and points the
' XObject's property at them, so they are the automation's inputs - the
' substation side writes them to ask for a maneuver.
Function IsCommand(p)

	IsCommand = (LCase(p.DataType & "") = "internaltag") And p.IsInterfaced()

End Function


' The tag one property gets in the substation's folder, wired to the
' interface tag standing for it. Nothing when it could not be made.
'
' A link carries a value one way, so the direction is the whole of the
' difference between the two kinds:
'
'   reading   automation -> interface -> substation, on a simple link
'   command   substation -> interface -> automation, on a reverse one
'
' A command left on a simple link is a command the automation never
' sees: the operator writes a tag that is downstream of the thing it was
' meant to drive, and the write is overwritten by the value it was
' supposed to change. Nothing errors, and the maneuver never starts.
Function BoundTag(folder, interfacePath, obj, p, problem, linkType)

	Set BoundTag = Nothing

	Dim source
	source = interfacePath & "." & p.Name

	' Checked before it is wired. The interface is emptied and rebuilt on
	' every Save, so tags built before that rebuild would wire to ones on
	' their way out - and a link to a tag that is not there fails quietly,
	' leaving an alarm that never raises or a command that goes nowhere.
	' Better said here than handed over looking like it works.
	Dim target
	Set target = Nothing
	On Error Resume Next
	Set target = Application.GetObject(source)
	On Error Goto 0

	If target Is Nothing Then
		problem = problem & vbCrLf & "  " & obj.Name & "." & p.Name & _
		          " has no interface tag at " & source
		Exit Function
	End If

	Dim tag
	Set tag = NewChild(folder, AUTOMATION_TAG_CLASS, p.Name)

	If tag Is Nothing Then
		problem = problem & vbCrLf & "  no tag for " & obj.Name & "." & p.Name
		Exit Function
	End If

	Dim where
	where = obj.Name & "." & p.Name

	SetMember tag, "DocString", SELF_MARK, where, problem

	On Error Resume Next
	Err.Clear

	tag.Links.CreateLink "Value", source, linkType

	If Err.Number <> 0 Then
		problem = problem & vbCrLf & "  " & obj.Name & "." & p.Name & _
		          " would not wire to " & source & " - " & Err.Description
		Err.Clear
	End If

	On Error Goto 0

	Set BoundTag = tag

End Function


' The alarm area for one object: the automation's own area for the
' objects that are the automation, and the device's name for the rest.
'
' Compared in lower case, because a class is named once in the manifest
' folder and TypeName answers with whatever case it was declared in.
Function AreaOf(obj)

	Select Case LCase(TypeName(obj))

		Case "xatm_tmtnm", "xatm_ta", "xatm_raseat"
			AreaOf = AUTOMATION_AREA

		Case Else
			AreaOf = obj.Name

	End Select

End Function

' One alarm for one reading: the bound tag, and the source on it that
' turns a change into an event.
Sub AlarmTag(folder, interfacePath, obj, p, problem)

	Dim tag
	Set tag = BoundTag(folder, interfacePath, obj, p, problem, LINK_SIMPLE)
	If tag Is Nothing Then Exit Sub

	Dim alarm
	Set alarm = NewChild(tag, ALARM_SOURCE_CLASS, ALARM_SOURCE_NAME)

	If alarm Is Nothing Then
		problem = problem & vbCrLf & "  " & obj.Name & "." & p.Name & _
		          " would not take a " & ALARM_SOURCE_CLASS
		Exit Sub
	End If

	' Each written on its own and reported on its own. What this version
	' of Power accepts is what decides the shape of an alarm, and a
	' property it does not know should cost that one setting rather than
	' the whole alarm.
	Dim where
	where = obj.Name & "." & p.Name
	
	SetMember alarm, "Digital",                  True,                where, problem
	SetMember alarm, "DigitalLimit",             p.AlarmLimit(),      where, problem
	SetMember alarm, "DigitalMessageText",       p.AlarmActiveText(), where, problem
	SetMember alarm, "DigitalReturnMessageText", p.AlarmNormalText(), where, problem
	SetMember alarm, "DigitalSeverity",          p.AlarmSeverity,     where, problem
	SetMember alarm, "AreaNameOverride",         AreaOf(obj),         where, problem
	SetMember alarm, "DigitalAckRequired", _
	          				  (p.AlarmSeverity <= ACK_FROM_SEVERITY), where, problem
	

	gAlarms = gAlarms + 1

End Sub


' One command for the substation side to write, on a reverse link so
' that writing it reaches the automation.
'
' No alarm source: a command is what the operator's screen asks for, not
' something the control room is told about, and an event per request
' would put the asking on the list as though the switchyard had done it.
Sub CommandTag(folder, interfacePath, obj, p, problem)

	Dim tag
	Set tag = BoundTag(folder, interfacePath, obj, p, problem, LINK_REVERSE)
	If tag Is Nothing Then Exit Sub

	gCommands = gCommands + 1

End Sub

' One setting onto the alarm source, late-bound the way the import
' writes a property, and named in the log when it will not take.
Sub SetMember(alarm, memberName, newValue, where, problem)

	gMemberValue = newValue

	On Error Resume Next
	Err.Clear

	Execute "alarm." & memberName & " = gMemberValue"

	If Err.Number <> 0 Then
		problem = problem & vbCrLf & "  " & where & " would not take " & _
		          memberName & " - " & Err.Description
		Err.Clear
	End If

	On Error Goto 0

End Sub

' Scratch cell for the late-bound write done through Execute, the way
' the interface rebuild keeps one.
Dim gMemberValue


' A child of a container by name, made where it is not already there.
Function ChildFolder(parent, folderName)

	Set ChildFolder = Nothing

	On Error Resume Next
	Set ChildFolder = parent.Item(folderName)
	On Error Goto 0

	If Not ChildFolder Is Nothing Then Exit Function

	Set ChildFolder = NewChild(parent, AUTOMATION_FOLDER_CLASS, folderName)

End Function


' Empties a folder. Taken out in a second pass, because a collection is
' not walked while what is in it is being removed.
Sub ClearFolder(folder)

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim item
	For Each item In folder
		doomed.Add doomed.Count, item.Name
	Next

	Dim n
	For Each n In doomed.Keys

		On Error Resume Next
		folder.DeleteObject doomed(n)
		On Error Goto 0

	Next

End Sub


' A new child of a container, Nothing when it would not take one -
' AddObject(ClassName, [Activate], [ObjectName]).
Function NewChild(parent, className, objectName)

	Set NewChild = Nothing

	On Error Resume Next
	Set NewChild = parent.AddObject(className, True, objectName)
	On Error Goto 0

End Function


' The data server the interface was built in. Named at the root of the
' data project, and looked for inside XATM_Data as well, because which
' of the two it sits under is a matter of how the project was laid out.
Function InterfaceRoot()

	Set InterfaceRoot = Nothing

	On Error Resume Next
	Set InterfaceRoot = Application.GetObject(INTERFACE_ROOT)
	On Error Goto 0

	If Not InterfaceRoot Is Nothing Then Exit Function

	On Error Resume Next
	Set InterfaceRoot = Application.GetObject(INTERFACE_ROOT_ALT)
	On Error Goto 0

End Function


' The manifest declared for a class, or Nothing when the class has none -
' the same lookup the export, the import and the interface do.
Function ManifestOf(className)

	Set ManifestOf = Nothing

	On Error Resume Next
	Set ManifestOf = Application.GetObject(CONFIG_DATA).Item(HELPER_FOLDER).Item(className).Value
	On Error Goto 0

End Function


' Said, logged, and the command cleared, for a build that got nowhere.
Sub Fail(message, ts)

	DocString = EXIT_FAILURE
	WriteLog message
	WriteEx Empty, ts

End Sub


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Name & "] - " & message
	End If
		
End Sub

<xatm_config_data.Config.BuildAlarms:BuildAlarms_OnStartRunning()>
Sub BuildAlarms_OnStartRunning()
	
	DocString = ""

End Sub

<xatm_config_data.Config.BuildDistribution:BuildDistribution_OnChangedValue()>
Sub BuildDistribution_OnChangedValue()

	If IsEmpty(Value) Or IsNull(Value) Then Exit Sub
	If Trim(CStr(Value)) = "" Then Exit Sub

	Dim ts
	ts = TimeStamp                         ' preserved for the silent clear

	' offset|path, split at the first bar. The offset is a number and
	' carries no bar of its own, so whatever a driver's path contains
	' travels whole - which a Split on every bar would not manage.
	Dim payload
	payload = Trim(CStr(Value))

	Dim bar
	bar = InStr(payload, "|")

	If bar = 0 Then
		Fail "the command carries no starting address (expected offset|path).", ts
		Exit Sub
	End If

	Dim offsetText, driverPath
	offsetText = Trim(Left(payload, bar - 1))
	driverPath = Trim(Mid(payload, bar + 1))

	If Not IsNumeric(offsetText) Then
		Fail "'" & offsetText & "' is not a starting address.", ts
		Exit Sub
	End If

	Dim driver
	Set driver = Nothing
	On Error Resume Next
	Set driver = Application.GetObject(driverPath)
	On Error Goto 0

	If driver Is Nothing Then
		Fail "there is nothing at " & driverPath & ".", ts
		Exit Sub
	End If

	' Where the signals come from. Read from the project rather than
	' assumed, because the interface sits at the root of the data project
	' in some layouts and inside XATM_Data in others.
	Dim interfaceServer
	Set interfaceServer = InterfaceRoot()

	If interfaceServer Is Nothing Then
		Fail "the project has no interface data server, so there is " & _
		     "nothing for the distribution to carry.", ts
		Exit Sub
	End If

	gInterfaceRoot = interfaceServer.PathName

	' Emptied and made again rather than brought up to date, the way the
	' interface and the substation's alarms are. Addresses are handed out
	' in walk order, so a half-kept folder would leave the numbering
	' saying one thing and the operation centre's point list another.
	Dim folder
	Set folder = ChildFolder(driver, DISTRIBUTION_FOLDER)

	If folder Is Nothing Then
		Fail "the driver would not take a " & DISTRIBUTION_FOLDER & " folder.", ts
		Exit Sub
	End If

	ClearFolder folder

	gFolders  = 0
	gMonitors = 0
	gCommands = 0

	' One run of addresses for both kinds. A monitored point and a
	' command share an address space in 104, so numbering them from
	' separate bases would only be two ways to land on one address.
	gFirstAddress = CLng(offsetText)
	gNextAddress  = gFirstAddress

	Dim problem
	problem = ""

	WalkFolder folder, Application.GetObject(DATA_ROOT), "", problem

	Dim built
	built = gMonitors + gCommands

	' --- and then the one thing most likely to go wrong ----------------
	'
	' Saved on its own and reported on its own, because this is where an
	' unlicensed copy stops: E3 refuses to save past a small number of
	' tags without a valid licence, and a distribution is easily hundreds.
	' The tags are made either way, so a Save that fails leaves a project
	' that looks built and comes back empty - which is worth saying
	' loudly rather than leaving as one line among the rest.
	' Save answers True or False and raises nothing, so the answer is what
	' decides. saved starts False so a call that never returns counts as a
	' refusal rather than passing for success.
	Dim saved
	saved = False

	On Error Resume Next
	Err.Clear

	saved = folder.Context("Container").Save()

	If Err.Number <> 0 Then Err.Clear
	On Error Goto 0

	If Not saved Then

		WriteLog "NOT SAVED - " & built & " tags were built and E3 would not " & _
		         "write them."
		WriteLog "NOT SAVED - E3 will not save past about " & LICENCE_TAG_HINT & _
		         " tags without a valid licence. The tags are in the project " & _
		         "now and will be gone when it is closed. Check the licence " & _
		         "before building again."

		DocString = EXIT_FAILURE
		WriteEx Empty, ts
		Exit Sub

	End If

	If built = 0 Then
		WriteLog "Distribution - nothing is interfaced to carry."
	Else
		WriteLog "Distribution - " & gMonitors & " monitored and " & gCommands & _
		         " commanded on " & gFolders & " objects, under " & driverPath & _
		         "." & DISTRIBUTION_FOLDER & ". Addresses " & gFirstAddress & _
		         ".." & (gNextAddress - 1) & "."
	End If

	If problem <> "" Then
		WriteLog "Distribution is incomplete:" & problem
		DocString = EXIT_FAILURE
		WriteEx Empty, ts
		Exit Sub
	End If

	DocString = EXIT_SUCCESS
	WriteEx Empty, ts                      ' clear without re-firing

End Sub


' The folder built under the chosen IODriver, and what it is made of.
Const DISTRIBUTION_FOLDER = "Automatismos"
Const IO_FOLDER_CLASS     = "IOFolder"
Const IO_TAG_CLASS        = "IOTag"

' What marks a tag as this configuration's to remake - the same mark the
' interface puts on the tags it builds.
Const SELF_MARK           = "$XATM$"

' Always in advise: the driver keeps the point up to date of its own
' accord rather than waiting to be asked, which is what a slave
' publishing to a control centre has to do.
Const ADVISE_ALWAYS       = 0

' What the driver is told each tag is - the mnemonic at the head of the
' configuration string.
'
' SP is a single point, which is what a discrete signal is: one bit,
' going up to the operation centre. SC is a single command coming down,
' and it is executed rather than selected, there being no
' select-before-operate between this level and level 3.
Const TYPE_MONITOR        = "SP"
Const TYPE_COMMAND        = "SC"
Const COMMAND_ACTION      = "Execute"

' How a link carries the value - the BindType argument of CreateLink.
' A discrete signal is copied from the interface into the tag the driver
' publishes; a command arrives on the tag and has to reach the interface,
' which is the other way about.
Const LINK_SIMPLE         = 0
Const LINK_REVERSE        = 5

' Said in the message when a save is refused. Not read by anything - it
' is the number an engineer needs to hear to recognise the cause.
Const LICENCE_TAG_HINT    = 20

Const DATA_ROOT           = "XATM_Data"
Const CONFIG_DATA         = "xatm_config_data"
Const HELPER_FOLDER       = "PropertiesHelper"

Const INTERFACE_ROOT      = "XATM_Interface"
Const INTERFACE_ROOT_ALT  = "XATM_Data.XATM_Interface"

Const EXIT_SUCCESS        = "EXIT_SUCCESS"
Const EXIT_FAILURE        = "EXIT_FAILURE"

' Where the interface was found this run, as the head of every path the
' distribution reads from.
Dim gInterfaceRoot

' Where the numbering started this run, and the next one to hand out.
Dim gFirstAddress
Dim gNextAddress

' What the last build came out at, for the line it writes to the log.
Dim gFolders
Dim gMonitors
Dim gCommands


' One folder of the data project: a tag for every object in it a
' manifest speaks for, and then the same for whatever is under it. The
' walk the interface rebuild does, over the same tree.
Sub WalkFolder(root, folder, path, problem)

	Dim item, bag

	For Each item In folder

		Set bag = ManifestOf(TypeName(item))

		If Not bag Is Nothing Then

			WalkObject root, path, item, bag, problem

		Else

			On Error Resume Next
			WalkFolder root, item, ChildPath(path, item.Name), problem
			On Error Goto 0

		End If

	Next

End Sub


' One name onto a path, and the first name is the whole of it.
Function ChildPath(path, objectName)

	If path = "" Then
		ChildPath = objectName
	Else
		ChildPath = path & "." & objectName
	End If

End Function


' The folder at a path under the distribution root, each level made
' where it is not already there.
Function EnsureFolder(root, path, problem)

	Set EnsureFolder = root
	If path = "" Then Exit Function

	Dim names
	names = Split(path, ".")

	Dim parent, child, i
	Set parent = root

	For i = 0 To UBound(names)

		Set child = Nothing

		On Error Resume Next
		Set child = parent.Item(names(i))
		On Error Goto 0

		If child Is Nothing Then Set child = NewChild(parent, IO_FOLDER_CLASS, names(i))

		If child Is Nothing Then
			problem = problem & vbCrLf & "  no folder for " & path
			Set EnsureFolder = Nothing
			Exit Function
		End If

		Set parent = child

	Next

	Set EnsureFolder = parent

End Function


' One object's folder, and in it a tag for each signal worth carrying up
' to the operation centre: a discrete reading, or a command coming back
' down. An object with neither gets no folder at all.
Sub WalkObject(root, path, obj, bag, problem)

	Dim key, p

	Dim wanted
	wanted = False

	For Each key In bag.Keys
		If IsDiscrete(bag(key)) Then wanted = True
		If IsCommand(bag(key))  Then wanted = True
	Next

	If Not wanted Then Exit Sub

	Dim parent
	Set parent = EnsureFolder(root, path, problem)
	If parent Is Nothing Then Exit Sub

	Dim folder
	Set folder = NewChild(parent, IO_FOLDER_CLASS, obj.Name)

	If folder Is Nothing Then
		problem = problem & vbCrLf & "  no folder for " & obj.Name
		Exit Sub
	End If

	gFolders = gFolders + 1

	' Where this object's signals are on the interface. The distribution
	' reads there and never from the XObject: the interface is where the
	' Elipse application meets the automation, and carrying a signal to
	' level 3 is the application's work.
	Dim interfacePath
	interfacePath = gInterfaceRoot & "." & ChildPath(path, obj.Name)

	For Each key In bag.Keys

		Set p = bag(key)

		If IsDiscrete(p) Then
			MonitorTag folder, interfacePath, obj, p, problem
		ElseIf IsCommand(p) Then
			CommandTag folder, interfacePath, obj, p, problem
		End If

	Next

End Sub


' A discrete signal: an interfaced Boolean the automation reads and the
' operation centre is shown. Not a command - those are InternalTag
' properties and go the other way - and not an IOTag, which is the
' engineer's own wiring to the IED.
Function IsDiscrete(p)

	IsDiscrete = False

	If LCase(p.DataType & "") <> "boolean" Then Exit Function
	If Not p.IsIOTagged() Then Exit Function

	IsDiscrete = True

End Function


' A command the interface carries: an InternalTag property a manifest
' interfaces. RebuildInterface leaves these unlinked and points the
' XObject's property at them, so writing one is how a maneuver is asked
' for - and level 3 asks through the driver.
Function IsCommand(p)

	IsCommand = False

	If LCase(p.DataType & "") <> "internaltag" Then Exit Function
	If Not p.IsIOTagged() Then Exit Function

	IsCommand = True

End Function


' The next address, and the one after it kept for the next caller.
Function NextAddress()

	NextAddress = gNextAddress
	gNextAddress = gNextAddress + 1

End Function


' The tag one signal gets under the driver, wired to the interface tag
' standing for it. Nothing when it could not be made.
'
' A link carries a value one way, so the direction is the whole of the
' difference between the two kinds:
'
'   discrete  interface -> tag -> operation centre, on a simple link,
'             and the driver is allowed to write it out
'   command   operation centre -> tag -> interface, on a reverse one,
'             and the driver is allowed to read it in
Function BoundIOTag(folder, interfacePath, obj, p, problem, linkType, configString)

	Set BoundIOTag = Nothing

	Dim source
	source = interfacePath & "." & p.Name

	' Checked before it is wired. The interface is emptied and rebuilt on
	' every Save, so tags built before that rebuild would wire to ones on
	' their way out - and a link to a tag that is not there fails quietly,
	' leaving a point the operation centre sees and nothing ever moves.
	Dim target
	Set target = Nothing
	On Error Resume Next
	Set target = Application.GetObject(source)
	On Error Goto 0

	If target Is Nothing Then
		problem = problem & vbCrLf & "  " & obj.Name & "." & p.Name & _
		          " has no interface tag at " & source
		Exit Function
	End If

	Dim tag
	Set tag = NewChild(folder, IO_TAG_CLASS, p.Name)

	If tag Is Nothing Then
		problem = problem & vbCrLf & "  no tag for " & obj.Name & "." & p.Name
		Exit Function
	End If

	Dim where
	where = obj.Name & "." & p.Name

	' What the driver is told this point is. Written on its own and
	' reported on its own, the way an alarm's settings are: a driver that
	' will not take one string should cost that point rather than the
	' whole distribution.
	SetMember tag, "ParamItem",      configString,  where, problem
	SetMember tag, "DocString",      SELF_MARK,     where, problem
	SetMember tag, "AdviseType",     ADVISE_ALWAYS, where, problem
	SetMember tag, "EnableDeadBand", False,         where, problem

	On Error Resume Next
	Err.Clear

	tag.Links.CreateLink "Value", source, linkType

	If Err.Number <> 0 Then
		problem = problem & vbCrLf & "  " & where & " would not wire to " & _
		          source & " - " & Err.Description
		Err.Clear
	End If

	On Error Goto 0

	Set BoundIOTag = tag

End Function


' One discrete signal on its way up: read off the interface, published
' by the driver. The driver writes it out and never reads it back.
Sub MonitorTag(folder, interfacePath, obj, p, problem)

	Dim configString
	configString = TYPE_MONITOR & ":" & NextAddress()

	Dim tag
	Set tag = BoundIOTag(folder, interfacePath, obj, p, problem, LINK_SIMPLE, configString)
	If tag Is Nothing Then Exit Sub

	Dim where
	where = obj.Name & "." & p.Name

	SetMember tag, "AllowRead",  False, where, problem
	SetMember tag, "AllowWrite", True,  where, problem

	gMonitors = gMonitors + 1

End Sub


' One command on its way down: taken off the driver, written through to
' the interface. The driver reads it in and never writes it out.
Sub CommandTag(folder, interfacePath, obj, p, problem)

	Dim configString
	configString = TYPE_COMMAND & ":" & NextAddress() & "." & COMMAND_ACTION

	Dim tag
	Set tag = BoundIOTag(folder, interfacePath, obj, p, problem, LINK_REVERSE, configString)
	If tag Is Nothing Then Exit Sub

	Dim where
	where = obj.Name & "." & p.Name

	SetMember tag, "AllowRead",  True,  where, problem
	SetMember tag, "AllowWrite", False, where, problem

	' The driver raises an event when the centre sends this, rather than
	' the value simply appearing on the tag. The same command asked for
	' twice running is the same value twice, and without this the second
	' one would look like nothing had happened at all.
	SetMember tag, "EnableDriverEvent", True, where, problem

	gCommands = gCommands + 1

End Sub


' One setting onto a tag, late-bound the way the import writes a
' property, and named in the log when it will not take.
Sub SetMember(tag, memberName, newValue, where, problem)

	gMemberValue = newValue

	On Error Resume Next
	Err.Clear

	Execute "tag." & memberName & " = gMemberValue"

	If Err.Number <> 0 Then
		problem = problem & vbCrLf & "  " & where & " would not take " & _
		          memberName & " - " & Err.Description
		Err.Clear
	End If

	On Error Goto 0

End Sub

' Scratch cell for the late-bound write done through Execute, the way
' the interface rebuild keeps one.
Dim gMemberValue


' A child of a container by name, made where it is not already there.
Function ChildFolder(parent, folderName)

	Set ChildFolder = Nothing

	On Error Resume Next
	Set ChildFolder = parent.Item(folderName)
	On Error Goto 0

	If Not ChildFolder Is Nothing Then Exit Function

	Set ChildFolder = NewChild(parent, IO_FOLDER_CLASS, folderName)

End Function


' Empties a folder. Taken out in a second pass, because a collection is
' not walked while what is in it is being removed.
Sub ClearFolder(folder)

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim item
	For Each item In folder
		doomed.Add doomed.Count, item.Name
	Next

	Dim n
	For Each n In doomed.Keys

		On Error Resume Next
		folder.DeleteObject doomed(n)
		On Error Goto 0

	Next

End Sub


' A new child of a container, Nothing when it would not take one -
' AddObject(ClassName, [Activate], [ObjectName]).
Function NewChild(parent, className, objectName)

	Set NewChild = Nothing

	On Error Resume Next
	Set NewChild = parent.AddObject(className, True, objectName)
	On Error Goto 0

End Function


' The data server the interface was built in.
Function InterfaceRoot()

	Set InterfaceRoot = Nothing

	On Error Resume Next
	Set InterfaceRoot = Application.GetObject(INTERFACE_ROOT)
	On Error Goto 0

	If Not InterfaceRoot Is Nothing Then Exit Function

	On Error Resume Next
	Set InterfaceRoot = Application.GetObject(INTERFACE_ROOT_ALT)
	On Error Goto 0

End Function


' The manifest declared for a class, or Nothing when the class has none.
Function ManifestOf(className)

	Set ManifestOf = Nothing

	On Error Resume Next
	Set ManifestOf = Application.GetObject(CONFIG_DATA).Item(HELPER_FOLDER).Item(className).Value
	On Error Goto 0

End Function


' Said, logged, and the command cleared, for a build that got nowhere.
Sub Fail(message, ts)

	DocString = EXIT_FAILURE
	WriteLog message
	WriteEx Empty, ts

End Sub


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Name & "] - " & message
	End If
	
End Sub

<xatm_config_data.Config.BuildDistribution:BuildDistribution_OnStartRunning()>
Sub BuildDistribution_OnStartRunning()

	DocString = ""
	
End Sub

<xatm_config_data.Config.ExportSignalsToCsv:ExportSignalsToCsv_OnChangedValue()>
Sub ExportSignalsToCsv_OnChangedValue()

	' Every signal the automation publishes, as two files the SCADA
	' developer pastes into the load plan the import library reads.
	'
	' This is the whole of what leaves here now. The alarms, the user
	' fields and the 104 addresses are made by that library from the
	' spreadsheet, so nothing in this project needs to make them - it
	' needs to say, exactly and in the developer's own columns, what
	' there is and where it lives.
	'
	' Run after Apply and Save and after nothing else: Save is what
	' builds the interface tags every row points at, and a row naming a
	' tag that is not there is a row that imports and never moves.
	'
	' Written inside the Sub and not above it: a scope has to open on
	' Sub, and E3 rejects one that opens on anything else - a comment
	' included.

	If IsEmpty(Value) Or IsNull(Value) Then Exit Sub
	If Trim(CStr(Value)) = "" Then Exit Sub

	Dim ts
	ts = TimeStamp                         ' preserved for the silent clear

	' What the operator picked in the Save dialog. One name, and two
	' files written beside it: the columns a discrete needs and the ones
	' a command needs are not the same list, so they cannot share a
	' sheet - and asking twice for what is one decision is asking twice.
	Dim chosenPath
	chosenPath = Trim(CStr(Value))

	Dim fso
	Set fso = Nothing
	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	On Error Goto 0

	If fso Is Nothing Then
		Fail "this machine will not give a FileSystemObject, so nothing " & _
		     "can be written.", ts
		Exit Sub
	End If

	Dim folderPath, baseName
	folderPath = ""
	baseName   = ""

	On Error Resume Next
	folderPath = fso.GetParentFolderName(chosenPath)
	baseName   = fso.GetBaseName(chosenPath)
	On Error Goto 0

	If folderPath = "" Or baseName = "" Then
		Fail chosenPath & " is not a file path this can write to.", ts
		Exit Sub
	End If

	If Not fso.FolderExists(folderPath) Then
		Fail "there is no folder at " & folderPath & ".", ts
		Exit Sub
	End If

	Dim root
	Set root = Nothing
	On Error Resume Next
	Set root = Application.GetObject(DATA_ROOT)
	On Error Goto 0

	If root Is Nothing Then
		Fail "the project has no " & DATA_ROOT & ".", ts
		Exit Sub
	End If

	' Written as ANSI and not Unicode - the third argument. Excel in a
	' Portuguese install reads a semicolon-separated ANSI file straight
	' into columns, and reads a UTF-8 one as one column of mojibake.
	Dim discFile, cmdFile
	Set discFile = Nothing
	Set cmdFile  = Nothing

	On Error Resume Next
	Dim discPath, cmdPath
	discPath = fso.BuildPath(folderPath, baseName & DISCRETE_SUFFIX)
	cmdPath  = fso.BuildPath(folderPath, baseName & COMMAND_SUFFIX)

	Set discFile = fso.CreateTextFile(discPath, True, False)
	Set cmdFile  = fso.CreateTextFile(cmdPath,  True, False)
	On Error Goto 0

	If discFile Is Nothing Or cmdFile Is Nothing Then
		If Not discFile Is Nothing Then discFile.Close
		If Not cmdFile  Is Nothing Then cmdFile.Close
		Fail "the two files would not open under " & folderPath & _
		     " - it may be read only, or open in Excel.", ts
		Exit Sub
	End If

	gDiscretes = 0
	gCommands  = 0
	gObjects   = 0
	gFolders   = 0

	Dim problem
	problem = ""

	discFile.WriteLine DISCRETE_HEADER
	cmdFile.WriteLine  COMMAND_HEADER

	ExportFolder root, "", discFile, cmdFile, problem

	On Error Resume Next
	discFile.Close
	cmdFile.Close
	On Error Goto 0

	If gDiscretes = 0 And gCommands = 0 Then
		WriteLog "Export - " & gObjects & " objects under " & gFolders & _
		         " folders of " & DATA_ROOT & ", and not one of them " & _
		         "carries a signal to level 3. Both files are headings " & _
		         "and no rows."
	Else
		WriteLog "Export - " & gDiscretes & " discretes and " & gCommands & _
		         " commands from " & gObjects & " objects, written beside " & _
		         chosenPath & "."
	End If

	If problem <> "" Then
		WriteLog "Export is incomplete:" & problem
		DocString = EXIT_FAILURE
		WriteEx Empty, ts
		Exit Sub
	End If

	DocString = EXIT_SUCCESS
	WriteEx Empty, ts                      ' clear without re-firing

End Sub


' What the two files are called and what stands at the head of each.
' The column names are the load plan's own, accents and all, so that
' a paste lands under the heading it belongs to.
' What is put on the end of the chosen name to tell the two apart.
Const DISCRETE_SUFFIX = "_Discretos.csv"
Const COMMAND_SUFFIX  = "_Comandos.csv"

Const DISCRETE_HEADER = "UserField SignalName;Equip;Equip ($[EQ]);Típico;Template;Descrição;MeasurementType;Texto0;Texto1;Conversão;Atua em;Atua em ($[EQ]);XATM Property"
Const COMMAND_HEADER  = "UserField SignalName;Equip;Equip ($[EQ]);Típico;Template;Descrição;CommandType;Ação;Conversão;Ação0;Ação1;Comandos Disparados;Comandos Disparados ($[EQ]);XATM Property"

' Semicolon, because a Portuguese Excel splits on it and because a
' decimal comma would tear a number in half if it did not.
Const SEP            = ";"

' Told to the import as the object it is standing in for. One row
' carries both: the name as it is, for a project that lists its
' equipment, and the placeholder, for one that expands a typical
' across every instance. The developer keeps the column they use.
Const EQ_PLACEHOLDER = "$[EQ]"

Const DATA_ROOT      = "XATM_Data"
Const CONFIG_DATA    = "xatm_config_data"
Const HELPER_FOLDER  = "PropertiesHelper"
Const INTERFACE_ROOT = "XATM_Interface"

Const EXIT_SUCCESS   = "EXIT_SUCCESS"
Const EXIT_FAILURE   = "EXIT_FAILURE"

Dim gDiscretes
Dim gCommands

' What the walk met on the way, so that a run with no rows can say
' whether it found no objects or found them and none had a point.
Dim gObjects
Dim gFolders


' One folder of the data project, then whatever is under it - the
' same walk the alarms and the distribution do, so the file comes out
' in the order the tree reads.
Sub ExportFolder(folder, path, discFile, cmdFile, problem)

	Dim item, bag, where

	For Each item In folder

		Set bag = ManifestOf(TypeName(item))
		where = ChildPath(path, item.Name)

		If Not bag Is Nothing Then

			gObjects = gObjects + 1

			' Kept rather than dropped. One property this cannot read would
			' otherwise abandon the object it is on and every object after
			' it in the folder, and the run would end saying only that
			' there was nothing to carry.
			On Error Resume Next
			Err.Clear

			ExportObject path, item, bag, discFile, cmdFile, problem

			If Err.Number <> 0 Then
				problem = problem & vbCrLf & "  " & where & " (" & _
				          TypeName(item) & ") - " & Err.Description
				Err.Clear
			End If

			On Error Goto 0

		Else

			' Anything without a manifest is taken for a folder and walked.
			' Something that is neither - a tag, or a class whose manifest
			' never loaded - raises here, and that is the whole reason a
			' run can come back empty with nothing to say about it. Said.
			gFolders = gFolders + 1

			On Error Resume Next
			Err.Clear

			ExportFolder item, where, discFile, cmdFile, problem

			If Err.Number <> 0 Then
				problem = problem & vbCrLf & "  " & where & " (" & _
				          TypeName(item) & ") is neither a folder nor a class " & _
				          "with a manifest - " & Err.Description
				Err.Clear
			End If

			On Error Goto 0

		End If

	Next

End Sub


' One name onto a path, and the first name is the whole of it.
Function ChildPath(path, objectName)

	If path = "" Then
		ChildPath = objectName
	Else
		ChildPath = path & "." & objectName
	End If

End Function


' Every property of one object that level 3 is given a point for: a
' Boolean is something the centre is shown and goes in the discretes,
' an InternalTag is something it asks for and goes in the commands.
'
' Asked of IsIOTagged and not of the alarm table, because the two are
' different questions - a signal can be carried without an operator
' ever being told about it.
Sub ExportObject(path, obj, bag, discFile, cmdFile, problem)

	Dim inner
	inner = ChildPath(path, obj.Name)

	Dim tipico
	tipico = TipicoOf(obj)

	Dim key, p, litPath, phPath

	For Each key In bag.Keys

		Set p = bag(key)

		If p.IsIOTagged() Then

			litPath = INTERFACE_ROOT & "." & inner & "." & p.Name
			phPath  = INTERFACE_ROOT & "." & _
			          ChildPath(path, EQ_PLACEHOLDER) & "." & p.Name

			Select Case LCase(p.DataType & "")

				Case "boolean"
					discFile.WriteLine DiscreteRow(obj, p, tipico, litPath, phPath)
					gDiscretes = gDiscretes + 1

				Case "internaltag"
					cmdFile.WriteLine CommandRow(obj, p, tipico, litPath, phPath)
					gCommands = gCommands + 1

			End Select

		End If

	Next

End Sub


' One discrete, in the developer's columns.
'
' Template is left empty on purpose. TM100PC and R086 are names out
' of the import library's own catalogue and there is no rule here
' that produces them - a plausible wrong one binds the signal to
' somebody else's template, which is worse than a blank the developer
' fills. XATM Property rides along at the end so the row can always
' be traced back to what made it.
Function DiscreteRow(obj, p, tipico, litPath, phPath)

	Dim t0, t1
	t0 = ""
	t1 = ""

	' Which text belongs to 0 and which to 1 is decided by the state
	' that raises the alarm, not by the order the pair is written in:
	' a precondition raises on 0 and reads the other way round.
	Dim parts
	If p.AlarmPair & "" <> "" Then

		parts = Split(p.AlarmPair, "|")

		If UBound(parts) >= 2 Then
			If Trim(parts(2)) = "1" Then
				t0 = parts(0)
				t1 = parts(1)
			Else
				t0 = parts(1)
				t1 = parts(0)
			End If
		End If

	End If

	Dim measurement
	measurement = ""
	If t0 <> "" Then measurement = MeasurementOf(t0, t1)

	DiscreteRow = Join(Array( _
		"", _
		Field(obj.Name), _
		EQ_PLACEHOLDER, _
		Field(tipico), _
		"", _
		Field(p.AlarmLabel), _
		Field(measurement), _
		Field(t0), _
		Field(t1), _
		"Booleano", _
		Field(litPath), _
		Field(phPath), _
		Field(p.Name)), SEP)

End Function


' One command, in the developer's columns.
'
' The description is the property's own Portuguese help. A command
' raises no alarm and so has no alarm label to lend, and the help is
' the only sentence in this project that says what the command does -
' which is what the column is for. It is a sentence and not a legend,
' so expect it to be shortened.
Function CommandRow(obj, p, tipico, litPath, phPath)

	CommandRow = Join(Array( _
		"", _
		Field(obj.Name), _
		EQ_PLACEHOLDER, _
		Field(tipico), _
		"", _
		Field(p.Help("pt-BR")), _
		"", _
		"", _
		"1", _
		"Executar", _
		"", _
		Field(litPath), _
		Field(phPath), _
		Field(p.Name)), SEP)

End Function


' What kind of equipment the load plan files this under.
Function TipicoOf(obj)

	Select Case LCase(TypeName(obj))

		Case "xatm_tmtnm"        : TipicoOf = "TMT"
		Case "xatm_ta"           : TipicoOf = "TA"
		Case "xatm_raseat"       : TipicoOf = "RASE"
		Case "xatm_breaker"      : TipicoOf = "DJ"
		Case "xatm_disconnector" : TipicoOf = "CS"
		Case "xatm_transformer"  : TipicoOf = "TR"
		Case Else                 : TipicoOf = ""

	End Select

End Function


' The measurement type, which the import spells as the two state
' texts joined - NORMAL_ATUADO, NATENDIDAS_ATENDIDAS.
Function MeasurementOf(t0, t1)

	MeasurementOf = MeasurementPart(t0) & "_" & MeasurementPart(t1)

End Function


' One half of it: no accents, no spaces, and NAO folded onto the word
' it negates the way the catalogue already writes NATENDIDAS.
Function MeasurementPart(text)

	Dim s
	s = Ascii(text)
	s = Replace(s, "NAO ", "N")
	s = Replace(s, " ", "")

	MeasurementPart = s

End Function


' The Portuguese letters without their accents.
'
' Written as Chr() and never as the letters themselves so that this
' scope stays plain ASCII: the dump this is kept in is cp1252, and an
' accent typed here is a byte that any tool guessing the encoding
' wrong will turn into a question mark.
Function Ascii(text)

	Dim s
	s = UCase(text & "")

	s = Replace(s, Chr(192), "A")    ' A grave
	s = Replace(s, Chr(193), "A")    ' A agudo
	s = Replace(s, Chr(194), "A")    ' A circunflexo
	s = Replace(s, Chr(195), "A")    ' A til
	s = Replace(s, Chr(199), "C")    ' C cedilha
	s = Replace(s, Chr(201), "E")    ' E agudo
	s = Replace(s, Chr(202), "E")    ' E circunflexo
	s = Replace(s, Chr(205), "I")    ' I agudo
	s = Replace(s, Chr(211), "O")    ' O agudo
	s = Replace(s, Chr(212), "O")    ' O circunflexo
	s = Replace(s, Chr(213), "O")    ' O til
	s = Replace(s, Chr(218), "U")    ' U agudo
	s = Replace(s, Chr(220), "U")    ' U trema

	Ascii = s

End Function


' One value on its way into the file. Quoted only when it has to be,
' which keeps the file readable in a text editor: a separator, a
' quote or a line break inside a value is what forces it, and a help
' sentence can carry all three.
Function Field(text)

	Dim s
	s = text & ""

	If InStr(s, SEP) > 0 Or InStr(s, """") > 0 Or _
	   InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
		s = """" & Replace(s, """", """""") & """"
	End If

	Field = s

End Function


' The manifest declared for a class, or Nothing where the class has
' none - the same lookup the export, the import and the alarms do.
Function ManifestOf(className)

	Set ManifestOf = Nothing

	On Error Resume Next
	Set ManifestOf = Application.GetObject(CONFIG_DATA).Item(HELPER_FOLDER).Item(className).Value
	On Error Goto 0

End Function


' Said, logged, and the command cleared, for an export that got
' nowhere.
Sub Fail(message, ts)

	DocString = EXIT_FAILURE
	WriteLog message
	WriteEx Empty, ts

End Sub


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Name & "] - " & message
	End If
	
End Sub

<xatm_config_data.Config.BindPositions:BindPositions_OnChangedValue()>
Sub BindPositions_OnChangedValue()

	' Every breaker and disconnector wired to the tags the substation side
	' already reads and operates it on.
	'
	' The engineer configures the PowerSubstation first - every device bound
	' to its measurements and its commands - and then had to bind the same
	' tags a second time on the automation, by hand, once per property. The
	' two sides already agree on the device names, so the second pass is
	' derivable from the first.
	'
	' Staged in the document and not written into the project. The panel is
	' built from the document and Save is what pushes it into E3, so a
	' property written straight into the object would be overwritten by the
	' next Save with whatever the document still held. This goes the same
	' way a tag picked by hand goes.

	If IsEmpty(Value) Or IsNull(Value) Then Exit Sub
	If Trim(CStr(Value)) = "" Then Exit Sub

	Dim ts
	ts = TimeStamp                         ' preserved for the silent clear

	' Which substation to take the tags from. It travels as the payload
	' because the server cannot ask: a domain can hold more than one, and
	' the button is where the choosing happens.
	Dim substationPath
	substationPath = Trim(CStr(Value))

	Dim substation
	Set substation = Nothing
	On Error Resume Next
	Set substation = Application.GetObject(substationPath)
	On Error Goto 0

	If substation Is Nothing Then
		Fail "there is nothing at " & substationPath & ".", ts
		Exit Sub
	End If

	' Every Power device of that substation, by name. Built once: the
	' alternative is walking it again for each of the twenty-odd automation
	' objects.
	Dim byName
	Set byName = CreateObject("Scripting.Dictionary")

	' Names are matched without case. The two sides are configured by
	' different people at different times and DJ03 against dj03 is not a
	' difference either of them meant.
	byName.CompareMode = 1

	gClashes = 0

	IndexClass BREAKER_CLASS, substationPath, byName
	IndexClass SWITCH_CLASS, substationPath, byName

	If byName.Count = 0 Then
		Fail substationPath & " holds no " & BREAKER_CLASS & " or " & _
		     SWITCH_CLASS & " to take the tags from.", ts
		Exit Sub
	End If

	Dim root
	Set root = Nothing
	On Error Resume Next
	Set root = Application.GetObject(DATA_ROOT)
	On Error Goto 0

	If root Is Nothing Then
		Fail "the project has no " & DATA_ROOT & ".", ts
		Exit Sub
	End If

	gBound   = 0
	gDevices = 0
	gMissing = 0
	gNoTag   = 0
	gNoCommand = 0
	gNoDefect  = 0
	gRefused = 0

	Dim problem
	problem = ""

	BindFolder root, byName, problem

	WriteLog gBound & " properties bound on " & gDevices & " devices, from " & _
	         substationPath & "."

	' Said as a count and not as a fault. A device that is only monitored
	' carries no command point at all, and listing each one would turn an
	' ordinary run into a failed one.
	If gNoCommand > 0 Then
		WriteLog gNoCommand & " devices have no " & COMMANDS & "." & COMMAND_POINT & _
		         " and were bound for reading only."
	End If

	If gNoDefect > 0 Then
		WriteLog gNoDefect & " devices carry none of the signals that block a " & _
		         "manoeuvre, and were left without a Defective expression."
	End If

	If gClashes > 0 Then
		WriteLog gClashes & " Power devices share a name " & "with another; the first of each was used."
	End If

	If problem <> "" Then
		WriteLog "Not everything was bound:" & problem
		DocString = EXIT_FAILURE
		WriteEx Empty, ts
		Exit Sub
	End If

	DocString = EXIT_SUCCESS
	WriteEx Empty, ts                      ' clear without re-firing

End Sub


' What the two sides are called.
Const BREAKER_CLASS = "PowerBreaker"
Const SWITCH_CLASS  = "PowerSwitch"

' Where the position lives on the Power side: the measurement named for
' the device state, and the Scada child under it that names the tag.
Const MEASUREMENTS  = "Measurements"
Const MEASUREMENT   = "0EST"
Const SCADA         = "Scada"

' And where the command lives: a collection of its own under a point of
' the same name, holding one PowerCommandUnit per action. Spelled out
' apart from the measurement even though the two read alike today - they
' are two points on the Power side and nothing binds them to stay equal.
Const COMMANDS      = "Commands"
Const COMMAND_POINT = "0EST"
Const OPEN_UNIT     = "Open"
Const CLOSE_UNIT    = "Close"

' The measurements that say a device cannot be worked, and the reading
' each one says it on. Written name;value the way EDP gave the list, so
' a fourth signal is one more entry and nothing else.
Const DEFECT_SIGNALS = "52ML;1|00LR;0|CBTL;0"

' The automation properties they are copied onto: the position onto both
' of the position pair, each command unit onto the one it names, and
' everything that blocks a manoeuvre onto the one expression.
Const PROP_OPEN      = "PositionOpen"
Const PROP_CLOSED    = "PositionClosed"
Const PROP_CMD_OPEN  = "CommandOpen"
Const PROP_CMD_CLOSE = "CommandClose"
Const PROP_DEFECTIVE = "Defective"

Const DATA_ROOT     = "XATM_Data"
Const FIND_OBJECTS  = "xatm_config_data.Config.FindObjects"
Const FOUND_OBJECTS = "xatm_config_data.Config.FoundObjects"
Const SET_PROPERTY  = "xatm_config_data.Config.SetProperty"
Const PROPERTY_VALUE = "xatm_config_data.Config.PropertyValue"
Const KIND_SOURCE   = "source"

Const EXIT_SUCCESS  = "EXIT_SUCCESS"
Const EXIT_FAILURE  = "EXIT_FAILURE"

Dim gBound
Dim gDevices
Dim gMissing
Dim gNoTag
Dim gNoCommand
Dim gNoDefect
Dim gRefused
Dim gClashes


' Every object of one class inside the chosen substation, added to the
' index by name.
'
' Asked of FindObjects rather than walked here: that tag already scans for
' a class and answers on FoundObjects, and a second walker would be a
' second thing to keep in step for no gain. It scans the whole domain
' though, so what comes back is filtered to the substation - see Under.
Sub IndexClass(className, substationPath, byName)

	Dim finder, answerTag
	Set finder = Nothing
	Set answerTag = Nothing

	On Error Resume Next
	Set finder = Application.GetObject(FIND_OBJECTS)
	Set answerTag = Application.GetObject(FOUND_OBJECTS)
	On Error Goto 0

	If finder Is Nothing Or answerTag Is Nothing Then
		WriteLog FIND_OBJECTS & " and " & FOUND_OBJECTS & _
		         " are what the search runs on, and one of them is missing."
		Exit Sub
	End If

	finder.WriteEx className

	If finder.DocString <> EXIT_SUCCESS Then
		WriteLog "The search for " & className & " got nowhere."
		Exit Sub
	End If

	Dim list, i, path, obj
	list = Split(CStr(answerTag.Value & ""), vbCrLf)

	For i = 0 To UBound(list)

		path = Trim(list(i))

		If path <> "" And Under(path, substationPath) Then

			Set obj = Nothing
			On Error Resume Next
			Set obj = Application.GetObject(path)
			On Error Goto 0

			If Not obj Is Nothing Then
				If byName.Exists(obj.Name) Then
					gClashes = gClashes + 1
				Else
					byName.Add obj.Name, path
				End If
			End If

		End If

	Next

End Sub


' Whether a path sits inside the substation that was chosen.
'
' A domain can hold more than one PowerSubstation - the picker on the
' button exists for exactly that - and two of them can each carry a DJ03.
' Without this the index would keep whichever the domain scan met first
' and bind the automation to another station's tag without a word about
' it.
Function Under(path, root)

	Under = False

	If Len(root) = 0 Then Exit Function
	If Len(path) <= Len(root) Then Exit Function

	Under = (LCase(Left(path, Len(root) + 1)) = LCase(root & "."))

End Function


' One folder of the data project, then whatever is under it - the walk the
' alarms and the export do, so a device is found where the tree says it.
Sub BindFolder(folder, byName, problem)

	Dim item, where

	For Each item In folder

		where = ""
		On Error Resume Next
		where = item.Name
		On Error Goto 0

		Select Case LCase(TypeName(item))

			Case "xatm_breaker", "xatm_disconnector"
				BindDevice item, byName, problem

			Case Else

				' Anything else is taken for a folder and walked. The error is
				' kept rather than dropped: swallowing it here would abandon
				' the rest of the folder and report nothing at all.
				On Error Resume Next
				Err.Clear

				BindFolder item, byName, problem

				If Err.Number <> 0 Then
					problem = problem & vbCrLf & "  " & where & " (" & _
					          TypeName(item) & ") - " & Err.Description
					Err.Clear
				End If

				On Error Goto 0

		End Select

	Next

End Sub


' One device: the Power object of the same name, and what it is read and
' operated on copied off it.
'
' The two halves are bound apart. A device configured for one and not the
' other is ordinary here - a disconnector that is only monitored has no
' command units at all - so a missing half is accounted for and the other
' half still lands.
Sub BindDevice(obj, byName, problem)

	If Not byName.Exists(obj.Name) Then
		gMissing = gMissing + 1
		problem = problem & vbCrLf & "  " & obj.Name & " - the domain has no " & _
		          BREAKER_CLASS & " or " & SWITCH_CLASS & " of that name."
		Exit Sub
	End If

	Dim objPath
	objPath = ""
	On Error Resume Next
	objPath = obj.PathName
	On Error Goto 0

	If objPath = "" Then Exit Sub

	Dim powerPath, before
	powerPath = byName(obj.Name)
	before = gBound

	BindPositionsOf obj, objPath, powerPath, problem
	BindCommandsOf obj, objPath, powerPath, problem
	BindDefectiveOf obj, objPath, powerPath, problem

	' Counted as a device only where something was actually bound onto it.
	If gBound > before Then gDevices = gDevices + 1

End Sub


' The position tag onto both of the position properties.
Sub BindPositionsOf(obj, objPath, powerPath, problem)

	Dim tagPath
	tagPath = ScadaTagOf(powerPath, MEASUREMENT)

	If tagPath = "" Then
		gNoTag = gNoTag + 1
		problem = problem & vbCrLf & "  " & obj.Name & " - nothing at " & _
		          MEASUREMENTS & "." & MEASUREMENT & "." & SCADA & ".Tag."
		Exit Sub
	End If

	StageOne objPath, PROP_OPEN, tagPath, obj.Name, problem
	StageOne objPath, PROP_CLOSED, tagPath, obj.Name, problem

End Sub


' Each command unit onto the command property it names.
'
' No command point at all is not a fault and is only counted - see the
' summary. The point with a unit missing off it, or a unit carrying no
' write tag, is half configured and is named.
Sub BindCommandsOf(obj, objPath, powerPath, problem)

	Dim openTag, closeTag

	If Not CommandTagsOf(powerPath, openTag, closeTag) Then
		gNoCommand = gNoCommand + 1
		Exit Sub
	End If

	If openTag = "" Then
		problem = problem & vbCrLf & "  " & obj.Name & " - no " & OPEN_UNIT & _
		          " unit with a write tag under " & COMMANDS & "." & COMMAND_POINT & "."
	Else
		StageOne objPath, PROP_CMD_OPEN, openTag, obj.Name, problem
	End If

	If closeTag = "" Then
		problem = problem & vbCrLf & "  " & obj.Name & " - no " & CLOSE_UNIT & _
		          " unit with a write tag under " & COMMANDS & "." & COMMAND_POINT & "."
	Else
		StageOne objPath, PROP_CMD_CLOSE, closeTag, obj.Name, problem
	End If

End Sub


' Everything that blocks a manoeuvre on this device, as one expression.
'
' Each of them is a measurement of its own and a device carries as many
' as it was configured with, so the expression is built out of the ones
' that are there. A device with none of them is left alone rather than
' given an expression that can never come true.
'
' Defective is a Boolean and holds no tag, so this is an expression and
' not an association - which is why .Value belongs on the end here and is
' stripped off the positions. It is put back rather than kept, so the
' expression reads the same whichever way the Power side spelled the tag.
Sub BindDefectiveOf(obj, objPath, powerPath, problem)

	Dim signals, i, piece, tagPath, terms
	signals = Split(DEFECT_SIGNALS, "|")
	terms = ""

	For i = 0 To UBound(signals)

		piece = Split(signals(i), ";")

		If UBound(piece) = 1 Then

			tagPath = ScadaTagOf(powerPath, Trim(piece(0)))

			If tagPath <> "" Then
				If terms <> "" Then terms = terms & " Or "
				terms = terms & tagPath & ".Value = " & Trim(piece(1))
			End If

		End If

	Next

	If terms = "" Then
		gNoDefect = gNoDefect + 1
		Exit Sub
	End If

	StageOne objPath, PROP_DEFECTIVE, terms, obj.Name, problem

End Sub


' The tag one Power measurement is read on, "" when there is none.
'
' The .Value on the end is taken off, whatever the caller wants the tag
' for: an IOTag property holds the tag and not one of its members, and an
' expression puts .Value back itself. The Power side spells the same tag
' both ways depending on who wired it, so neither caller can be handed
' what is written there.
Function ScadaTagOf(powerPath, measurementName)

	ScadaTagOf = ""

	Dim tagText
	tagText = ""

	On Error Resume Next
	Err.Clear

	tagText = CStr(Application.GetObject(powerPath) _
	          .Item(MEASUREMENTS).Item(measurementName).Item(SCADA).Tag & "")

	If Err.Number <> 0 Then
		tagText = ""
		Err.Clear
	End If

	On Error Goto 0

	ScadaTagOf = WithoutValue(Trim(tagText))

End Function


Function WithoutValue(text)

	Dim s
	s = Trim(text & "")

	If Len(s) > 6 Then
		If LCase(Right(s, 6)) = ".value" Then s = Left(s, Len(s) - 6)
	End If

	WithoutValue = s

End Function


' The two tags one Power device is operated on, and whether it carries a
' command point at all - which is the difference between a device that
' cannot be operated and one that was configured badly.
Function CommandTagsOf(powerPath, openTag, closeTag)

	CommandTagsOf = False
	openTag = ""
	closeTag = ""

	Dim point
	Set point = Nothing

	On Error Resume Next
	Set point = Application.GetObject(powerPath).Item(COMMANDS).Item(COMMAND_POINT)
	On Error Goto 0

	If point Is Nothing Then Exit Function

	CommandTagsOf = True

	Dim seen, unit, i
	seen = 0

	On Error Resume Next

	For Each unit In point

		' Counted on what came out and not on the turn of the loop. Where the
		' enumeration is refused outright the count has to stay at nothing, or
		' the fallback below reads as unnecessary and never runs.
		If IsObject(unit) Then
			seen = seen + 1
			TakeUnit unit, openTag, closeTag
		End If

	Next

	' A collection that will not enumerate still answers by index. Walked
	' from 0 through the count rather than 1 through it, because that spans
	' a collection numbered either way and the one subscript off the end of
	' whichever it is costs nothing under the guard.
	If seen = 0 Then
		Err.Clear
		For i = 0 To CountOf(point)
			Set unit = Nothing
			Set unit = point.Item(i)
			If Not (unit Is Nothing) Then TakeUnit unit, openTag, closeTag
			Err.Clear
		Next
	End If

	Err.Clear
	On Error Goto 0

End Function


' One command unit read onto whichever of the two tags its CommandName
' calls for.
'
' Taken off the name and not off the order they arrive in: a device that
' carries them the other way round, or carries only one of them, still
' lands on the right property. A unit with nothing to write on is left
' out, so the caller can tell it apart from one that is not there.
Sub TakeUnit(unit, openTag, closeTag)

	Dim action, tagText
	action = ""
	tagText = ""

	On Error Resume Next
	action = LCase(Trim(CStr(unit.CommandName & "")))
	tagText = WithoutValue(Trim(CStr(unit.OperateWriteTag & "")))
	Err.Clear
	On Error Goto 0

	If tagText = "" Then Exit Sub

	If action = LCase(OPEN_UNIT) Then
		openTag = tagText
	ElseIf action = LCase(CLOSE_UNIT) Then
		closeTag = tagText
	End If

End Sub


' How many a collection holds, 0 where it will not say.
Function CountOf(collection)

	CountOf = 0

	On Error Resume Next
	CountOf = CLng(collection.Count)
	Err.Clear
	On Error Goto 0

End Function


' One property staged in the document, on the two tags every other edit on
' this panel travels on: the value on one, and the command that reads it
' on the other.
'
' The key is the object's own path, which is what the document writes on
' each object element - so nothing here has to know about the id: form the
' panel rows use.
Sub StageOne(objPath, propName, tagPath, deviceName, problem)

	Dim command, valueTag
	Set command = Nothing
	Set valueTag = Nothing

	On Error Resume Next
	Set command = Application.GetObject(SET_PROPERTY)
	Set valueTag = Application.GetObject(PROPERTY_VALUE)
	On Error Goto 0

	If command Is Nothing Or valueTag Is Nothing Then
		gRefused = gRefused + 1
		problem = problem & vbCrLf & "  " & deviceName & "." & propName & _
		          " - " & SET_PROPERTY & " and " & PROPERTY_VALUE & _
		          " are what carry an edit, and one of them is missing."
		Exit Sub
	End If

	' The value first, so it is standing by when the command fires, and on a
	' tag of its own so nothing in a path has to be escaped or split.
	valueTag.WriteEx tagPath

	command.WriteEx KIND_SOURCE & "|" & objPath & "|" & propName

	If command.DocString <> EXIT_SUCCESS Then
		gRefused = gRefused + 1
		problem = problem & vbCrLf & "  " & deviceName & "." & propName & _
		          " - the document would not take " & tagPath & "."
		Exit Sub
	End If

	gBound = gBound + 1

End Sub


' Said, logged, and the command cleared, for a run that got nowhere.
Sub Fail(message, ts)

	DocString = EXIT_FAILURE
	WriteLog message
	WriteEx Empty, ts

End Sub


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Name & "] - " & message
	End If

End Sub

<xatm_config_data.Config.BindPositions:BindPositions_OnStartRunning()>
Sub BindPositions_OnStartRunning()

	DocString = ""

End Sub


<xatm_config_data.Config.FindObjects:FindObjects_OnChangedValue()>
Sub FindObjects_OnChangedValue()

	If IsEmpty(Value) Or IsNull(Value) Then Exit Sub   ' self-cleared write, ignore
	If Trim(CStr(Value)) = "" Then Exit Sub

	Dim ts
	ts = TimeStamp                         ' preserved for the silent clear

	' What to look for, as a class name. One scanner rather than one per
	' class: the walk is the whole domain either way, and two of them
	' would be two things to keep in step for no gain.
	gWanted = UCase(Trim(CStr(Value)))

	' Where the answer is left. A DocString says how it went and nothing
	' more, so what was found travels on a tag of its own - the way
	' SetProperty takes its payload off PropertyValue rather than out of
	' the command.
	Dim answerTag
	Set answerTag = Nothing
	On Error Resume Next
	Set answerTag = Parent.Item(ANSWER_TAG)
	On Error Goto 0

	If answerTag Is Nothing Then
		DocString = EXIT_FAILURE
		WriteLog "There is no " & ANSWER_TAG & " tag to answer on."
		WriteEx Empty, ts
		Exit Sub
	End If

	Set gFound = CreateObject("Scripting.Dictionary")
	Set gSeen  = CreateObject("Scripting.Dictionary")

	ScanForClass Application, 0

	' One path per line. A path holds no newline, so the caller gets its
	' list back by splitting and needs no escaping either way.
	Dim answer
	answer = Join(gFound.Keys, vbCrLf)

	answerTag.WriteEx answer

	DocString = EXIT_SUCCESS
	WriteLog "Found " & gFound.Count & " of class " & gWanted & "."

	Set gFound = Nothing
	Set gSeen  = Nothing

	WriteEx Empty, ts                      ' clear without re-firing

End Sub


' The tag the answer is left on.
Const ANSWER_TAG   = "FoundObjects"

Const EXIT_SUCCESS = "EXIT_SUCCESS"
Const EXIT_FAILURE = "EXIT_FAILURE"

' How deep the walk goes. A substation or a driver sits a few levels
' into a data server, well inside this. The cap is what stops a
' container that answers For Each with something holding itself from
' scanning forever.
Const MAX_DEPTH    = 12

' What is being looked for this run, upper-cased once so the comparison
' in the walk is a comparison and not two conversions.
Dim gWanted

' What the scan has found, and what it has already been through. Held
' here because the walk recurses and threading two dictionaries through
' every call buys nothing - the same reason the tree keeps its images.
Dim gFound
Dim gSeen


' Everything under a container, to the depth the cap allows.
'
' The whole domain is walked from Application rather than from a known
' data server: which project an object was put in is the engineer's
' choice, and a scan that only looks where this one expects would answer
' "none" for a project that is laid out differently.
Sub ScanForClass(container, depth)

	If depth > MAX_DEPTH Then Exit Sub

	' A tag, a property, anything that is not a container refuses For
	' Each - which is how the walk finds the leaves, so the error is
	' the answer here and not a fault to report.
	On Error Resume Next

	Dim item
	For Each item In container

		If Err.Number <> 0 Then
			Err.Clear
			Exit For
		End If

		Consider item, depth

		If Err.Number <> 0 Then Err.Clear

	Next

	If Err.Number <> 0 Then Err.Clear
	On Error Goto 0

End Sub


' One object: kept when it is what was asked for, walked into when not.
Sub Consider(item, depth)

	Dim className, path
	className = ""
	path      = ""

	On Error Resume Next
	className = TypeName(item)
	path      = item.PathName
	On Error Goto 0

	If UCase(className) = gWanted Then

		' An object that will not say where it is cannot be handed to
		' the caller, which has only a path to find it by again.
		If path <> "" Then
			If Not gFound.Exists(path) Then gFound.Add path, True
		End If

		' Nothing of the kind being looked for is nested in another of
		' its kind - not a substation in a substation, nor a driver in
		' a driver - so a match is not walked into.
		Exit Sub

	End If

	' Walked once however many places it is reachable from. A domain is
	' a graph rather than a tree - a server is under Application and
	' under the objects that link to it - and without this the scan
	' does the same subtree again for each way in.
	If path <> "" Then
		If gSeen.Exists(path) Then Exit Sub
		gSeen.Add path, True
	End If

	ScanForClass item, depth + 1

End Sub


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Name & "] - " & message
	End If

End Sub

<xatm_config_data.Config.FindObjects:FindObjects_OnStartRunning()>
Sub FindObjects_OnStartRunning()

	DocString = ""

End Sub

<xatm_config_data.Config.ImportXml:ImportXml_OnChangedValue()>
Sub ImportXml_OnChangedValue()

	Dim ts
	ts = TimeStamp

	Dim contentTag
	Set contentTag = Nothing
	On Error Resume Next
	Set contentTag = Parent.Item("XMLContent")
	On Error Goto 0

	If contentTag Is Nothing Then
		Fail "there is no XMLContent tag to read from.", ts
		Exit Sub
	End If

	Dim doc
	Set doc = NewDomDocument()

	If Not doc.loadXML(StripBom(CStr(contentTag.Value))) Then
		Fail "XMLContent does not parse - " & OneLine(doc.parseError.reason), ts
		Exit Sub
	End If

	If doc.documentElement.nodeName <> ROOT_ELEMENT Then
		Fail "XMLContent is not a " & ROOT_ELEMENT & " document.", ts
		Exit Sub
	End If

	Dim report, problem
	report  = ""
	problem = ""

	Dim root
	root = doc.documentElement.getAttribute("root")

	Dim child
	For Each child In doc.documentElement.childNodes
		If IsFolderElement(child) Then
			ImportFolder child, root & "." & child.getAttribute("name"), report, problem
		End If
	Next

	If problem <> "" Then
		Fail problem, ts
		Exit Sub
	End If

	If report = "" Then
		WriteLog "The project already matched the document."
	Else
		WriteLog "Written into the project:" & report
	End If

	DocString = EXIT_SUCCESS
	WriteEx Empty, ts

End Sub


Const ROOT_ELEMENT  = "xatm-config"
Const CONFIG_DATA   = "xatm_config_data"
Const HELPER_FOLDER = "PropertiesHelper"
Const NODE_ELEMENT  = 1
Const EXIT_SUCCESS  = "EXIT_SUCCESS"


' Scratch cells for the late-bound writes done through Execute, the way
' ReadProperty has one for its reads.
Dim gWriteValue
Dim gWriteObject


' Brings one folder of the project to what the document says is in it.
'
' Objects are matched on Id and not on name. The Id is the contract the
' automation goes by, so a rename stays a rename instead of reading as a
' delete and a create - which would take every tag the operator had
' associated with the old object down with it.
Sub ImportFolder(folderElement, folderPath, report, problem)

	Dim folder
	Set folder = Nothing
	On Error Resume Next
	Set folder = Application.GetObject(E3Path(folderPath))
	On Error Goto 0

	If folder Is Nothing Then
		problem = "the project has no folder at " & folderPath & "."
		Exit Sub
	End If

	' --- what the document says belongs here ---------------------------

	Dim wanted
	Set wanted = CreateObject("Scripting.Dictionary")

	Dim child, key
	For Each child In folderElement.childNodes

		If IsObjectElement(child) Then

			key = ObjectKey(PropertyValue(child, "Id"), child.getAttribute("name"))

			If wanted.Exists(key) Then
				problem = "two objects in " & folderPath & " answer to " & key & "."
				Exit Sub
			End If

			wanted.Add key, child

		End If

	Next

	' --- what the project has -----------------------------------------

	Dim present
	Set present = CreateObject("Scripting.Dictionary")

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim obj
	For Each obj In folder

		If IsDevice(obj) Then

			key = ObjectKey(IdOf(obj), obj.Name)

			If wanted.Exists(key) Then
				present.Add key, obj
			Else
				doomed.Add doomed.Count, obj.Name
			End If

		End If

	Next

	' Taken out in a second pass - a collection is not walked while what
	' is in it is being removed.
	Dim n
	For Each n In doomed.Keys

		On Error Resume Next
		folder.DeleteObject doomed(n)
		On Error Goto 0

		report = report & vbCrLf & "  removed " & doomed(n) & " from " & folderPath

	Next

	' --- bring each object to what the document says ------------------

	Dim existing
	For Each key In wanted.Keys

		Set existing = Nothing
		If present.Exists(key) Then Set existing = present(key)

		ImportObject folder, folderPath, wanted(key), existing, report, problem
		If problem <> "" Then Exit Sub

	Next

	' --- and the folders under it -------------------------------------

	For Each child In folderElement.childNodes
		If IsFolderElement(child) Then
			ImportFolder child, folderPath & "." & child.getAttribute("name"), report, problem
			If problem <> "" Then Exit Sub
		End If
	Next

End Sub


' One object: created where the project has not got it, brought to the
' name the document gives it, and then filled in either way.
Sub ImportObject(folder, folderPath, objectElement, existing, report, problem)

	Dim name, className
	name      = objectElement.getAttribute("name")
	className = objectElement.getAttribute("type")

	Dim obj
	Set obj = Nothing

	If existing Is Nothing Then

		Set obj = NewDevice(folder, className, name)

		If obj Is Nothing Then
			problem = "could not create " & name & ", a " & className & ", in " & folderPath & "."
			Exit Sub
		End If

		report = report & vbCrLf & "  created " & name & " in " & folderPath

	ElseIf existing.Name <> name Then

		' Renaming in place is tried first, because an object that keeps its
		' identity keeps everything associated with it - the tags on a
		' breaker above all. E3 may refuse it at runtime, and then the only
		' way to the new name is to build the object again, which loses them.
		Dim oldName
		oldName = existing.Name

		On Error Resume Next
		existing.Name = name
		On Error Goto 0

		If existing.Name = name Then

			Set obj = existing
			report = report & vbCrLf & "  renamed " & oldName & " to " & name & " in " & folderPath

		Else

			On Error Resume Next
			folder.DeleteObject oldName
			On Error Goto 0

			Set obj = NewDevice(folder, className, name)

			If obj Is Nothing Then
				problem = "could not build " & oldName & " again as " & name & " in " & _
				          folderPath & " - and it has been taken out."
				Exit Sub
			End If

			report = report & vbCrLf & "  built " & oldName & " again as " & name & " in " & _
			         folderPath & " - E3 would not rename it, so what was linked to it is gone"

		End If

	Else

		Set obj = existing

	End If

	Dim property
	For Each property In objectElement.selectNodes("property")
		WriteProperty obj, property, report, folderPath & "." & name
	Next

	' The document says nothing about a reading, so a reading goes back to
	' the default its class declares. Container.Save writes whatever the
	' live objects hold - so without this, a step failure latched during a
	' test, or a block left set, is saved into xatm_dados.prj as though
	' somebody had configured it that way.
	ResetReadings obj, className, report, folderPath & "." & name

End Sub


' Puts every property that is not configuration back to the value its
' class declares for it. What the document carries has already been
' written by then, so this only touches what the document has nothing to
' say about.
Sub ResetReadings(obj, className, report, where)

	Dim bag
	Set bag = ManifestOf(className)
	If bag Is Nothing Then Exit Sub

	Dim key, p
	For Each key In bag.Keys

		Set p = bag(key)

		If Not p.IsSaved() Then
			WriteValue obj, p.Name, p.DataType, p.InitialValue, report, where
		End If

	Next

End Sub


' A new object in a data folder, active so its properties can be written
' - AddObject(ClassName, [Activate], [ObjectName]).
Function NewDevice(folder, className, name)

	Set NewDevice = Nothing

	On Error Resume Next
	Set NewDevice = folder.AddObject(className, True, name)
	On Error Goto 0

End Function


' Writes one property, from whichever of the two attributes carries it.
'
' value  is what the property holds of its own - a number, a boolean, the
'        path of another object
' source is the tag it is wired to, which is where an IOTag property keeps
'        what it is worth instead
'
' A property with neither is unset in the document, and absent means leave
' it alone rather than clear it.
'
' Only value was read here for a long time, so no tag association in the
' document ever reached the project. Picking a tag by hand went unnoticed
' because that button writes the object itself as well as staging the edit;
' anything that only staged - a bulk bind, a document applied to a project
' that had not been wired yet - was dropped without a word, and the import
' then reported the project as already matching.
Sub WriteProperty(obj, property, report, where)

	Dim a
	Set a = property.getAttributeNode("value")

	If Not (a Is Nothing) Then
		WriteValue obj, property.getAttribute("name"), property.getAttribute("type"), _
		           a.value, report, where
		Exit Sub
	End If

	Dim s
	Set s = property.getAttributeNode("source")
	If s Is Nothing Then Exit Sub

	WriteSource obj, property.getAttribute("name"), property.getAttribute("type"), _
	            s.value, report, where

End Sub


' Puts one value onto an object, late-bound, the way ReadProperty takes
' one off. A property typed xatm_ holds the path of another object, so it
' is an association and has to be Set rather than assigned.
Sub WriteValue(obj, name, dataType, newValue, report, where)

	Dim failed
	failed = ""

	gWriteValue = newValue

	On Error Resume Next

	If IsObjectType(dataType) Then

		Set gWriteObject = Application.GetObject(E3Path(CStr(gWriteValue)))
		Execute "obj." & name & " = gWriteObject.PathName"

	Else

		Execute "obj." & name & " = gWriteValue"

	End If

	If Err.Number <> 0 Then
		failed = Err.Description
		Err.Clear
	End If

	On Error Goto 0

	If failed <> "" Then
		report = report & vbCrLf & " could not set " & name & " on " & where & " - " & failed
	End If

End Sub


' Wires one property to what the document gives as its source.
'
' The attribute carries two different things and the type says which -
' the same split SourceOf reads them back on:
'
'   iotag, internaltag   a tag path, and the property holds the tag
'   anything else        an expression, which lives on the object's Link
'                        for that property and not on the property
'
' A Boolean like Defective is configured by expression, so taking every
' source for a path would send an expression to be looked up as though it
' were one and report it missing.
Sub WriteSource(obj, name, dataType, path, report, where)

	' An empty path is not skipped. It is what the panel's clear button
	' leaves behind, and it means take the source off - so it goes the same
	' way a real one goes, and each of the two below knows what to do with
	' nothing.
	If IsLinkType(dataType) Then
		WireTag obj, name, CStr(path), report, where
	Else
		WireExpression obj, name, CStr(path), report, where
	End If

End Sub


' A tag onto a property that holds one.
'
' The path is written as it stands, and nothing is looked up first.
'
' It used to be resolved here so the property could be given the tag's
' own PathName, and that lookup is what kept failing on paths that were
' real - the checking was doing the harm the checking was for. The paths
' come off the catalog, which is E3's own account of what is in the
' domain, so a path on the catalog is taken to name something.
'
' What that gives up: E3 decides now. A path it will not take raises on
' the assignment and is reported below; one it takes quietly is wired
' whether or not anything is at the far end of it.
Sub WireTag(obj, name, path, report, where)

	Dim failed
	failed = ""

	gWriteValue = CStr(path)

	On Error Resume Next

	Execute "obj." & name & " = gWriteValue"

	If Err.Number <> 0 Then
		failed = Err.Description
		Err.Clear
	End If

	On Error Goto 0

	If failed <> "" Then
		report = report & vbCrLf & " could not wire " & name & " on " & where & _
		         " to " & path & " - " & failed
	End If

End Sub


' An expression onto the object's Link for the property.
'
' Made once and moved after that. CreateLink over a property that already
' carries a link is the case with no good answer here, so where there is
' one already its Source is written through instead - which is the member
' SourceOf reads the expression back off.
Sub WireExpression(obj, name, expression, report, where)

	Dim link
	Set link = Nothing

	On Error Resume Next
	Set link = obj.Links.Item(name)
	On Error Goto 0

	Dim failed
	failed = ""

	On Error Resume Next
	Err.Clear

	If link Is Nothing Then
		' Nothing wired and nothing to wire: a clear on a property that was
		' never bound, which is already the state being asked for.
		If expression = "" Then Exit Sub
		obj.Links.CreateLink name, expression
	Else
		link.Source = expression
	End If

	If Err.Number <> 0 Then
		failed = Err.Description
		Err.Clear
	End If

	On Error Goto 0

	If failed <> "" Then
		report = report & vbCrLf & " could not put " & expression & " on " & _
		         name & " of " & where & " - " & failed
	End If

End Sub


' What an object answers to across a rename - its Id where it has one,
' and its name where it has not. A TMTNM carries no Id, and SyncAutomation
' already goes by the number on the end of its name.
Function ObjectKey(id, name)

	If CStr(id) = "" Then
		ObjectKey = "name:" & LCase(name)
	Else
		ObjectKey = "id:" & CStr(id)
	End If

End Function


' The Id of an object in the project, "" where the class declares none.
Function IdOf(obj)

	IdOf = ""

	On Error Resume Next
	IdOf = CStr(obj.Id)
	On Error Goto 0

End Function


' The children of a folder the document speaks for. A class with a
' manifest is a device the export writes and the import owns; anything
' else in the folder is E3's own and is left alone.
Function IsDevice(obj)

	IsDevice = Not (ManifestOf(TypeName(obj)) Is Nothing)

End Function


' A property whose value is the path of another object rather than a
' value of its own - the manifest names the class, as xatm_Transformer.
Function IsObjectType(dataType)

	IsObjectType = (LCase(Left(CStr(dataType), 5)) = "xatm_")

End Function


' A property whose configuration is an association, not a value. Two of
' them: an IOTag is wired to a tag out in the project, an InternalTag to
' one of the object's own. Both export as a source and neither carries a
' value, because what is configured is the tag it points at.
'
' The export keeps its own copy of this and so does the layout button.
' Three copies of six lines is not a thing to be pleased about, but a
' scope cannot call into another scope's, and WriteSource has to tell a
' tag path from an expression before it can write either one.
Function IsLinkType(dataType)

	IsLinkType = False

	Select Case LCase(dataType)
		Case "iotag", "internaltag" : IsLinkType = True
	End Select

End Function


Function IsFolderElement(element)

	IsFolderElement = (element.nodeType = NODE_ELEMENT) And (element.nodeName = "folder")

End Function


Function IsObjectElement(element)

	IsObjectElement = (element.nodeType = NODE_ELEMENT) And (element.nodeName = "object")

End Function


' The value attribute of a named property - "" when the property is
' absent or unset.
Function PropertyValue(objectNode, propertyName)

	PropertyValue = ""

	Dim p
	Set p = objectNode.selectSingleNode("property[@name='" & propertyName & "']")
	If p Is Nothing Then Exit Function

	Dim a
	Set a = p.getAttributeNode("value")
	If a Is Nothing Then Exit Function

	PropertyValue = a.value

End Function


' The manifest declared for a class, or Nothing when the class has none -
' the same lookup the export does.
Function ManifestOf(className)

	Set ManifestOf = Nothing

	On Error Resume Next
	Set ManifestOf = Application.GetObject(CONFIG_DATA).Item(HELPER_FOLDER).Item(className).Value
	On Error Goto 0

End Function


' A path E3 will resolve. A name that starts with a digit or an
' underscore, or carries a space or a hyphen, has to be bracketed, and
' bracketing every piece is always right. One already bracketed is left
' alone - PathName may have bracketed it, and [[52-01]] resolves no
' better than 52-01 does.
Function E3Path(path)

	E3Path = ""
	If path = "" Then Exit Function

	Dim pieces
	pieces = SplitPath(path)

	Dim i
	For i = 0 To UBound(pieces)
		If Left(pieces(i), 1) <> "[" Then
			pieces(i) = "[" & pieces(i) & "]"
		End If
	Next

	E3Path = Join(pieces, ".")

End Function


' The pieces of a path, split on the dots between them - the ones outside
' brackets, since a bracketed name may carry a dot of its own.
Function SplitPath(path)

	Dim pieces()
	ReDim pieces(0)
	pieces(0) = ""

	Dim i, n, depth, ch
	n = 0
	depth = 0

	For i = 1 To Len(path)

		ch = Mid(path, i, 1)

		If ch = "[" Then
			depth = depth + 1
			pieces(n) = pieces(n) & ch
		ElseIf ch = "]" Then
			depth = depth - 1
			pieces(n) = pieces(n) & ch
		ElseIf ch = "." And depth = 0 Then
			n = n + 1
			ReDim Preserve pieces(n)
			pieces(n) = ""
		Else
			pieces(n) = pieces(n) & ch
		End If

	Next

	SplitPath = pieces

End Function


' Says why nothing was written, and clears the command.
Sub Fail(reason, ts)

	DocString = "EXIT_FAILURE"
	WriteLog "Not written into the project - " & reason
	WriteEx Empty, ts

End Sub


' MSXML 6 where the machine has it, the version-independent progid
' otherwise.
Function NewDomDocument()

	Set NewDomDocument = Nothing

	On Error Resume Next
	Set NewDomDocument = CreateObject("MSXML2.DOMDocument.6.0")
	On Error Goto 0

	If NewDomDocument Is Nothing Then
		Set NewDomDocument = CreateObject("MSXML2.DOMDocument")
	End If

	NewDomDocument.async = False
	NewDomDocument.preserveWhiteSpace = True

	On Error Resume Next
	NewDomDocument.setProperty "SelectionLanguage", "XPath"
	On Error Goto 0

End Function


' A byte-order mark survives a round trip through a utf-8 file and makes
' loadXML fail on what looks like perfectly good XML.
Function StripBom(text)

	StripBom = text

	If IsEmpty(text) Or IsNull(text) Then Exit Function
	If Len(text) = 0 Then Exit Function

	If AscW(Left(text, 1)) = &HFEFF Then
		StripBom = Mid(text, 2)
	End If

End Function


' A parse error reason comes with the line break MSXML puts on the end of
' it, which a one-line log entry has no room for.
Function OneLine(text)

	OneLine = Trim(Replace(Replace(CStr(text), vbCr, " "), vbLf, " "))

End Function


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Name & "] - " & message
	End If
	
End Sub

<xatm_config_data.Config.ImportXml:ImportXml_OnStartRunning()>
Sub ImportXml_OnStartRunning()
	
	DocString = ""
	
End Sub

<xatm_config_data.Config.SaveXML:SaveXML_OnChangedValue()>
Sub SaveXML_OnChangedValue()
	
	Dim ts
	ts = TimeStamp                         ' preserved for the silent clear
	
	Dim xmlContent
	xmlContent = Parent.Item("XMLContent").Value
	
	Dim filePath
	filePath = Parent.Item("XMLFilePath").Value
	
	If SaveXml(xmlContent, filePath) Then
	
		WriteLog "XML file saved to " & filePath
		
		DocString = "EXIT_SUCCESS"
	
	Else
	
		DocString = "EXIT_FAILURE"
			
	End If
		
	WriteEx Empty, ts
	
	' Sends a signal to the treeview that the document has changed, 
	' so it can refresh itself
	Parent.Item("UpdateTreeviewSignal").WriteEx True
	
End Sub

Const XML_VERSION   = "1"
Const DATA_ROOT     = "XATM_Data"
Const ROOT_ELEMENT  = "xatm-config"
Const CONFIG_DATA   = "xatm_config_data"
Const HELPER_FOLDER = "PropertiesHelper"

' Writes a document to disk as UTF-8, and answers whether it went.
' Content that is not a document is not written - it would replace the
' last good file with something nothing can read back.
Function SaveXml(content, path)

	SaveXml = False

	' Content that came back from a utf-8 file carries a mark the parser
	' rejects, and the stream puts a mark of its own in on the way out.
	Dim text
	text = StripBom(content)

	Dim problem
	problem = XmlProblem(text)

	If problem <> "" Then
		WriteLog "Not saved to " & path & " - " & problem
		Exit Function
	End If

	Dim stream
	Set stream = CreateObject("ADODB.Stream")

	On Error Resume Next

	stream.Type     = 2			' adTypeText
	stream.Charset  = "utf-8"
	stream.Open
	stream.WriteText text
	stream.SaveToFile path, 2	' adSaveCreateOverWrite
	stream.Close

	If Err.Number <> 0 Then
		problem = Err.Description
		Err.Clear
		On Error Goto 0
		WriteLog "Not saved to " & path & " - " & problem
		Exit Function
	End If

	On Error Goto 0

	SaveXml = True

End Function


' Why the content is not a document worth writing, "" when it is one.
' An export that fails part way through comes back empty or truncated,
' and either one parses as something the import can never use.
Function XmlProblem(content)

	If IsEmpty(content) Or IsNull(content) Then
		XmlProblem = "the export is empty"
		Exit Function
	End If

	If Trim(CStr(content)) = "" Then
		XmlProblem = "the export is empty"
		Exit Function
	End If

	Dim doc
	Set doc = Nothing

	On Error Resume Next
	Set doc = CreateObject("MSXML2.DOMDocument.6.0")
	On Error Goto 0

	If doc Is Nothing Then
		Set doc = CreateObject("MSXML2.DOMDocument")
	End If

	doc.async = False

	If Not doc.loadXML(CStr(content)) Then
		XmlProblem = "it does not parse as XML - " & OneLine(doc.parseError.reason) & _
		             " (line " & doc.parseError.line & ")"
		Exit Function
	End If

	If doc.documentElement.nodeName <> ROOT_ELEMENT Then
		XmlProblem = "it is not a " & ROOT_ELEMENT & " document, but a " & _
		             doc.documentElement.nodeName & " one"
		Exit Function
	End If

	XmlProblem = ""

End Function


' A byte-order mark survives a round trip through a utf-8 file and makes
' loadXML fail on what looks like perfectly good XML.
Function StripBom(text)

	StripBom = text

	If IsEmpty(text) Or IsNull(text) Then Exit Function
	If Len(text) = 0 Then Exit Function

	If AscW(Left(text, 1)) = &HFEFF Then
		StripBom = Mid(text, 2)
	End If

End Function


' A parse error reason comes with the line break MSXML puts on the end
' of it, which a one-line log entry has no room for.
Function OneLine(text)

	OneLine = Trim(Replace(Replace(CStr(text), vbCr, " "), vbLf, " "))

End Function


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Name & "] - " & message
	End If

End Sub

<xatm_config_data.Config.SaveXML:SaveXML_OnStartRunning()>
Sub SaveXML_OnStartRunning()

	DocString = ""
		
End Sub

<xatm_config_data.Config.SetProperty:SetProperty_OnChangedValue()>
Sub SetProperty_OnChangedValue()
	
	If Trim(Value) = "" Then 
		DocString = "EMPTY_VALUE"
		Exit Sub      ' self-cleared write, ignore
	End If
	
	Dim ts
	ts = TimeStamp                         ' preserved for the silent clear

	Dim command
	command = CStr(Value)

	' kind|path|name. Only tokens of ours travel here, so none of them can
	' carry the bar they are split on. The value the operator typed is not
	' among them - it comes off PropertyValue, whole and unparsed, and can
	' be anything at all.
	Dim parts
	parts = Split(command, "|")

	If UBound(parts) <> 2 Then
		Refuse "malformed command '" & command & "' (expected kind|path|name).", ts
		Exit Sub
	End If

	Dim kind, key, name
	kind = LCase(Trim(parts(0)))
	key  = Trim(parts(1))
	name = Trim(parts(2))

	Select Case kind

		Case KIND_NAME, KIND_PROPERTY, KIND_SOURCE

		Case Else
			Refuse "unknown kind '" & parts(0) & "'.", ts
			Exit Sub

	End Select

	If key = "" Then
		Refuse "the command names no object.", ts
		Exit Sub
	End If

	Dim contentTag
	Set contentTag = Nothing
	On Error Resume Next
	Set contentTag = Parent.Item("XMLContent")
	On Error Goto 0

	If contentTag Is Nothing Then
		Refuse "there is no XMLContent tag to write to.", ts
		Exit Sub
	End If

	' Written before the command that fires this, so it is already standing
	' by. A Variant of its own, so an empty one stays empty and a value
	' carrying anything at all arrives as it was typed.
	Dim valueTag
	Set valueTag = Nothing
	On Error Resume Next
	Set valueTag = Parent.Item(VALUE_TAG)
	On Error Goto 0

	If valueTag Is Nothing Then
		Refuse "there is no " & VALUE_TAG & " tag to read the value from.", ts
		Exit Sub
	End If

	Dim newValue
	newValue = valueTag.Value

	Dim doc
	Set doc = NewDomDocument()

	If Not doc.loadXML(StripBom(CStr(contentTag.Value))) Then
		Refuse "XMLContent does not parse - " & OneLine(doc.parseError.reason), ts
		Exit Sub
	End If

	Dim objectNode
	Set objectNode = FindObject(doc, key)

	If objectNode Is Nothing Then
		Refuse "the document has nothing at " & key & ".", ts
		Exit Sub
	End If

	Dim problem

	If kind = KIND_NAME Then
		problem = SetName(objectNode, newValue)
	ElseIf kind = KIND_SOURCE Then
		problem = SetSource(objectNode, name, newValue)
	Else
		problem = SetValue(objectNode, name, newValue)
	End If

	If problem <> "" Then
		Refuse problem, ts
		Exit Sub
	End If

	' The buffer the Save button commits. Nothing has reached E3 or the
	' file yet, and nothing will until it is pressed.
	contentTag.WriteEx DocumentText(doc)

	DocString = "EXIT_SUCCESS"
	WriteLog "Set " & name & " of " & key & "."

	WriteEx "", ts                         ' clear without re-firing

End Sub


Const ROOT_ELEMENT  = "xatm-config"
Const NODE_ELEMENT  = 1

' What the command is asking to change: the object itself, one of the
' properties declared on it, or where that property gets its value from.
'
' The last is what wires an IOTag. It is a kind of its own because it
' writes a different attribute: an association lives in source, not in
' value, and a tag written into value would come back as a string.
Const KIND_NAME     = "name"
Const KIND_PROPERTY = "property"
Const KIND_SOURCE   = "source"

' The tag the value is left on. It is a Variant and not part of the
' command, so nothing the operator types has to be escaped or split.
Const VALUE_TAG     = "PropertyValue"

' What marks a key as an Id rather than a path.
Const ID_PREFIX     = "id:"


' The object a key stands for, Nothing when the document has none. An
' object with an Id is addressed by it and not by its path, so a rename
' does not strand the rows that were built before it.
Function FindObject(doc, key)

	Set FindObject = Nothing

	' The Len is tested and not just the comparison. Were ID_PREFIX missing
	' from this scope it would be Empty, Len would be 0, and Left(key, 0)
	' would be "" - which equals Empty in VBScript. Every key would then
	' take the Id branch with the prefix still on the front of it, match
	' nothing, and say nothing about why. Testing the Len first costs a
	' missing const the Id keys only, and leaves the path ones working.
	If Len(ID_PREFIX) > 0 And Left(key, Len(ID_PREFIX)) = ID_PREFIX Then
		Set FindObject = doc.selectSingleNode( _
		                 "//object[property[@name='Id']/@value='" & _
		                 Mid(key, Len(ID_PREFIX) + 1) & "']")
	Else
		Set FindObject = doc.selectSingleNode("//object[@path=""" & key & """]")
	End If

End Function


' The value attribute of a named property - "" when the property is
' absent or unset.
Function PropertyValue(objectNode, propertyName)

	PropertyValue = ""

	Dim p
	Set p = objectNode.selectSingleNode("property[@name='" & propertyName & "']")
	If p Is Nothing Then Exit Function

	Dim a
	Set a = p.getAttributeNode("value")
	If a Is Nothing Then Exit Function

	PropertyValue = a.value

End Function


' Sets a declared property, or takes the attribute off when the value is
' cleared. The export drops it for a property that is unset, so writing
' value="" would say something else entirely on the way back in.
Function SetValue(objectNode, name, value)

	SetValue = ""

	Dim property
	Set property = objectNode.selectSingleNode("property[@name=""" & name & """]")

	If property Is Nothing Then
		SetValue = objectNode.getAttribute("name") & " has no property called " & name & "."
		Exit Function
	End If

	If IsEmpty(value) Or IsNull(value) Then
		property.removeAttribute "value"
	ElseIf CStr(value) = "" Then
		property.removeAttribute "value"
	Else
		property.setAttribute "value", CStr(value)
	End If

End Function


' Sets where a property gets its value from, or takes the attribute off
' when it is cleared - the source attribute, which is what an IOTag is
' wired by. The value attribute is left alone: an IOTag has none, and the
' export writes none for it.
'
' The project itself is not touched here. What wires the tag has already
' written it onto the object, because the import has no way to make an
' association; this is the document being brought level with what the
' project holds, so the panel is built from the same thing twice.
Function SetSource(objectNode, name, value)

	SetSource = ""

	Dim property
	Set property = objectNode.selectSingleNode("property[@name=""" & name & """]")

	If property Is Nothing Then
		SetSource = objectNode.getAttribute("name") & " has no property called " & name & "."
		Exit Function
	End If

	' Written empty rather than taken off, and the difference between the
	' two is the whole of how a source gets removed.
	'
	' No attribute means the property was never configured, and the import
	' has to leave those alone - otherwise every unwired IOTag in the
	' document would be cleared on its object at every save.
	'
	' source="" means somebody took the source off on purpose, and the
	' import reads it as: unwire this. The next export writes no attribute
	' at all, because by then the project agrees, and the two states fold
	' back into the one.
	If IsEmpty(value) Or IsNull(value) Then
		property.setAttribute "source", ""
	Else
		property.setAttribute "source", CStr(value)
	End If

End Function


' Renames the object and rebuilds the path it is known by, the name being
' the last piece of it.
'
' TODO: a property somewhere else in the document may hold the old path -
' an automation's Transformer does - and those are not followed yet.
Function SetName(objectNode, value)

	SetName = ""

	Dim name
	name = Trim(CStr(value) & "")

	If name = "" Then
		SetName = "an object cannot be left unnamed."
		Exit Function
	End If

	' E3 will not hold two of a name in one folder, and two objects at one
	' path reads worse in the document than it does in E3. Told apart by
	' path, which is unique, rather than by object identity, which two
	' lookups of the same node need not share.
	Dim twin
	Set twin = objectNode.parentNode.selectSingleNode("object[@name=""" & name & """]")

	If Not (twin Is Nothing) Then
		If twin.getAttribute("path") <> objectNode.getAttribute("path") Then
			SetName = "there is already an object called " & name & " in that folder."
			Exit Function
		End If
	End If

	objectNode.setAttribute "name", name
	objectNode.setAttribute "path", RebuildPath(objectNode)

End Function


' Where the object sits, as a path - the root the document was exported
' from, the folders down to it, then its own name. The same shape
' ObjectPath builds for an object the layout adds.
Function RebuildPath(objectNode)

	Dim path
	path = objectNode.getAttribute("name")

	Dim n
	Set n = objectNode.parentNode

	Do While n.nodeType = NODE_ELEMENT And n.nodeName = "folder"
		path = n.getAttribute("name") & "." & path
		Set n = n.parentNode
	Loop

	RebuildPath = objectNode.ownerDocument.documentElement.getAttribute("root") & "." & path

End Function


' Says why nothing was written, and clears the command without re-firing
' so the same edit can be sent again.
Sub Refuse(reason, ts)

	DocString = "EXIT_FAILURE"
	WriteLog "Not set - " & reason
	WriteEx "", ts

End Sub


' The document as text, headed the way ExportXml writes it - MSXML
' regenerates the declaration from its own encoding and ignores what the
' node says, so it is restored here rather than in the tree.
Function DocumentText(doc)

	Dim text
	text = doc.xml

	If Left(text, 5) = "<?xml" Then
		text = Mid(text, InStr(text, "?>") + 2)
	End If

	Do While Left(text, 1) = vbCr Or Left(text, 1) = vbLf
		text = Mid(text, 2)
	Loop

	DocumentText = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & text

End Function


' A byte-order mark survives a round trip through a utf-8 file and makes
' loadXML fail on what looks like perfectly good XML.
Function StripBom(text)

	StripBom = text

	If IsEmpty(text) Or IsNull(text) Then Exit Function
	If Len(text) = 0 Then Exit Function

	If AscW(Left(text, 1)) = &HFEFF Then
		StripBom = Mid(text, 2)
	End If

End Function


' A parse error reason comes with the line break MSXML puts on the end of
' it, which a one-line log entry has no room for.
Function OneLine(text)

	OneLine = Trim(Replace(Replace(CStr(text), vbCr, " "), vbLf, " "))

End Function


' MSXML 6 where the machine has it, the version-independent progid
' otherwise.
Function NewDomDocument()

	Set NewDomDocument = Nothing

	On Error Resume Next
	Set NewDomDocument = CreateObject("MSXML2.DOMDocument.6.0")
	On Error Goto 0

	If NewDomDocument Is Nothing Then
		Set NewDomDocument = CreateObject("MSXML2.DOMDocument")
	End If

	NewDomDocument.async = False
	NewDomDocument.preserveWhiteSpace = True

	On Error Resume Next
	NewDomDocument.setProperty "SelectionLanguage", "XPath"
	On Error Goto 0

End Function


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Name & "] - " & message
	End If
	
End Sub

<xatm_config_data.Config.SetProperty:SetProperty_OnStartRunning()>
Sub SetProperty_OnStartRunning()

	DocString = ""

End Sub

<xatm_config_data.Config.UpdateTreeviewSignal:UpdateTreeviewSignal_OnStartRunning()>
Sub UpdateTreeviewSignal_OnStartRunning()
	
	WriteEx False
		
End Sub

<xatm_config_data.Config.UpdateTreeviewSignal:UpdateTreeviewSignal_OnTrue()>
Sub UpdateTreeviewSignal_OnTrue()
	
	WriteEx False
		
End Sub

<xatm_config_data.Config.XMLBuilderAfterDelay:XMLBuilderAfterDelay_Functions()>
Sub XMLBuilderAfterDelay_Functions()
End Sub

Sub BuildXML()

	Dim xmlContent
	xmlContent = ExportXml()

	Parent.Item("XMLContent").WriteEx xmlContent
	
	Parent.Item("SaveXML").WriteEx True
	
End Sub

Const XML_VERSION   = "1"
Const DATA_ROOT     = "XATM_Data"
Const ROOT_ELEMENT  = "xatm-config"
Const CONFIG_DATA   = "xatm_config_data"
Const HELPER_FOLDER = "PropertiesHelper"


' Scratch cell for the late-bound property reads done through Execute.
Dim gReadValue


' ------------------------------------------------------------
'  PUBLIC API
' ------------------------------------------------------------

' The whole XATM_Data tree as an XML document.
Function ExportXml()

	Dim root
	Set root = Application.GetObject(DATA_ROOT)

	Dim xml
	xml = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf

	xml = xml & "<" & ROOT_ELEMENT & " version=""" & XML_VERSION & """" & _
	                    " exported=""" & IsoNow() & """" & _
	                    " root=""" & EscapeXml(root.Name) & """>" & vbCrLf

	xml = xml & ExportFolder(root, vbTab)

	xml = xml & "</" & ROOT_ELEMENT & ">" & vbCrLf

	ExportXml = xml

End Function


' ------------------------------------------------------------
'  WALK
' ------------------------------------------------------------

' Children of a folder. An object of a manifested class is exported
' and not descended into - its sub-tags are internals, not config.
Function ExportFolder(folder, indent)

	Dim xml
	xml = ""

	Dim item
	For Each item In folder

		Dim bag
		Set bag = ManifestOf(TypeName(item))

		If Not bag Is Nothing Then

			xml = xml & ExportObject(item, bag, indent)

		ElseIf TypeName(item) = "InternalTag" Then

			xml = xml & indent & "<tag name=""" & EscapeXml(item.Name) & """" & _
			      ValueAttribute(item.Value) & "/>" & vbCrLf

		Else

			' Not a device and not a tag - treat it as a subfolder.
			' A folder E3 will not enumerate simply comes back empty.
			Dim inner
			inner = ""

			On Error Resume Next
			inner = ExportFolder(item, indent & vbTab)
			On Error Goto 0

			If inner = "" Then
				xml = xml & indent & "<folder name=""" & EscapeXml(item.Name) & """/>" & vbCrLf
			Else
				xml = xml & indent & "<folder name=""" & EscapeXml(item.Name) & """>" & vbCrLf
				xml = xml & inner
				xml = xml & indent & "</folder>" & vbCrLf
			End If

		End If

	Next

	ExportFolder = xml

End Function


' One object, in the order its properties are declared in the manifest.
Function ExportObject(obj, bag, indent)

	Dim xml
	xml = indent & "<object name=""" & EscapeXml(obj.Name) & """" & _
	                    " type=""" & EscapeXml(TypeName(obj)) & """" & _
	                    " path=""" & EscapeXml(obj.PathName) & """>" & vbCrLf

	Dim key
	For Each key In bag.Keys

		Dim p
		Set p = bag(key)

		' Where the value comes from, and the value itself. A property has
		' one or the other: an IOTag is an association and carries no value
		' of its own, and a bound property's value is whatever its expression
		' last worked out to.
		Dim value, source
		source = SourceOf(obj, p.Name, p.DataType)
		
		If IsLinkType(p.DataType) Then
			value = Empty
		Else
			value = ReadProperty(obj, p.Name)
		End If

		xml = xml & indent & vbTab & "<property name=""" & EscapeXml(p.Name) & """" & _
		                                  " type=""" & EscapeXml(p.DataType) & """" & _
		                                  SourceAttribute(source) & _
		                                  ValueAttribute(value) & "/>" & vbCrLf

	Next

	xml = xml & indent & "</object>" & vbCrLf

	ExportObject = xml

End Function


' Where a property gets its value from, "" when it is a value of its own.
'
' Two shapes, and they are read differently. An IOTag property holds the
' tag object itself, so its source is that object's path - which is what
' ReadProperty already hands back for an object-valued property, and what
' the export used to throw away. Anything else may carry an E3 Link, and
' a Link's Source is the expression the operator wrote.
Function SourceOf(obj, propertyName, dataType)

	SourceOf = ""

	If IsLinkType(dataType) Then

		SourceOf = ReadProperty(obj, propertyName)

	Else

		On Error Resume Next
		SourceOf = obj.Links.Item(propertyName).Source
		On Error Goto 0

	End If

	If IsEmpty(SourceOf) Or IsNull(SourceOf) Then SourceOf = ""

End Function


' The manifest declared for a class, or Nothing when the class has
' none - which is how a folder tells itself apart from an object.
Function ManifestOf(className)

	Set ManifestOf = Nothing

	On Error Resume Next
	Set ManifestOf = Application.GetObject(CONFIG_DATA).Item(HELPER_FOLDER).Item(className).Value
	On Error Goto 0

End Function


' A property whose configuration is an association, not a value. Two of
' them: an IOTag is wired to a tag out in the project, an InternalTag to
' one of the object's own - the automation's command interface. Both export as a
' source and neither carries a value, because what is configured is the
' tag it points at, and whatever that tag happens to be holding at the
' time is not configuration at all.
Function IsLinkType(dataType)

	IsLinkType = False

	Select Case LCase(dataType)
		Case "iotag", "internaltag" : IsLinkType = True
	End Select

End Function

' Value of a property, read late-bound because E3 has no
' GetProperty. An object-typed property exports as the path of
' the object it points at.
Function ReadProperty(obj, propertyName)

	gReadValue = Empty

	On Error Resume Next

	Execute "Set gReadValue = obj." & propertyName

	If Err.Number <> 0 Then
		Err.Clear
		Execute "gReadValue = obj." & propertyName
	End If

	On Error Goto 0

	If IsObject(gReadValue) Then

		If gReadValue Is Nothing Then
			ReadProperty = Empty
		Else
			ReadProperty = gReadValue.PathName
		End If

	Else

		ReadProperty = gReadValue

	End If

End Function


' ------------------------------------------------------------
'  XML
' ------------------------------------------------------------

' The source attribute, dropped when the property has none - so a
' property with a value of its own reads exactly as it did before.
'
' Escaped like everything else: an expression is the first thing in this
' document at all likely to carry a < or an &.
Function SourceAttribute(source)

	SourceAttribute = ""

	If IsNull(source) Then Exit Function
	If IsEmpty(source) Then Exit Function
	If CStr(source) = "" Then Exit Function

	SourceAttribute = " source=""" & EscapeXml(CStr(source)) & """"

End Function


' The value attribute, dropped when the property is unset so that
' an absent attribute reads back as Empty on import.
Function ValueAttribute(value)

	ValueAttribute = ""

	If IsNull(value) Then Exit Function
	If IsEmpty(value) Then Exit Function

	ValueAttribute = " value=""" & EscapeXml(CStr(value)) & """"

End Function


' & first, or the escapes introduced below would be escaped again.
Function EscapeXml(text)

	Dim s
	s = CStr(text)

	s = Replace(s, "&", "&amp;")
	s = Replace(s, "<", "&lt;")
	s = Replace(s, ">", "&gt;")
	s = Replace(s, """", "&quot;")
	s = Replace(s, "'", "&apos;")

	EscapeXml = s

End Function


Function IsoNow()

	Dim d
	d = Now

	IsoNow = Year(d) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d)) & _
	         "T" & Pad2(Hour(d)) & ":" & Pad2(Minute(d)) & ":" & Pad2(Second(d))

End Function


Function Pad2(n)

	Pad2 = Right("0" & n, 2)

End Function

Sub Foo()
	
End Sub

<xatm_config_data.Config.XMLBuilderAfterDelay:XMLBuilderAfterDelay_OnStartRunning()>
Sub XMLBuilderAfterDelay_OnStartRunning()

	Value = 3	' delay in seconds
	
End Sub

<xatm_config_data.Config.XMLBuilderAfterDelay:XMLBuilderAfterDelay_TickCountdown()>
Sub XMLBuilderAfterDelay_TickCountdown()
	
	If Value > 0 Then
			
		Value = Value - 1
	
	Else
		
		' stop
		Value = -1
		
		BuildXML
		
	End If
	
End Sub

<xatm_config_data.Config.XMLFilePath:XMLFilePath_OnStartRunning()>
Sub XMLFilePath_OnStartRunning()

	Dim filePath
	filePath = ConfigFilePath()

	' Somewhere writable rather than nowhere. This tag holds no
	' configured path to fall back on - it is worked out fresh at every
	' start - so a start that works nothing out leaves the export with
	' nowhere at all to go.
	If filePath = "" Then filePath = FallbackFilePath()

	If filePath = "" Then
		WriteLog "Neither the data project nor a temporary folder could be " & _
		         "had, so " & Name & " is empty and the export has nowhere " & _
		         "to be written."
		Exit Sub
	End If

	WriteEx filePath

	WriteLog Name & " is " & filePath

End Sub


' What is looked for, and what is written beside it.
'
' The data project is found rather than configured because it is not the
' same folder twice: xatm_data.prj sits in a folder of its own per
' substation - se_gul for Guarulhos - and a path typed in by hand is
' right until this configuration is copied to the next site.
Const DATA_PROJECT_FILE = "xatm_data.prj"
Const XML_FILE_NAME     = "AutomationConfig.xml"

' Where the export goes when the data project was not found -
' GetSpecialFolder's TemporaryFolder.
Const TEMPORARY_FOLDER  = 2

' How far under the starting folder to look. The domain file sits at the
' top and each substation is a folder under it, so the project is one
' level down - and a couple more costs nothing and allows for a project
' laid out with a little more nesting.
Const MAX_SEARCH_DEPTH  = 3

' Every folder holding a data project, in the order they were come
' across. A dictionary rather than a single answer so that a domain
' carrying two of them is something that can be said out loud rather
' than silently resolved to whichever was met first.
Dim gFound


' The full path to write the export to, "" when the data project was not
' found.
'
' Asked of the file system rather than of E3. A container knows where it
' was loaded from, but not under a property name written down anywhere
' that could be found - so this goes the plain way instead.
Function ConfigFilePath()

	ConfigFilePath = ""

	Dim base
	base = StartFolder()

	If base = "" Then
		WriteLog "The current directory could not be read, so " & _
		         DATA_PROJECT_FILE & " was not looked for."
		Exit Function
	End If

	Dim fso
	Set fso = Nothing
	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	On Error Goto 0

	If fso Is Nothing Then
		WriteLog "There is no file system object to look with."
		Exit Function
	End If

	If Not fso.FolderExists(base) Then
		WriteLog "The current directory, " & base & ", is not a folder."
		Exit Function
	End If

	Set gFound = CreateObject("Scripting.Dictionary")

	SearchFolder fso.GetFolder(base), 0

	If gFound.Count = 0 Then
		WriteLog DATA_PROJECT_FILE & " was not found under " & base & _
		         " within " & MAX_SEARCH_DEPTH & " folders."
		Set gFound = Nothing
		Exit Function
	End If

	Dim folders
	folders = gFound.Keys

	' Said out loud and not resolved quietly. Two data projects under one
	' domain is a domain this configuration cannot speak for, and writing
	' the export into whichever folder was met first would be a choice
	' nobody made.
	If gFound.Count > 1 Then
		WriteLog "More than one " & DATA_PROJECT_FILE & " is under " & base & _
		         " - " & Join(folders, ", ") & ". The first is being used, " & _
		         "which may not be the one meant."
	End If

	ConfigFilePath = fso.BuildPath(folders(0), XML_FILE_NAME)

	Set gFound = Nothing

End Function


' Somewhere the export can always be written, for when the data project
' was not found. "" only when there is no temporary folder either, which
' is a machine with larger problems than this one.
'
' The file turning up in TEMP is itself the sign that something went
' wrong - nobody would put it there on purpose - and the line logged
' just above it says which of the ways it got there.
Function FallbackFilePath()

	FallbackFilePath = ""

	Dim fso
	Set fso = Nothing
	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	On Error Goto 0

	If fso Is Nothing Then Exit Function

	Dim tempFolder
	tempFolder = ""

	On Error Resume Next
	tempFolder = fso.GetSpecialFolder(TEMPORARY_FOLDER).Path
	On Error Goto 0

	If Trim(tempFolder & "") = "" Then Exit Function

	FallbackFilePath = fso.BuildPath(tempFolder, XML_FILE_NAME)

	WriteLog "Falling back to the temporary folder - " & FallbackFilePath

End Function


' Where E3 was started from, which is the folder the domain file sits in.
Function StartFolder()

	StartFolder = ""

	Dim shell
	Set shell = Nothing

	On Error Resume Next
	Err.Clear

	Set shell = CreateObject("WScript.Shell")
	If Not shell Is Nothing Then StartFolder = Trim(shell.CurrentDirectory & "")

	If Err.Number <> 0 Then
		StartFolder = ""
		Err.Clear
	End If

	On Error Goto 0

End Function


' One folder and what is under it, to the depth the cap allows. A folder
' that will not be read is stepped over rather than reported: a domain
' folder holds all sorts, and one of them being unreadable says nothing
' about where the data project is.
Sub SearchFolder(folder, depth)

	If depth > MAX_SEARCH_DEPTH Then Exit Sub

	On Error Resume Next

	Dim theFile
	For Each theFile In folder.Files
		If LCase(theFile.Name) = LCase(DATA_PROJECT_FILE) Then
			If Not gFound.Exists(folder.Path) Then gFound.Add folder.Path, True
		End If
	Next

	If Err.Number <> 0 Then Err.Clear

	Dim child
	For Each child In folder.SubFolders
		If Not Skipped(child) Then SearchFolder child, depth + 1
	Next

	If Err.Number <> 0 Then Err.Clear

	On Error Goto 0

End Sub


' Folders not worth walking into. A working copy carries .git, which
' holds thousands of files and no data project, and E3 puts nothing of
' ours anywhere hidden.
Function Skipped(folder)

	Skipped = True

	Dim folderName
	folderName = ""

	On Error Resume Next
	folderName = folder.Name
	On Error Goto 0

	If folderName = "" Then Exit Function
	If Left(folderName, 1) = "." Then Exit Function

	Dim attributes
	attributes = 0

	On Error Resume Next
	attributes = folder.Attributes
	On Error Goto 0

	If (attributes And 2) <> 0 Then Exit Function     ' hidden
	If (attributes And 4) <> 0 Then Exit Function     ' system

	Skipped = False

End Function


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Name & "] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Name & "] - " & message
	End If
	
End Sub

<xatm_config_data.PropertiesHelper.xatm_Breaker:xatm_Breaker_OnStartRunning()>
Sub xatm_Breaker_OnStartRunning()

	Dim bag
	Set bag = CreateObject("Scripting.Dictionary")

	AddProperty bag, "Id", "Integer", Empty, _
		"Unique numeric identifier used by the automation to locate this breaker.", _
		"Identificador numérico único usado pelo automatismo para localizar este disjuntor."

	AddProperty bag, "PositionOpen", "IOTag", Empty, _
		"Open-position input of relay 1. Double-point word, or the 52b contact when UseDoublePoints is False.", _
		"Entrada de posição aberto do relé 1. Palavra de duplo ponto, ou contato 52b quando UseDoublePoints for False."

	AddProperty bag, "PositionClosed", "IOTag", Empty, _
		"Closed-position input of relay 1. Double-point word, or the 52a contact when UseDoublePoints is False.", _
		"Entrada de posição fechado do relé 1. Palavra de duplo ponto, ou contato 52a quando UseDoublePoints for False."

	AddProperty bag, "PositionOpenAlt", "IOTag", Empty, _
		"Open-position input of relay 2. Leave empty when there is no redundant relay.", _
		"Entrada de posição aberto do relé 2. Deixe vazio quando não houver relé redundante."

	AddProperty bag, "PositionClosedAlt", "IOTag", Empty, _
		"Closed-position input of relay 2. Leave empty when there is no redundant relay.", _
		"Entrada de posição fechado do relé 2. Deixe vazio quando não houver relé redundante."

	AddProperty bag, "UseDoublePoints", "Boolean", True, _
		"True - position comes from a double-point word. False - position comes from the 52a/52b contact pair.", _
		"True - a posição vem de uma palavra de duplo ponto. False - a posição vem do par de contatos 52a/52b."

	AddProperty bag, "RawValueOpen", "Integer", 2, _
		"Raw value from the driver that means OPEN.", _
		"Valor bruto do driver que significa ABERTO."

	AddProperty bag, "RawValueClosed", "Integer", 1, _
		"Raw value from the driver that means CLOSED.", _
		"Valor bruto do driver que significa FECHADO."

	AddProperty bag, "NormalState", "EdbSwitchState", 1, _
		"Normal state of the equipment. Reference for normalisation and for Simulation Mode.", _
		"Estado normal do equipamento. Referência para a normalização e para o Modo Simulação."

	AddProperty bag, "CommandTimeout", "Integer", 40, _
		"Command supervision window in seconds. The command is re-sent once at half the window and fails when it expires.", _
		"Janela de supervisão do comando em segundos. O comando é reenviado uma vez na metade da janela e falha ao expirar."

	AddProperty bag, "RawValueCommandOpen", "Integer", 0, _
		"Raw value written to the open (trip) output.", _
		"Valor bruto escrito na saída de abertura (trip)."

	AddProperty bag, "RawValueCommandClose", "Integer", 1, _
		"Raw value written to the close output.", _
		"Valor bruto escrito na saída de fechamento."

	AddProperty bag, "CommandOpen", "IOTag", Empty, _
		"Open (trip) output of relay 1. Falls back to relay 2 when empty.", _
		"Saída de abertura (trip) do relé 1. Recorre ao relé 2 quando vazio."

	AddProperty bag, "CommandSBOOpen", "IOTag", Empty, _
		"Select tag for the open command on relay 1. Optional - Select-Before-Operate only runs when filled.", _
		"Tag de seleção do comando de abertura no relé 1. Opcional - o Select-Before-Operate só ocorre quando preenchido."

	AddProperty bag, "CommandOpenAlt", "IOTag", Empty, _
		"Open (trip) output of relay 2. Used when relay 1 is unavailable.", _
		"Saída de abertura (trip) do relé 2. Usada quando o relé 1 está indisponível."

	AddProperty bag, "CommandSBOOpenAlt", "IOTag", Empty, _
		"Select tag for the open command on relay 2. Optional.", _
		"Tag de seleção do comando de abertura no relé 2. Opcional."

	AddProperty bag, "CommandClose", "IOTag", Empty, _
		"Close output of relay 1. Falls back to relay 2 when empty.", _
		"Saída de fechamento do relé 1. Recorre ao relé 2 quando vazio."

	AddProperty bag, "CommandSBOClose", "IOTag", Empty, _
		"Select tag for the close command on relay 1. Optional - Select-Before-Operate only runs when filled.", _
		"Tag de seleção do comando de fechamento no relé 1. Opcional - o Select-Before-Operate só ocorre quando preenchido."

	AddProperty bag, "CommandCloseAlt", "IOTag", Empty, _
		"Close output of relay 2. Used when relay 1 is unavailable.", _
		"Saída de fechamento do relé 2. Usada quando o relé 1 está indisponível."

	AddProperty bag, "CommandSBOCloseAlt", "IOTag", Empty, _
		"Select tag for the close command on relay 2. Optional.", _
		"Tag de seleção do comando de fechamento no relé 2. Opcional."
	
	' --- what the plant and the automation say back ---------------------
	'
	' None of the four is configured with a value. Defective is configured
	' with an expression, and the other three are written by the breaker
	' and by the automation as they go.

	AddProperty bag, "Defective", "Boolean", False, _
		"Equipment is defective and must not be operated. Bound to an expression - not in remote, spring discharged, whatever the panel reports.", _
		"Equipamento com defeito e que não deve ser operado. Vinculada a uma expressão - fora de remoto, mola descarregada, o que o painel indicar."

	' Whether the bay is carrying load, which is how a reclosing scheme
	' confirms a close that the position contacts did not report.
	'
	' A boolean and not three currents: the only question ever asked of it
	' is whether the sum is above a threshold, and the specification and the
	' reference logic disagree on that threshold - 0,1 pu against 1 A. Left
	' as an expression, that stays a site decision instead of one the
	' library has to arbitrate.
	AddProperty bag, "HasLoadCurrent", "Boolean", False, _
		"Load current is flowing through this breaker. Bound to an expression - the sum of the phase currents above whatever threshold the substation uses.", _
		"Há corrente de carga neste disjuntor. Vinculada a uma expressão - a soma das correntes de fase acima do limiar que a subestação usar."

	AddProperty bag, "Position", "Integer", 0, _
		"Position as the automation reads it: 1 open, 2 closed, 0 neither. Always these three, whatever raw values the protocol carries - the raw ones are configured above and translated into this.", _
		"Posição como o automatismo a lê: 1 aberto, 2 fechado, 0 indefinido. Sempre estes três, quaisquer que sejam os valores brutos do protocolo - os brutos são configurados acima e traduzidos para esta."

	AddProperty bag, "CommandOpenFailed", "Boolean", False, _
		"Latched failure of an open command sent by the automation. Set when the breaker does not reach the open position inside CommandTimeout, cleared by Reset.", _
		"Falha selada de um comando de abertura enviado pelo automatismo. Marcada quando o disjuntor não atinge a posição aberta dentro de CommandTimeout, apagada pelo Reset."

	AddProperty bag, "CommandCloseFailed", "Boolean", False, _
		"Latched failure of a close command sent by the automation. Set when the breaker does not reach the closed position inside CommandTimeout, cleared by Reset.", _
		"Falha selada de um comando de fechamento enviado pelo automatismo. Marcada quando o disjuntor não atinge a posição fechada dentro de CommandTimeout, apagada pelo Reset."

	' The command outputs are write-only, so they get no EXPOSE_VALUE -
	' there is nothing to read back. Forcing one sends the raw value
	' configured for it rather than flipping a boolean, because a protocol
	' may want 65 to close.
	'
	' The select outputs are forced the same way and for the same
	' reason: where the IED wants select before operate, a command
	' cannot be tried from the panel at all unless the select can be
	' sent first.
	'
	' Id is shown and never edited: it is what the automation locates this
	' breaker by, and what the panel and the import address it by.
	SetExposure bag, "Id",                   EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED

	SetExposure bag, "PositionOpen",         EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "PositionClosed",       EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "PositionOpenAlt",      EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "PositionClosedAlt",    EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED

	SetExposure bag, "UseDoublePoints",      EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueOpen",         EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueClosed",       EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "NormalState",          EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "CommandTimeout",       EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueCommandOpen",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueCommandClose", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED

	SetExposure bag, "CommandOpen",          EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOOpen",       EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandOpenAlt",       EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOOpenAlt",    EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandClose",         EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOClose",      EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandCloseAlt",      EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOCloseAlt",   EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	
	
	' Defective is configured the way a transformer's relays are: an
	' expression is written for it on the panel, it can be forced for a
	' test, and it is never saved - the expression is the configuration
	' and the reading is whatever the switchyard was doing at the time.
	'
	' No EXPOSE_IOTAG. It is derived rather than measured - an expression
	' over what the panel reports, not in remote, spring discharged - and
	' level 3 is already told those raw conditions. A point for it would
	' put the same condition on the operation centre's list twice, which
	' is the same reason it is off the alarm table below.
	SetExposure bag, "Defective", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE

	' Position is shown and never touched. Reading it back on the panel is
	' how an engineer sees that the raw values above were configured the
	' right way round, which is worth a row of its own; forcing it would
	' only make the panel lie about the switchyard.
	' Exposed the way Defective is: an expression is the configuration, a
	' force is how it gets tested without a real trip, and it is not saved.
	'
	' No EXPOSE_IOTAG, for Defective's reason: it is derived, and level 3
	' is already told the raw currents it is derived from.
	SetExposure bag, "HasLoadCurrent", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE

	SetExposure bag, "Position", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_INTERFACE

	' The two command failures are latches the automation sets and Reset
	' clears - the breaker's own StepExecutionFailed. Nothing about them
	' is configuration, so they are interfaced and nothing else.
	SetExposure bag, "CommandOpenFailed",  EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "CommandCloseFailed", EXPOSE_INTERFACE + EXPOSE_IOTAG
	
	
	' --- what the operator is alarmed on --------------------------------
	'
	' The two command failures and nothing else.
	'
	' Defective is deliberately off the list: the PowerSubstation alarms
	' that condition already, and a second source on it puts one event on
	' the operator's list twice. Position is a state rather than an event
	' - the panel draws it, and an alarm per position change would be
	' noise the control room learns to ignore.
	SetAlarm bag, "CommandOpenFailed",  "FALHA ABERTURA",   PAIR_ACTUATED, SEV_HIGH
	SetAlarm bag, "CommandCloseFailed", "FALHA FECHAMENTO", PAIR_ACTUATED, SEV_HIGH
		
	Set Value = bag

End Sub

' What the configuration screen may do with a property, and whether its
' value is a setting at all. A bitmask: a property can be bound to an
' expression and forced, or shown and not edited, and so on.
'
' AddProperty leaves every property at EXPOSE_NONE, so nothing appears on
' the screen and nothing is written to the project until the table at the
' foot of the manifest says so. Both defaults fail closed.
Const EXPOSE_NONE       = 0
Const EXPOSE_VIEW       = 1     ' a row appears for it
Const EXPOSE_VALUE      = 2     ' its value is shown - never for a write-only command
Const EXPOSE_EDIT       = 4     ' the value can be typed
Const EXPOSE_EXPRESSION = 8     ' it can be bound to an expression
Const EXPOSE_FORCE      = 16    ' it can be forced at runtime, and is never saved
Const EXPOSE_SAVED      = 32    ' its value is configuration, not a reading
Const EXPOSE_INTERFACE  = 64    ' the Elipse application is given a tag for it
Const EXPOSE_IOTAG      = 128   ' level 3 is given a point for it, over 104


' What the operator reads, and which state says it: normal|active|limit.
'
' One string rather than three fields because the polarity belongs with
' the words. Preconditions is True while the maneuver is permitted, so
' it alarms on 0 where the rest alarm on 1 - and a pair that carried
' only the two words would let someone reword the message without ever
' seeing that the state raising it was the healthy one.
Const PAIR_ACTUATED     = "NORMAL|ATUADO|1"
Const PAIR_BLOCKED      = "LIBERADO|BLOQUEADO|1"
Const PAIR_PRECONDITION = "ATENDIDAS|NÃO ATENDIDAS|0"
Const PAIR_RUNNING      = "CONCLUÍDO|EM ANDAMENTO|1"

' DigitalSeverity, as Power numbers it.
'
' The scale runs backwards: the smaller the number the worse the alarm,
' and -2 is the most severe value here rather than the least. Anything
' comparing two of these has to be read twice - "worse than medium" is
' a < and not a >. The manifests ask for medium unless a signal earns
' otherwise; the overlay is what moves a single alarm off its default.
Const SEV_CRITICAL = -2
Const SEV_HIGH     =  0
Const SEV_MEDIUM   =  1
Const SEV_LOW      =  2

Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	' What the operator is told when this property changes, and which of
	' its two states does the telling. Empty on every property until the
	' alarm table at the foot of the manifest names it - the same way the
	' exposure table is what makes a property appear on the panel. Both
	' default to silence.
	Public AlarmLabel
	Public AlarmPair
	Public AlarmSeverity
	
	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

	End Function

	' Asked of the property rather than of the caller, so the flags stay in
	' the one scope that declares them. The instances travel to whatever
	' scope reads the manifest and answer there just the same.
	Public Function Shows()
		Shows = Has(EXPOSE_VIEW)
	End Function

	Public Function ShowsValue()
		ShowsValue = Has(EXPOSE_VALUE)
	End Function

	Public Function CanEdit()
		CanEdit = Has(EXPOSE_EDIT)
	End Function

	Public Function CanBind()
		CanBind = Has(EXPOSE_EXPRESSION)
	End Function

	Public Function CanForce()
		CanForce = Has(EXPOSE_FORCE)
	End Function

	Public Function IsSaved()
		IsSaved = Has(EXPOSE_SAVED)
	End Function
	
	Public Function IsInterfaced()
		IsInterfaced = Has(EXPOSE_INTERFACE)
	End Function

	' Whether level 3 is given a point for it.
	'
	' A separate question from IsInterfaced, and asked separately. The
	' interface is where the Elipse application meets the automation; the
	' distribution is what leaves the station. Everything distributed is
	' interfaced - the distribution reads off the interface - but not
	' everything interfaced is distributed, and conflating the two left no
	' way to say so except a list of names kept somewhere else.
	Public Function IsIOTagged()
		IsIOTagged = Has(EXPOSE_IOTAG)
	End Function

	' An unnamed property raises nothing. Empty and "" compare equal in
	' VBScript, so a property the alarm table never mentions answers no
	' here without needing a flag of its own.
	Public Function IsAlarmed()
		IsAlarmed = (AlarmLabel <> "")
	End Function

	' The message either way, in the pattern the control room reads:
	' a label and the state, joined by a dash.
	Public Function AlarmNormalText()
		AlarmNormalText = AlarmLabel & " - " & PairPart(0)
	End Function

	Public Function AlarmActiveText()
		AlarmActiveText = AlarmLabel & " - " & PairPart(1)
	End Function

	Public Function AlarmLimit()
		AlarmLimit = (CLng("0" & PairPart(2)) <> 0)
	End Function

	Private Function PairPart(i)

		PairPart = ""

		Dim parts
		parts = Split(AlarmPair & "", "|")

		If i <= UBound(parts) Then PairPart = parts(i)

	End Function
	
	' Empty And anything is 0, so a property nobody classified answers no
	' to all of these.
	Private Function Has(flag)
		Has = ((Exposure And flag) <> 0)
	End Function

End Class

Sub AddProperty(bag, name, dataType, initialValue, helpEn, helpPt)

	Dim p
	Set p = New PropertyInfo

	p.Name         = name
	p.DataType     = dataType
	p.InitialValue = initialValue
	p.Exposure     = EXPOSE_NONE
	p.HelpEn       = helpEn
	p.HelpPt       = helpPt

	bag.Add LCase(name), p
	
End Sub

' What the screen may do with a property. Set apart from AddProperty so
' the classifications read as a table, and so changing one never means
' touching the help text - which is where the accents live.
' What the operator is alarmed on. Set apart from SetExposure for the
' reason that one is set apart from AddProperty: the alarms read as a
' table of their own, and a property left out of it raises nothing.
'
' A curated list and never a sweep of what is interfaced. An interface
' tag exists so a screen can draw a value, which is a different question
' from whether an operator should be told about it.
Sub SetAlarm(bag, propertyName, label, pair, severity)

	Dim k
	k = LCase(propertyName)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).AlarmLabel    = label
	bag(k).AlarmPair     = pair
	bag(k).AlarmSeverity = severity

End Sub
Sub SetExposure(bag, name, exposure)

	Dim k
	k = LCase(name)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).Exposure = exposure
		
End Sub

<xatm_config_data.PropertiesHelper.xatm_Disconnector:xatm_Disconnector_OnStartRunning()>
Sub xatm_Disconnector_OnStartRunning()

	Dim bag
	Set bag = CreateObject("Scripting.Dictionary")

	AddProperty bag, "Id", "Integer", Empty, _
		"Unique numeric identifier used by the automation to locate this breaker.", _
		"Identificador numérico único usado pelo automatismo para localizar este disjuntor."

	AddProperty bag, "PositionOpen", "IOTag", Empty, _
		"Open-position input of relay 1. Double-point word, or the 52b contact when UseDoublePoints is False.", _
		"Entrada de posição aberto do relé 1. Palavra de duplo ponto, ou contato 52b quando UseDoublePoints for False."

	AddProperty bag, "PositionClosed", "IOTag", Empty, _
		"Closed-position input of relay 1. Double-point word, or the 52a contact when UseDoublePoints is False.", _
		"Entrada de posição fechado do relé 1. Palavra de duplo ponto, ou contato 52a quando UseDoublePoints for False."

	AddProperty bag, "PositionOpenAlt", "IOTag", Empty, _
		"Open-position input of relay 2. Leave empty when there is no redundant relay.", _
		"Entrada de posição aberto do relé 2. Deixe vazio quando não houver relé redundante."

	AddProperty bag, "PositionClosedAlt", "IOTag", Empty, _
		"Closed-position input of relay 2. Leave empty when there is no redundant relay.", _
		"Entrada de posição fechado do relé 2. Deixe vazio quando não houver relé redundante."

	AddProperty bag, "UseDoublePoints", "Boolean", True, _
		"True - position comes from a double-point word. False - position comes from the 52a/52b contact pair.", _
		"True - a posição vem de uma palavra de duplo ponto. False - a posição vem do par de contatos 52a/52b."

	AddProperty bag, "RawValueOpen", "Integer", 2, _
		"Raw value from the driver that means OPEN.", _
		"Valor bruto do driver que significa ABERTO."

	AddProperty bag, "RawValueClosed", "Integer", 1, _
		"Raw value from the driver that means CLOSED.", _
		"Valor bruto do driver que significa FECHADO."

	AddProperty bag, "NormalState", "EdbSwitchState", 1, _
		"Normal state of the equipment. Reference for normalisation and for Simulation Mode.", _
		"Estado normal do equipamento. Referência para a normalização e para o Modo Simulação."

	' --- whether it can be operated at all ------------------------------
	'
	' A disconnector is not always motorised. One switched by hand is read
	' and never commanded, and the automation has to know the difference
	' before it puts a step on one: a sequence that waits for a hand-
	' operated disconnector to travel waits until CommandTimeout expires
	' and then reports a step failure that never had a chance of not
	' happening.

	AddProperty bag, "Motorized", "Boolean", True, _
		"True when the disconnector has a motor drive the automation can command. False for a hand-operated one, which is read and never commanded.", _
		"True quando a seccionadora tem comando motorizado que o automatismo pode acionar. False para manobra manual, que é apenas lida e nunca comandada."

	AddProperty bag, "CommandTimeout", "Integer", 40, _
		"Command supervision window in seconds. The command is re-sent once at half the window and fails when it expires.", _
		"Janela de supervisão do comando em segundos. O comando é reenviado uma vez na metade da janela e falha ao expirar."

	AddProperty bag, "RawValueCommandOpen", "Integer", 0, _
		"Raw value written to the open (trip) output.", _
		"Valor bruto escrito na saída de abertura (trip)."

	AddProperty bag, "RawValueCommandClose", "Integer", 1, _
		"Raw value written to the close output.", _
		"Valor bruto escrito na saída de fechamento."

	AddProperty bag, "CommandOpen", "IOTag", Empty, _
		"Open (trip) output of relay 1. Falls back to relay 2 when empty.", _
		"Saída de abertura (trip) do relé 1. Recorre ao relé 2 quando vazio."

	AddProperty bag, "CommandSBOOpen", "IOTag", Empty, _
		"Select tag for the open command on relay 1. Optional - Select-Before-Operate only runs when filled.", _
		"Tag de seleção do comando de abertura no relé 1. Opcional - o Select-Before-Operate só ocorre quando preenchido."

	AddProperty bag, "CommandOpenAlt", "IOTag", Empty, _
		"Open (trip) output of relay 2. Used when relay 1 is unavailable.", _
		"Saída de abertura (trip) do relé 2. Usada quando o relé 1 está indisponível."

	AddProperty bag, "CommandSBOOpenAlt", "IOTag", Empty, _
		"Select tag for the open command on relay 2. Optional.", _
		"Tag de seleção do comando de abertura no relé 2. Opcional."

	AddProperty bag, "CommandClose", "IOTag", Empty, _
		"Close output of relay 1. Falls back to relay 2 when empty.", _
		"Saída de fechamento do relé 1. Recorre ao relé 2 quando vazio."

	AddProperty bag, "CommandSBOClose", "IOTag", Empty, _
		"Select tag for the close command on relay 1. Optional - Select-Before-Operate only runs when filled.", _
		"Tag de seleção do comando de fechamento no relé 1. Opcional - o Select-Before-Operate só ocorre quando preenchido."

	AddProperty bag, "CommandCloseAlt", "IOTag", Empty, _
		"Close output of relay 2. Used when relay 1 is unavailable.", _
		"Saída de fechamento do relé 2. Usada quando o relé 1 está indisponível."

	AddProperty bag, "CommandSBOCloseAlt", "IOTag", Empty, _
		"Select tag for the close command on relay 2. Optional.", _
		"Tag de seleção do comando de fechamento no relé 2. Opcional."
	
	' --- what the plant and the automation say back ---------------------
	'
	' None of the four is configured with a value. Defective is configured
	' with an expression, and the other three are written by the breaker
	' and by the automation as they go.

	AddProperty bag, "Defective", "Boolean", False, _
		"Equipment is defective and must not be operated. Bound to an expression - not in remote, spring discharged, whatever the panel reports.", _
		"Equipamento com defeito e que não deve ser operado. Vinculada a uma expressão - fora de remoto, mola descarregada, o que o painel indicar."

	AddProperty bag, "Position", "Integer", 0, _
		"Position as the automation reads it: 1 open, 2 closed, 0 neither. Always these three, whatever raw values the protocol carries - the raw ones are configured above and translated into this.", _
		"Posição como o automatismo a lê: 1 aberto, 2 fechado, 0 indefinido. Sempre estes três, quaisquer que sejam os valores brutos do protocolo - os brutos são configurados acima e traduzidos para esta."

	AddProperty bag, "CommandOpenFailed", "Boolean", False, _
		"Latched failure of an open command sent by the automation. Set when the breaker does not reach the open position inside CommandTimeout, cleared by Reset.", _
		"Falha selada de um comando de abertura enviado pelo automatismo. Marcada quando o disjuntor não atinge a posição aberta dentro de CommandTimeout, apagada pelo Reset."

	AddProperty bag, "CommandCloseFailed", "Boolean", False, _
		"Latched failure of a close command sent by the automation. Set when the breaker does not reach the closed position inside CommandTimeout, cleared by Reset.", _
		"Falha selada de um comando de fechamento enviado pelo automatismo. Marcada quando o disjuntor não atinge a posição fechada dentro de CommandTimeout, apagada pelo Reset."

	' The command outputs are write-only, so they get no EXPOSE_VALUE -
	' there is nothing to read back. Forcing one sends the raw value
	' configured for it rather than flipping a boolean, because a protocol
	' may want 65 to close.
	'
	' Id is shown and never edited: it is what the automation locates this
	' breaker by, and what the panel and the import address it by.
	SetExposure bag, "Id",                   EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED

	SetExposure bag, "PositionOpen",         EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "PositionClosed",       EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "PositionOpenAlt",      EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "PositionClosedAlt",    EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED

	SetExposure bag, "UseDoublePoints",      EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueOpen",         EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueClosed",       EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "NormalState",          EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "Motorized",            EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "CommandTimeout",       EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueCommandOpen",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueCommandClose", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED

	SetExposure bag, "CommandOpen",          EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOOpen",       EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandOpenAlt",       EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOOpenAlt",    EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandClose",         EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOClose",      EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandCloseAlt",      EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOCloseAlt",   EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	
	
	' Defective is configured the way a transformer's relays are: an
	' expression is written for it on the panel, it can be forced for a
	' test, and it is never saved - the expression is the configuration
	' and the reading is whatever the switchyard was doing at the time.
	'
	' No EXPOSE_IOTAG. It is derived rather than measured - an expression
	' over what the panel reports, not in remote, spring discharged - and
	' level 3 is already told those raw conditions. A point for it would
	' put the same condition on the operation centre's list twice, which
	' is the same reason it is off the alarm table below.
	SetExposure bag, "Defective", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE

	' Position is shown and never touched. Reading it back on the panel is
	' how an engineer sees that the raw values above were configured the
	' right way round, which is worth a row of its own; forcing it would
	' only make the panel lie about the switchyard.
	SetExposure bag, "Position", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_INTERFACE

	' The two command failures are latches the automation sets and Reset
	' clears - the breaker's own StepExecutionFailed. Nothing about them
	' is configuration, so they are interfaced and nothing else.
	SetExposure bag, "CommandOpenFailed",  EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "CommandCloseFailed", EXPOSE_INTERFACE + EXPOSE_IOTAG


	' --- what the operator is alarmed on --------------------------------
	'
	' The two command failures and nothing else.
	'
	' Defective is deliberately off the list: the PowerSubstation alarms
	' that condition already, and a second source on it puts one event on
	' the operator's list twice. Position is a state rather than an event
	' - the panel draws it, and an alarm per position change would be
	' noise the control room learns to ignore.
	SetAlarm bag, "CommandOpenFailed",  "FALHA ABERTURA",   PAIR_ACTUATED, SEV_HIGH
	SetAlarm bag, "CommandCloseFailed", "FALHA FECHAMENTO", PAIR_ACTUATED, SEV_HIGH
	
	Set Value = bag

End Sub

' What the configuration screen may do with a property, and whether its
' value is a setting at all. A bitmask: a property can be bound to an
' expression and forced, or shown and not edited, and so on.
'
' AddProperty leaves every property at EXPOSE_NONE, so nothing appears on
' the screen and nothing is written to the project until the table at the
' foot of the manifest says so. Both defaults fail closed.
Const EXPOSE_NONE       = 0
Const EXPOSE_VIEW       = 1     ' a row appears for it
Const EXPOSE_VALUE      = 2     ' its value is shown - never for a write-only command
Const EXPOSE_EDIT       = 4     ' the value can be typed
Const EXPOSE_EXPRESSION = 8     ' it can be bound to an expression
Const EXPOSE_FORCE      = 16    ' it can be forced at runtime, and is never saved
Const EXPOSE_SAVED      = 32    ' its value is configuration, not a reading
Const EXPOSE_INTERFACE  = 64    ' the Elipse application is given a tag for it
Const EXPOSE_IOTAG      = 128   ' level 3 is given a point for it, over 104


' What the operator reads, and which state says it: normal|active|limit.
'
' One string rather than three fields because the polarity belongs with
' the words. Preconditions is True while the maneuver is permitted, so
' it alarms on 0 where the rest alarm on 1 - and a pair that carried
' only the two words would let someone reword the message without ever
' seeing that the state raising it was the healthy one.
Const PAIR_ACTUATED     = "NORMAL|ATUADO|1"
Const PAIR_BLOCKED      = "LIBERADO|BLOQUEADO|1"
Const PAIR_PRECONDITION = "ATENDIDAS|NÃO ATENDIDAS|0"
Const PAIR_RUNNING      = "CONCLUÍDO|EM ANDAMENTO|1"

' DigitalSeverity, as Power numbers it.
'
' The scale runs backwards: the smaller the number the worse the alarm,
' and -2 is the most severe value here rather than the least. Anything
' comparing two of these has to be read twice - "worse than medium" is
' a < and not a >. The manifests ask for medium unless a signal earns
' otherwise; the overlay is what moves a single alarm off its default.
Const SEV_CRITICAL = -2
Const SEV_HIGH     =  0
Const SEV_MEDIUM   =  1
Const SEV_LOW      =  2


Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	' What the operator is told when this property changes, and which of
	' its two states does the telling. Empty on every property until the
	' alarm table at the foot of the manifest names it - the same way the
	' exposure table is what makes a property appear on the panel. Both
	' default to silence.
	Public AlarmLabel
	Public AlarmPair
	Public AlarmSeverity

	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

	End Function

	' Asked of the property rather than of the caller, so the flags stay in
	' the one scope that declares them. The instances travel to whatever
	' scope reads the manifest and answer there just the same.
	Public Function Shows()
		Shows = Has(EXPOSE_VIEW)
	End Function

	Public Function ShowsValue()
		ShowsValue = Has(EXPOSE_VALUE)
	End Function

	Public Function CanEdit()
		CanEdit = Has(EXPOSE_EDIT)
	End Function

	Public Function CanBind()
		CanBind = Has(EXPOSE_EXPRESSION)
	End Function

	Public Function CanForce()
		CanForce = Has(EXPOSE_FORCE)
	End Function

	Public Function IsSaved()
		IsSaved = Has(EXPOSE_SAVED)
	End Function
	
	Public Function IsInterfaced()
		IsInterfaced = Has(EXPOSE_INTERFACE)
	End Function

	' Whether level 3 is given a point for it.
	'
	' A separate question from IsInterfaced, and asked separately. The
	' interface is where the Elipse application meets the automation; the
	' distribution is what leaves the station. Everything distributed is
	' interfaced - the distribution reads off the interface - but not
	' everything interfaced is distributed, and conflating the two left no
	' way to say so except a list of names kept somewhere else.
	Public Function IsIOTagged()
		IsIOTagged = Has(EXPOSE_IOTAG)
	End Function

	' An unnamed property raises nothing. Empty and "" compare equal in
	' VBScript, so a property the alarm table never mentions answers no
	' here without needing a flag of its own.
	Public Function IsAlarmed()
		IsAlarmed = (AlarmLabel <> "")
	End Function

	' The message either way, in the pattern the control room reads:
	' a label and the state, joined by a dash.
	Public Function AlarmNormalText()
		AlarmNormalText = AlarmLabel & " - " & PairPart(0)
	End Function

	Public Function AlarmActiveText()
		AlarmActiveText = AlarmLabel & " - " & PairPart(1)
	End Function

	Public Function AlarmLimit()
		AlarmLimit = (CLng("0" & PairPart(2)) <> 0)
	End Function

	Private Function PairPart(i)

		PairPart = ""

		Dim parts
		parts = Split(AlarmPair & "", "|")

		If i <= UBound(parts) Then PairPart = parts(i)

	End Function
	
	' Empty And anything is 0, so a property nobody classified answers no
	' to all of these.
	Private Function Has(flag)
		Has = ((Exposure And flag) <> 0)
	End Function

End Class

Sub AddProperty(bag, name, dataType, initialValue, helpEn, helpPt)

	Dim p
	Set p = New PropertyInfo

	p.Name         = name
	p.DataType     = dataType
	p.InitialValue = initialValue
	p.Exposure     = EXPOSE_NONE
	p.HelpEn       = helpEn
	p.HelpPt       = helpPt

	bag.Add LCase(name), p
	
End Sub

' What the screen may do with a property. Set apart from AddProperty so
' the classifications read as a table, and so changing one never means
' touching the help text - which is where the accents live.
' What the operator is alarmed on. Set apart from SetExposure for the
' reason that one is set apart from AddProperty: the alarms read as a
' table of their own, and a property left out of it raises nothing.
'
' A curated list and never a sweep of what is interfaced. An interface
' tag exists so a screen can draw a value, which is a different question
' from whether an operator should be told about it.
Sub SetAlarm(bag, propertyName, label, pair, severity)

	Dim k
	k = LCase(propertyName)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).AlarmLabel    = label
	bag(k).AlarmPair     = pair
	bag(k).AlarmSeverity = severity

End Sub
Sub SetExposure(bag, name, exposure)

	Dim k
	k = LCase(name)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).Exposure = exposure
	
End Sub

<xatm_config_data.PropertiesHelper.xatm_RASEAT:xatm_RASEAT_OnStartRunning()>
Sub xatm_RASEAT_OnStartRunning()

	Dim bag
	Set bag = CreateObject("Scripting.Dictionary")

	Dim i

	' One automation for the whole high-voltage entry, holding no
	' equipment of its own.
	'
	' The station runs with one incomer closed and the other open, so a
	' busbar relay takes down whichever was carrying it - and the steps
	' work out which that was from the memorised positions, resolving the
	' equipment the way every other sequence here does: GetDeviceById
	' against what the layout declares. A reference property would only be
	' a second place for the same fact to be written down, and a second
	' place to get it wrong.
	AddProperty bag, "Enabled", "Boolean", True, _
		"Master enable of this automation. Start requests are rejected and a running sequence stops while it is False.", _
		"Habilitação geral deste automatismo. Pedidos de partida são recusados e a sequência em andamento para enquanto estiver False."

	AddProperty bag, "Running", "Boolean", False, _
		"True while a sequence is in progress - the specification's religamento em curso.", _
		"True enquanto uma sequência está em andamento - o religamento em curso da especificação."

	' --- what holds a start back ------------------------------------------
	'
	' The same four the transfer answers to, and for the same reasons: a
	' person withholding it, a failure withholding it, the switchyard
	' withholding it, and the switchyard permitting it.
	AddProperty bag, "Preconditions", "Boolean", True, _
		"Field conditions that have to hold before a reclosing may start. Bound to an expression - True while the maneuver is permitted.", _
		"Condições de campo que devem valer antes de um religamento partir. Vinculada a uma expressão - True enquanto a manobra é permitida."

	AddProperty bag, "AutomaticBlock", "Boolean", False, _
		"Field conditions that block the reclosing. Bound to an expression - True keeps a sequence from starting, alongside OperatorBlock and GeneralBlock.", _
		"Condições de campo que bloqueiam o religamento. Vinculada a uma expressão - True impede a partida, junto com OperatorBlock e GeneralBlock."

	AddProperty bag, "OperatorBlock", "Boolean", False, _
		"Operator lock. Blocks the start until the operator releases it - the specification's RA blocked against RA in service.", _
		"Bloqueio do operador. Impede a partida até que o operador libere - o RA bloqueado contra RA em serviço da especificação."

	AddProperty bag, "GeneralBlock", "Boolean", False, _
		"General interlock. Blocks the start, and is latched by a step failure until Reset clears it.", _
		"Intertravamento geral. Impede a partida e é selado por uma falha de passo até que o Reset o apague."

	' One lamp for a question that takes four properties to answer.
	'
	' The parts stay published - a control room that wants to know why is
	' entitled to. This says only whether the automation could be asked to
	' run at all, which is what somebody deciding whether to reach for it
	' needs, and what the client's own signal list carries: one point per
	' maneuver, LIBERADO or BLOQUEADO.
	'
	' Recomputed on a tick by the Gates tag. Not saved and not editable:
	' it is a reading, and the only thing that writes it is the library.
	'
	' Running is not in it, nor another automation being in progress. Those
	' say busy, which is a different thing from barred and passes on its own.
	AddProperty bag, "Blocked", "Boolean", False, _
		"True when the high-voltage reclosing could not start - disabled, blocked by the operator or the general interlock, or barred by the field.", _
		"True quando o religamento de alta tensão não pode partir - desabilitado, bloqueado pelo operador ou pelo intertravamento geral, ou barrado pelo campo."


	AddProperty bag, "CommandOperatorBlock", "InternalTag", Empty, _
		"Operator lock command. Written by the operator from level 1, 2 or 3 to set or release OperatorBlock.", _
		"Comando de bloqueio do operador. Escrito pelo operador do nível 1, 2 ou 3 para marcar ou liberar o OperatorBlock."

	AddProperty bag, "CommandReset", "InternalTag", Empty, _
		"Reset command. Clears the latched step failures and the general block, so that a sequence can be started again.", _
		"Comando de reset. Apaga as falhas seladas de passo e o bloqueio geral, para que uma sequência possa partir novamente."

	' --- what the sequence did --------------------------------------------
	'
	' The two results the specification asks to be supervised, alongside
	' the religamento em curso that Running carries. Which step failed is
	' StepExecutionFailed's to say, and which breaker or transformer it
	' was is named in the log - the steps know both, having resolved them
	' themselves, so a property per failure would only repeat it.
	AddProperty bag, "Successful", "Boolean", False, _
		"The reclosing succeeded - the breaker confirmed closed, or load current confirmed it.", _
		"O religamento foi bem-sucedido - o disjuntor confirmou fechado, ou a corrente de carga o confirmou."

	AddProperty bag, "Unsuccessful", "Boolean", False, _
		"The reclosing failed.", _
		"O religamento foi malsucedido."

	For i = 1 To 6

		AddProperty bag, "StepExecutionFailed" & i, "Boolean", False, _
			"Latched failure of step " & i & ". Set when the step does not execute and the automation goes to general block, cleared by Reset.", _
			"Falha selada do passo " & i & ". Marcada quando o passo não executa e o automatismo entra em bloqueio geral, apagada pelo Reset."

	Next

	' The step timings are not properties. They are Consts in Main_Main,
	' the way the transfer keeps its own: numbers the sequence is written
	' around rather than dials a panel offers.

	' --- what the screen may do, and what leaves the station ---------------

	SetExposure bag, "Enabled", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "Blocked", EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "Running", EXPOSE_INTERFACE + EXPOSE_IOTAG

	SetExposure bag, "Preconditions",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "AutomaticBlock", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "OperatorBlock",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "GeneralBlock",   EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG

	SetExposure bag, "CommandOperatorBlock", EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "CommandReset",         EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG

	SetExposure bag, "Successful",   EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "Unsuccessful", EXPOSE_INTERFACE + EXPOSE_IOTAG

	For i = 1 To 6
		SetExposure bag, "StepExecutionFailed" & i, EXPOSE_INTERFACE + EXPOSE_IOTAG
	Next

	' --- what the operator is alarmed on -----------------------------------

	SetAlarm bag, "GeneralBlock",   "BLOQUEIO GERAL RA",           PAIR_BLOCKED,      SEV_HIGH
	SetAlarm bag, "OperatorBlock",  "BLOQUEIO OPERADOR RA",        PAIR_BLOCKED,      SEV_MEDIUM
	SetAlarm bag, "AutomaticBlock", "BLOQUEIO AUTOMÁTICO RA",   PAIR_BLOCKED,      SEV_MEDIUM
	SetAlarm bag, "Preconditions",  "PRECONDIÇÕES RA",       PAIR_PRECONDITION, SEV_MEDIUM
	SetAlarm bag, "Running",        "RELIGAMENTO AT",              PAIR_RUNNING,      SEV_LOW
	SetAlarm bag, "Successful",     "RELIGAMENTO AT BEM SUCEDIDO", PAIR_ACTUATED,     SEV_LOW
	SetAlarm bag, "Unsuccessful",   "RELIGAMENTO AT MAL SUCEDIDO", PAIR_ACTUATED,     SEV_HIGH

	SetAlarm bag, "Blocked",       "RELIGAMENTO AT", PAIR_BLOCKED, SEV_MEDIUM

	For i = 1 To 6
		SetAlarm bag, "StepExecutionFailed" & i, "FALHA PASSO " & i, PAIR_ACTUATED, SEV_HIGH
	Next

	Set Value = bag

End Sub

' What the configuration screen may do with a property, and whether its
' value is a setting at all. A bitmask: a property can be bound to an
' expression and forced, or shown and not edited, and so on.
'
' AddProperty leaves every property at EXPOSE_NONE, so nothing appears on
' the screen and nothing is written to the project until the table at the
' foot of the manifest says so. Both defaults fail closed.
Const EXPOSE_NONE       = 0
Const EXPOSE_VIEW       = 1     ' a row appears for it
Const EXPOSE_VALUE      = 2     ' its value is shown - never for a write-only command
Const EXPOSE_EDIT       = 4     ' the value can be typed
Const EXPOSE_EXPRESSION = 8     ' it can be bound to an expression
Const EXPOSE_FORCE      = 16    ' it can be forced at runtime, and is never saved
Const EXPOSE_SAVED      = 32    ' its value is configuration, not a reading
Const EXPOSE_INTERFACE  = 64    ' the Elipse application is given a tag for it
Const EXPOSE_IOTAG      = 128   ' level 3 is given a point for it, over 104


' What the operator reads, and which state says it: normal|active|limit.
'
' One string rather than three fields because the polarity belongs with
' the words. Preconditions is True while the maneuver is permitted, so
' it alarms on 0 where the rest alarm on 1 - and a pair that carried
' only the two words would let someone reword the message without ever
' seeing that the state raising it was the healthy one.
Const PAIR_ACTUATED     = "NORMAL|ATUADO|1"
Const PAIR_BLOCKED      = "LIBERADO|BLOQUEADO|1"
Const PAIR_PRECONDITION = "ATENDIDAS|NÃO ATENDIDAS|0"
Const PAIR_RUNNING      = "CONCLUÍDO|EM ANDAMENTO|1"

' DigitalSeverity, as Power numbers it.
'
' The scale runs backwards: the smaller the number the worse the alarm,
' and -2 is the most severe value here rather than the least. Anything
' comparing two of these has to be read twice - "worse than medium" is
' a < and not a >. The manifests ask for medium unless a signal earns
' otherwise; the overlay is what moves a single alarm off its default.
Const SEV_CRITICAL = -2
Const SEV_HIGH     =  0
Const SEV_MEDIUM   =  1
Const SEV_LOW      =  2
Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	' What the operator is told when this property changes, and which of
	' its two states does the telling. Empty on every property until the
	' alarm table at the foot of the manifest names it - the same way the
	' exposure table is what makes a property appear on the panel. Both
	' default to silence.
	Public AlarmLabel
	Public AlarmPair
	Public AlarmSeverity

	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

	End Function

	' Asked of the property rather than of the caller, so the flags stay in
	' the one scope that declares them. The instances travel to whatever
	' scope reads the manifest and answer there just the same.
	Public Function Shows()
		Shows = Has(EXPOSE_VIEW)
	End Function

	Public Function ShowsValue()
		ShowsValue = Has(EXPOSE_VALUE)
	End Function

	Public Function CanEdit()
		CanEdit = Has(EXPOSE_EDIT)
	End Function

	Public Function CanBind()
		CanBind = Has(EXPOSE_EXPRESSION)
	End Function

	Public Function CanForce()
		CanForce = Has(EXPOSE_FORCE)
	End Function

	Public Function IsSaved()
		IsSaved = Has(EXPOSE_SAVED)
	End Function
	
	Public Function IsInterfaced()
		IsInterfaced = Has(EXPOSE_INTERFACE)
	End Function

	' Whether level 3 is given a point for it.
	'
	' A separate question from IsInterfaced, and asked separately. The
	' interface is where the Elipse application meets the automation; the
	' distribution is what leaves the station. Everything distributed is
	' interfaced - the distribution reads off the interface - but not
	' everything interfaced is distributed, and conflating the two left no
	' way to say so except a list of names kept somewhere else.
	Public Function IsIOTagged()
		IsIOTagged = Has(EXPOSE_IOTAG)
	End Function

	' An unnamed property raises nothing. Empty and "" compare equal in
	' VBScript, so a property the alarm table never mentions answers no
	' here without needing a flag of its own.
	Public Function IsAlarmed()
		IsAlarmed = (AlarmLabel <> "")
	End Function

	' The message either way, in the pattern the control room reads:
	' a label and the state, joined by a dash.
	Public Function AlarmNormalText()
		AlarmNormalText = AlarmLabel & " - " & PairPart(0)
	End Function

	Public Function AlarmActiveText()
		AlarmActiveText = AlarmLabel & " - " & PairPart(1)
	End Function

	Public Function AlarmLimit()
		AlarmLimit = (CLng("0" & PairPart(2)) <> 0)
	End Function

	Private Function PairPart(i)

		PairPart = ""

		Dim parts
		parts = Split(AlarmPair & "", "|")

		If i <= UBound(parts) Then PairPart = parts(i)

	End Function
	
	
	' Empty And anything is 0, so a property nobody classified answers no
	' to all of these.
	Private Function Has(flag)
		Has = ((Exposure And flag) <> 0)
	End Function

End Class

Sub AddProperty(bag, name, dataType, initialValue, helpEn, helpPt)

	Dim p
	Set p = New PropertyInfo

	p.Name         = name
	p.DataType     = dataType
	p.InitialValue = initialValue
	p.Exposure     = EXPOSE_NONE
	p.HelpEn       = helpEn
	p.HelpPt       = helpPt

	bag.Add LCase(name), p
	
End Sub

' What the screen may do with a property. Set apart from AddProperty so
' the classifications read as a table, and so changing one never means
' touching the help text - which is where the accents live.
' What the operator is alarmed on. Set apart from SetExposure for the
' reason that one is set apart from AddProperty: the alarms read as a
' table of their own, and a property left out of it raises nothing.
'
' A curated list and never a sweep of what is interfaced. An interface
' tag exists so a screen can draw a value, which is a different question
' from whether an operator should be told about it.
Sub SetAlarm(bag, propertyName, label, pair, severity)

	Dim k
	k = LCase(propertyName)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).AlarmLabel    = label
	bag(k).AlarmPair     = pair
	bag(k).AlarmSeverity = severity

End Sub
Sub SetExposure(bag, name, exposure)

	Dim k
	k = LCase(name)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).Exposure = exposure
		
End Sub

<xatm_config_data.PropertiesHelper.xatm_TA:xatm_TA_OnStartRunning()>
Sub xatm_TA_OnStartRunning()

	Dim bag
	Set bag = CreateObject("Scripting.Dictionary")

	' The automatic transfer, and only that.
	'
	' Everything an operator asks for lives on xatm_TMTNM. This one is
	' asked for by a trip, through the transformer's own trigger, and
	' never by a person - which is why it has no command to start it,
	' one pair of gates rather than five, and no property naming a
	' transformer to treat as out of service. The field is read for
	' that when the sequence starts.

	AddProperty bag, "Enabled", "Boolean", True, _
		"Master enable of this automation. Start requests are rejected and a running sequence stops while it is False.", _
		"Habilitação geral deste automatismo. Pedidos de partida são recusados e a sequência em andamento para enquanto estiver False."

	AddProperty bag, "Running", "Boolean", False, _
		"True while a sequence is in progress. Read by the other automation objects for mutual exclusion, so only one runs at a time.", _
		"True enquanto uma sequência está em andamento. Lido pelos demais automatismos para exclusão mútua, de modo que apenas um execute por vez."

	AddProperty bag, "Transformer", "xatm_Transformer", Empty, _
		"Transformer XObject this automation instance is bound to.", _
		"XObject do transformador ao qual esta instância do automatismo está vinculada."
	
	' Whether step 1 sends the open command, or only confirms the breaker.
	'
	' The lockout relay trips that breaker itself on most panels, and where
	' it does there is nothing for step 1 to command - only something to
	' confirm. An order sent into a breaker the protection is already
	' opening is a second order for one movement, so clear it there.
	'
	' True by default, which is what every station did before this existed,
	' and the forgiving direction: step 1 leaves early when the breaker is
	' already open, so a redundant order only ever reaches one still closed,
	' where a wait nobody satisfies runs the step to its timeout.
	'
	' Read at step 1 and nowhere else. It says how this transformer's panel
	' is wired, so it belongs to the instance and not to the request - the
	' undervoltage trigger reaches the same step and finds the same answer.
	AddProperty bag, "SendLowVoltageBreakerOpenCommand", "Boolean", True, _
		"True when step 1 sends the open command to the low-voltage breaker. Clear it where the lockout relay already trips that breaker and step 1 need only confirm it.", _
		"True quando o passo 1 envia o comando de abertura ao disjuntor de baixa tensão. Desmarque onde o relé de bloqueio já abre esse disjuntor e ao passo 1 basta confirmá-lo."


	AddProperty bag, "OperatorBlock", "Boolean", False, _
		"Operator lock. Blocks the start until the operator releases it.", _
		"Bloqueio do operador. Impede a partida até que o operador libere."

	AddProperty bag, "GeneralBlock", "Boolean", False, _
		"General interlock. Blocks the start, and is latched by a step failure until Reset clears it.", _
		"Intertravamento geral. Impede a partida e é selado por uma falha de passo até que o Reset o apague."
		
	
	' --- the gates, one pair for each maneuver --------------------------
	'
	' A pair for every command, because what has to hold before a transfer
	' may start is not the same from one contingency to the next: TM with
	' TR2 out closes a different path, over different equipment, from a
	' different starting state. One expression could not answer for all of
	' them, and a single pair is what forced the question.
	'
	' Preconditions<gate> is True while that maneuver is permitted;
	' AutomaticBlock<gate> is True while the switchyard bars it. The two
	' polarities the single pair had, kept.
	'
	' The gate name is the mode, and the transformer the maneuver assumes
	' is out where there is one: TM, TM200, NM, NM400. Exactly the names
	' the command tags carry, so a command and the gates that let it
	' through are read off one string and cannot drift apart.
	'
	Dim gm, gi, gate, gateLabel, outEN, outPT, whatEN, whatPT

	AddProperty bag, "Preconditions", "Boolean", True, _
		"Field conditions that have to hold before an automatic transfer may start. Bound to an expression - True while the maneuver is permitted.", _
		"Condições de campo que devem valer antes que uma transferência automática possa partir. Vinculada a uma expressão - True enquanto a manobra é permitida."

	AddProperty bag, "AutomaticBlock", "Boolean", False, _
		"Field conditions that block the automatic transfer. Bound to an expression - True keeps it from starting, alongside OperatorBlock and GeneralBlock.", _
		"Condições de campo que bloqueiam a transferência automática. Vinculada a uma expressão - True impede a partida, junto com OperatorBlock e GeneralBlock."

	' One lamp for a question that takes four properties to answer.
	'
	' The parts stay published - a control room that wants to know why is
	' entitled to. This says only whether the automation could be asked to
	' run at all, which is what somebody deciding whether to reach for it
	' needs, and what the client's own signal list carries: one point per
	' maneuver, LIBERADO or BLOQUEADO.
	'
	' Recomputed on a tick by the Gates tag. Not saved and not editable:
	' it is a reading, and the only thing that writes it is the library.
	'
	' Running is not in it, nor another automation being in progress. Those
	' say busy, which is a different thing from barred and passes on its own.
	AddProperty bag, "Blocked", "Boolean", False, _
		"True when an automatic transfer could not start - disabled, blocked by the operator or the general interlock, or barred by the field.", _
		"True quando uma transferência automÁtica não pode partir - desabilitado, bloqueado pelo operador ou pelo intertravamento geral, ou barrado pelo campo."


	' --- the command interface ------------------------------------------
	'
	' No CommandStart of any kind. Over IEC 60870-5-104 an address means
	' one maneuver, and there is no maneuver here for a person to ask
	' for: a point that started an automatic transfer by hand would be a
	' way to run the restoration scheme on a busbar that is still live.

	AddProperty bag, "CommandReset", "InternalTag", Empty, _
		"Reset command. Clears the latched step failures and the general block, so that a sequence can be started again.", _
		"Comando de reset. Apaga as falhas seladas de passo e o bloqueio geral, para que uma sequência possa partir novamente."

	AddProperty bag, "CommandOperatorBlock", "InternalTag", Empty, _
		"Operator lock command. Written by the operator's screen to set or release OperatorBlock.", _
		"Comando de bloqueio do operador. Escrito pela tela do operador para marcar ou liberar o OperatorBlock."

	Dim i
	For i = 1 To 6

		AddProperty bag, "StepExecutionFailed" & i, "Boolean", False, _
			"Latched failure of step " & i & ". Set when the step does not execute and the automation goes to global lockout, cleared by Reset.", _
			"Falha selada do passo " & i & ". Marcada quando o passo não executa e o automatismo entra em bloqueio geral, apagada pelo Reset."

	Next

	' What the screen may do with each of these. Anything left out stays
	' EXPOSE_NONE.
	'
	' The same answers the manual automation gives, for the properties
	' the two have in common - the reasoning is written out there.
	SetExposure bag, "Enabled",        EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "Transformer",    EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "SendLowVoltageBreakerOpenCommand", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "OperatorBlock",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "GeneralBlock",   EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG

	SetExposure bag, "Preconditions",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "AutomaticBlock", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG

	SetExposure bag, "Blocked", EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "Running", EXPOSE_INTERFACE + EXPOSE_IOTAG

	For i = 1 To 6
		SetExposure bag, "StepExecutionFailed" & i, EXPOSE_INTERFACE + EXPOSE_IOTAG
	Next

	SetExposure bag, "CommandReset",         EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "CommandOperatorBlock", EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG


	' --- what the operator is alarmed on --------------------------------
	'
	' The same three the manual automation raises about itself, and the
	' one pair of gates.
	SetAlarm bag, "GeneralBlock",   "BLOQUEIO GERAL",      PAIR_BLOCKED,      SEV_HIGH
	SetAlarm bag, "OperatorBlock",  "BLOQUEIO OPERADOR",   PAIR_BLOCKED,      SEV_MEDIUM
	SetAlarm bag, "Running",        "AUTOMATISMO",         PAIR_RUNNING,      SEV_LOW

	SetAlarm bag, "Preconditions",  "PRECONDIÇÕES TA",        PAIR_PRECONDITION, SEV_MEDIUM
	SetAlarm bag, "AutomaticBlock", "BLOQUEIO AUTOMÁTICO TA", PAIR_BLOCKED,      SEV_MEDIUM

	SetAlarm bag, "Blocked",       "TRANSFERÊNCIA AUTOMÁTICA", PAIR_BLOCKED, SEV_MEDIUM

	For i = 1 To 6
		SetAlarm bag, "StepExecutionFailed" & i, "FALHA PASSO " & i, PAIR_ACTUATED, SEV_HIGH
	Next

	Set Value = bag

End Sub

' What the configuration screen may do with a property, and whether its
' value is a setting at all. A bitmask: a property can be bound to an
' expression and forced, or shown and not edited, and so on.
'
' AddProperty leaves every property at EXPOSE_NONE, so nothing appears on
' the screen and nothing is written to the project until the table at the
' foot of the manifest says so. Both defaults fail closed.
Const EXPOSE_NONE       = 0
Const EXPOSE_VIEW       = 1     ' a row appears for it
Const EXPOSE_VALUE      = 2     ' its value is shown - never for a write-only command
Const EXPOSE_EDIT       = 4     ' the value can be typed
Const EXPOSE_EXPRESSION = 8     ' it can be bound to an expression
Const EXPOSE_FORCE      = 16    ' it can be forced at runtime, and is never saved
Const EXPOSE_SAVED      = 32    ' its value is configuration, not a reading
Const EXPOSE_INTERFACE  = 64    ' the Elipse application is given a tag for it
Const EXPOSE_IOTAG      = 128   ' level 3 is given a point for it, over 104


' What the operator reads, and which state says it: normal|active|limit.
'
' One string rather than three fields because the polarity belongs with
' the words. Preconditions is True while the maneuver is permitted, so
' it alarms on 0 where the rest alarm on 1 - and a pair that carried
' only the two words would let someone reword the message without ever
' seeing that the state raising it was the healthy one.
Const PAIR_ACTUATED     = "NORMAL|ATUADO|1"
Const PAIR_BLOCKED      = "LIBERADO|BLOQUEADO|1"
Const PAIR_PRECONDITION = "ATENDIDAS|NÃO ATENDIDAS|0"
Const PAIR_RUNNING      = "CONCLUÍDO|EM ANDAMENTO|1"

' DigitalSeverity, as Power numbers it.
'
' The scale runs backwards: the smaller the number the worse the alarm,
' and -2 is the most severe value here rather than the least. Anything
' comparing two of these has to be read twice - "worse than medium" is
' a < and not a >. The manifests ask for medium unless a signal earns
' otherwise; the overlay is what moves a single alarm off its default.
Const SEV_CRITICAL = -2
Const SEV_HIGH     =  0
Const SEV_MEDIUM   =  1
Const SEV_LOW      =  2



Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	' What the operator is told when this property changes, and which of
	' its two states does the telling. Empty on every property until the
	' alarm table at the foot of the manifest names it - the same way the
	' exposure table is what makes a property appear on the panel. Both
	' default to silence.
	Public AlarmLabel
	Public AlarmPair
	Public AlarmSeverity

	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

	End Function

	' Asked of the property rather than of the caller, so the flags stay in
	' the one scope that declares them. The instances travel to whatever
	' scope reads the manifest and answer there just the same.
	Public Function Shows()
		Shows = Has(EXPOSE_VIEW)
	End Function

	Public Function ShowsValue()
		ShowsValue = Has(EXPOSE_VALUE)
	End Function

	Public Function CanEdit()
		CanEdit = Has(EXPOSE_EDIT)
	End Function

	Public Function CanBind()
		CanBind = Has(EXPOSE_EXPRESSION)
	End Function

	Public Function CanForce()
		CanForce = Has(EXPOSE_FORCE)
	End Function

	Public Function IsSaved()
		IsSaved = Has(EXPOSE_SAVED)
	End Function
	
	Public Function IsInterfaced()
		IsInterfaced = Has(EXPOSE_INTERFACE)
	End Function

	' Whether level 3 is given a point for it.
	'
	' A separate question from IsInterfaced, and asked separately. The
	' interface is where the Elipse application meets the automation; the
	' distribution is what leaves the station. Everything distributed is
	' interfaced - the distribution reads off the interface - but not
	' everything interfaced is distributed, and conflating the two left no
	' way to say so except a list of names kept somewhere else.
	Public Function IsIOTagged()
		IsIOTagged = Has(EXPOSE_IOTAG)
	End Function

	' An unnamed property raises nothing. Empty and "" compare equal in
	' VBScript, so a property the alarm table never mentions answers no
	' here without needing a flag of its own.
	Public Function IsAlarmed()
		IsAlarmed = (AlarmLabel <> "")
	End Function

	' The message either way, in the pattern the control room reads:
	' a label and the state, joined by a dash.
	Public Function AlarmNormalText()
		AlarmNormalText = AlarmLabel & " - " & PairPart(0)
	End Function

	Public Function AlarmActiveText()
		AlarmActiveText = AlarmLabel & " - " & PairPart(1)
	End Function

	Public Function AlarmLimit()
		AlarmLimit = (CLng("0" & PairPart(2)) <> 0)
	End Function

	Private Function PairPart(i)

		PairPart = ""

		Dim parts
		parts = Split(AlarmPair & "", "|")

		If i <= UBound(parts) Then PairPart = parts(i)

	End Function

	' Empty And anything is 0, so a property nobody classified answers no
	' to all of these.
	Private Function Has(flag)
		Has = ((Exposure And flag) <> 0)
	End Function

End Class

Sub AddProperty(bag, name, dataType, initialValue, helpEn, helpPt)

	Dim p
	Set p = New PropertyInfo

	p.Name         = name
	p.DataType     = dataType
	p.InitialValue = initialValue
	p.Exposure     = EXPOSE_NONE
	p.HelpEn       = helpEn
	p.HelpPt       = helpPt

	bag.Add LCase(name), p
	
End Sub

' What the screen may do with a property. Set apart from AddProperty so
' the classifications read as a table, and so changing one never means
' touching the help text - which is where the accents live.
' What the operator is alarmed on. Set apart from SetExposure for the
' reason that one is set apart from AddProperty: the alarms read as a
' table of their own, and a property left out of it raises nothing.
'
' A curated list and never a sweep of what is interfaced. An interface
' tag exists so a screen can draw a value, which is a different question
' from whether an operator should be told about it.
Sub SetAlarm(bag, propertyName, label, pair, severity)

	Dim k
	k = LCase(propertyName)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).AlarmLabel    = label
	bag(k).AlarmPair     = pair
	bag(k).AlarmSeverity = severity

End Sub
Sub SetExposure(bag, name, exposure)

	Dim k
	k = LCase(name)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).Exposure = exposure
			
End Sub

<xatm_config_data.PropertiesHelper.xatm_TMTNM:xatm_TMTNM_OnStartRunning()>
Sub xatm_TMTNM_OnStartRunning()

	Dim bag
	Set bag = CreateObject("Scripting.Dictionary")

	AddProperty bag, "Enabled", "Boolean", True, _
		"Master enable of this automation. Start requests are rejected and a running sequence stops while it is False.", _
		"Habilitação geral deste automatismo. Pedidos de partida são recusados e a sequência em andamento para enquanto estiver False."

	AddProperty bag, "Running", "Boolean", False, _
		"True while a sequence is in progress. Read by the other automation objects for mutual exclusion, so only one runs at a time.", _
		"True enquanto uma sequência está em andamento. Lido pelos demais automatismos para exclusão mútua, de modo que apenas um execute por vez."

	AddProperty bag, "Transformer", "xatm_Transformer", Empty, _
		"Transformer XObject this automation instance is bound to.", _
		"XObject do transformador ao qual esta instância do automatismo está vinculada."
	

	' Which pair of busbars this automation moves, and whether it moves one
	' at all.
	'
	' Empty, and this is an ordinary transformer automation: TM takes its
	' transformer's load off, NM brings it back, and the numbered variants
	' say which other transformer is out of service.
	'
	' Set, and it is a busbar automation instead. TM hands that busbar to
	' the far side of the ring, NM brings it back, there is no contingency
	' to declare, and Transformer is left unbound - it is not any
	' transformer's automation.
	'
	' Text and not a link, unlike every other property here that names
	' equipment, because a busbar is not an XObject there is anything to
	' bind to. And one property rather than two: a flag reading "this is a
	' busbar automation" beside a name saying which busbar would be two ways
	' of writing one fact, and two ways of writing one fact can be made to
	' disagree.
	'
	' The six-busbar ring defines B1A-B4A and B2B-B3A. A two-busbar layout
	' defines neither, and SyncAutomation creates no such instance there.
	AddProperty bag, "BusbarPair", "String", "", _
		"The pair of busbars this automation transfers between, B1A-B4A or B2B-B3A, which makes TM and NM the busbar transfer and its normalisation. Empty on an ordinary transformer automation.", _
		"O par de barramentos que este automatismo transfere, B1A-B4A ou B2B-B3A, o que faz de TM e NM a transferência de barra e a sua normalização. Vazio num automatismo de transformador comum."

	AddProperty bag, "OperatorBlock", "Boolean", False, _
		"Operator lock. Blocks the start until the operator releases it.", _
		"Bloqueio do operador. Impede a partida até que o operador libere."

	AddProperty bag, "GeneralBlock", "Boolean", False, _
		"General interlock. Blocks the start, and is latched by a step failure until Reset clears it.", _
		"Intertravamento geral. Impede a partida e é selado por uma falha de passo até que o Reset o apague."
	
	' Whether a failed maneuver unwinds itself instead of stopping where it
	' failed.
	'
	' Every command a run issues is written down as it goes. Unwinding walks
	' that record backwards and puts each device back, one per pass, and
	' touches nothing else.
	'
	' Only what was operated. A step that found its device already in the
	' position it wanted wrote nothing, so the revert has no reason to visit
	' it - which matters, because a device that was already there is one
	' somebody else put there, and driving it back would undo their work
	' rather than this run's.
	'
	' A run that operated nothing is not reverted; there is nothing to put
	' back, and saying so is a better account than a sequence of no-ops.
	'
	' False by default. Operating breakers on their own after something has
	' already gone wrong is not a thing to switch on for a station without
	' being asked.
	AddProperty bag, "RevertOnFailure", "Boolean", False, _
		"True when a failed maneuver unwinds itself instead of stopping. It puts back only the devices it actually operated, in reverse order, and leaves alone anything it found already in position.", _
		"True quando uma manobra que falha se desfaz em vez de parar. Ela devolve apenas os equipamentos que realmente operou, na ordem inversa, e não toca no que encontrou já na posição desejada."


	' What the last run did, kept until the next one starts.
	'
	' Cleared at the start of a maneuver and not at the end of one, so the
	' control room goes on seeing how the last attempt went for as long as
	' nothing new has been asked for. The same shape the reclosing uses.
	'
	' A revert ends Unsuccessful. It finished its sequence, but the maneuver
	' somebody asked for did not happen, and that is what these two are for.
	AddProperty bag, "Successful", "Boolean", False, _
		"True when the last maneuver completed. Cleared when the next one starts, so what the last attempt did stays readable until then.", _
		"True quando a última manobra foi concluída. Apagada quando a próxima parte, de modo que o resultado da anterior permaneça legível até lá."

	AddProperty bag, "Unsuccessful", "Boolean", False, _
		"True when the last maneuver did not complete - it failed and stopped, or it failed and unwound itself.", _
		"True quando a última manobra não foi concluída - falhou e parou, ou falhou e se desfez."

	AddProperty bag, "Reverting", "Boolean", False, _
		"True while a failed maneuver is unwinding itself. Cleared when the sequence ends, which it ends unsuccessfully.", _
		"True enquanto uma manobra que falhou está se desfazendo. Apagada quando a sequência termina - e ela termina mal sucedida."
	
	' --- the gates, one pair for each maneuver --------------------------
	'
	' A pair for every command, because what has to hold before a transfer
	' may start is not the same from one contingency to the next: TM with
	' TR2 out closes a different path, over different equipment, from a
	' different starting state. One expression could not answer for all of
	' them, and a single pair is what forced the question.
	'
	' Preconditions<gate> is True while that maneuver is permitted;
	' AutomaticBlock<gate> is True while the switchyard bars it. The two
	' polarities the single pair had, kept.
	'
	' The gate name is the mode, and the transformer the maneuver assumes
	' is out where there is one: TM, TM200, NM, NM400. Exactly the names
	' the command tags carry, so a command and the gates that let it
	' through are read off one string and cannot drift apart.
	'
	Dim gm, gi, gate, gateLabel, outEN, outPT, whatEN, whatPT

	For Each gm In Array("TM", "NM")

		If gm = "TM" Then
			whatEN = "a manual transfer"
			whatPT = "uma transferência manual"
		Else
			whatEN = "a manual normalisation"
			whatPT = "uma normalização manual"
		End If

		For gi = 0 To 4

			If gi = 0 Then
				gate  = gm
				outEN = ""
				outPT = ""
			Else
				gate  = gm & (gi * 100)
				outEN = " with transformer " & (gi * 100) & " out of service"
				outPT = " com o transformador " & (gi * 100) & " impedido"
			End If

			AddProperty bag, "Preconditions" & gate, "Boolean", True, _
				"Field conditions that have to hold before " & whatEN & outEN & " may start. Bound to an expression - True while the maneuver is permitted.", _
				"Condições de campo que devem valer antes que " & whatPT & outPT & " possa partir. Vinculada a uma expressão - True enquanto a manobra é permitida."

			AddProperty bag, "AutomaticBlock" & gate, "Boolean", False, _
				"Field conditions that block " & whatEN & outEN & ". Bound to an expression - True keeps it from starting, alongside OperatorBlock and GeneralBlock.", _
				"Condições de campo que bloqueiam " & whatPT & outPT & ". Vinculada a uma expressão - True impede a partida, junto com OperatorBlock e GeneralBlock."

		Next

	Next


	' One lamp for a question that takes four properties to answer.
	'
	' The parts stay published - a control room that wants to know why is
	' entitled to. This says only whether the automation could be asked to
	' run at all, which is what somebody deciding whether to reach for it
	' needs, and what the client's own signal list carries: one point per
	' maneuver, LIBERADO or BLOQUEADO.
	'
	' Recomputed on a tick by the Gates tag. Not saved and not editable:
	' it is a reading, and the only thing that writes it is the library.
	'
	' Running is not in it, nor another automation being in progress. Those
	' say busy, which is a different thing from barred and passes on its own.
	'
	' Every contingency has to be barred before either of these is,
	' because a transfer with TR3 out being unavailable says nothing
	' about the plain transfer. On an instance that moves a busbar
	' there is one gate rather than five, and only it is read.
	AddProperty bag, "BlockedTM", "Boolean", False, _
		"True when no manual transfer could start - disabled, blocked by the operator or the general interlock, or every contingency of it barred by the field.", _
		"True quando nenhuma transferência manual pode partir - desabilitada, bloqueada pelo operador ou pelo intertravamento geral, ou toda contingência dela barrada pelo campo."

	AddProperty bag, "BlockedNM", "Boolean", False, _
		"True when no manual normalisation could start - disabled, blocked by the operator or the general interlock, or every contingency of it barred by the field.", _
		"True quando nenhuma normalização manual pode partir - desabilitada, bloqueada pelo operador ou pelo intertravamento geral, ou toda contingência dela barrada pelo campo."


	' --- the command interface -----------------------------------------

	AddProperty bag, "CommandReset", "InternalTag", Empty, _
		"Reset command. Clears the latched step failures and the general block, so that a sequence can be started again.", _
		"Comando de reset. Apaga as falhas seladas de passo e o bloqueio geral, para que uma sequência possa partir novamente."

	' One command per maneuver, rather than one command saying which.
	'
	' Level 3 reaches these over IEC 60870-5-104, where a command is a point
	' address carrying no argument: the address has to mean one maneuver by
	' itself. The E3 operation screen works the same way round - a button
	' writes a tag, it does not compose an argument. So the variant lives in
	' the tag, and the tag name says which transformer the maneuver assumes
	' is out of service.
	'
	' CommandStartTM is the maneuver with nothing out; CommandStartTM<n00>
	' is the same trigger with transformer n00 out, which is what the spec
	' calls "TM TR<trigger>-TR<n>" - so the point list and the spec sections
	' name the same thing.
	'
	' The one whose Id is the bound transformer's own is declared with the
	' rest and never used: Start rejects a maneuver that is its own
	' impediment. Declaring it keeps the set identical on every instance,
	' so no point list has to reason about which member is missing.
	AddProperty bag, "CommandStartTM", "InternalTag", Empty, _
		"Start command for a manual transfer with no transformer out of service.", _
		"Comando de partida da transferência manual sem transformador impedido."

	AddProperty bag, "CommandStartNM", "InternalTag", Empty, _
		"Start command for a manual normalisation with no transformer out of service.", _
		"Comando de partida da normalização manual sem transformador impedido."

	Dim trn
	For trn = 1 To 4

		AddProperty bag, "CommandStartTM" & (trn * 100), "InternalTag", Empty, _
			"Start command for a manual transfer with transformer " & (trn * 100) & " out of service.", _
			"Comando de partida da transferência manual com o transformador " & (trn * 100) & " impedido."

		AddProperty bag, "CommandStartNM" & (trn * 100), "InternalTag", Empty, _
			"Start command for a manual normalisation with transformer " & (trn * 100) & " out of service.", _
			"Comando de partida da normalização manual com o transformador " & (trn * 100) & " impedido."

	Next

	AddProperty bag, "CommandOperatorBlock", "InternalTag", Empty, _
		"Operator lock command. Written by the operator's screen to set or release OperatorBlock.", _
		"Comando de bloqueio do operador. Escrito pela tela do operador para marcar ou liberar o OperatorBlock."

	Dim i
	For i = 1 To 6

		AddProperty bag, "StepExecutionFailed" & i, "Boolean", False, _
			"Latched failure of step " & i & ". Set when the step does not execute and the automation goes to global lockout, cleared by Reset.", _
			"Falha selada do passo " & i & ". Marcada quando o passo não executa e o automatismo entra em bloqueio geral, apagada pelo Reset."

	Next
	
	' What the screen may do with each of these, and which are settings
	' rather than readings. Anything left out stays EXPOSE_NONE.
	'
	' The readings - the blocks and every gate - are not saved: their
	' expression is the configuration, and the reading is whatever the
	' switchyard was doing at the time.
	'
	' EXPOSE_INTERFACE is a different question from the rest, and answered
	' separately: not what the panel may do with a property, but whether
	' the Elipse application is given a tag of its own for it. What a
	' screen draws or acts on is interfaced; what only an engineer sets -
	' the transformer this instance drives - is not.
	SetExposure bag, "Enabled",        EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "Transformer",    EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	' Interfaced, so a screen can tell an operator what the TM and NM
	' buttons in front of them actually do, but not carried to level 3: it
	' is configuration and never changes at runtime, so the point list says
	' it once rather than the station repeating it forever.
	SetExposure bag, "BusbarPair",     EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED + EXPOSE_INTERFACE
	SetExposure bag, "OperatorBlock",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "GeneralBlock",   EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG

	' Every gate is shown, bound to an expression, forceable for a test and
	' interfaced - what the single pair was - and carried to level 3 as
	' well, so the control centre reads which maneuver is barred and not
	' merely that one of them is.

	For Each gm In Array("TM", "NM")
		For gi = 0 To 4
			If gi = 0 Then gate = gm Else gate = gm & (gi * 100)
			SetExposure bag, "Preconditions"  & gate, EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG
			SetExposure bag, "AutomaticBlock" & gate, EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE + EXPOSE_IOTAG
		Next
	Next

	' Running and the six step latches are the automation talking about
	' itself. Neither belongs on the configuration panel - there is
	' nothing about them to configure - but both are what a screen draws,
	' so they are interfaced and nothing else.
	SetExposure bag, "RevertOnFailure", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED

	SetExposure bag, "BlockedTM",   EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "BlockedNM",   EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "Running",      EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "Successful",   EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "Unsuccessful", EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "Reverting",    EXPOSE_INTERFACE + EXPOSE_IOTAG

	For i = 1 To 6
		SetExposure bag, "StepExecutionFailed" & i, EXPOSE_INTERFACE + EXPOSE_IOTAG
	Next

	' The command interface is shown and nothing more. No EXPOSE_EDIT and
	' no EXPOSE_FORCE, because either one would let the panel start a real
	' maneuver on a live switchyard; no EXPOSE_VALUE, because a command is
	' written and never read, and reading one back shows the last thing
	' written to it as though it were a state.
	'
	' EXPOSE_SAVED all the same: what these hold is which tag the object
	' talks through, which is configuration - and without it ResetReadings
	' would put every one of them back to Empty on the next import.
	SetExposure bag, "CommandReset",         EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "CommandStartTM",       EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "CommandStartNM",       EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG

	For trn = 1 To 4
		SetExposure bag, "CommandStartTM" & (trn * 100), EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG
		SetExposure bag, "CommandStartNM" & (trn * 100), EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG
	Next
	SetExposure bag, "CommandOperatorBlock", EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE + EXPOSE_IOTAG


	' --- what the operator is alarmed on --------------------------------
	'
	' The automation talking about itself, which is what the control room
	' wants from it: a step that would not execute, whatever is holding a
	' start back, and whether a sequence is on.
	'
	' Running is here as a status and not a fault - low severity, and a
	' pair reading PARADO and EM OPERAÇÃO rather than NORMAL and
	' ATUADO, because nothing about a sequence in progress is abnormal.
	'
	' Preconditions is the one that raises on False. It is True while the
	' maneuver is permitted, so PAIR_PRECONDITION carries a limit of 0.
	SetAlarm bag, "GeneralBlock",   "BLOQUEIO GERAL",      PAIR_BLOCKED,      SEV_HIGH
	SetAlarm bag, "OperatorBlock",  "BLOQUEIO OPERADOR",   PAIR_BLOCKED,      SEV_MEDIUM
	SetAlarm bag, "Running",        "AUTOMATISMO",         PAIR_RUNNING,      SEV_LOW
	SetAlarm bag, "Successful",    "MANOBRA BEM SUCEDIDA",  PAIR_ACTUATED, SEV_LOW
	SetAlarm bag, "Unsuccessful",  "MANOBRA MAL SUCEDIDA",  PAIR_ACTUATED, SEV_HIGH
	SetAlarm bag, "Reverting",     "REVERSÃO DE MANOBRA", PAIR_ACTUATED, SEV_HIGH
	
	' Every gate raises too, the way the single pair did.
	'
	' This is also what puts them in the substation folder at all:
	' AlarmObject builds a tag for a reading that is alarmed and for a
	' command, and for nothing else. An unalarmed gate is interfaced and
	' carried to level 3, but has no tag under the PowerSubstation - so the
	' alarm table is what decides whether the control room can see it, not
	' merely whether it is told about it.
	'
	' The label names the maneuver the way the control room says it: TM-TR2
	' is a manual transfer with TR2 impeded, which is the spec's own
	' "TM TR<trigger>-TR2".

	For Each gm In Array("TM", "NM")
		For gi = 0 To 4

			If gi = 0 Then
				gate      = gm
				gateLabel = gm
			Else
				gate      = gm & (gi * 100)
				gateLabel = gm & "-TR" & gi
			End If

			SetAlarm bag, "Preconditions"  & gate, "PRECONDIÇÕES " & gateLabel,        PAIR_PRECONDITION, SEV_MEDIUM
			SetAlarm bag, "AutomaticBlock" & gate, "BLOQUEIO AUTOMÁTICO " & gateLabel, PAIR_BLOCKED,      SEV_MEDIUM

		Next
	Next

	SetAlarm bag, "BlockedTM",     "TRANSFERÊNCIA",    PAIR_BLOCKED, SEV_MEDIUM
	SetAlarm bag, "BlockedNM",     "NORMALIZAÇÃO",   PAIR_BLOCKED, SEV_MEDIUM

	For i = 1 To 6
		SetAlarm bag, "StepExecutionFailed" & i, "FALHA PASSO " & i, PAIR_ACTUATED, SEV_HIGH
	Next

	Set Value = bag

End Sub

' What the configuration screen may do with a property, and whether its
' value is a setting at all. A bitmask: a property can be bound to an
' expression and forced, or shown and not edited, and so on.
'
' AddProperty leaves every property at EXPOSE_NONE, so nothing appears on
' the screen and nothing is written to the project until the table at the
' foot of the manifest says so. Both defaults fail closed.
Const EXPOSE_NONE       = 0
Const EXPOSE_VIEW       = 1     ' a row appears for it
Const EXPOSE_VALUE      = 2     ' its value is shown - never for a write-only command
Const EXPOSE_EDIT       = 4     ' the value can be typed
Const EXPOSE_EXPRESSION = 8     ' it can be bound to an expression
Const EXPOSE_FORCE      = 16    ' it can be forced at runtime, and is never saved
Const EXPOSE_SAVED      = 32    ' its value is configuration, not a reading
Const EXPOSE_INTERFACE  = 64    ' the Elipse application is given a tag for it
Const EXPOSE_IOTAG      = 128   ' level 3 is given a point for it, over 104


' What the operator reads, and which state says it: normal|active|limit.
'
' One string rather than three fields because the polarity belongs with
' the words. Preconditions is True while the maneuver is permitted, so
' it alarms on 0 where the rest alarm on 1 - and a pair that carried
' only the two words would let someone reword the message without ever
' seeing that the state raising it was the healthy one.
Const PAIR_ACTUATED     = "NORMAL|ATUADO|1"
Const PAIR_BLOCKED      = "LIBERADO|BLOQUEADO|1"
Const PAIR_PRECONDITION = "ATENDIDAS|NÃO ATENDIDAS|0"
Const PAIR_RUNNING      = "CONCLUÍDO|EM ANDAMENTO|1"

' DigitalSeverity, as Power numbers it.
'
' The scale runs backwards: the smaller the number the worse the alarm,
' and -2 is the most severe value here rather than the least. Anything
' comparing two of these has to be read twice - "worse than medium" is
' a < and not a >. The manifests ask for medium unless a signal earns
' otherwise; the overlay is what moves a single alarm off its default.
Const SEV_CRITICAL = -2
Const SEV_HIGH     =  0
Const SEV_MEDIUM   =  1
Const SEV_LOW      =  2



Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	' What the operator is told when this property changes, and which of
	' its two states does the telling. Empty on every property until the
	' alarm table at the foot of the manifest names it - the same way the
	' exposure table is what makes a property appear on the panel. Both
	' default to silence.
	Public AlarmLabel
	Public AlarmPair
	Public AlarmSeverity

	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

	End Function

	' Asked of the property rather than of the caller, so the flags stay in
	' the one scope that declares them. The instances travel to whatever
	' scope reads the manifest and answer there just the same.
	Public Function Shows()
		Shows = Has(EXPOSE_VIEW)
	End Function

	Public Function ShowsValue()
		ShowsValue = Has(EXPOSE_VALUE)
	End Function

	Public Function CanEdit()
		CanEdit = Has(EXPOSE_EDIT)
	End Function

	Public Function CanBind()
		CanBind = Has(EXPOSE_EXPRESSION)
	End Function

	Public Function CanForce()
		CanForce = Has(EXPOSE_FORCE)
	End Function

	Public Function IsSaved()
		IsSaved = Has(EXPOSE_SAVED)
	End Function
	
	Public Function IsInterfaced()
		IsInterfaced = Has(EXPOSE_INTERFACE)
	End Function

	' Whether level 3 is given a point for it.
	'
	' A separate question from IsInterfaced, and asked separately. The
	' interface is where the Elipse application meets the automation; the
	' distribution is what leaves the station. Everything distributed is
	' interfaced - the distribution reads off the interface - but not
	' everything interfaced is distributed, and conflating the two left no
	' way to say so except a list of names kept somewhere else.
	Public Function IsIOTagged()
		IsIOTagged = Has(EXPOSE_IOTAG)
	End Function

	' An unnamed property raises nothing. Empty and "" compare equal in
	' VBScript, so a property the alarm table never mentions answers no
	' here without needing a flag of its own.
	Public Function IsAlarmed()
		IsAlarmed = (AlarmLabel <> "")
	End Function

	' The message either way, in the pattern the control room reads:
	' a label and the state, joined by a dash.
	Public Function AlarmNormalText()
		AlarmNormalText = AlarmLabel & " - " & PairPart(0)
	End Function

	Public Function AlarmActiveText()
		AlarmActiveText = AlarmLabel & " - " & PairPart(1)
	End Function

	Public Function AlarmLimit()
		AlarmLimit = (CLng("0" & PairPart(2)) <> 0)
	End Function

	Private Function PairPart(i)

		PairPart = ""

		Dim parts
		parts = Split(AlarmPair & "", "|")

		If i <= UBound(parts) Then PairPart = parts(i)

	End Function

	' Empty And anything is 0, so a property nobody classified answers no
	' to all of these.
	Private Function Has(flag)
		Has = ((Exposure And flag) <> 0)
	End Function

End Class

Sub AddProperty(bag, name, dataType, initialValue, helpEn, helpPt)

	Dim p
	Set p = New PropertyInfo

	p.Name         = name
	p.DataType     = dataType
	p.InitialValue = initialValue
	p.Exposure     = EXPOSE_NONE
	p.HelpEn       = helpEn
	p.HelpPt       = helpPt

	bag.Add LCase(name), p
	
End Sub

' What the screen may do with a property. Set apart from AddProperty so
' the classifications read as a table, and so changing one never means
' touching the help text - which is where the accents live.
' What the operator is alarmed on. Set apart from SetExposure for the
' reason that one is set apart from AddProperty: the alarms read as a
' table of their own, and a property left out of it raises nothing.
'
' A curated list and never a sweep of what is interfaced. An interface
' tag exists so a screen can draw a value, which is a different question
' from whether an operator should be told about it.
Sub SetAlarm(bag, propertyName, label, pair, severity)

	Dim k
	k = LCase(propertyName)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).AlarmLabel    = label
	bag(k).AlarmPair     = pair
	bag(k).AlarmSeverity = severity

End Sub
Sub SetExposure(bag, name, exposure)

	Dim k
	k = LCase(name)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).Exposure = exposure

End Sub

<xatm_config_data.PropertiesHelper.xatm_Transformer:xatm_Transformer_OnStartRunning()>
Sub xatm_Transformer_OnStartRunning()

	Dim bag
	Set bag = CreateObject("Scripting.Dictionary")

	AddProperty bag, "OutOfService", "Boolean", False, _
		"Transformer is out of service (impediment). The automation reads it to plan the maneuver around this transformer.", _
		"Transformador fora de serviço (impedimento). O automatismo o lê para planejar a manobra sem este transformador."

	AddProperty bag, "Id", "Integer", Empty, _
		"Unique numeric identifier used by the automation to locate this transformer.", _
		"Identificador numérico único usado pelo automatismo para localizar este transformador."

	AddProperty bag, "LockingOutRelay", "Boolean", False, _
		"Locking out relay (86) of the transformer is actuated.", _
		"Relé de bloqueio (86) do transformador atuado."

	AddProperty bag, "UndervoltageRelay", "Boolean", False, _
		"Undervoltage relay (27) of the transformer is actuated.", _
		"Relé de subtensão (27) do transformador atuado."

	' What the transformer's own IED publishes about it, received rather
	' than worked out here. The specification is explicit that the IED is
	' the emitter and this controller a receiver, so the five conditions
	' behind each one - the secondary breaker off, both high-voltage
	' disconnectors open, ANSI 27 actuated, the currents at zero - stay in
	' the IED and arrive as one bit.
	'
	' The relay-failure check the specification asks for goes into the
	' expression rather than into a property of its own: what is wanted is
	' the message AND the messenger being healthy, which is one fact.
	' The busbar chave rele for this transformer, and what starts the
	' high-voltage reclosing.
	'
	' Not the LockingOutRelay above. That one is the transformer's own 86
	' and trips its secondary breaker alone; this one trips the incomer as
	' well, which is what takes the station down and what there is to
	' reclose from.
	'
	' One property and not CR-1 and CR-2 apart. They differ in what they do
	' to the medium-voltage transfer - one starts it, the other bars it -
	' but that difference is written into the transfer's own AutomaticBlock
	' rather than read here. All this has to say is that this transformer
	' took the entry down, so the reclosing knows to wait for it to isolate.
	
	AddProperty bag, "Isolated", "Boolean", False, _
		"Transformer is electrically isolated from the high-voltage busbars. Bound to an expression - the GOOSE its IED publishes, and that IED not having failed.", _
		"Transformador isolado eletricamente das barras de alta tensão. Vinculada a uma expressão - a mensagem GOOSE publicada pelo seu IED, e esse IED sem defeito."

	AddProperty bag, "Maintenance", "Boolean", False, _
		"Transformer is under maintenance - isolated, and with the command locks applied so it stays that way. Bound to an expression.", _
		"Transformador em manutenção - isolado e com os bloqueios de comando aplicados para que assim permaneça. Vinculada a uma expressão."

	AddProperty bag, "UndervoltageDelay", "Integer", 25, _
		"Time in seconds the undervoltage (27) condition has to persist before the automation acts on it.", _
		"Tempo em segundos que a condição de subtensão (27) deve permanecer antes que o automatismo atue."

	' The three readings take an expression and can be forced for a test,
	' and none of them is saved - the expression is the configuration.
	SetExposure bag, "Id",                EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "UndervoltageDelay", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED

	SetExposure bag, "OutOfService",      EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE
	SetExposure bag, "LockingOutRelay",   EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE
	SetExposure bag, "UndervoltageRelay", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE
	SetExposure bag, "Isolated",          EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE
	SetExposure bag, "Maintenance",       EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE

	Set Value = bag

End Sub

' What the configuration screen may do with a property, and whether its
' value is a setting at all. A bitmask: a property can be bound to an
' expression and forced, or shown and not edited, and so on.
'
' AddProperty leaves every property at EXPOSE_NONE, so nothing appears on
' the screen and nothing is written to the project until the table at the
' foot of the manifest says so. Both defaults fail closed.
Const EXPOSE_NONE       = 0
Const EXPOSE_VIEW       = 1     ' a row appears for it
Const EXPOSE_VALUE      = 2     ' its value is shown - never for a write-only command
Const EXPOSE_EDIT       = 4     ' the value can be typed
Const EXPOSE_EXPRESSION = 8     ' it can be bound to an expression
Const EXPOSE_FORCE      = 16    ' it can be forced at runtime, and is never saved
Const EXPOSE_SAVED      = 32    ' its value is configuration, not a reading
Const EXPOSE_INTERFACE  = 64    ' the Elipse application is given a tag for it
Const EXPOSE_IOTAG      = 128   ' level 3 is given a point for it, over 104


' What the operator reads, and which state says it: normal|active|limit.
'
' One string rather than three fields because the polarity belongs with
' the words. Preconditions is True while the maneuver is permitted, so
' it alarms on 0 where the rest alarm on 1 - and a pair that carried
' only the two words would let someone reword the message without ever
' seeing that the state raising it was the healthy one.
Const PAIR_ACTUATED     = "NORMAL|ATUADO|1"
Const PAIR_BLOCKED      = "LIBERADO|BLOQUEADO|1"
Const PAIR_PRECONDITION = "ATENDIDAS|NÃO ATENDIDAS|0"
Const PAIR_RUNNING      = "CONCLUÍDO|EM ANDAMENTO|1"

' DigitalSeverity, as Power numbers it.
'
' The scale runs backwards: the smaller the number the worse the alarm,
' and -2 is the most severe value here rather than the least. Anything
' comparing two of these has to be read twice - "worse than medium" is
' a < and not a >. The manifests ask for medium unless a signal earns
' otherwise; the overlay is what moves a single alarm off its default.
Const SEV_CRITICAL = -2
Const SEV_HIGH     =  0
Const SEV_MEDIUM   =  1
Const SEV_LOW      =  2
Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	' What the operator is told when this property changes, and which of
	' its two states does the telling. Empty on every property until the
	' alarm table at the foot of the manifest names it - the same way the
	' exposure table is what makes a property appear on the panel. Both
	' default to silence.
	Public AlarmLabel
	Public AlarmPair
	Public AlarmSeverity

	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

	End Function

	' Asked of the property rather than of the caller, so the flags stay in
	' the one scope that declares them. The instances travel to whatever
	' scope reads the manifest and answer there just the same.
	Public Function Shows()
		Shows = Has(EXPOSE_VIEW)
	End Function

	Public Function ShowsValue()
		ShowsValue = Has(EXPOSE_VALUE)
	End Function

	Public Function CanEdit()
		CanEdit = Has(EXPOSE_EDIT)
	End Function

	Public Function CanBind()
		CanBind = Has(EXPOSE_EXPRESSION)
	End Function

	Public Function CanForce()
		CanForce = Has(EXPOSE_FORCE)
	End Function

	Public Function IsSaved()
		IsSaved = Has(EXPOSE_SAVED)
	End Function
	
	Public Function IsInterfaced()
		IsInterfaced = Has(EXPOSE_INTERFACE)
	End Function

	' Whether level 3 is given a point for it.
	'
	' A separate question from IsInterfaced, and asked separately. The
	' interface is where the Elipse application meets the automation; the
	' distribution is what leaves the station. Everything distributed is
	' interfaced - the distribution reads off the interface - but not
	' everything interfaced is distributed, and conflating the two left no
	' way to say so except a list of names kept somewhere else.
	Public Function IsIOTagged()
		IsIOTagged = Has(EXPOSE_IOTAG)
	End Function

	' An unnamed property raises nothing. Empty and "" compare equal in
	' VBScript, so a property the alarm table never mentions answers no
	' here without needing a flag of its own.
	Public Function IsAlarmed()
		IsAlarmed = (AlarmLabel <> "")
	End Function

	' The message either way, in the pattern the control room reads:
	' a label and the state, joined by a dash.
	Public Function AlarmNormalText()
		AlarmNormalText = AlarmLabel & " - " & PairPart(0)
	End Function

	Public Function AlarmActiveText()
		AlarmActiveText = AlarmLabel & " - " & PairPart(1)
	End Function

	Public Function AlarmLimit()
		AlarmLimit = (CLng("0" & PairPart(2)) <> 0)
	End Function

	Private Function PairPart(i)

		PairPart = ""

		Dim parts
		parts = Split(AlarmPair & "", "|")

		If i <= UBound(parts) Then PairPart = parts(i)

	End Function
	
	
	' Empty And anything is 0, so a property nobody classified answers no
	' to all of these.
	Private Function Has(flag)
		Has = ((Exposure And flag) <> 0)
	End Function

End Class

Sub AddProperty(bag, name, dataType, initialValue, helpEn, helpPt)

	Dim p
	Set p = New PropertyInfo

	p.Name         = name
	p.DataType     = dataType
	p.InitialValue = initialValue
	p.Exposure     = EXPOSE_NONE
	p.HelpEn       = helpEn
	p.HelpPt       = helpPt

	bag.Add LCase(name), p
	
End Sub

' What the screen may do with a property. Set apart from AddProperty so
' the classifications read as a table, and so changing one never means
' touching the help text - which is where the accents live.
' What the operator is alarmed on. Set apart from SetExposure for the
' reason that one is set apart from AddProperty: the alarms read as a
' table of their own, and a property left out of it raises nothing.
'
' A curated list and never a sweep of what is interfaced. An interface
' tag exists so a screen can draw a value, which is a different question
' from whether an operator should be told about it.
Sub SetAlarm(bag, propertyName, label, pair, severity)

	Dim k
	k = LCase(propertyName)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).AlarmLabel    = label
	bag(k).AlarmPair     = pair
	bag(k).AlarmSeverity = severity

End Sub
Sub SetExposure(bag, name, exposure)

	Dim k
	k = LCase(name)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).Exposure = exposure
		
End Sub

<xatm_config_data.PropertiesHelper.xatm_Version:xatm_Version_OnStartRunning()>
Sub xatm_Version_OnStartRunning()

	Dim bag
	Set bag = CreateObject("Scripting.Dictionary")

	' Which library is running, published so it can be asked from outside.
	'
	' Nothing here is configuration. Every one is written by the class at
	' start from constants inside xatm_lib.lib, so none is saved and none is
	' edited - swapping the library file is what changes them.
	'
	' Whether the build is a demo lives elsewhere, on xatm_Build, which sits
	' inside every device because that is where the command gate needs it.
	' Two different facts, two different homes.
	AddProperty bag, "Version", "String", "", _
		"Library version as it is written, for a person to read.", _
		"Versão da biblioteca como é escrita, para uma pessoa ler."

	' The same version as three numbers, which is what travels. An interface
	' tag and a 104 point carry a value rather than a sentence, and an
	' operation centre auditing a fleet wants to compare numbers instead of
	' parsing text at the far end.
	AddProperty bag, "Major", "Integer", 0, _
		"Major part of the library version.", _
		"Parte maior da versão da biblioteca."

	AddProperty bag, "Minor", "Integer", 0, _
		"Minor part of the library version.", _
		"Parte menor da versão da biblioteca."

	AddProperty bag, "Patch", "Integer", 0, _
		"Patch part of the library version.", _
		"Parte de correção da versão da biblioteca."

	AddProperty bag, "ReleaseNotes", "String", "", _
		"What changed in this release.", _
		"O que mudou nesta versão."

	' Shown and never edited, and only the three numbers leave the station.
	' The written version and the notes are for the panel: a sentence and a
	' page of notes are not what a 104 point is for.
	SetExposure bag, "Version",      EXPOSE_VIEW + EXPOSE_VALUE
	SetExposure bag, "ReleaseNotes", EXPOSE_VIEW + EXPOSE_VALUE

	SetExposure bag, "Major", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "Minor", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_INTERFACE + EXPOSE_IOTAG
	SetExposure bag, "Patch", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_INTERFACE + EXPOSE_IOTAG

	' Not alarmed. A version is not an event.

	Set Value = bag

End Sub

' What the configuration screen may do with a property, and whether its
' value is a setting at all. A bitmask: a property can be bound to an
' expression and forced, or shown and not edited, and so on.
'
' AddProperty leaves every property at EXPOSE_NONE, so nothing appears on
' the screen and nothing is written to the project until the table at the
' foot of the manifest says so. Both defaults fail closed.
Const EXPOSE_NONE       = 0
Const EXPOSE_VIEW       = 1     ' a row appears for it
Const EXPOSE_VALUE      = 2     ' its value is shown - never for a write-only command
Const EXPOSE_EDIT       = 4     ' the value can be typed
Const EXPOSE_EXPRESSION = 8     ' it can be bound to an expression
Const EXPOSE_FORCE      = 16    ' it can be forced at runtime, and is never saved
Const EXPOSE_SAVED      = 32    ' its value is configuration, not a reading
Const EXPOSE_INTERFACE  = 64    ' the Elipse application is given a tag for it
Const EXPOSE_IOTAG      = 128   ' level 3 is given a point for it, over 104


' What the operator reads, and which state says it: normal|active|limit.
'
' One string rather than three fields because the polarity belongs with
' the words. Preconditions is True while the maneuver is permitted, so
' it alarms on 0 where the rest alarm on 1 - and a pair that carried
' only the two words would let someone reword the message without ever
' seeing that the state raising it was the healthy one.
Const PAIR_ACTUATED     = "NORMAL|ATUADO|1"
Const PAIR_BLOCKED      = "LIBERADO|BLOQUEADO|1"
Const PAIR_PRECONDITION = "ATENDIDAS|NÃO ATENDIDAS|0"
Const PAIR_RUNNING      = "CONCLUÍDO|EM ANDAMENTO|1"

' DigitalSeverity, as Power numbers it.
'
' The scale runs backwards: the smaller the number the worse the alarm,
' and -2 is the most severe value here rather than the least. Anything
' comparing two of these has to be read twice - "worse than medium" is
' a < and not a >. The manifests ask for medium unless a signal earns
' otherwise; the overlay is what moves a single alarm off its default.
Const SEV_CRITICAL = -2
Const SEV_HIGH     =  0
Const SEV_MEDIUM   =  1
Const SEV_LOW      =  2

Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	' What the operator is told when this property changes, and which of
	' its two states does the telling. Empty on every property until the
	' alarm table at the foot of the manifest names it - the same way the
	' exposure table is what makes a property appear on the panel. Both
	' default to silence.
	Public AlarmLabel
	Public AlarmPair
	Public AlarmSeverity
	
	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

	End Function

	' Asked of the property rather than of the caller, so the flags stay in
	' the one scope that declares them. The instances travel to whatever
	' scope reads the manifest and answer there just the same.
	Public Function Shows()
		Shows = Has(EXPOSE_VIEW)
	End Function

	Public Function ShowsValue()
		ShowsValue = Has(EXPOSE_VALUE)
	End Function

	Public Function CanEdit()
		CanEdit = Has(EXPOSE_EDIT)
	End Function

	Public Function CanBind()
		CanBind = Has(EXPOSE_EXPRESSION)
	End Function

	Public Function CanForce()
		CanForce = Has(EXPOSE_FORCE)
	End Function

	Public Function IsSaved()
		IsSaved = Has(EXPOSE_SAVED)
	End Function
	
	Public Function IsInterfaced()
		IsInterfaced = Has(EXPOSE_INTERFACE)
	End Function

	' Whether level 3 is given a point for it.
	'
	' A separate question from IsInterfaced, and asked separately. The
	' interface is where the Elipse application meets the automation; the
	' distribution is what leaves the station. Everything distributed is
	' interfaced - the distribution reads off the interface - but not
	' everything interfaced is distributed, and conflating the two left no
	' way to say so except a list of names kept somewhere else.
	Public Function IsIOTagged()
		IsIOTagged = Has(EXPOSE_IOTAG)
	End Function

	' An unnamed property raises nothing. Empty and "" compare equal in
	' VBScript, so a property the alarm table never mentions answers no
	' here without needing a flag of its own.
	Public Function IsAlarmed()
		IsAlarmed = (AlarmLabel <> "")
	End Function

	' The message either way, in the pattern the control room reads:
	' a label and the state, joined by a dash.
	Public Function AlarmNormalText()
		AlarmNormalText = AlarmLabel & " - " & PairPart(0)
	End Function

	Public Function AlarmActiveText()
		AlarmActiveText = AlarmLabel & " - " & PairPart(1)
	End Function

	Public Function AlarmLimit()
		AlarmLimit = (CLng("0" & PairPart(2)) <> 0)
	End Function

	Private Function PairPart(i)

		PairPart = ""

		Dim parts
		parts = Split(AlarmPair & "", "|")

		If i <= UBound(parts) Then PairPart = parts(i)

	End Function
	
	' Empty And anything is 0, so a property nobody classified answers no
	' to all of these.
	Private Function Has(flag)
		Has = ((Exposure And flag) <> 0)
	End Function

End Class

Sub AddProperty(bag, name, dataType, initialValue, helpEn, helpPt)

	Dim p
	Set p = New PropertyInfo

	p.Name         = name
	p.DataType     = dataType
	p.InitialValue = initialValue
	p.Exposure     = EXPOSE_NONE
	p.HelpEn       = helpEn
	p.HelpPt       = helpPt

	bag.Add LCase(name), p
	
End Sub

' What the screen may do with a property. Set apart from AddProperty so
' the classifications read as a table, and so changing one never means
' touching the help text - which is where the accents live.
' What the operator is alarmed on. Set apart from SetExposure for the
' reason that one is set apart from AddProperty: the alarms read as a
' table of their own, and a property left out of it raises nothing.
'
' A curated list and never a sweep of what is interfaced. An interface
' tag exists so a screen can draw a value, which is a different question
' from whether an operator should be told about it.
Sub SetAlarm(bag, propertyName, label, pair, severity)

	Dim k
	k = LCase(propertyName)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).AlarmLabel    = label
	bag(k).AlarmPair     = pair
	bag(k).AlarmSeverity = severity

End Sub
Sub SetExposure(bag, name, exposure)

	Dim k
	k = LCase(name)

	If Not bag.Exists(k) Then Exit Sub

	bag(k).Exposure = exposure
				
End Sub

<xatm_config_data.SimulationMode:SimulationMode_OnChangedValue()>
Sub SimulationMode_OnChangedValue()


    ScanFolder Application.GetObject("XATM_Data.Substation"), Value

End Sub

' Returns True if at least one position tag has good quality - breaker does NOT need sim mode.
Function AnyPositionTagHealthy(device)

    If TagHealthy(device.PositionOpen)      Then AnyPositionTagHealthy = True : Exit Function
    If TagHealthy(device.PositionClosed)    Then AnyPositionTagHealthy = True : Exit Function
    If TagHealthy(device.PositionOpenAlt)   Then AnyPositionTagHealthy = True : Exit Function
    If TagHealthy(device.PositionClosedAlt) Then AnyPositionTagHealthy = True : Exit Function
    AnyPositionTagHealthy = False

End Function

Function TagHealthy(tag)
    
    On Error Resume Next
    Dim q
    q = tag.Quality
    TagHealthy = (Err.Number = 0) And (q >= 192)
    On Error GoTo 0

End Function

Sub ScanFolder(folder, simEnabled)
    
    Dim obj
    For Each obj In folder
        Select Case UCase(TypeName(obj))

            ' A disconnector carries the same four position properties a
            ' breaker does, so one branch serves both. Without it the Else
            ' below walked into a disconnector as though it were a folder,
            ' which quietly left it out of simulation altogether.
            Case "XATM_BREAKER", "XATM_DISCONNECTOR"
    
                Dim shouldSim
                
                If simEnabled Then
                    shouldSim = Not AnyPositionTagHealthy(obj)
                Else
                    shouldSim = False
                End If
                
                obj.Item("Data").Item("SimulationModeEnabled").WriteEx shouldSim
			
			Case "XATM_TRANSFORMER"
				
				obj.Item("Data").Item("SimulationModeEnabled").WriteEx simEnabled
			
            Case Else
                On Error Resume Next
                ScanFolder obj, simEnabled
                On Error GoTo 0
        End Select
    Next
		
End Sub

<xatm_config_screens.Config.TreeView:TreeView_NodeClick(Node)>
Sub TreeView_NodeClick(Node)

	' What the node stands for - AddBranch put it on the node when it built
	' the tree.
	Dim key
	key = ""

	On Error Resume Next
	key = Node.Tag
	On Error Goto 0

	' The rows of the last selection go either way, so a click that shows
	' nothing leaves nothing behind.
	ClearPropertyRows

	If key = "" Then Exit Sub
	
	Dim content
	content = Empty

	On Error Resume Next
	content = Application.GetObject("xatm_config_data.Config.XMLContent").Value
	On Error Goto 0

	If IsEmpty(content) Or IsNull(content) Then Exit Sub

	Dim doc
	Set doc = NewDomDocument()

	If Not doc.loadXML(StripBom(CStr(content))) Then Exit Sub

	' A folder is a node too, and has no properties of its own to show -
	' the document answers which kind this is by having an object for the
	' key, or not.
	Dim objectNode
	Set objectNode = FindObject(doc, key)

	If objectNode Is Nothing Then Exit Sub

	BuildPropertyRows objectNode, key

End Sub


Const ROW_CLASS = "xatm_PropertyRow"

' The corner the rows stack from, in pixels - Himetric converts. A row
' keeps the size it was drawn at, so there is no width or height here.
Const ROW_LEFT_PX = 920
Const ROW_TOP_PX  = 56
Const ROW_GAP_PX  = 0


' What the object's own name is shown as. It has no manifest entry, so
' the row has no help text for it and nothing declares its type - and its
' exposure is stated here instead: shown, with a readout, and editable.
Const NAME_PROPERTY = "Name"
Const NAME_TYPE     = "String"
Const NAME_EXPOSURE = 7          ' EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT

' The language a row's tip is written in. Every manifest carries both, and
' this is the one place that picks between them - a tag would do instead,
' the day the screen has to switch without a restart.
Const HELP_LANG = "pt-BR"

' The object's own name has no manifest entry, so its help is stated here
' the way its exposure is.
Const NAME_HELP_EN = "The name this object is known by in the project."
Const NAME_HELP_PT = "Nome pelo qual este objeto é conhecido no projeto."

' Where the manifests live - the same place the export and the import
' read them from.
Const CONFIG_DATA   = "xatm_config_data"
Const HELPER_FOLDER = "PropertiesHelper"

' The folder whose objects are named by the layout rather than by hand.
Const AUTOMATION_FOLDER = "Automation"

' Which of the two an edit is asking to change, told to the row when it
' is built so it never has to work it out from the property name.
Const KIND_NAME     = "name"
Const KIND_PROPERTY = "property"

' What marks a key as an Id rather than a path.
'Const ID_PREFIX     = "id:"

' E3 places and sizes objects in himetric, not pixels.
Function Himetric(pixels)

	Himetric = CLng((pixels * 2540) / 96)

End Function


' The object a key stands for, Nothing when the document has none - which
' is the answer for a folder, since only objects carry properties.
'
' A copy: AddBranch builds the keys in the button's script and the click
' reads them here, and one E3 object has no way to call a procedure in
' another. The two have to say the same thing about a key.
Function FindObject(doc, key)

	Set FindObject = Nothing

	' The Len is tested and not just the comparison. Were ID_PREFIX missing
	' from this scope it would be Empty, Len would be 0, and Left(key, 0)
	' would be "" - which equals Empty in VBScript. Every key would then
	' take the Id branch with the prefix still on the front of it, match
	' nothing, and say nothing about why. Testing the Len first costs a
	' missing const the Id keys only, and leaves the path ones working.
	If Len(ID_PREFIX) > 0 And Left(key, Len(ID_PREFIX)) = ID_PREFIX Then
		Set FindObject = doc.selectSingleNode( _
		                 "//object[property[@name='Id']/@value='" & _
		                 Mid(key, Len(ID_PREFIX) + 1) & "']")
	Else
		Set FindObject = doc.selectSingleNode("//object[@path=""" & key & """]")
	End If

End Function


' A path E3 will resolve. A name that starts with a digit or an
' underscore, or carries a space or a hyphen, has to be bracketed, and
' bracketing every piece is always right. One already bracketed is left
' alone - PathName may have bracketed it, and [[52-01]] resolves no
' better than 52-01 does.
'
' A copy. The row it is bracketing for is built in this script, and one
' E3 object has no way to call a procedure defined in another.
Function E3Path(path)

	E3Path = ""
	If path = "" Then Exit Function

	Dim pieces
	pieces = SplitPath(path)

	Dim i
	For i = 0 To UBound(pieces)
		If Left(pieces(i), 1) <> "[" Then
			pieces(i) = "[" & pieces(i) & "]"
		End If
	Next

	E3Path = Join(pieces, ".")

End Function


' The pieces of a path, split on the dots between them - the ones outside
' brackets, since a bracketed name may carry a dot of its own.
Function SplitPath(path)

	Dim pieces()
	ReDim pieces(0)
	pieces(0) = ""

	Dim i, n, depth, ch
	n = 0
	depth = 0

	For i = 1 To Len(path)

		ch = Mid(path, i, 1)

		If ch = "[" Then
			depth = depth + 1
			pieces(n) = pieces(n) & ch
		ElseIf ch = "]" Then
			depth = depth - 1
			pieces(n) = pieces(n) & ch
		ElseIf ch = "." And depth = 0 Then
			n = n + 1
			ReDim Preserve pieces(n)
			pieces(n) = ""
		Else
			pieces(n) = pieces(n) & ch
		End If

	Next

	SplitPath = pieces

End Function


' A row per property the manifest says to show, in the order the export
' wrote them - which is the order the manifest declares them in.
'
' The document says what a property holds; the manifest says what may be
' done with it. A property it does not classify gets no row at all, which
' is how the latches and the internals stay off the panel.
Sub BuildPropertyRows(objectNode, key)

	Dim objectType
	objectType = Attribute(objectNode, "type")

	' Where the object is, for a row that wants to link to it live.
	' Source is a key and not a path - id:700 for anything with an Id -
	' because a rename must not strand the panel. A link needs the real
	' thing, bracketed here where E3Path is, so the control never has to
	' know which names E3 needs brackets around.
	Dim objectPath
	objectPath = E3Path(Attribute(objectNode, "path") & "")

	Dim bag
	Set bag = ManifestOf(objectType)

	Dim properties
	Set properties = objectNode.selectNodes("property")

	Dim i, property, row, y, p
	y = Himetric(ROW_TOP_PX)

	' The object's own name goes first. It is the object rather than
	' something declared on it, so it is neither a property of the document
	' nor an entry in the manifest, and its exposure is stated here.
	'
	' Not under Automation, though: SyncAutomation finds an automation by the
	' number on the end of its name, so the name is structure there and not
	' the operator's to change.
	If Not IsAutomation(objectNode) Then
		Set row = NewRow(KIND_NAME, key, objectPath, objectType, NAME_PROPERTY, NAME_TYPE, _
		                 Attribute(objectNode, "name"), Empty, NAME_EXPOSURE, NameHelp(), y)
		y = y + row.Height + Himetric(ROW_GAP_PX)
	End If

	For i = 0 To properties.length - 1

		Set property = properties.item(i)

		Set p = Nothing
		If Not (bag Is Nothing) Then
			Dim k
			k = LCase(Attribute(property, "name") & "")
			If bag.Exists(k) Then Set p = bag(k)
		End If

		' Nothing the manifest does not vouch for. A property it has never
		' heard of is one the export wrote and nobody classified, and showing
		' it would be guessing at what the operator may do with it.
		If Not (p Is Nothing) Then
			If p.Shows() Then

				Set row = NewRow(KIND_PROPERTY, key, objectPath, objectType, Attribute(property, "name"), _
				                 Attribute(property, "type"), Attribute(property, "value"), _
				                 Attribute(property, "source"), p.Exposure, p.Help(HELP_LANG), y)

				' A row comes out at the size it was drawn, so the next one goes
				' under whatever that turned out to be.
				y = y + row.Height + Himetric(ROW_GAP_PX)

			End If
		End If

	Next

End Sub


' The tip the name row shows, in whichever language HELP_LANG picks. It
' is the only row whose help is not in a manifest, the name being the
' object rather than something declared on it.
Function NameHelp()

	If HELP_LANG = "pt-BR" Then
		NameHelp = NAME_HELP_PT
	Else
		NameHelp = NAME_HELP_EN
	End If

End Function


' The manifest declared for a class, or Nothing when the class has none -
' the same lookup the export and the import do.
Function ManifestOf(className)

	Set ManifestOf = Nothing

	On Error Resume Next
	Set ManifestOf = Application.GetObject(CONFIG_DATA).Item(HELPER_FOLDER).Item(className).Value
	On Error Goto 0

End Function


' True for anything under the Automation folder, however deep. Asked of
' the element and not of the path, so a folder renamed on the way down
' cannot fool it.
Function IsAutomation(objectNode)

	IsAutomation = Not (objectNode.selectSingleNode( _
	               "ancestor::folder[@name='" & AUTOMATION_FOLDER & "']") Is Nothing)

End Function


' One row on the screen. Added inactive so every property is set before
' the control goes up and reads them, and handed back so the caller can
' step past the height it came out at.
Function NewRow(kind, key, objectPath, objectType, propertyName, propertyType, value, source, exposure, help, y)

	Dim row
	Set row = Screen.AddObject(ROW_CLASS, False)
	
	row.BringToFront()
	
	row.X = Himetric(ROW_LEFT_PX)
	row.Y = y
	
	row.Kind         = kind
	row.Source       = key
	row.ObjectPath   = objectPath
	row.ObjectType   = objectType
	row.PropertyName = propertyName
	row.PropertyType = propertyType
	row.Value        = value

	' Where the property gets its value from, when it is not a value of its
	' own: the path of the tag an IOTag is associated with, or the
	' expression a bound property carries. Named apart from Source, which
	' is the object this row belongs to, the way it is on every other
	' control in xatm_libconfig.
	'
	' It comes from PathName and so already resolves - a driver called 61850
	' arrives as [61850].DJ03_POS, bracketed where it had to be. Nothing
	' here should bracket it again.
	row.PropertySource = source

	' What the row may offer: a readout, a field to type in, the expression
	' dialog, a force. The row reads the flags rather than working any of it
	' out from the type - two Booleans on a transformer want completely
	' different controls.
	row.Exposure     = exposure
	
	' What the row shows as its tip. The manifest wrote it, in both
	' languages, next to the property it describes - which is where it stays
	' accurate, being written by whoever declared what the property is for.
	row.Help         = help

	row.Activate()

	Set NewRow = row

End Function

' Takes the rows of the last selection off the screen. Collected first
' and deleted second - a collection is not walked while what is in it is
' being removed.
Sub ClearPropertyRows()

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim obj, n
	n = 0

	For Each obj In Screen
		If UCase(TypeName(obj)) = UCase(ROW_CLASS) Then
			doomed.Add n, obj.Name
			n = n + 1
		End If
	Next

	For Each n In doomed.Keys
		On Error Resume Next
		Screen.DeleteObject doomed(n)
		On Error Goto 0
	Next

End Sub


' An attribute of an element, Empty when it carries none. The export
' drops the value attribute of a property that is unset, so an absent one
' has to come back Empty and not as "".
Function Attribute(element, name)

	Attribute = Empty

	Dim a
	Set a = element.getAttributeNode(name)
	If a Is Nothing Then Exit Function

	Attribute = a.value

End Function


' MSXML 6 where the machine has it, the version-independent progid
' otherwise.
Function NewDomDocument()

	Set NewDomDocument = Nothing

	On Error Resume Next
	Set NewDomDocument = CreateObject("MSXML2.DOMDocument.6.0")
	On Error Goto 0

	If NewDomDocument Is Nothing Then
		Set NewDomDocument = CreateObject("MSXML2.DOMDocument")
	End If

	NewDomDocument.async = False
	NewDomDocument.preserveWhiteSpace = True

	On Error Resume Next
	NewDomDocument.setProperty "SelectionLanguage", "XPath"
	On Error Goto 0

End Function


' A byte-order mark survives a round trip through a utf-8 file and
' makes loadXML fail on what looks like perfectly good XML.
Function StripBom(text)

	StripBom = text

	If Len(text) = 0 Then Exit Function

	If AscW(Left(text, 1)) = &HFEFF Then
		StripBom = Mid(text, 2)
	End If

End Function

Sub Foo()	

End Sub

<xatm_config_screens.Config.TreeView:TreeView_OnStartRunning()>
Sub TreeView_OnStartRunning()

	Application.GetObject("xatm_config_data.Config.UpdateTreeviewSignal").WriteEx True
	
End Sub

<xatm_config_screens.Config.TreeView:TreeView_UpdateTreeview()>
Sub TreeView_UpdateTreeview()
	
	' Builds the tree by hand, from the export as it stands in the tag. The
	' same thing Config_OnPreShow does on open and btnApply does after an
	' Apply - here it says out loud what went wrong instead of doing nothing.

	Dim tree
	Set tree = Screen.Item(TREE_CONTROL)
	
	ClearPropertyRows
	
	If tree Is Nothing Then
		MsgBox "No tree control named '" & TREE_CONTROL & "' was found on this screen.", _
		       vbExclamation, "Tree"
		Exit Sub
	End If

	Dim content
	content = Empty

	On Error Resume Next
	content = Application.GetObject("xatm_config_data.Config.XMLContent").Value
	On Error Goto 0

	If IsEmpty(content) Or IsNull(content) Then
		'MsgBox "'XMLContent' is empty - run the XML export first.", vbExclamation, "Tree"
		Exit Sub
	End If

	Dim doc
	Set doc = NewDomDocument()

	If Not doc.loadXML(StripBom(CStr(content))) Then
		MsgBox "'XMLContent' does not parse as XML:" & vbCrLf & vbCrLf & _
		       doc.parseError.reason, vbCritical, "Tree"
		Exit Sub
	End If

	PopulateTree tree, doc

	'MsgBox tree.Nodes.Count & " nodes built from the export.", vbInformation, "Tree"

End Sub


Const NODE_ELEMENT = 1

' Name of the tree control on the screen, and the Add relationship that
' hangs a node under another one (tvwChild).
Const TREE_CONTROL = "TreeView"
Const TVW_CHILD    = 4

' Name of the ImageList control that carries the tree's icons, and the
' list itself once the tree has taken it. Held here because AddBranch
' recurses, and threading the list through every call buys nothing.
Const IMAGES_CONTROL = "TreeImages"
Dim gTreeImages

' What an engineer puts in the DocString of a folder or a device to keep
' it, and everything under it, off the tree.
Const NO_SHOW      = "$NoShow$"

' Fills the tree from the document: one node per folder and per object.
' The two folders at the top of the document - Automation and Substation -
' become the two roots. A property or a tag configures a node, it is not
' a node of its own.
Sub PopulateTree(tree, doc)

	tree.Nodes.Clear
	
	' The icons, taken while the tree is empty. The control keeps the
	' images its nodes were given, so a list handed over after a node is
	' on the tree is a list that node never sees.
	Set gTreeImages = BoundImages(tree)

	If doc.documentElement Is Nothing Then Exit Sub

	' Where the document was exported from - the head of every path in it.
	Dim root
	root = doc.documentElement.getAttribute("root")
	If IsNull(root) Then root = ""

	Dim children, i
	children = BranchChildren(doc.documentElement)

	For i = 0 To UBound(children)
		AddBranch tree, children(i), 0, root
	Next

End Sub


' One node for an element, then the same for the elements under it.
' A parentIndex of 0 makes a root - VBScript cannot leave out the two
' arguments Add wants for that, so the call is split in two instead.
Sub AddBranch(tree, element, parentIndex, parentPath)

	Dim path
	path = ElementPath(element, parentPath)

	' Nothing under a hidden element is walked either, so marking a folder
	' takes everything in it off the tree along with it.
	If IsHidden(path) Then Exit Sub

	Dim node

	If parentIndex = 0 Then
		Set node = tree.Nodes.Add()
	Else
		Set node = tree.Nodes.Add(parentIndex, TVW_CHILD)
	End If

	node.Text = NodeLabel(element)

	' The icon the ImageList carries for this kind of node, where it
	' carries one.
	SetNodeImage node, element

	' What NodeClick reads back to know what was selected, and what every
	' row built from it is addressed by.
	node.Tag = NodeKey(element, path)

	' Nothing is expanded on the way in, and a node the control has just
	' been given is collapsed - so the tree opens on the two roots alone.
	Dim children, i
	children = BranchChildren(element)

	For i = 0 To UBound(children)
		AddBranch tree, children(i), node.Index, path
	Next

End Sub


' The elements that are nodes of the tree.
Function IsBranch(element)

	IsBranch = False

	If element.nodeType <> NODE_ELEMENT Then Exit Function

	Select Case element.nodeName
		Case "folder", "object" : IsBranch = True
	End Select

End Function


' The children of an element that are nodes, in the order the tree shows
' them rather than the order anything holds them in. Only the tree is
' ordered by this - E3 and the document are left to hold them however
' they like, and neither is asked to change.
Function BranchChildren(element)

	Dim items()
	ReDim items(0)

	Dim n
	n = -1

	Dim child
	For Each child In element.childNodes

		If IsBranch(child) Then
			n = n + 1
			ReDim Preserve items(n)
			Set items(n) = child
		End If

	Next

	If n < 0 Then
		BranchChildren = Array()
		Exit Function
	End If

	SortBranches items, n

	BranchChildren = items

End Function


' Insertion sort - it is one folder's worth of children and VBScript has
' no sort of its own. Stable, so everything sharing a key keeps the
' order the document had it in, which is what holds the folders still.
Sub SortBranches(items, last)

	Dim i, j, key, current

	For i = 1 To last

		Set current = items(i)
		key = SortKey(current)

		j = i - 1

		Do While j >= 0
			If SortKey(items(j)) <= key Then Exit Do
			Set items(j + 1) = items(j)
			j = j - 1
		Loop

		Set items(j + 1) = current

	Next

End Sub


' What a node sorts on.
'
' Folders share one key and so keep the order the document has them in,
' ahead of everything else - Incomer, Transformer, Busbar, Feeder is a
' deliberate order and alphabetising it would lose it.
'
' An object sorts on its Id where its class declares one and on its name
' where it does not. That is what puts the devices in Id order, which
' pairs each transformer with its own breaker, and the automations in
' name order, an automation having no Id to go by.
Function SortKey(element)

	If element.nodeName = "folder" Then
		SortKey = "0"
		Exit Function
	End If

	Dim id
	id = PropertyValue(element, "Id")

	If id <> "" Then

		If IsNumeric(id) Then
			' Padded, or 1000 would sort in front of 700.
			SortKey = "1" & Right("000000000000" & CLng(id), 12)
		Else
			SortKey = "1" & LCase(id)
		End If

		Exit Function

	End If

	SortKey = "2" & LCase(NodeText(element))

End Function


' The value attribute of a named property - "" when the property is
' absent or unset.
Function PropertyValue(objectNode, propertyName)

	PropertyValue = ""

	Dim p
	Set p = objectNode.selectSingleNode("property[@name='" & propertyName & "']")
	If p Is Nothing Then Exit Function

	Dim a
	Set a = p.getAttributeNode("value")
	If a Is Nothing Then Exit Function

	PropertyValue = a.value

End Function


' What marks a key as an Id rather than a path.
Const ID_PREFIX = "id:"


' What a node is addressed by, and with it every row built from that
' node. An object whose class declares an Id is addressed by the Id,
' because renaming it changes its path and anything still holding the old
' one goes stale - the node, and all twenty rows on the panel. A folder,
' and an object with no Id, has only its path to go by; neither can be
' renamed from this screen.
Function NodeKey(element, path)

	NodeKey = path

	If element.nodeName <> "object" Then Exit Function

	Dim id
	id = PropertyValue(element, "Id")
	If id = "" Then Exit Function

	NodeKey = ID_PREFIX & id

End Function


' The E3 path of an element. An object carries its own; a folder is named
' and nothing more, so its path is its parent's with its name on the end.
Function ElementPath(element, parentPath)

	Dim path
	path = element.getAttribute("path")

	If Not IsNull(path) Then
		ElementPath = CStr(path)
		Exit Function
	End If

	If parentPath = "" Then
		ElementPath = element.getAttribute("name")
	Else
		ElementPath = parentPath & "." & element.getAttribute("name")
	End If

End Function


' True when the object at a path is marked to stay off the tree. The mark
' is read from E3 rather than from the document, because it is set by hand
' in Elipse Studio and an export taken before that was set would not carry
' it. Something the document names that E3 has not got - an object Apply
' has only just added - cannot have been marked, so it stays on the tree.
Function IsHidden(path)

	IsHidden = False

	Dim obj
	Set obj = Nothing
	On Error Resume Next
	Set obj = Application.GetObject(E3Path(path))
	On Error Goto 0

	If obj Is Nothing Then Exit Function

	Dim text
	text = ""
	On Error Resume Next
	text = obj.DocString
	On Error Goto 0

	If IsNull(text) Or IsEmpty(text) Then Exit Function

	IsHidden = (InStr(1, CStr(text), NO_SHOW, vbTextCompare) > 0)

End Function


' What a node is called. Every folder and object carries a name, so the
' element name is only ever a fallback for a malformed document.
Function NodeText(element)

	Dim name
	name = element.getAttribute("name")

	If IsNull(name) Then
		NodeText = "<" & element.nodeName & ">"
	Else
		NodeText = name
	End If

End Function


' What a node shows. A device is known on the one-line diagram by its
' Id, not by its name, so the Id leads and the name follows - the Id is
' what the operator reads to know which device the panel is editing.
' Padded to three digits, so the Ids read straight down the column. A
' folder, and an object whose class declares no Id, shows its name
' alone.
'
' Only what is shown carries the Id: the node is still addressed by
' NodeKey and still sorts on SortKey, and neither reads this.
Function NodeLabel(element)

	NodeLabel = NodeText(element)

	If element.nodeName <> "object" Then Exit Function

	Dim id
	id = PropertyValue(element, "Id")
	If id = "" Then Exit Function

	id = Right("000" & id, 3)

	NodeLabel = id & ": " & NodeLabel

End Function

' The Key of the image a node shows, "" when nothing in the list claims
' it.
'
' What claims a node is the image's Tag, not its Key: an object is
' claimed by the Tag naming its class - "xatm_Breaker" - and a folder by
' the Tag naming the folder - "Substation". A Tag may name more than
' one, separated by commas - "Transformer,Busbar" - so the one picture
' that suits both bays is inserted once and not once per folder.
'
' Holding the mapping in the Tag is what lets an engineer add an icon,
' or point one at another class, from the ImageList's property page,
' with no line here changing.
Function NodeImage(element)

	NodeImage = ""

	If gTreeImages Is Nothing Then Exit Function

	Dim claim

	If element.nodeName = "object" Then
		claim = element.getAttribute("type")
	Else
		claim = element.getAttribute("name")
	End If

	If IsNull(claim) Then Exit Function

	claim = Trim(CStr(claim))
	If claim = "" Then Exit Function

	Dim img

	For Each img In gTreeImages.ListImages

		If Claims(img, claim) Then
			NodeImage = img.Key
			Exit Function
		End If

	Next

End Function


' Whether an image's Tag names this claim. The Tag is read as a list -
' a list of one being the ordinary case - so an empty Tag claims nothing
' and a picture that suits two kinds of node need not be inserted twice.
Function Claims(img, claim)

	Claims = False

	Dim imgTag
	imgTag = ""

	On Error Resume Next
	imgTag = img.Tag
	On Error Goto 0

	If IsNull(imgTag) Or IsEmpty(imgTag) Then Exit Function

	Dim named, i
	named = Split(CStr(imgTag), ",")

	For i = 0 To UBound(named)

		If StrComp(Trim(named(i)), claim, vbTextCompare) = 0 Then
			Claims = True
			Exit Function
		End If

	Next

End Function


' Gives a node its icon, where something in the ImageList claims it. A
' node nothing claims keeps the blank the tree gives it.
Sub SetNodeImage(node, element)

	Dim imageKey
	imageKey = NodeImage(element)
	If imageKey = "" Then Exit Sub

	' A tree that never took the list refuses the Key, and one node short
	' of an icon is not worth stopping the build for.
	On Error Resume Next
	node.Image = imageKey
	On Error Goto 0

End Sub

' The tree control on the screen, Nothing when the screen has none. E3
' forwards the members of an ActiveX, but not in every version, so
' whichever of the two answers Nodes is the control.
Function TreeControl(controlName)

	Set TreeControl = Nothing

	Dim ctl
	Set ctl = Nothing
	On Error Resume Next
	Set ctl = Screen.Item(controlName)
	On Error Goto 0

	If ctl Is Nothing Then Exit Function

	Dim nodes
	Set nodes = Nothing
	On Error Resume Next
	Set nodes = ctl.Nodes
	On Error Goto 0

	If Not nodes Is Nothing Then
		Set TreeControl = ctl
		Exit Function
	End If

	On Error Resume Next
	Set TreeControl = ctl.Object
	On Error Goto 0

End Function

' Hands the tree its icons and answers the list it took - Nothing when
' the screen carries no ImageList, or when the tree would not take it,
' and then no node asks for an icon and the tree draws as it always has.
'
' Called from PopulateTree, on an empty tree, and nowhere else.
Function BoundImages(tree)

	Set BoundImages = Nothing

	Dim images
	Set images = ImageListControl(IMAGES_CONTROL)
	If images Is Nothing Then Exit Function

	Dim taken
	taken = False

	On Error Resume Next
	Err.Clear
	Set tree.ImageList = images
	taken = (Err.Number = 0)
	On Error Goto 0

	If taken Then Set BoundImages = images

End Function


' The ImageList control on the screen, Nothing when the screen has none.
' Answered the two ways TreeControl is answered, for the same reason: E3
' forwards the members of an ActiveX in some versions and hands out the
' control itself in others.
Function ImageListControl(controlName)

	Set ImageListControl = Nothing

	Dim ctl
	Set ctl = Nothing
	On Error Resume Next
	Set ctl = Screen.Item(controlName)
	On Error Goto 0

	If ctl Is Nothing Then Exit Function

	Dim images
	Set images = Nothing
	On Error Resume Next
	Set images = ctl.ListImages
	On Error Goto 0

	If Not images Is Nothing Then
		Set ImageListControl = ctl
		Exit Function
	End If

	On Error Resume Next
	Set ImageListControl = ctl.Object
	On Error Goto 0

End Function


' MSXML 6 where the machine has it, the version-independent progid
' otherwise.
Function NewDomDocument()

	Set NewDomDocument = Nothing

	On Error Resume Next
	Set NewDomDocument = CreateObject("MSXML2.DOMDocument.6.0")
	On Error Goto 0

	If NewDomDocument Is Nothing Then
		Set NewDomDocument = CreateObject("MSXML2.DOMDocument")
	End If

	NewDomDocument.async = False
	NewDomDocument.preserveWhiteSpace = True

	On Error Resume Next
	NewDomDocument.setProperty "SelectionLanguage", "XPath"
	On Error Goto 0

End Function


' A byte-order mark survives a round trip through a utf-8 file and
' makes loadXML fail on what looks like perfectly good XML.
Function StripBom(text)

	StripBom = text

	If Len(text) = 0 Then Exit Function

	If AscW(Left(text, 1)) = &HFEFF Then
		StripBom = Mid(text, 2)
	End If

End Function


' A path E3 will resolve. A name that starts with a digit or an
' underscore, or carries a space or a hyphen, has to be bracketed -
' Substation.Transformer.52-01 does not resolve and
' [Substation].[Transformer].[52-01] does. Bracketing every piece is
' always right, so no piece has to be judged on its own.
'
' A piece already bracketed is left as it is: the document takes its
' object paths from PathName, which may have bracketed them already, and
' [[52-01]] resolves no better than 52-01 does.
Function E3Path(path)

	E3Path = ""
	If path = "" Then Exit Function

	Dim pieces
	pieces = SplitPath(path)

	Dim i
	For i = 0 To UBound(pieces)
		If Left(pieces(i), 1) <> "[" Then
			pieces(i) = "[" & pieces(i) & "]"
		End If
	Next

	E3Path = Join(pieces, ".")

End Function


' The pieces of a path, split on the dots between them - the ones outside
' brackets, since a bracketed name is allowed to carry a dot of its own.
Function SplitPath(path)

	Dim pieces()
	ReDim pieces(0)
	pieces(0) = ""

	Dim i, n, depth, ch
	n = 0
	depth = 0

	For i = 1 To Len(path)

		ch = Mid(path, i, 1)

		If ch = "[" Then
			depth = depth + 1
			pieces(n) = pieces(n) & ch
		ElseIf ch = "]" Then
			depth = depth - 1
			pieces(n) = pieces(n) & ch
		ElseIf ch = "." And depth = 0 Then
			n = n + 1
			ReDim Preserve pieces(n)
			pieces(n) = ""
		Else
			pieces(n) = pieces(n) & ch
		End If

	Next

	SplitPath = pieces

End Function

Sub Foo()
			
End Sub

<xatm_config_screens.Config.btnBindPositions:btnBindPositions_Click()>
Sub btnBindPositions_Click()

	' Asked before it runs, because it does not ask per device.
	'
	' Every breaker and disconnector that finds a Power device of its name
	' has its position, command and Defective properties overwritten. A tag
	' or an expression somebody put in by hand goes with them, which is the
	' whole point on a first pass and a loss on a project already tuned.
	'
	' So the box warns rather than asks - vbExclamation and not vbQuestion,
	' and it says out loud that there is no undo on this screen, because
	' there is not one: what this overwrites comes back only by being
	' configured again by hand, one property at a time.
	'
	' In Portuguese, alone among the dialogs here, because this is the one
	' that can cost somebody an afternoon and it should be read rather than
	' translated at a glance.

	Dim substation
	substation = ChosenSubstation()

	If substation = "" Then Exit Sub

	If MsgBox( _
		"ATENÇÃO - esta rotina sobrescreve o que já está " & _
		"configurado." & vbCrLf & vbCrLf & _
		"Vincular os sinais de todos os disjuntores e seccionadoras a " & _
		"partir da PowerSubstation?" & vbCrLf & vbCrLf & _
		"PositionOpen, PositionClosed, CommandOpen, CommandClose e " & _
		"Defective serão substituídos em todo dispositivo que tiver um " & _
		"equivalente de mesmo nome na PowerSubstation - inclusive onde a " & _
		"tag ou a expressão foi escolhida à mão." & vbCrLf & vbCrLf & _
		"Não há desfazer nesta tela: o que for sobrescrito só volta " & _
		"reconfigurando à mão, um por um." & vbCrLf & vbCrLf & _
		"Nada chega ao projeto até você clicar em Salvar.", _
		vbYesNo + vbExclamation + vbDefaultButton2, BIND_TITLE) <> vbYes Then Exit Sub

	Dim binder
	Set binder = Nothing

	On Error Resume Next
	Set binder = Application.GetObject(BIND_POSITIONS)
	On Error Goto 0

	If binder Is Nothing Then
		MsgBox "Este projeto não tem a tag " & BIND_POSITIONS & ", " & _
		       "então não há como vincular nada.", vbCritical, BIND_TITLE
		Exit Sub
	End If

	binder.WriteEx substation

	If binder.DocString <> EXIT_SUCCESS Then
		MsgBox "Nem tudo foi vinculado - o console diz quais dispositivos " & _
		       "e por quê.", vbExclamation, BIND_TITLE
		Exit Sub
	End If

	MsgBox "Os sinais de posição, comando e Defective foram " & _
	       "preenchidos." & vbCrLf & vbCrLf & _
	       "Revise no painel e clique em Salvar para gravar no projeto.", _
	       vbInformation, BIND_TITLE

End Sub

Const BIND_POSITIONS = "xatm_config_data.Config.BindPositions"
Const EXIT_SUCCESS   = "EXIT_SUCCESS"
Const BIND_TITLE     = "Vincular sinais"

' The tag that does the looking, and the one it answers on. Written out
' again here because one E3 object cannot call another's - the alarms
' button keeps the same pair for the same reason.
Const FIND_OBJECTS     = "xatm_config_data.Config.FindObjects"
Const FOUND_OBJECTS    = "xatm_config_data.Config.FoundObjects"
Const SUBSTATION_CLASS = "PowerSubstation"


' The substation to take the tags from, or "" when there is none or the
' operator backed out of choosing.
Function ChosenSubstation()

	ChosenSubstation = ""

	Dim finder
	Set finder = Nothing
	On Error Resume Next
	Set finder = Application.GetObject(FIND_OBJECTS)
	On Error Goto 0

	If finder Is Nothing Then
		MsgBox "Este projeto não tem a tag " & FIND_OBJECTS & ", " & _
		       "então não há como procurar as subestações.", _
		       vbCritical, BIND_TITLE
		Exit Function
	End If

	finder.WriteEx SUBSTATION_CLASS

	If finder.DocString <> EXIT_SUCCESS Then
		MsgBox "Não foi possível procurar as subestações - o " & _
		       "console diz por quê.", vbCritical, BIND_TITLE
		Exit Function
	End If

	Dim answer
	answer = ""
	On Error Resume Next
	answer = CStr(Application.GetObject(FOUND_OBJECTS).Value)
	On Error Goto 0

	If Trim(answer) = "" Then
		MsgBox "Nenhuma " & SUBSTATION_CLASS & " foi encontrada neste " & _
		       "domínio.", vbExclamation, BIND_TITLE
		Exit Function
	End If

	Dim found
	found = Split(answer, vbCrLf)

	If UBound(found) = 0 Then
		ChosenSubstation = found(0)
		Exit Function
	End If

	ChosenSubstation = PickSubstation(found)

End Function


' Which of several, asked as a number against a numbered list - the same
' box the alarms button puts up, because a viewer has nothing better.
Function PickSubstation(found)

	PickSubstation = ""

	Dim prompt, i
	prompt = "Foi encontrada mais de uma " & SUBSTATION_CLASS & "." & _
	         vbCrLf & "Digite o número da que deve fornecer os sinais:" & _
	         vbCrLf & vbCrLf

	For i = 0 To UBound(found)
		prompt = prompt & "  " & (i + 1) & " - " & found(i) & vbCrLf
	Next

	Dim reply
	reply = InputBox(prompt, BIND_TITLE, "1")

	' Cancel and an empty box both come back "", and neither is a choice.
	If Trim(reply) = "" Then Exit Function

	If Not IsNumeric(reply) Then
		MsgBox "'" & reply & "' não é um dos números da lista.", _
		       vbExclamation, BIND_TITLE
		Exit Function
	End If

	Dim choice
	choice = CLng(reply)

	If choice < 1 Or choice > UBound(found) + 1 Then
		MsgBox "Não existe " & choice & " na lista.", vbExclamation, BIND_TITLE
		Exit Function
	End If

	PickSubstation = found(choice - 1)

End Function

Sub EndOfScope()
End Sub


<xatm_config_screens.Config.btnAlarms:btnAlarms_Click()>
Sub btnAlarms_Click()
	
	If Not SupersededByExport(ALARMS_TITLE, "alarms") Then Exit Sub
	
	Dim substation
	substation = ChosenSubstation()

	If substation = "" Then Exit Sub

	' The build runs on the server, for the reason the search did: a
	' screen cannot make objects in a data server. The substation it
	' chose travels as the payload.
	Dim builder
	Set builder = Nothing
	On Error Resume Next
	Set builder = Application.GetObject(BUILD_ALARMS)
	On Error Goto 0

	If builder Is Nothing Then
		MsgBox "This project has no " & BUILD_ALARMS & " tag, so the alarms " & _
		       "cannot be built.", vbCritical, ALARMS_TITLE
		Exit Sub
	End If

	builder.WriteEx substation

	If builder.DocString <> EXIT_SUCCESS Then
		MsgBox "The alarms were not built, or were built only in part - the " & _
		       "console says which.", vbExclamation, ALARMS_TITLE
		Exit Sub
	End If

	MsgBox "The alarms were built on:" & vbCrLf & vbCrLf & substation, _
	       vbInformation, ALARMS_TITLE

End Sub


' The tag that does the looking, and the one it answers on.
'
' A screen is a viewer resource: it holds no server objects and cannot
' walk them, so it cannot find a PowerSubstation itself. It writes to
' the command tag instead and reads the answer back off the other - the
' way the Save button gets its work done on the server.
Const FIND_OBJECTS     = "xatm_config_data.Config.FindObjects"
Const FOUND_OBJECTS    = "xatm_config_data.Config.FoundObjects"
Const SUBSTATION_CLASS = "PowerSubstation"
Const BUILD_ALARMS     = "xatm_config_data.Config.BuildAlarms"

Const EXIT_SUCCESS     = "EXIT_SUCCESS"
Const ALARMS_TITLE     = "Alarms"


' Both of these built what the signal spreadsheet now builds, and
' both are kept because a bench project with no SCADA developer on
' it still wants its alarms. Asked rather than refused, and
' defaulted to No: running one of them on a project that has been
' exported leaves two sets of the same objects, one of them nobody
' will maintain.
'
' Named the same in both scopes and written out twice, because a
' screen scope cannot call into another one.
Function SupersededByExport(title, what)

	SupersededByExport = (MsgBox( _
		"The " & what & " are now built from the signal spreadsheet, not here." & _
		vbCrLf & vbCrLf & _
		"Use Export signals to CSV and hand the file to the SCADA developer - " & _
		"their import creates the alarms, the user fields and the addresses." & _
		vbCrLf & vbCrLf & _
		"Build them here anyway?", _
		vbYesNo + vbExclamation + vbDefaultButton2, title) = vbYes)

End Function


' The substation to build the alarms on, or "" when there is none to
' build on or the operator backed out of choosing.
Function ChosenSubstation()

	ChosenSubstation = ""

	Dim finder
	Set finder = Nothing
	On Error Resume Next
	Set finder = Application.GetObject(FIND_OBJECTS)
	On Error Goto 0

	If finder Is Nothing Then
		MsgBox "This project has no " & FIND_OBJECTS & " tag, so the " & _
		       "substations cannot be looked for.", vbCritical, ALARMS_TITLE
		Exit Function
	End If

	finder.WriteEx SUBSTATION_CLASS

	If finder.DocString <> EXIT_SUCCESS Then
		MsgBox "The substations could not be looked for - the console says why.", _
		       vbCritical, ALARMS_TITLE
		Exit Function
	End If

	Dim answer
	answer = ""
	On Error Resume Next
	answer = CStr(Application.GetObject(FOUND_OBJECTS).Value)
	On Error Goto 0

	If Trim(answer) = "" Then
		MsgBox "No PowerSubstation was found in this domain.", _
		       vbExclamation, ALARMS_TITLE
		Exit Function
	End If

	Dim found
	found = Split(answer, vbCrLf)

	If UBound(found) = 0 Then
		ChosenSubstation = found(0)
		Exit Function
	End If

	ChosenSubstation = PickSubstation(found)

End Function


' Which of several, asked as a number against a numbered list.
'
' An InputBox is what a viewer has to ask with, so the list goes into
' the prompt and the answer comes back as free text - checked here,
' because anything at all can be typed into one.
Function PickSubstation(found)

	PickSubstation = ""

	Dim prompt, i
	prompt = "More than one PowerSubstation was found." & vbCrLf & _
	         "Type the number of the one to build the alarms on:" & vbCrLf & vbCrLf

	For i = 0 To UBound(found)
		prompt = prompt & "  " & (i + 1) & " - " & found(i) & vbCrLf
	Next

	Dim reply
	reply = InputBox(prompt, ALARMS_TITLE, "1")

	' Cancel and an empty box both come back "", and neither is a choice.
	If Trim(reply) = "" Then Exit Function

	If Not IsNumeric(reply) Then
		MsgBox "'" & reply & "' is not one of the numbers on the list.", _
		       vbExclamation, ALARMS_TITLE
		Exit Function
	End If

	Dim choice
	choice = CLng(reply)

	If choice < 1 Or choice > UBound(found) + 1 Then
		MsgBox "There is no " & choice & " on the list.", _
		       vbExclamation, ALARMS_TITLE
		Exit Function
	End If

	PickSubstation = found(choice - 1)

End Function

' E3 will not take a Function as the last thing in a scope.
Sub EndOfScope()
		
End Sub

<xatm_config_screens.Config.btnApply:btnApply_Click()>
Sub btnApply_Click()

	Dim transformerType, busbarType, incomerType
	transformerType = SelectedLayout("SelectLayoutTransformer")
	busbarType      = SelectedLayout("SelectLayoutBusbar")
	incomerType      = SelectedLayout("SelectLayoutIncomer")

	' Asked apart from the pairing below. No TM or TA step operates income
	' equipment, so the incomer layout cannot make a combination
	' unsupported - but a name this screen does not know would quietly
	' build nothing at all, which is worth saying out loud.
	If Not IsIncomerLayout(incomerType) Then
		MsgBox "'" & incomerType & "' is not an incomer layout this screen can " & _
		       "apply. Nothing was changed.", vbExclamation, "Apply"
		Exit Sub
	End If
	If Not IsSupportedCombination(transformerType, busbarType) Then

		If MsgBox(UnsupportedCombinationText(transformerType, busbarType) & vbCrLf & vbCrLf & _
		          "Apply it anyway?", vbExclamation + vbYesNo + vbDefaultButton2, "Apply") <> vbYes Then
			Exit Sub
		End If

	End If
	
	Dim contentTag
	Set contentTag = Nothing
	On Error Resume Next
	Set contentTag = Application.GetObject(CONFIG_DATA).Item("Config").Item("XMLContent")
	On Error Goto 0

	If contentTag Is Nothing Then
		MsgBox "No 'XMLContent' tag was found under '" & CONFIG_DATA & ".XML'.", _
		       vbExclamation, "Apply"
		Exit Sub
	End If

	If IsEmpty(contentTag.Value) Or CStr(contentTag.Value) = "" Then
		MsgBox "'XMLContent' is empty - run the XML export before applying a layout.", _
		       vbExclamation, "Apply"
		Exit Sub
	End If

	Dim doc
	Set doc = NewDomDocument()

	If Not doc.loadXML(StripBom(CStr(contentTag.Value))) Then
		MsgBox "'XMLContent' does not parse as XML:" & vbCrLf & vbCrLf & _
		       doc.parseError.reason, vbCritical, "Apply"
		Exit Sub
	End If

	gNewLine = DocumentNewLine(doc)

	Dim removed, added, failed
	removed = ""
	added   = ""
	failed  = ""

	' The Transformer folder is fed by both axes: what the transformer
	' layout declares, plus the high-voltage disconnectors that exist
	' only because there is an incomerto connect to.
	SyncFolder doc, TRANSFORMER_PATH, _
	           BothLists(TransformerDevices(transformerType), _
	                     TransformerHVDevices(transformerType, incomerType)), _
	           removed, added, failed

	SyncFolder doc, BUSBAR_PATH, BusbarDevices(busbarType), removed, added, failed

	' Skipped rather than reported when there is nothing to put there and
	' no folder to put it in. A substation that never had an incomerhas no
	' Incomerfolder either, and saying so at every Apply would be noise.
	If Not (IncomerWanted(incomerType) Or HasFolder(doc, INCOMER_PATH)) Then
		' nothing to do
	Else
		SyncFolder doc, INCOMER_PATH, IncomerDevices(incomerType), removed, added, failed
	End If
	
	' TODO: change function names to SyncAutomation
	SyncAutomation doc, AutomationCount(transformerType), busbarType, removed, added, failed

	' One automatic transfer per transformer, alongside the manual ones.
	' The count is the same because the question is the same - how many
	' transformers the layout has - and it is asked again rather than
	' shared, so that the two families stay free to differ.
	SyncTA doc, AutomationCount(transformerType), removed, added, failed
	
	' After the incomer folder above, so the breakers a reclosing will
	' operate are already in the document by the time it appears - and
	' after SyncAutomation, which shares the folder and prunes only its
	' own class.
	SyncRASEAT doc, RASEATCount(incomerType), removed, added, failed
	
	SetLayoutTag doc, "Transformer", transformerType
	SetLayoutTag doc, "Busbar", busbarType
	SetLayoutTag doc, "Incomer", incomerType

	contentTag.WriteEx DocumentText(doc)
	
	Application.GetObject("xatm_config_data.Config.UpdateTreeviewSignal").WriteEx True
	
	MsgBox ApplyReport(transformerType, busbarType, incomerType, removed, added, failed), vbInformation, "Apply"

End Sub

Const CONFIG_DATA      = "xatm_config_data"
Const HELPER_FOLDER    = "PropertiesHelper"
Const AUTOMATION_PATH  = "/xatm-config/folder[@name='Automation']"
Const LAYOUT_PATH      = "/xatm-config/folder[@name='Automation']/folder[@name='Layout']"
Const TRANSFORMER_PATH = "/xatm-config/folder[@name='Substation']/folder[@name='Transformer']"
Const BUSBAR_PATH      = "/xatm-config/folder[@name='Substation']/folder[@name='Busbar']"
Const INCOMER_PATH     = "/xatm-config/folder[@name='Substation']/folder[@name='Incomer']"

Const TMTNM_CLASS       = "xatm_TMTNM"
Const TA_CLASS          = "xatm_TA"
Const RASEAT_CLASS      = "xatm_RASEAT"
Const TRANSFORMER_CLASS = "xatm_Transformer"

Const NODE_ELEMENT = 1
Const NODE_TEXT    = 3


' Line break of the document being applied, so a node added to it breaks
' its lines the same way the rest of the document already does.
Dim gNewLine


' ------------------------------------------------------------
'  LAYOUTS
' ------------------------------------------------------------

' Every device the transformer layout expects to find in the Transformer
' folder, as "class:Id" - the transformers themselves and their
' secondary breakers.
Function TransformerDevices(layoutType)

	Select Case UCase(layoutType)

		Case "4TR4LV"
			TransformerDevices = Array("xatm_Transformer:100", "xatm_Transformer:200", _
			                           "xatm_Transformer:300", "xatm_Transformer:400", _
			                           "xatm_Breaker:120", "xatm_Breaker:220", _
			                           "xatm_Breaker:320", "xatm_Breaker:420")

		Case "2TR2LV"
			TransformerDevices = Array("xatm_Transformer:100", "xatm_Transformer:200", _
			                           "xatm_Breaker:120", "xatm_Breaker:220")

		Case Else
			TransformerDevices = Array()

	End Select

End Function


' Every device the incomerlayout expects to find in the Incomerfolder,
' as "class:Id" - the two line bays, each a breaker with its line
' disconnector and one disconnector per incomerbusbar.
'
' The tens digit is the bay and the units digit the role. Nothing is
' numbered 0: an unconfigured Id reads as Empty, and VBScript makes
' Empty equal to 0 - so a device with Id 0 collides with every device
' nobody has configured yet, both in FindInFolder's comparison and as a
' Dictionary key in BuildDeviceDictionary.
Function IncomerDevices(layoutType)

	Select Case UCase(layoutType)

		Case "2BR2BB"
			IncomerDevices = Array("xatm_Breaker:10", "xatm_Disconnector:12", _
			                      "xatm_Disconnector:14", "xatm_Disconnector:15", _
			                      "xatm_Breaker:20", "xatm_Disconnector:22", _
			                      "xatm_Disconnector:24", "xatm_Disconnector:25")

		' No incomersupervised by this automation - and with it, no
		' high-voltage disconnectors on the transformers either.
		Case "NONE"
			IncomerDevices = Array()

		Case Else
			IncomerDevices = Array()

	End Select

End Function


' The transformers' high-voltage disconnectors, which belong to neither
' layout on its own.
'
' They sit in the Transformer folder, but how many there are is
' transformers x incomerbusbars - so the transformer layout cannot name
' them and neither can the incomerlayout. They are the devices that
' exist only where the two axes cross, which is why 4TR4LV does not
' mention them in its name.
Function TransformerHVDevices(transformerType, incomerType)

	TransformerHVDevices = Array()

	If UCase(incomerType) <> "2BR2BB" Then Exit Function

	Dim count
	count = TransformerCount(transformerType)
	If count < 1 Then Exit Function

	Dim list()
	ReDim list(count * 2 - 1)

	Dim i, base
	For i = 1 To count

		base = i * 100
		list((i - 1) * 2)     = "xatm_Disconnector:" & (base + HV_TO_BAR_A)
		list((i - 1) * 2 + 1) = "xatm_Disconnector:" & (base + HV_TO_BAR_B)

	Next

	TransformerHVDevices = list

End Function

' The two offsets a transformer's high-voltage disconnectors are
' numbered at, one per incomerbusbar - the lower offset to bar A, the
' higher to bar B.
'
' The incomerbays pair the same way round: +30 and +4 land on bar A,
' +80 and +5 on bar B. Bar B is the upper of the two on the drawing,
' which is the opposite of what the label order suggests.
Const HV_TO_BAR_A = 30
Const HV_TO_BAR_B = 80


' How many transformers a layout has. The same count as the automations
' - one manual automation per transformer - but asked as its own question, because the
' high-voltage disconnectors are per transformer and not per automation.
Function TransformerCount(layoutType)

	Select Case UCase(layoutType)
		Case "4TR4LV" : TransformerCount = 4
		Case "2TR2LV" : TransformerCount = 2
		Case Else     : TransformerCount = 0
	End Select

End Function


' Two device lists as one, for a folder fed by more than one layout.
Function BothLists(first, second)

	If UBound(first) < 0 Then
		BothLists = second
		Exit Function
	End If

	If UBound(second) < 0 Then
		BothLists = first
		Exit Function
	End If

	Dim list()
	ReDim list(UBound(first) + UBound(second) + 1)

	Dim i
	For i = 0 To UBound(first)
		list(i) = first(i)
	Next
	For i = 0 To UBound(second)
		list(UBound(first) + 1 + i) = second(i)
	Next

	BothLists = list

End Function

' Every device the busbar layout expects to find in the Busbar folder,
' as "class:Id" - the bus-tie and bus-section breakers.
Function BusbarDevices(layoutType)

	Select Case UCase(layoutType)

		Case "6BB6TIERING"
			BusbarDevices = Array("xatm_Breaker:700", "xatm_Breaker:710", _
			                      "xatm_Breaker:720", "xatm_Breaker:730", _
			                      "xatm_Breaker:740", "xatm_Breaker:900")

		Case "2BB1TIE"
			BusbarDevices = Array("xatm_Breaker:700")

		Case Else
			BusbarDevices = Array()

	End Select

End Function


' The two letters a device name starts with. A device the layout adds is
' named prefix & Id, which is unique because the Id is.
Function TypePrefix(className)

	Select Case LCase(className)

		Case "xatm_breaker"      : TypePrefix = "DJ"
		Case "xatm_transformer"  : TypePrefix = "TR"
		Case "xatm_disconnector" : TypePrefix = "CS"    ' Chave seccionadora

		' A class with no tag of its own falls back to the first two
		' letters of its name - XATM_Feeder is FE.
		Case Else                : TypePrefix = UCase(Left(Replace(LCase(className), "xatm_", ""), 2))

	End Select

End Function


' Transformer n is Id n00, the way the layouts hand the Ids out, so a
' automation the layout adds knows which transformer it drives.
Function AutomationTransformerId(n)

	AutomationTransformerId = n * 100

End Function


' One manual automation per transformer, so the transformer layout sets
' the count.
' How many high-voltage reclosing automations the incomer layout wants.
'
' One for the pair, not one per incomer: the station runs with a single
' incomer closed, and the sequence works out which that was rather than
' a second instance racing it for the answer.
'
' Asked of the incomer layout and not the transformer one, because it is
' the incomer axis that decides whether there are entry bays at all - and
' with NONE there is nothing to reclose.
Function RASEATCount(incomerType)

	Select Case UCase(incomerType)
		Case "2BR2BB" : RASEATCount = 1
		Case Else     : RASEATCount = 0
	End Select

End Function

Function AutomationCount(layoutType)

	Select Case UCase(layoutType)
		Case "4TR4LV" : AutomationCount = 4
		Case "2TR2LV" : AutomationCount = 2
		Case Else     : AutomationCount = 0
	End Select

End Function


Function IsTransformerLayout(layoutType)

	Select Case UCase(layoutType)
		Case "4TR4LV", "2TR2LV" : IsTransformerLayout = True
		Case Else               : IsTransformerLayout = False
	End Select

End Function


Function IsBusbarLayout(layoutType)

	Select Case UCase(layoutType)
		Case "6BB6TIERING", "2BB1TIE" : IsBusbarLayout = True
		Case Else                     : IsBusbarLayout = False
	End Select

End Function


' NONE is a layout like any other here - it is the answer for a
' substation whose incomerthis automation does not supervise, and is
' not the same thing as a name nobody recognises.
Function IsIncomerLayout(layoutType)

	Select Case UCase(layoutType)
		Case "2BR2BB", "NONE" : IsIncomerLayout = True
		Case Else             : IsIncomerLayout = False
	End Select

End Function


' The two pairings the automation logic in xatm_lib has step sequences
' for. A device tree can be built for any other pairing, but no
' automation will run on it.
Function IsSupportedCombination(transformerType, busbarType)

	Select Case UCase(transformerType) & "_" & UCase(busbarType)
		Case "4TR4LV_6BB6TIERING", "2TR2LV_2BB1TIE" : IsSupportedCombination = True
		Case Else                                   : IsSupportedCombination = False
	End Select

End Function


' What is wrong with a combination, in the operator's terms.
Function UnsupportedCombinationText(transformerType, busbarType)

	Dim text

	If Not IsTransformerLayout(transformerType) Then
		text = "'" & transformerType & "' is not a transformer layout this screen can apply."
	ElseIf Not IsBusbarLayout(busbarType) Then
		text = "'" & busbarType & "' is not a busbar layout this screen can apply."
	Else
		text = transformerType & " + " & busbarType & " is not a combination the automation logic implements."
	End If

	UnsupportedCombinationText = text & vbCrLf & vbCrLf & _
	                             "Only 4TR4LV + 6BB6TIERING and 2TR2LV + 2BB1TIE have step sequences. " & _
	                             "Any other pairing builds the device tree, but every automation on it " & _
	                             "refuses to run."

End Function


' The layout a selector is showing. The names live in the control's
' ImageList, so Index resolves through it - whether Index carries the
' position in the list or the name itself, which is what OnPreShow
' writes into it.
Function SelectedLayout(controlName)

	SelectedLayout = ""

	Dim ctl
	Set ctl = Nothing
	On Error Resume Next
	Set ctl = Screen.Item(controlName)
	On Error Goto 0

	If ctl Is Nothing Then Exit Function

	Dim names
	names = Split(ctl.ImageList, ",")

	Dim idx
	idx = ctl.Index

	If IsNumeric(idx) Then
		If CLng(idx) >= 0 And CLng(idx) <= UBound(names) Then
			SelectedLayout = Trim(names(CLng(idx)))
			Exit Function
		End If
	End If

	Dim i
	For i = 0 To UBound(names)
		If Trim(names(i)) = Trim(CStr(idx)) Then
			SelectedLayout = Trim(names(i))
			Exit Function
		End If
	Next

End Function


' ------------------------------------------------------------
'  DOCUMENT
' ------------------------------------------------------------

' Brings a device folder to the layout. An object whose Id the layout
' does not name is dropped, an Id it names that the folder does not have
' is created, and an object that stays is left alone - it keeps its
' name, its Id and everything configured on it.
Sub SyncFolder(doc, folderPath, devices, removed, added, failed)

	Dim folder
	Set folder = doc.selectSingleNode(folderPath)

	If folder Is Nothing Then

		' The document is an export of the project, so a folder missing
		' here is a folder missing in E3 - and the import creates objects
		' but never folders. Saying which one and where it belongs turns
		' this from a puzzle into a one-line job.
		failed = failed & vbCrLf & "  the " & FolderLabel(folderPath) & _
		         " folder is not in the document - create it in the project " & _
		         "(XATM_Data.Substation." & FolderLabel(folderPath) & ") and export again"
		Exit Sub

	End If

	' --- Drop what the layout does not name ---------------------------

	Dim wanted
	Set wanted = CreateObject("Scripting.Dictionary")

	Dim i, parts
	For i = 0 To UBound(devices)
		parts = Split(devices(i), ":")
		wanted(Trim(parts(1))) = Trim(parts(0))
	Next

	Dim present
	Set present = CreateObject("Scripting.Dictionary")

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim nodes, node, id, n
	Set nodes = folder.selectNodes("object")

	For n = 0 To nodes.length - 1

		Set node = nodes.item(n)
		id = PropertyValue(node, "Id")

		If wanted.Exists(id) Then
			present(id) = True
		Else
			removed = removed & vbCrLf & "  " & node.getAttribute("name") & _
			          " (Id " & id & ") from " & FolderLabel(folderPath)
			doomed.Add n, node
		End If

	Next

	' Removed in a second pass - the node list is not walked while the
	' tree under it is being changed.
	For Each n In doomed.Keys
		Set node = doomed(n)
		DropNode node
	Next

	' --- Create what the layout names and the folder does not have ----

	Dim className, name
	For i = 0 To UBound(devices)

		parts     = Split(devices(i), ":")
		className = Trim(parts(0))
		id        = Trim(parts(1))

		If Not present.Exists(id) Then

			name = TypePrefix(className) & id

			Set node = NewObject(folder, className, name, Array("Id", id))

			If node Is Nothing Then
				failed = failed & vbCrLf & "  " & name & " (Id " & id & ") in " & _
				         FolderLabel(folderPath) & " - no manifest for " & className
			Else
				added = added & vbCrLf & "  " & name & " (Id " & id & ") in " & _
				        FolderLabel(folderPath)
			End If

		End If

	Next

End Sub


' Brings the Automation folder to the layout. These instances carry no
' Id property, so they go by the number in the name - TMTNM1..TMTNMn stay,
' the rest are dropped, and a missing one is created bound to the
' transformer of the same number.
Sub SyncAutomation(doc, keepCount, busbarType, removed, added, failed)

	Dim folder
	Set folder = doc.selectSingleNode(AUTOMATION_PATH)

	If folder Is Nothing Then
		failed = failed & vbCrLf & "  the Automation folder is not in the document"
		Exit Sub
	End If

	' Two families of manual automation share this folder. The transformer ones are
	' TMTNM1..TMTNMn, one per transformer, and are known by the number on the
	' end of the name; the busbar ones are TMTNM_B1A and TMTNM_B2B, exist only
	' where the busbar layout has somewhere to hand a busbar to, and are
	' known by name because there is nothing about them to count.
	Dim busbars
	busbars = BusbarAutomations(busbarType)

	Dim kept, keptBusbar
	Set kept       = CreateObject("Scripting.Dictionary")
	Set keptBusbar = CreateObject("Scripting.Dictionary")

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim nodes, node, n, num, objectName
	Set nodes = folder.selectNodes("object[@type='" & TMTNM_CLASS & "']")
	
	For n = 0 To nodes.length - 1

		Set node   = nodes.item(n)
		objectName = node.getAttribute("name")
		num        = TrailingNumber(objectName)

		' Asked in this order because a busbar automation has no trailing
		' number at all - TrailingNumber answers 0 for TMTNM_B1A - so it
		' would fall through to the prune and be dropped on the next apply.
		If BusbarWanted(busbars, objectName) Then
			keptBusbar(LCase(objectName)) = True
		ElseIf num >= 1 And num <= keepCount Then
			kept(num) = True
		Else
			removed = removed & vbCrLf & "  " & objectName & " from Automation"
			doomed.Add n, node
		End If

	Next

	For Each n In doomed.Keys
		Set node = doomed(n)
		DropNode node
	Next

	Dim name, path, overrides
	For n = 1 To keepCount

		If Not kept.Exists(n) Then

			name = "TMTNM" & n

			' The transformer folder is brought to the layout first, so a
			' transformer this automation drives is already there to bind to.
			path = TransformerPath(doc, AutomationTransformerId(n))

			If path = "" Then
				overrides = Array()
			Else
				overrides = Array("Transformer", path)
			End If

			Set node = NewObject(folder, TMTNM_CLASS, name, overrides)

			If node Is Nothing Then
				failed = failed & vbCrLf & "  " & name & " in Automation - no manifest for " & TMTNM_CLASS
			ElseIf path = "" Then
				added = added & vbCrLf & "  " & name & " in Automation, with no transformer to bind to"
			Else
				added = added & vbCrLf & "  " & name & " in Automation, bound to " & path
			End If

		End If

	Next

	' And the busbar ones, which are the same class carrying a pair.
	'
	' Nothing to bind. A busbar is not an XObject, so which pair this one
	' moves is written into a property rather than linked, and Transformer
	' is left unset - it is not any transformer's automation, and the
	' trigger inside xatm_Transformer passes it over for that reason.
	Dim pair
	For Each pair In busbars

		name = BusbarAutomationName(pair)

		If Not keptBusbar.Exists(LCase(name)) Then

			Set node = NewObject(folder, TMTNM_CLASS, name, Array("BusbarPair", pair))

			If node Is Nothing Then
				failed = failed & vbCrLf & "  " & name & " in Automation - no manifest for " & TMTNM_CLASS
			Else
				added = added & vbCrLf & "  " & name & " in Automation, transferring " & pair
			End If

		End If

	Next

End Sub


' The busbar pairs a busbar layout can transfer between.
'
' Only on the six-busbar ring, and only two of the six: B1A can be handed
' across the ring closer to B4A, and B2B to B3A, because those are the two
' whose neighbouring tie is the one normally left open. A two-busbar layout
' has nowhere to hand anything to and gets none of these automations.
Function BusbarAutomations(busbarType)

	Select Case UCase(busbarType)

		Case "6BB6TIERING"
			BusbarAutomations = Array("B1A-B4A", "B2B-B3A")

		Case Else
			BusbarAutomations = Array()

	End Select

End Function


' What one of them is called: TMTNM_B1A, TMTNM_B2B.
'
' Named for the busbar that moves rather than numbered, because there is
' nothing to count - a layout either defines the pair or defines neither -
' and because a number on the end would put it in the same series as the
' transformer automations, which the prune above walks by number.
Function BusbarAutomationName(pair)

	Dim parts
	parts = Split(pair, "-")

	BusbarAutomationName = "TMTNM_" & parts(0)

End Function


' Whether a name already in the folder is one of the busbar automations
' this layout asks for.
Function BusbarWanted(busbars, objectName)

	BusbarWanted = False

	Dim pairName
	For Each pairName In busbars

		If LCase(objectName) = LCase(BusbarAutomationName(pairName)) Then
			BusbarWanted = True
			Exit Function
		End If

	Next

End Function


' Brings the Automation folder to one automatic transfer per transformer.
'
' Shaped like the manual sync above and kept apart from it for the reason
' SyncRASEAT is kept apart: selecting by type is what lets three classes
' share one folder without any of them pruning another's objects.
'
' Bound to a transformer, and this one really is that transformer's
' automation - it answers its trip, and the trigger inside xatm_Transformer
' finds it by the binding. A manual automation bound to the same
' transformer is a different thing and is no longer looked at there.
Sub SyncTA(doc, keepCount, removed, added, failed)

	Dim folder
	Set folder = doc.selectSingleNode(AUTOMATION_PATH)

	If folder Is Nothing Then
		failed = failed & vbCrLf & "  the Automation folder is not in the document"
		Exit Sub
	End If

	Dim kept
	Set kept = CreateObject("Scripting.Dictionary")

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim nodes, node, n, num
	Set nodes = folder.selectNodes("object[@type='" & TA_CLASS & "']")

	For n = 0 To nodes.length - 1

		Set node = nodes.item(n)
		num = TrailingNumber(node.getAttribute("name"))

		If num >= 1 And num <= keepCount Then
			kept(num) = True
		Else
			removed = removed & vbCrLf & "  " & node.getAttribute("name") & " from Automation"
			doomed.Add n, node
		End If

	Next

	For Each n In doomed.Keys
		Set node = doomed(n)
		DropNode node
	Next

	Dim objectName, path, overrides
	For n = 1 To keepCount

		If Not kept.Exists(n) Then

			objectName = "TA" & n

			' The transformer folder is brought to the layout first, so the
			' transformer this one answers for is already there to bind to.
			path = TransformerPath(doc, AutomationTransformerId(n))

			If path = "" Then
				overrides = Array()
			Else
				overrides = Array("Transformer", path)
			End If

			Set node = NewObject(folder, TA_CLASS, objectName, overrides)

			If node Is Nothing Then
				failed = failed & vbCrLf & "  " & objectName & " in Automation - no manifest for " & TA_CLASS
			ElseIf path = "" Then
				added = added & vbCrLf & "  " & objectName & " in Automation, with no transformer to bind to"
			Else
				added = added & vbCrLf & "  " & objectName & " in Automation, bound to " & path
			End If

		End If

	Next

End Sub


' Brings the Automation folder to the number of reclosing automations the
' incomer layout wants - which is one, or none.
'
' The same shape as SyncAutomation above and deliberately not folded into
' it: that one binds each instance to a transformer it has to go and find,
' and this one binds nothing. RASEAT holds no equipment; its steps
' resolve breakers through GetDeviceById and the layout when they run.
'
' Selecting by type is what lets the two live in one folder without
' either pruning the other: SyncAutomation sees only xatm_TMTNM nodes, and
' this sees only xatm_RASEAT ones.
Sub SyncRASEAT(doc, keepCount, removed, added, failed)

	Dim folder
	Set folder = doc.selectSingleNode(AUTOMATION_PATH)

	If folder Is Nothing Then
		failed = failed & vbCrLf & "  the Automation folder is not in the document"
		Exit Sub
	End If

	Dim kept
	Set kept = CreateObject("Scripting.Dictionary")

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim nodes, node, n, num
	Set nodes = folder.selectNodes("object[@type='" & RASEAT_CLASS & "']")

	For n = 0 To nodes.length - 1

		Set node = nodes.item(n)
		num = TrailingNumber(node.getAttribute("name"))

		If num >= 1 And num <= keepCount Then
			kept(num) = True
		Else
			removed = removed & vbCrLf & "  " & node.getAttribute("name") & " from Automation"
			doomed.Add n, node
		End If

	Next

	For Each n In doomed.Keys
		Set node = doomed(n)
		DropNode node
	Next

	Dim name
	For n = 1 To keepCount

		If Not kept.Exists(n) Then

			name = "RASEAT" & n

			' Numbered even though there is only ever one, so the object is
			' keyed the way the manual ones are: they carry no Id, and both the import
			' and the loop above go by the number on the end of the name.
			Set node = NewObject(folder, RASEAT_CLASS, name, Array())

			If node Is Nothing Then
				failed = failed & vbCrLf & "  " & name & " in Automation - no manifest for " & RASEAT_CLASS
			Else
				added = added & vbCrLf & "  " & name & " in Automation"
			End If

		End If

	Next

End Sub

' The path of the transformer carrying an Id, "" when the document has
' no such transformer.
Function TransformerPath(doc, id)

	TransformerPath = ""

	Dim nodes, node, n
	Set nodes = doc.selectNodes(TRANSFORMER_PATH & "/object[@type='" & TRANSFORMER_CLASS & "']")

	For n = 0 To nodes.length - 1

		Set node = nodes.item(n)

		If PropertyValue(node, "Id") = CStr(id) Then
			TransformerPath = node.getAttribute("path")
			Exit Function
		End If

	Next

End Function


' ------------------------------------------------------------
'  NEW OBJECTS
' ------------------------------------------------------------

' Creates an object in a folder, with the properties its class declares
' at the initial values the manifest gives them - the same manifest the
' export reads, so a created object is shaped like an exported one.
' Nothing when the class has no manifest, which is the only way this
' fails.
'
' Overrides is a flat name, value, name, value list of the properties
' the caller sets itself, the Id among them.
Function NewObject(folder, className, objectName, overrides)

	Set NewObject = Nothing

	Dim bag
	Set bag = ManifestOf(className)
	If bag Is Nothing Then Exit Function

	Dim doc
	Set doc = folder.ownerDocument

	Dim indent
	indent = ChildIndent(folder)

	Dim node
	Set node = doc.createElement("object")

	node.setAttribute "name", objectName
	node.setAttribute "type", className
	node.setAttribute "path", ObjectPath(folder, objectName)

	Dim key, p, property, value
	For Each key In bag.Keys

		Set p = bag(key)

		Set property = doc.createElement("property")
		property.setAttribute "name", p.Name
		property.setAttribute "type", p.DataType

		value = Override(overrides, p.Name)

		' A link is an association the operator makes, never an initial
		' value, so it is left unset unless the caller names it.
		If IsEmpty(value) And Not IsLinkType(p.DataType) Then value = p.InitialValue

		If Not (IsEmpty(value) Or IsNull(value)) Then
			property.setAttribute "value", CStr(value)
		End If

		node.appendChild doc.createTextNode(gNewLine & indent & vbTab)
		node.appendChild property

	Next

	node.appendChild doc.createTextNode(gNewLine & indent)

	Place folder, node, LastOfType(folder, className), indent

	Set NewObject = node

End Function


' The value an override list gives for a property, Empty when it names
' none.
Function Override(overrides, propertyName)

	Override = Empty

	Dim i
	For i = 0 To UBound(overrides) - 1 Step 2
		If LCase(overrides(i)) = LCase(propertyName) Then
			Override = overrides(i + 1)
			Exit Function
		End If
	Next

End Function


' Puts an element in a folder after the last object of its own class, so
' the folder stays grouped by type, with the line break and indentation
' in front of it that ExportXml would have written.
Sub Place(folder, node, after, indent)

	Dim doc
	Set doc = folder.ownerDocument

	Dim anchor
	Set anchor = Nothing

	If after Is Nothing Then
		If Not folder.lastChild Is Nothing Then
			If folder.lastChild.nodeType = NODE_TEXT Then Set anchor = folder.lastChild
		End If
	Else
		If Not after.nextSibling Is Nothing Then Set anchor = after.nextSibling
	End If

	If anchor Is Nothing Then

		' Nothing to insert in front of - the folder was empty or ends on
		' an element, so the closing tag needs a line of its own too.
		folder.appendChild doc.createTextNode(gNewLine & indent)
		folder.appendChild node
		folder.appendChild doc.createTextNode(gNewLine & Left(indent, Len(indent) - 1))

	Else

		' Anchor is the whitespace in front of the next element, or in
		' front of the closing tag when there is no next element.
		folder.insertBefore doc.createTextNode(gNewLine & indent), anchor
		folder.insertBefore node, anchor

	End If

End Sub


' The last object of a class in a folder, Nothing when it holds none.
Function LastOfType(folder, className)

	Set LastOfType = Nothing

	Dim nodes
	Set nodes = folder.selectNodes("object[@type='" & className & "']")

	If nodes.length > 0 Then
		Set LastOfType = nodes.item(nodes.length - 1)
	End If

End Function


' The indentation ExportXml gives a child of this folder - one tab per
' element the child sits inside.
Function ChildIndent(folder)

	Dim depth, n
	depth = 0

	Set n = folder

	Do While Not n Is Nothing
		If n.nodeType <> NODE_ELEMENT Then Exit Do
		depth = depth + 1
		Set n = n.parentNode
	Loop

	ChildIndent = String(depth, vbTab)

End Function


' The E3 path an object in this folder would have - the root the
' document was exported from, then the folders down to the object.
Function ObjectPath(folder, objectName)

	Dim path
	path = objectName

	Dim n
	Set n = folder

	Do While n.nodeType = NODE_ELEMENT And n.nodeName = "folder"
		path = n.getAttribute("name") & "." & path
		Set n = n.parentNode
	Loop

	ObjectPath = folder.ownerDocument.documentElement.getAttribute("root") & "." & path

End Function


' The manifest declared for a class, or Nothing when the class has
' none - the same lookup the export does.
Function ManifestOf(className)

	Set ManifestOf = Nothing

	On Error Resume Next
	Set ManifestOf = Application.GetObject(CONFIG_DATA).Item(HELPER_FOLDER).Item(className).Value
	On Error Goto 0

End Function


' A property whose configuration is an association, not a value. Two of
' them: an IOTag is wired to a tag out in the project, an InternalTag to
' one of the object's own - the automation's command interface. Both export as a
' source and neither carries a value, because what is configured is the
' tag it points at, and whatever that tag happens to be holding at the
' time is not configuration at all.
Function IsLinkType(dataType)

	IsLinkType = False

	Select Case LCase(dataType)
		Case "iotag", "internaltag" : IsLinkType = True
	End Select

End Function


' The line break the document already uses. MSXML normalises CRLF to LF
' when it loads a string, and a document that mixes the two reads badly
' everywhere it is opened afterwards.
Function DocumentNewLine(doc)

	If InStr(doc.xml, vbCrLf) > 0 Then
		DocumentNewLine = vbCrLf
	Else
		DocumentNewLine = vbLf
	End If

End Function


' Records the applied selection on the Layout tags of the document.
Sub SetLayoutTag(doc, tagName, value)

	Dim node
	Set node = doc.selectSingleNode(LAYOUT_PATH & "/tag[@name='" & tagName & "']")

	If node Is Nothing Then Exit Sub

	node.setAttribute "value", value

End Sub


' Whether an incomerlayout asks for any equipment at all. NONE is a
' valid answer that happens to want nothing.
Function IncomerWanted(incomerType)

	IncomerWanted = (UBound(IncomerDevices(incomerType)) >= 0)

End Function


' Whether the document has a folder at a path.
Function HasFolder(doc, folderPath)

	HasFolder = Not (doc.selectSingleNode(folderPath) Is Nothing)

End Function

' Removes an element together with the indentation in front of it, so
' pruning leaves no blank lines behind.
Sub DropNode(node)

	Dim ws
	Set ws = node.previousSibling

	If Not ws Is Nothing Then
		If ws.nodeType = NODE_TEXT Then
			If IsWhitespace(ws.text) Then
				ws.parentNode.removeChild ws
			End If
		End If
	End If

	node.parentNode.removeChild node

End Sub


' True when a text node holds nothing but layout. Trim only takes spaces
' off, and what separates two elements is tabs and line breaks.
Function IsWhitespace(text)

	Dim s
	s = Replace(text, vbTab, "")
	s = Replace(s, vbCr, "")
	s = Replace(s, vbLf, "")

	IsWhitespace = (Trim(s) = "")

End Function


' The value attribute of a named property - "" when the property is
' absent or unset, which never matches a layout Id.
Function PropertyValue(objectNode, propertyName)

	PropertyValue = ""

	Dim p
	Set p = objectNode.selectSingleNode("property[@name='" & propertyName & "']")
	If p Is Nothing Then Exit Function

	Dim a
	Set a = p.getAttributeNode("value")
	If a Is Nothing Then Exit Function

	PropertyValue = a.value

End Function


' The document as text, headed the way ExportXml writes it.
'
' MSXML serialises a string-loaded document with a bare
' <?xml version="1.0"?> and no line break after it - it regenerates the
' declaration from its own encoding and ignores whatever the
' declaration node says, so the heading is restored here rather than in
' the tree.
Function DocumentText(doc)

	Dim text
	text = doc.xml

	If Left(text, 5) = "<?xml" Then
		text = Mid(text, InStr(text, "?>") + 2)
	End If

	Do While Left(text, 1) = vbCr Or Left(text, 1) = vbLf
		text = Mid(text, 2)
	Loop

	DocumentText = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & text

End Function


' A byte-order mark survives a round trip through a utf-8 file and
' makes loadXML fail on what looks like perfectly good XML.
Function StripBom(text)

	StripBom = text

	If Len(text) = 0 Then Exit Function

	If AscW(Left(text, 1)) = &HFEFF Then
		StripBom = Mid(text, 2)
	End If

End Function


' MSXML 6 where the machine has it, the version-independent progid
' otherwise.
Function NewDomDocument()

	Set NewDomDocument = Nothing

	On Error Resume Next
	Set NewDomDocument = CreateObject("MSXML2.DOMDocument.6.0")
	On Error Goto 0

	If NewDomDocument Is Nothing Then
		Set NewDomDocument = CreateObject("MSXML2.DOMDocument")
	End If

	NewDomDocument.async = False
	NewDomDocument.preserveWhiteSpace = True

	On Error Resume Next
	NewDomDocument.setProperty "SelectionLanguage", "XPath"
	On Error Goto 0

End Function


' ------------------------------------------------------------
'  REPORT
' ------------------------------------------------------------

Function ApplyReport(transformerType, busbarType, incomerType, removed, added, failed)

	Dim text
	text = "Applied " & transformerType & " + " & busbarType & " + " & incomerType & "."

	If removed <> "" Then
		text = text & vbCrLf & vbCrLf & "Removed from the document:" & removed
	End If

	If added <> "" Then
		text = text & vbCrLf & vbCrLf & "Added to the document:" & added
	End If

	If failed <> "" Then
		text = text & vbCrLf & vbCrLf & _
		       "Expected by the layout and could not be added:" & failed
	End If

	If removed = "" And added = "" And failed = "" Then
		text = text & vbCrLf & vbCrLf & "The document already matched the layout."
	End If

	ApplyReport = text

End Function


' Last folder name of an XPath, for the operator-facing report.
Function FolderLabel(folderPath)

	Dim parts
	parts = Split(folderPath, "@name='")

	FolderLabel = Split(parts(UBound(parts)), "'")(0)

End Function


' Trailing digits of a name - TMTNM12 is 12, a name without them is 0.
Function TrailingNumber(name)

	Dim i, digits
	digits = ""

	For i = Len(name) To 1 Step -1
		If IsNumeric(Mid(name, i, 1)) Then
			digits = Mid(name, i, 1) & digits
		Else
			Exit For
		End If
	Next

	If digits = "" Then
		TrailingNumber = 0
	Else
		TrailingNumber = CLng(digits)
	End If

End Function

' E3 will not take a Function as the last thing in a scope.
Sub EndOfScope()

End Sub

<xatm_config_screens.Config.btnDistribution:btnDistribution_Click()>
Sub btnDistribution_Click()

	If Not SupersededByExport(DISTRIBUTION_TITLE, "signals") Then Exit Sub

	Dim driver
	driver = ChosenObject(DRIVER_CLASS)

	If driver = "" Then Exit Sub

	Dim offset
	offset = ChosenOffset()

	If offset = "" Then Exit Sub

	' The build runs on the server, for the reason the search did: a
	' screen cannot make objects in a driver. The driver it chose
	' travels as the payload.
	Dim builder
	Set builder = Nothing
	On Error Resume Next
	Set builder = Application.GetObject(BUILD_DISTRIBUTION)
	On Error Goto 0

	If builder Is Nothing Then
		MsgBox "This project has no " & BUILD_DISTRIBUTION & " tag, so the " & _
		       "signals cannot be distributed.", vbCritical, DISTRIBUTION_TITLE
		Exit Sub
	End If

	builder.WriteEx offset & "|" & driver

	If builder.DocString <> EXIT_SUCCESS Then
		MsgBox "The signals were not distributed, or only in part - the " & _
		       "console says which." & vbCrLf & vbCrLf & _
		       "If the tags were built but not saved, this copy of E3 is " & _
		       "refusing to save that many without a licence.", _
		       vbExclamation, DISTRIBUTION_TITLE
		Exit Sub
	End If

	MsgBox "The signals were distributed through:" & vbCrLf & vbCrLf & driver, _
	       vbInformation, DISTRIBUTION_TITLE

End Sub


' The tag that does the looking, and the one it answers on.
'
' A screen is a viewer resource: it holds no server objects and cannot
' walk them, so it cannot find an IODriver itself. It writes the class
' it wants to the command tag and reads the answer back off the other -
' the way the Alarms button finds its substation.
Const FIND_OBJECTS       = "xatm_config_data.Config.FindObjects"
Const FOUND_OBJECTS      = "xatm_config_data.Config.FoundObjects"

' What carries the signals out to the operation centre. The level 3
' centre is spoken to through a driver in slave mode, and which driver
' that is, is the engineer's to choose.
Const DRIVER_CLASS       = "IODriver"

Const BUILD_DISTRIBUTION = "xatm_config_data.Config.BuildDistribution"

Const EXIT_SUCCESS       = "EXIT_SUCCESS"
Const DISTRIBUTION_TITLE = "Distribution"

' What the address box offers before anything is typed. The operation
' centre's point list decides the real one, and every substation
' answering to that centre needs a range of its own.
Const DEFAULT_OFFSET     = "3000"


' Both of these built what the signal spreadsheet now builds, and
' both are kept because a bench project with no SCADA developer on
' it still wants its alarms. Asked rather than refused, and
' defaulted to No: running one of them on a project that has been
' exported leaves two sets of the same objects, one of them nobody
' will maintain.
'
' Named the same in both scopes and written out twice, because a
' screen scope cannot call into another one.
Function SupersededByExport(title, what)

	SupersededByExport = (MsgBox( _
		"The " & what & " are now built from the signal spreadsheet, not here." & _
		vbCrLf & vbCrLf & _
		"Use Export signals to CSV and hand the file to the SCADA developer - " & _
		"their import creates the alarms, the user fields and the addresses." & _
		vbCrLf & vbCrLf & _
		"Build them here anyway?", _
		vbYesNo + vbExclamation + vbDefaultButton2, title) = vbYes)

End Function


' One object of a class, chosen by the operator where there is more than
' one. "" when there is none to choose from or the choosing was backed
' out of.
Function ChosenObject(className)

	ChosenObject = ""

	Dim finder
	Set finder = Nothing
	On Error Resume Next
	Set finder = Application.GetObject(FIND_OBJECTS)
	On Error Goto 0

	If finder Is Nothing Then
		MsgBox "This project has no " & FIND_OBJECTS & " tag, so nothing " & _
		       "can be looked for.", vbCritical, DISTRIBUTION_TITLE
		Exit Function
	End If

	finder.WriteEx className

	If finder.DocString <> EXIT_SUCCESS Then
		MsgBox "The " & className & " objects could not be looked for - the " & _
		       "console says why.", vbCritical, DISTRIBUTION_TITLE
		Exit Function
	End If

	Dim answer
	answer = ""
	On Error Resume Next
	answer = CStr(Application.GetObject(FOUND_OBJECTS).Value)
	On Error Goto 0

	If Trim(answer) = "" Then
		MsgBox "No " & className & " was found in this domain.", _
		       vbExclamation, DISTRIBUTION_TITLE
		Exit Function
	End If

	Dim found
	found = Split(answer, vbCrLf)

	If UBound(found) = 0 Then
		ChosenObject = found(0)
		Exit Function
	End If

	ChosenObject = PickObject(found, className)

End Function


' Which of several, asked as a number against a numbered list.
'
' An InputBox is what a viewer has to ask with, so the list goes into
' the prompt and the answer comes back as free text - checked here,
' because anything at all can be typed into one.
Function PickObject(found, className)

	PickObject = ""

	Dim prompt, i
	prompt = "More than one " & className & " was found." & vbCrLf & _
	         "Type the number of the one to distribute through:" & vbCrLf & vbCrLf

	For i = 0 To UBound(found)
		prompt = prompt & "  " & (i + 1) & " - " & found(i) & vbCrLf
	Next

	Dim reply
	reply = InputBox(prompt, DISTRIBUTION_TITLE, "1")

	' Cancel and an empty box both come back "", and neither is a choice.
	If Trim(reply) = "" Then Exit Function

	If Not IsNumeric(reply) Then
		MsgBox "'" & reply & "' is not one of the numbers on the list.", _
		       vbExclamation, DISTRIBUTION_TITLE
		Exit Function
	End If

	Dim choice
	choice = CLng(reply)

	If choice < 1 Or choice > UBound(found) + 1 Then
		MsgBox "There is no " & choice & " on the list.", _
		       vbExclamation, DISTRIBUTION_TITLE
		Exit Function
	End If

	PickObject = found(choice - 1)

End Function


' Where the point numbering starts, asked rather than fixed.
'
' An InputBox is what a viewer has to ask with, so the answer comes back
' as free text and is checked here - anything at all can be typed in.
Function ChosenOffset()

	ChosenOffset = ""

	Dim reply
	reply = InputBox("Address to start the point numbering from:" & vbCrLf & vbCrLf & _
	                 "Every signal is numbered upwards from here, monitored " & _
	                 "points and commands alike.", DISTRIBUTION_TITLE, DEFAULT_OFFSET)

	' Cancel and an empty box both come back "", and neither is an answer.
	If Trim(reply) = "" Then Exit Function

	If Not IsNumeric(reply) Then
		MsgBox "'" & reply & "' is not a number.", vbExclamation, DISTRIBUTION_TITLE
		Exit Function
	End If

	Dim offset
	offset = CLng(reply)

	If offset < 0 Then
		MsgBox "An address cannot be negative.", vbExclamation, DISTRIBUTION_TITLE
		Exit Function
	End If

	ChosenOffset = CStr(offset)

End Function

' E3 will not take a Function as the last thing in a scope.
Sub EndOfScope()
		
End Sub

<xatm_config_screens.Config.btnExport:btnExport_Click()>
Sub btnExport_Click()

	' Where the two files go, asked here and not in the export: that
	' runs as a data server event and a data server has no screen. The
	' path travels as the payload, the way the substation does.
	'
	' Written inside the Sub because a scope has to open on Sub.

	Dim chosenPath
	chosenPath = ChosenFile()

	If chosenPath = "" Then Exit Sub

	Dim exporter
	Set exporter = Nothing
	On Error Resume Next
	Set exporter = Application.GetObject(EXPORT_SIGNALS)
	On Error Goto 0

	If exporter Is Nothing Then
		MsgBox "This project has no " & EXPORT_SIGNALS & " tag, so the " & _
		       "signals cannot be exported.", vbCritical, EXPORT_TITLE
		Exit Sub
	End If

	exporter.WriteEx chosenPath

	If exporter.DocString <> EXIT_SUCCESS Then
		MsgBox "The signals were not exported, or only in part - the " & _
		       "console says which.", vbExclamation, EXPORT_TITLE
		Exit Sub
	End If

	MsgBox "The signals were written beside:" & vbCrLf & vbCrLf & chosenPath & _
		   vbCrLf & vbCrLf & "as _Discretos.csv and _Comandos.csv. Open them in " & _
		   "Excel and paste the columns into the load plan.", _
		   vbInformation, EXPORT_TITLE

End Sub

Const EXPORT_SIGNALS = "xatm_config_data.Config.ExportSignalsToCsv"
Const EXIT_SUCCESS   = "EXIT_SUCCESS"
Const EXPORT_TITLE   = "Export signals"


' Where the files go, asked with the Viewer's own Save dialog so that a
' folder is browsed for rather than typed out - ShowFilePicker(Open,
' Filename, Extension, Flags, Filter), with Open False for a Save box
' and Filename carrying the answer back out.
'
' Empty when the dialog is refused. Refusing is a decision and is not
' asked again: the InputBox below is what happens when this build has
' no picker at all, and not what happens when someone pressed Cancel.
Function ChosenFile()

	ChosenFile = ""

	Dim chosen
	chosen = DefaultPath()

	Dim picked, hasPicker
	picked    = False
	hasPicker = True

	On Error Resume Next
	Err.Clear

	picked = Application.ShowFilePicker(False, chosen, "csv||", _
	                                    PATHMUSTEXIST + OVERWRITEPROMPT, FILE_FILTER)

	If Err.Number <> 0 Then
		hasPicker = False
		Err.Clear
	End If

	On Error Goto 0

	If hasPicker Then
		If picked Then ChosenFile = Trim(chosen & "")
		Exit Function
	End If

	ChosenFile = Trim(InputBox( _
		"Full path of the file to write. Two are made from it, one ending " & _
		"_Discretos.csv and one _Comandos.csv." & vbCrLf & vbCrLf & _
		"They are overwritten if they are already there, so close them " & _
		"in Excel first.", _
		EXPORT_TITLE, DefaultPath()) & "")

End Function


' What the dialog opens on, and what the fallback box is filled with.
Function DefaultPath()

	Dim folderPath
	folderPath = StartFolder()

	If folderPath = "" Then
		DefaultPath = DEFAULT_NAME
		Exit Function
	End If

	If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

	DefaultPath = folderPath & DEFAULT_NAME

End Function


' The flags ShowFilePicker takes, summed. The path has to exist
' because nothing here creates one, and overwriting is confirmed
' because a second export onto the same name is the ordinary case and
' losing an edited file to it is not.
Const PATHMUSTEXIST   = 32
Const OVERWRITEPROMPT = 128

' Description and pattern in pairs, the whole ending in two bars.
Const FILE_FILTER = "CSV (*.csv)|*.csv|All files (*.*)|*.*||"
Const DEFAULT_NAME = "Sinais.csv"


' What the box opens on: where E3 was started from, which is the folder
' the domain file sits in. The same answer XMLFilePath opens on, and
' written out again here because a screen scope cannot call into a data
' server one.
Function StartFolder()

	StartFolder = ""

	Dim shell
	Set shell = Nothing

	On Error Resume Next
	Err.Clear

	Set shell = CreateObject("WScript.Shell")
	If Not shell Is Nothing Then StartFolder = Trim(shell.CurrentDirectory & "")

	If Err.Number <> 0 Then
		StartFolder = ""
		Err.Clear
	End If

	On Error Goto 0

End Function

' E3 will not take a Function as the last thing in a scope.
Sub EndOfScope()
	
End Sub

<xatm_config_screens.Config.btnSave:btnSave_Click()>
Sub btnSave_Click()

	' 1. The document into the project.
	Dim importer
	Set importer = Application.GetObject(IMPORT_XML)
	importer.WriteEx True

	Dim imported
	imported = (importer.DocString = EXIT_SUCCESS)

	' 2. The interface the Elipse application binds to, made again from
	' what the project holds now.
	'
	' After the import, so a rename made on this screen has already
	' landed. Whether or not the import had anything to write, and even
	' where it refused: the interface is built from the project and not
	' from the document, so what the document did is not its business, and
	' a project left without its tags is worse than one whose tags are a
	' rename behind.
	RebuildInterface

	' 3. The layout the screen is showing, and then the project file. The
	' objects the import just wrote are in the same container, so one Save
	' persists them.
	'
	' Not the interface: that is a data server of its own and a container
	' of its own, and RebuildInterface has already saved it. This runs on
	' the way out of a failed import all the same, because the import may
	' have written some of the document before it stopped.
	Dim layoutFolder
	Set layoutFolder = Application.GetObject("XATM_Data.Automation.Layout")
	
	layoutFolder.Item("Incomer").WriteEx Screen.Item("SelectLayoutIncomer").Index
	layoutFolder.Item("Transformer").WriteEx Screen.Item("SelectLayoutTransformer").Index
	layoutFolder.Item("Busbar").WriteEx Screen.Item("SelectLayoutBusbar").Index

	' Save answers True or False and raises nothing. Unchecked, a refused
	' save reported success and the file went out against a project that
	' had not taken the changes - which is the one thing the import is
	' gated on below, and no less true here.
	If Not layoutFolder.Context("Container").Save() Then
		MsgBox "The project was not saved - E3 would not write it, and the " & _
		       "XML file has been left alone. A licence that will not carry " & _
		       "this many tags is the usual reason.", vbCritical, "Save"
		Exit Sub
	End If

	' The file is the one thing the import does gate: written against a
	' project that does not match it, it is worse than no file at all.
	If Not imported Then
		MsgBox "The project and the interface were saved, but the document " & _
		       "could not be written into the project, so the XML file was " & _
		       "left alone - the console says why.", vbCritical, "Save"
		Exit Sub
	End If

	' 4. And the file.
	Dim saver
	Set saver = Application.GetObject(SAVE_XML)
	saver.WriteEx True

	If saver.DocString <> EXIT_SUCCESS Then
		MsgBox "The project was saved, but the XML file was not - the console " & _
		       "says why.", vbExclamation, "Save"
		Exit Sub
	End If

End Sub

Const IMPORT_XML   = "xatm_config_data.Config.ImportXml"
Const SAVE_XML     = "xatm_config_data.Config.SaveXML"
Const EXIT_SUCCESS = "EXIT_SUCCESS"


' ------------------------------------------------------------
'  THE INTERFACE
' ------------------------------------------------------------

' Where the Elipse application meets the automation: one internal tag per
' interfaced property, in a folder named after the object it belongs to,
' so a detached project binds to XATM_Interface and never reaches into an
' XObject.
Const INTERFACE_ROOT     = "XATM_Interface"
Const INTERFACE_ROOT_ALT = "XATM_Data.XATM_Interface"
Const INTERFACE_FOLDER   = "DataFolder"
Const INTERFACE_TAG      = "InternalTag"

Const DATA_ROOT     = "XATM_Data"
Const CONFIG_DATA   = "xatm_config_data"
Const HELPER_FOLDER = "PropertiesHelper"

Const SELF_MARK 	= "$XATM$"

' Scratch cell for the late-bound write done through Execute, the way the
' import keeps one.
Dim gInterfaceTag

' What the last rebuild came out at, for the line it writes to the log.
Dim gFolders
Dim gTags


' Builds the interface again from what the project holds this moment.
'
' Everything under the root goes first and is made again rather than
' brought up to date. A device can be renamed on this screen, and a
' folder named after the name it used to have is indistinguishable from
' one nobody wants any more - so nothing tries to tell them apart. There
' is no history in a tag like these to lose by rebuilding it.
Sub RebuildInterface()

	Dim root
	Set root = InterfaceRoot()

	If root Is Nothing Then
		MsgBox "The interface was not built - the project has no data server " & _
		       "at " & INTERFACE_ROOT & ".", vbExclamation, "Save"
		Exit Sub
	End If

	ClearFolder root

	gFolders = 0
	gTags    = 0

	Dim problem
	problem = ""

	InterfaceFolder root, Application.GetObject(DATA_ROOT), "", problem

	' Written here, and not left to the Save that persists XATM_Data: the
	' interface is a data server of its own, and the container holding the
	' automation objects is not the container holding these tags.
	'
	' Save answers True or False and raises nothing, so the answer is what
	' decides rather than Err.
	Dim saved
	saved = False

	On Error Resume Next
	Err.Clear

	saved = root.Save()

	If Err.Number <> 0 Then Err.Clear
	On Error Goto 0

	If Not saved Then
		problem = problem & vbCrLf & "  not saved - E3 would not write the " & _
		          "interface data server."
	End If

	' Said out loud even when it all went through. A routine that builds
	' forty tags and mentions only the ones it could not build reads the
	' same whether it ran or never got called at all, and the log is the
	' only place the difference shows.
	If gFolders = 0 Then
		WriteLog "Interface - nothing is marked for it."
	Else
		WriteLog "Interface - " & gFolders & " objects, " & gTags & " tags."
	End If

	If problem <> "" Then
		WriteLog "Interface is incomplete:" & problem
		MsgBox "The interface is incomplete:" & vbCrLf & problem, vbExclamation, "Save"
	End If

End Sub


' The data server the interface is built in. Named at the root of the
' data project, and looked for inside XATM_Data as well because which of
' the two it sits under is a matter of how the project was laid out.
Function InterfaceRoot()

	Set InterfaceRoot = Nothing

	On Error Resume Next
	Set InterfaceRoot = Application.GetObject(INTERFACE_ROOT)
	On Error Goto 0

	If Not InterfaceRoot Is Nothing Then Exit Function

	On Error Resume Next
	Set InterfaceRoot = Application.GetObject(INTERFACE_ROOT_ALT)
	On Error Goto 0

End Function


' Empties a folder. Taken out in a second pass, because a collection is
' not walked while what is in it is being removed.
Sub ClearFolder(folder)

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim item
	For Each item In folder
		doomed.Add doomed.Count, item.Name
	Next

	Dim n
	For Each n In doomed.Keys

		On Error Resume Next
		folder.DeleteObject doomed(n)
		On Error Goto 0

	Next

End Sub


' One folder of the project: an interface for every object in it a
' manifest speaks for, and then the same for whatever is under it. A
' class with no manifest is not a device, so it is treated as a folder -
' and one E3 will not enumerate simply adds nothing.
'
' path is where the walk has got to, as the chain of folder names under
' the interface root that mirrors it - "" at XATM_Data itself, then
' "Substation", then "Substation.Busbar". The interface comes out shaped
' like the data it speaks for, so a tie breaker is found where the tree
' says it is rather than in a heap with everything else.
Sub InterfaceFolder(root, folder, path, problem)

	Dim item, bag

	For Each item In folder

		Set bag = ManifestOf(TypeName(item))

		If Not bag Is Nothing Then

			InterfaceObject root, path, item, bag, problem

		Else

			On Error Resume Next
			InterfaceFolder root, item, ChildPath(path, item.Name), problem
			On Error Goto 0

		End If

	Next

End Sub


' One name onto a path, and the first name is the whole of it.
Function ChildPath(path, name)

	If path = "" Then
		ChildPath = name
	Else
		ChildPath = path & "." & name
	End If

End Function


' The folder at a path under the interface root, each level made where it
' is not already there.
'
' Made on the way to something rather than mirrored up front, because a
' folder is only wanted where something goes into it: the data project
' also holds Layout and every folder an engineer adds, and an interface
' full of empty folders says nothing about what is on it.
Function EnsureFolder(root, path, problem)

	Set EnsureFolder = root
	If path = "" Then Exit Function

	Dim names
	names = Split(path, ".")

	Dim parent, child, i
	Set parent = root

	For i = 0 To UBound(names)

		Set child = Nothing

		On Error Resume Next
		Set child = parent.Item(names(i))
		On Error Goto 0

		If child Is Nothing Then Set child = NewChild(parent, INTERFACE_FOLDER, names(i))

		If child Is Nothing Then
			problem = problem & vbCrLf & "  no folder for " & path
			Set EnsureFolder = Nothing
			Exit Function
		End If

		Set parent = child

	Next

	Set EnsureFolder = parent

End Function


' One object's folder, and a tag in it for each property the manifest
' interfaces. An object with nothing interfaced gets no folder at all,
' rather than an empty one to wonder about.
Sub InterfaceObject(root, path, obj, bag, problem)

	Dim key, p

	Dim wanted
	wanted = False

	For Each key In bag.Keys
		If bag(key).IsInterfaced() Then wanted = True
	Next

	If Not wanted Then Exit Sub

	Dim parent
	Set parent = EnsureFolder(root, path, problem)
	If parent Is Nothing Then Exit Sub

	Dim folder
	Set folder = NewChild(parent, INTERFACE_FOLDER, obj.Name)

	If folder Is Nothing Then
		problem = problem & vbCrLf & "  no folder for " & obj.Name
		Exit Sub
	End If

	gFolders = gFolders + 1

	For Each key In bag.Keys

		Set p = bag(key)
		If p.IsInterfaced() Then InterfaceProperty folder, obj, p, problem

	Next

End Sub


' One tag for one property, and what ties the two together.
'
' A reading is linked to the property and follows it: the tag is a window
' on what the automation holds, and writing to it would be writing to the
' link. A command is the other way about - nothing in this project drives
' it, the detached project writes it - so it is left unlinked and the
' object's property is pointed at it instead.
'
' An IOTag is never interfaced, whatever a manifest says: a breaker is
' commanded and read through its IED, and those tags are the engineer's
' to configure.
Sub InterfaceProperty(folder, obj, p, problem)

	If LCase(p.DataType & "") = "iotag" Then Exit Sub

	Dim tag
	Set tag = NewChild(folder, INTERFACE_TAG, p.Name)
	tag.DocString = SELF_MARK
	
	If tag Is Nothing Then
		problem = problem & vbCrLf & "  no tag for " & obj.Name & "." & p.Name
		Exit Sub
	End If

	gTags = gTags + 1

	If LCase(p.DataType & "") = "internaltag" Then

		gInterfaceTag = tag.PathName

		On Error Resume Next
		Execute "obj." & p.Name & " = gInterfaceTag"

		If Err.Number <> 0 Then
			problem = problem & vbCrLf & "  " & obj.Name & "." & p.Name & _
			          " would not take its tag - " & Err.Description
			Err.Clear
		End If

		On Error Goto 0

		Exit Sub

	End If

	On Error Resume Next
	tag.Links.CreateLink "Value", obj.PathName & "." & p.Name

	If Err.Number <> 0 Then
		problem = problem & vbCrLf & "  " & obj.Name & "." & p.Name & _
		          " would not link - " & Err.Description
		Err.Clear
	End If

	On Error Goto 0

End Sub


' A new child of a folder, Nothing when it would not take one -
' AddObject(ClassName, [Activate], [ObjectName]).
Function NewChild(parent, className, name)

	Set NewChild = Nothing

	On Error Resume Next
	Set NewChild = parent.AddObject(className, True, name)
	On Error Goto 0

End Function


' The manifest declared for a class, or Nothing when the class has none -
' the same lookup the export and the import do.
Function ManifestOf(className)

	Set ManifestOf = Nothing

	On Error Resume Next
	Set ManifestOf = Application.GetObject(CONFIG_DATA).Item(HELPER_FOLDER).Item(className).Value
	On Error Goto 0

End Function


' The console the automation logs to, and the E3 trace either way.
Sub WriteLog(message)

	Dim consoleLogEngine
	Set consoleLogEngine = Nothing

	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[Interface] - " & message
	On Error Goto 0

	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[Interface] - " & message
	End If
		
End Sub

<xatm_config_screens.Config:Config_OnPreShow(Arg)>
Sub Config_OnPreShow(Arg)
	
	' ================================
	' LAYOUT
	' ================================
	
	Dim imageList, i
	
	' ---- INCOMER ----
	Dim incomerType
	incomerType = Application.GetObject("XATM_Data.Automation.Layout.Incomer").Value
	
	
	imageList = Split(Item("SelectLayoutIncomer").ImageList, ",")
	
	For i = 0 To UBound(imageList)
		
		If Trim(imageList(i)) = incomerType Then
			
			Item("SelectLayoutIncomer").Index = Trim(imageList(i))
			Exit For
			
		End If
		
	Next
	
	' ---- TRANSFORMER ----
	Dim transformerType
	transformerType = Application.GetObject("XATM_Data.Automation.Layout.Transformer").Value
	
	
	imageList = Split(Item("SelectLayoutTransformer").ImageList, ",")
	
	For i = 0 To UBound(imageList)
		
		If Trim(imageList(i)) = transformerType Then
			
			Item("SelectLayoutTransformer").Index = Trim(imageList(i))
			Exit For
			
		End If
		
	Next
	
	' ---- BUSBAR ----
	Dim busbarType
	busbarType = Application.GetObject("XATM_Data.Automation.Layout.Busbar").Value
	
	imageList = Split(Item("SelectLayoutBusbar").ImageList, ",")
	
	For i = 0 To UBound(imageList)
		
		If Trim(imageList(i)) = busbarType Then
			
			Item("SelectLayoutBusbar").Index = Trim(imageList(i))
			Exit For
			
		End If
		
	Next
		
End Sub

<xatm_config_screens.DomainBrowser.TreeView:TreeView_Expand(Node)>
Sub TreeView_Expand(Node)

	' The children of one node, fetched the moment it is opened.
	'
	' Runs once per node: the placeholder is taken out and never comes
	' back, so a node already filled falls out at the first line.

	If Not OnlyPlaceholder(Node) Then Exit Sub

	Dim doc
	Set doc = CatalogDoc()
	If doc Is Nothing Then Exit Sub

	' The root stands for the document itself; everything else is found by
	' the path its Tag carries.
	Dim xmlParent
	Set xmlParent = Nothing

	On Error Resume Next
	If Node.Parent Is Nothing Then
		Set xmlParent = doc.documentElement
	Else
		Set xmlParent = FindXmlByPath(doc, TreeNodeTagPath(Node.Tag))
	End If
	On Error GoTo 0

	If xmlParent Is Nothing Then
		Application.Trace "TreeView expand found nothing at " & CStr(Node.Tag)
		Exit Sub
	End If

	Dim tree
	Set tree = Screen.Item("TreeView")

	' A folder can hold hundreds, and each Add repaints.
	On Error Resume Next
	tree.Visible = False
	On Error GoTo 0

	On Error Resume Next
	tree.Nodes.Remove "p" & Node.Key
	Err.Clear
	On Error GoTo 0

	Dim useImages
	useImages = BindTreeImageList(tree)

	Dim child, added
	added = 0

	For Each child In xmlParent.childNodes
		If child.nodeName = "Object" Then
			AddChildNode tree, Node.Key, child, useImages
			added = added + 1
		End If
	Next

	On Error Resume Next
	tree.Visible = True
	On Error GoTo 0

	Application.Trace "TreeView expanded " & CStr(Node.Text) & " with " & _
	                  CStr(added) & " nodes."

End Sub


' The catalogue, parsed once and kept.
'
' Held for the life of the screen rather than parsed per expand: the
' document is the same every time and parsing twenty thousand objects is
' the one cost that would undo the whole point of loading lazily.
Dim gDoc

Function CatalogDoc()

	Set CatalogDoc = Nothing

	If IsObject(gDoc) Then
		If Not (gDoc Is Nothing) Then
			Set CatalogDoc = gDoc
			Exit Function
		End If
	End If

	Dim xmlText
	xmlText = CatalogXmlText()
	If Len(xmlText) = 0 Then
		Application.Trace "Catalog XML is empty; nothing to expand."
		Exit Function
	End If

	Dim doc
	Set doc = CreateObject("MSXML2.DOMDocument.6.0")
	doc.async = False
	doc.loadXML xmlText

	If doc.parseError.errorCode <> 0 Then
		Application.Trace "Catalog XML parse failed: " & doc.parseError.reason
		Exit Function
	End If

	Set gDoc = doc
	Set CatalogDoc = gDoc

End Function


' Whether a node is still holding its stand-in, and so has never been
' opened. One child, keyed with a p in front, is the whole test.
Function OnlyPlaceholder(Node)

	OnlyPlaceholder = False

	On Error Resume Next
	If Node.Children <> 1 Then Exit Function
	If Left(CStr(Node.Child.Key), 1) = "p" Then OnlyPlaceholder = True
	Err.Clear
	On Error GoTo 0

End Function


' The catalogue entry at one path.
'
' A plain walk of every Object rather than an XPath: a path can carry a
' bracket or a quote, and there is no escaping either inside an XPath
' literal. One pass over the document per expand is cheap beside the
' twenty thousand adds this replaced.
Function FindXmlByPath(doc, path)

	Set FindXmlByPath = Nothing
	If Len(path) = 0 Then Exit Function

	Dim nodes, i
	Set nodes = doc.getElementsByTagName("Object")

	For i = 0 To nodes.length - 1
		If nodes.item(i).getAttribute("path") = path Then
			Set FindXmlByPath = nodes.item(i)
			Exit Function
		End If
	Next

End Function


' One child under an opened node, with a stand-in of its own where the
' catalogue says it has children to come.
Sub AddChildNode(tree, parentKey, xmlNode, useImages)

	gSeq = gSeq + 1

	Dim key
	key = "x" & CStr(gSeq)

	Dim text
	text = XmlNodeText(xmlNode)

	Dim imageKey
	imageKey = ""
	If useImages Then imageKey = XmlNodeImageKey(xmlNode, tree.ImageList)

	Dim treeNode
	Set treeNode = Nothing

	On Error Resume Next
	If useImages And Len(CStr(imageKey)) > 0 Then
		Set treeNode = tree.Nodes.Add(parentKey, 4, key, text, imageKey, imageKey)
	End If
	Err.Clear
	If treeNode Is Nothing Then Set treeNode = tree.Nodes.Add(parentKey, 4, key, text)
	If Err.Number <> 0 Then
		Application.Trace "TreeView add failed: " & xmlNode.getAttribute("path") & _
		                  " - " & Err.Description
		Err.Clear
		On Error GoTo 0
		Exit Sub
	End If

	treeNode.Tag = xmlNode.getAttribute("path")
	treeNode.ToolTipText = xmlNode.getAttribute("path")
	Err.Clear
	On Error GoTo 0

	If HasObjectChildren(xmlNode) Then
		On Error Resume Next
		tree.Nodes.Add key, 4, "p" & key, "..."
		Err.Clear
		On Error GoTo 0
	End If

End Sub

' Keys have to be unique across the whole control, and the root already
' took n1 in the other scope - so these are x1, x2, and cannot collide.
Dim gSeq


Function HasObjectChildren(xmlNode)

	HasObjectChildren = False

	Dim child
	For Each child In xmlNode.childNodes
		If child.nodeName = "Object" Then
			HasObjectChildren = True
			Exit Function
		End If
	Next

End Function


' What the node is called, and which picture it gets. Named apart from
' the copies in the screen's scope because that scope cannot be seen from
' here, and a name shared with a local would shadow it.
Function XmlNodeText(xmlNode)

	XmlNodeText = xmlNode.getAttribute("name")
	If Len(XmlNodeText) = 0 Then XmlNodeText = xmlNode.getAttribute("path")
	If Len(XmlNodeText) = 0 Then XmlNodeText = "[object]"

End Function


Function XmlNodeImageKey(xmlNode, imageList)

	Dim className
	className = xmlNode.getAttribute("type")

	Dim imageValue
	imageValue = ImageListValueForTag(imageList, className)

	If Len(CStr(imageValue)) = 0 Then
		imageValue = ImageListValueForTag(imageList, "default")
	End If

	If Len(CStr(imageValue)) > 0 Then
		XmlNodeImageKey = imageValue
	ElseIf Len(className) = 0 Then
		XmlNodeImageKey = "Object"
	Else
		XmlNodeImageKey = className
	End If

End Function


Function BindTreeImageList(tree)

	BindTreeImageList = False

	On Error Resume Next
	Dim imageList
	Set imageList = Screen.Item("TreeImages")
	If Err.Number <> 0 Then
		Err.Clear
		Exit Function
	End If

	Set tree.ImageList = imageList
	If Err.Number = 0 Then BindTreeImageList = True
	Err.Clear
	On Error GoTo 0

End Function

' A Function may not be the last thing in a scope.
Sub EndOfScope()
	
End Sub

<xatm_config_screens.DomainBrowser.TreeView:TreeView_NodeClick(Node)>
Sub TreeView_NodeClick(Node)

	Dim objectPath
	objectPath = TreeNodeTagPath(Node.Tag)
	If Len(objectPath) = 0 Then
		objectPath = Node.FullPath
	End If

	Dim className
	className = TreeNodeTagClassName(Node.Tag)

	Screen.Item("txtPath").Value = objectPath
	FillPropertiesList className
		
End Sub

Function TreeNodeTag(objectPath, className)

	TreeNodeTag = objectPath & vbTab & className

End Function

Function TreeNodeTagPath(tagText)

	Dim parts
	parts = Split(CStr(tagText), vbTab)

	TreeNodeTagPath = parts(0)

End Function

' The catalogue, off the one tag that holds it - the same tag the
' OnPreShow copy reads, and for the reason written out there. This copy
' used to fall back to C:\temp\appbrowser\elipse-catalog.xml, a file
' nothing writes any more.
Function CatalogXmlText()

	Dim xmlText
	xmlText = ""

	Dim xmlTag
	Set xmlTag = Nothing

	On Error Resume Next
	Set xmlTag = Application.GetObject(CATALOG_CONTENT)
	On Error GoTo 0

	If Not (xmlTag Is Nothing) Then xmlText = CStr(xmlTag.Value & "")

	CatalogXmlText = xmlText

End Function

' Where the catalogue lives. Written out again because this is another
' tag's scope and cannot see the copy over in OnPreShow.
Const CATALOG_CONTENT = "xatm_config_data.Catalog.XMLContent"


Function Pad2(value)

	If value < 10 Then
		Pad2 = "0" & CStr(value)
	Else
		Pad2 = CStr(value)
	End If

End Function

Function TreeNodeTagClassName(tagText)

	TreeNodeTagClassName = ""

	Dim parts
	parts = Split(CStr(tagText), vbTab)
	If UBound(parts) >= 1 Then
		TreeNodeTagClassName = parts(1)
		Exit Function
	End If

	TreeNodeTagClassName = CatalogClassNameByPath(TreeNodeTagPath(tagText))

End Function

Function ImageTagHasClass(tagText, className)

	ImageTagHasClass = False

	Dim expected
	expected = LCase(Trim(className))
	If Len(expected) = 0 Then Exit Function

	Dim tags
	tags = Split(tagText, ",")

	Dim i
	For i = 0 To UBound(tags)
		Dim actual
		actual = LCase(Trim(tags(i)))
		If actual = expected Then
			ImageTagHasClass = True
			Exit Function
		End If
	Next

End Function

Function CatalogClassNameByPath(objectPath)

	CatalogClassNameByPath = ""
	If Len(objectPath) = 0 Then Exit Function

	Dim xmlText
	xmlText = CatalogXmlText()
	If Len(xmlText) = 0 Then Exit Function

	On Error Resume Next
	Dim doc
	Set doc = CreateObject("MSXML2.DOMDocument.6.0")
	doc.async = False
	doc.loadXML xmlText
	If Err.Number <> 0 Then
		Application.Trace "Catalog XML load for class lookup failed: " & Err.Description
		Err.Clear
		Exit Function
	End If
	On Error GoTo 0

	If doc.parseError.errorCode <> 0 Then
		Application.Trace "Catalog XML parse for class lookup failed: " & doc.parseError.reason
		Exit Function
	End If

	CatalogClassNameByPath = CatalogNodeClassNameByPath(doc.documentElement, objectPath)

End Function

Function ImageListValueForTag(imageList, tagName)

	ImageListValueForTag = ""
	If Len(tagName) = 0 Then Exit Function

	On Error Resume Next
	Dim image
	For Each image In imageList.ListImages
		If Err.Number <> 0 Then
			Application.Trace "ImageList traversal failed: " & Err.Description
			Err.Clear
			Exit For
		End If

		Dim tagText
		tagText = CStr(image.Tag)
		
		If ImageTagHasClass(tagText, tagName) Then
			If Len(CStr(image.Key)) > 0 Then
				ImageListValueForTag = image.Key
			Else
				ImageListValueForTag = image.Index
			End If
			Exit Function
		End If
	Next
	Err.Clear
	On Error GoTo 0

End Function

Function CatalogNodeClassNameByPath(xmlNode, objectPath)

	CatalogNodeClassNameByPath = ""

	Dim child
	For Each child In xmlNode.childNodes
		If child.nodeName = "Object" Then
			If child.getAttribute("path") = objectPath Then
				CatalogNodeClassNameByPath = child.getAttribute("type")
				Exit Function
			End If

			Dim className
			className = CatalogNodeClassNameByPath(child, objectPath)
			If Len(className) > 0 Then
				CatalogNodeClassNameByPath = className
				Exit Function
			End If
		End If
	Next

End Function

Sub FillPropertiesList(className)

	On Error Resume Next
	Dim list
	Set list = Screen.Item("lstProperties")
	If Err.Number <> 0 Then
		Application.Trace "Properties list was not found: " & Err.Description
		Err.Clear
		Exit Sub
	End If
	On Error GoTo 0

	ConfigurePropertiesListView list
	ClearPropertiesList list

	If LCase(className) = "domain" Or Len(className) = 0 Then Exit Sub

	Dim imageList
	Set imageList = BindPropertiesImageList(list)

	AddCommonPropertyNames list, imageList

	Select Case LCase(className)
		Case "iotag"
			AddIOTagPropertyNames list, imageList
		Case "internaltag"
			AddInternalTagPropertyNames list, imageList
	End Select

End Sub

Sub ConfigurePropertiesListView(list)

	On Error Resume Next
	list.View = 3
	list.FullRowSelect = True
	list.HideColumnHeaders = True
	list.ColumnHeaders.Clear
	list.ColumnHeaders.Add , , "Property", 5000
	If Err.Number <> 0 Then
		Application.Trace "Properties ListView configure failed: " & Err.Description
		Err.Clear
	End If
	On Error GoTo 0

End Sub

Sub ClearPropertiesList(list)

	On Error Resume Next
	list.ListItems.Clear
	If Err.Number <> 0 Then
		Application.Trace "Properties ListView clear failed: " & Err.Description
		Err.Clear
	End If
	On Error GoTo 0

End Sub

Function BindPropertiesImageList(list)

	Set BindPropertiesImageList = Nothing

	On Error Resume Next
	Dim imageList
	Set imageList = Screen.Item("TreeImages")
	
	If Err.Number <> 0 Then
		'Application.Trace "Properties ImageList was not found: " & Err.Description
		
		MsgBox "Properties ImageList was not found: " & Err.Description
		Err.Clear
		Exit Function
	End If

	Set list.SmallIcons = imageList
	If Err.Number <> 0 Then
		Err.Clear
		Exit Function
	End If
	On Error GoTo 0

	Set BindPropertiesImageList = imageList

End Function

Sub AddCommonPropertyNames(list, imageList)

	AddPropertyToList list, imageList, "Name", "string"
	AddPropertyToList list, imageList, "DocString", "string"
	AddPropertyToList list, imageList, "PathName", "string"
	AddPropertyToList list, imageList, "PathVolume", "string"
	AddPropertyToList list, imageList, "PathContainer", "string"

End Sub

Sub AddIOTagPropertyNames(list, imageList)

	AddPropertyToList list, imageList, "AdviseType", "variant"
	AddPropertyToList list, imageList, "AllowRead", "variant"
	AddPropertyToList list, imageList, "AllowWrite", "variant"

	Dim i
	For i = 0 To 31
		AddPropertyToList list, imageList, "Bit" & Pad2(i), "boolean"
	Next

	AddPropertyToList list, imageList, "DeviceHigh", "variant"
	AddPropertyToList list, imageList, "DeviceLow", "variant"
	AddPropertyToList list, imageList, "EnableDeadBand", "variant"
	AddPropertyToList list, imageList, "EnableDriverEvent", "variant"
	AddPropertyToList list, imageList, "EnableScaling", "variant"
	AddPropertyToList list, imageList, "EU", "variant"
	AddPropertyToList list, imageList, "EUHigh", "variant"
	AddPropertyToList list, imageList, "EULow", "variant"
	AddPropertyToList list, imageList, "N1", "variant"
	AddPropertyToList list, imageList, "N2", "variant"
	AddPropertyToList list, imageList, "N3", "variant"
	AddPropertyToList list, imageList, "N4", "variant"
	AddPropertyToList list, imageList, "ParamDevice", "variant"
	AddPropertyToList list, imageList, "ParamItem", "variant"
	AddPropertyToList list, imageList, "PercentDeadBand", "variant"
	AddPropertyToList list, imageList, "Quality", "variant"
	AddPropertyToList list, imageList, "RawValue", "variant"
	AddPropertyToList list, imageList, "Scan", "variant"
	AddPropertyToList list, imageList, "TimeStamp", "variant"
	AddPropertyToList list, imageList, "UseBitFields", "variant"
	AddPropertyToList list, imageList, "Value", "variant"

End Sub

Sub AddInternalTagPropertyNames(list, imageList)

	AddPropertyToList list, imageList, "Quality", "variant"
	AddPropertyToList list, imageList, "Retentive", "variant"
	AddPropertyToList list, imageList, "TimeStamp", "variant"
	AddPropertyToList list, imageList, "Value", "variant"

End Sub

Sub AddPropertyToList(list, imageList, propertyName, propertyType)

	Dim imageKey
	imageKey = ""

	If Not (imageList Is Nothing) Then
	
		imageKey = ImageListValueForTag(imageList, propertyType)
		
		If Len(CStr(imageKey)) = 0 Then
			imageKey = ImageListValueForTag(imageList, "variant")
		End If
		If Len(CStr(imageKey)) = 0 Then
			imageKey = ImageListValueForTag(imageList, "default")
		End If
	End If

	On Error Resume Next
	Dim item
	If Len(CStr(imageKey)) > 0 Then
		Set item = list.ListItems.Add(, , propertyName, , imageKey)
	Else
		Set item = list.ListItems.Add(, , propertyName)
	End If
	If Err.Number <> 0 Then
		Application.Trace "Properties ListView add failed: " & propertyName & " - " & Err.Description
		Err.Clear
		Exit Sub
	End If

	item.Tag = propertyType
	On Error GoTo 0

End Sub

<xatm_config_screens.DomainBrowser.btnCancel:btnCancel_Click()>
Sub btnCancel_Click()
	
	Application.GetObject("xatm_config_data.Catalog.SelectedPath").WriteEx ""
	Screen.Close(-1)

End Sub

<xatm_config_screens.DomainBrowser.btnOk:btnOk_Click()>
Sub btnOk_Click()

	Application.GetObject("xatm_config_data.Catalog.SelectedPath").WriteEx Screen.Item("txtPath").Value
	
	Screen.Close(0)

End Sub

<xatm_config_screens.DomainBrowser:DomainBrowser_OnPreShow(Arg)>
Sub DomainBrowser_OnPreShow(Arg)
	
	Application.GetObject("xatm_config_data.Catalog.SelectedPath").WriteEx ""
	
End Sub

Sub FillTree()

	Dim tree
	Set tree = Item("TreeView")

	Dim xmlText
	xmlText = CatalogXmlText()

	FillTreeViewFromCatalogXml tree, xmlText
	
End Sub

' The catalogue, off the one tag that holds it.
'
' It used to try domainbrowser_Data.Catalog.XMLContent first and come
' here only if that was missing or empty - so an export, which writes
' this tag, changed nothing the tree showed for as long as the other tag
' held anything at all. The other copy of this function, over in the
' TreeView's scope, fell back to a file in C:\temp instead, which let the
' two scopes disagree about the same node.
'
' Handed back as it came. The caller parses it and says so when it will
' not parse, so loading it here to check was a second parse that only
' turned a bad catalogue into an empty one.
Function CatalogXmlText()

	CatalogXmlText = ""

	Dim xmlTag
	Set xmlTag = Nothing

	On Error Resume Next
	Set xmlTag = Application.GetObject(CATALOG_CONTENT)
	On Error GoTo 0

	If xmlTag Is Nothing Then
		Application.Trace "There is no " & CATALOG_CONTENT & " to read the catalog from."
		Exit Function
	End If

	CatalogXmlText = CStr(xmlTag.Value & "")

End Function


' Where the catalogue lives, which is where ExportElipseCatalog puts it.
Const CATALOG_CONTENT = "xatm_config_data.Catalog.XMLContent"


Function ImageListValueForTag(imageList, tagName)

	ImageListValueForTag = ""
	If Len(tagName) = 0 Then Exit Function

	On Error Resume Next
	Dim image
	For Each image In imageList.ListImages
		If Err.Number <> 0 Then
			Application.Trace "ImageList traversal failed: " & Err.Description
			Err.Clear
			Exit For
		End If

		Dim tagText
		tagText = CStr(image.Tag)

		If ImageTagHasClass(tagText, tagName) Then
			If Len(CStr(image.Key)) > 0 Then
				ImageListValueForTag = image.Key
			Else
				ImageListValueForTag = image.Index
			End If
			Exit Function
		End If
	Next
	Err.Clear
	On Error GoTo 0

End Function

Function ImageTagHasClass(tagText, className)

	ImageTagHasClass = False

	Dim expected
	expected = LCase(Trim(className))
	If Len(expected) = 0 Then Exit Function

	Dim tags
	tags = Split(tagText, ",")

	Dim i
	For i = 0 To UBound(tags)
		Dim actual
		actual = LCase(Trim(tags(i)))
		If actual = expected Then
			ImageTagHasClass = True
			Exit Function
		End If
	Next

End Function

Function BindTreeImageList(tree)

	BindTreeImageList = False

	On Error Resume Next
	Dim imageList
	Set imageList = Item("TreeImages")
	If Err.Number <> 0 Then
		Application.Trace "TreeView ImageList was not found: " & Err.Description
		Err.Clear
		Exit Function
	End If

	Set tree.ImageList = imageList
	If Err.Number <> 0 Then
		Application.Trace "TreeView ImageList bind failed: " & Err.Description
		Err.Clear
	Else
		BindTreeImageList = True
	End If
	Set imageList = Nothing
	On Error GoTo 0

End Function


Function ElipseDomainName()

	ElipseDomainName = ""

	On Error Resume Next
	Dim domainObj
	Set domainObj = Application.GetObject("[?Server].Domain")
	If Err.Number <> 0 Then
		Err.Clear
		Exit Function
	End If

	ElipseDomainName = UCase(domainObj.Domain)
	If Err.Number <> 0 Then
		ElipseDomainName = ""
		Err.Clear
	End If
	Set domainObj = Nothing
	On Error GoTo 0

End Function

Function AddTreeRootNode(tree, text, sequence, useImages)

	AddTreeRootNode = ""
	If Len(text) = 0 Then Exit Function

	sequence = sequence + 1

	Dim key
	key = "n" & CStr(sequence)

	Dim imageKey
	imageKey = ""
	If useImages Then
		imageKey = ImageListValueForTag(tree.ImageList, "folder")
		If Len(CStr(imageKey)) = 0 Then
			imageKey = ImageListValueForTag(tree.ImageList, "default")
		End If
	End If

	On Error Resume Next
	Dim treeNode
	Set treeNode = Nothing
	If useImages And Len(CStr(imageKey)) > 0 Then
		Set treeNode = tree.Nodes.Add(, , key, text, imageKey, imageKey)
	End If
	If Err.Number <> 0 Then
		Application.Trace "TreeView domain image key failed: " & imageKey & " - " & Err.Description
		Err.Clear
	End If
	If treeNode Is Nothing Then
		Set treeNode = tree.Nodes.Add(, , key, text)
	End If
	If Err.Number <> 0 Then
		Application.Trace "TreeView domain root add failed: " & text & " - " & Err.Description
		Err.Clear
		Exit Function
	End If

	treeNode.Tag = text
	treeNode.ToolTipText = text
	' treeNode.Expanded = True
	Err.Clear
	On Error GoTo 0

	AddTreeRootNode = key

End Function

Sub FillTreeViewFromCatalogXml(tree, xmlText)

	' The domain root, and nothing under it.
	'
	' The catalogue runs to twenty-odd thousand objects, and building them
	' all cost the same whether or not anybody looked: hiding the control
	' and collapsing the nodes only stopped it repainting, never stopped it
	' adding. So it adds one node now, hangs a placeholder under it, and
	' TreeView_Expand puts the real children in when the operator opens it.
	'
	' The expansion lives in the TreeView's own scope, which cannot see any
	' of this - a tag's scopes see only each other. It reads the catalogue
	' again there and finds its place by the path this leaves in Tag.
	HideTree tree

	On Error Resume Next
	tree.Nodes.Clear
	If Err.Number <> 0 Then
		Application.Trace "TreeView clear failed: " & Err.Description
		Err.Clear
	End If
	On Error GoTo 0

	If Len(xmlText) = 0 Then
		Application.Trace "Catalog XML is empty; TreeView was not filled."
		ShowTree tree
		Exit Sub
	End If

	Dim doc
	Set doc = CreateObject("MSXML2.DOMDocument.6.0")
	doc.async = False
	doc.loadXML xmlText

	If doc.parseError.errorCode <> 0 Then
		Application.Trace "Catalog XML parse failed: " & doc.parseError.reason
		ShowTree tree
		Exit Sub
	End If

	Dim sequence
	sequence = 0

	Dim useImages
	useImages = BindTreeImageList(tree)

	Dim domainName
	domainName = ElipseDomainName()
	If Len(domainName) = 0 Then domainName = "Domain"

	Dim rootKey
	rootKey = AddTreeRootNode(tree, domainName, sequence, useImages)

	If Len(rootKey) > 0 Then AddPlaceholder tree, rootKey

	ShowTree tree

	Application.Trace "TreeView opened on the domain root; children load on expand."

End Sub


' The stand-in that gives a node its expander.
'
' A node with no children shows no plus sign, so there would be nothing
' to click and the tree would look like an empty domain. One throwaway
' child is what makes it openable; TreeView_Expand takes it out again and
' puts the real ones in.
'
' Keyed with a p in front, which is how the other scope tells a stand-in
' from a real node without having to store anything.
Sub AddPlaceholder(tree, parentKey)

	On Error Resume Next
	tree.Nodes.Add parentKey, 4, "p" & parentKey, "..."
	Err.Clear
	On Error GoTo 0

End Sub


' Suppressing the repaint, and putting it back. Two lines each, but they
' are called from five places between them and a Visible left False is a
' blank panel, so they are named rather than repeated.
Sub HideTree(tree)

	On Error Resume Next
	tree.Visible = False
	On Error GoTo 0

End Sub


Sub ShowTree(tree)

	On Error Resume Next
	tree.Visible = True
	On Error GoTo 0
	
End Sub

<xatm_config_screens.DomainBrowser:DomainBrowser_OnShow()>
Sub DomainBrowser_OnShow()
	
	FillTree
		
End Sub

<xatm_config_screens.Footer.btnExpandConsole:btnExpandConsole_Click()>
Sub btnExpandConsole_Click()

	Const CONSOLE_SCREEN = "xatm_config_screens.ConsoleLogExpanded"
	Const CONSOLE_FRAME  = "xatm_ConsoleLog"
	Const CONSOLE_TITLE  = "Console"
	Const CONSOLE_LEFT   = 0
	Const CONSOLE_TOP    = 0
	Dim CONSOLE_WIDTH 	 : CONSOLE_WIDTH  = 1366 + 24
	Dim CONSOLE_HEIGHT   : CONSOLE_HEIGHT = 768 + 48

	Dim FLAGS : FLAGS = 1 + 2 + 4 + 8 + 16 + 32 + 64 + 2048

	Dim panel
	Set panel = Nothing

	On Error Resume Next
	Set panel = Application.GetFrame(CONSOLE_FRAME)
	On Error Goto 0

	If panel Is Nothing Then
		MsgBox "The console has nowhere to open. This project has no frame " & _
		       "called " & CONSOLE_FRAME & ". Add one to the frameset.", _
		       vbExclamation, CONSOLE_TITLE
		Exit Sub
	End If

	panel.MoveFrame CONSOLE_LEFT, CONSOLE_TOP, CONSOLE_WIDTH, CONSOLE_HEIGHT
	panel.SetFrameOptions CONSOLE_TITLE, FLAGS
	
	panel.OpenScreen CONSOLE_SCREEN & "?4?0", 0
	
End Sub

<xatm_config_screens.Footer.ListBox:ListBox_MouseUp(Button, Shift, X, Y)>
Sub ListBox_MouseUp(Button, Shift, X, Y)
	
	If Button <> 2 Then
		Exit Sub
	End If
	
	Dim userOption
	userOption = Application.SelectMenu("Clear...")
	
	If userOption = 1 Then
		
		Dim contentTag
		Set contentTag = Nothing
		On Error Resume Next
		Set contentTag = Application.GetObject("xatm_config_data.ConsoleLogEngine.Data.Content")
		On Error Goto 0
		
		If contentTag Is Nothing Then
			MsgBox "No 'Console Log Engine' was bound to '" & Me.PathName & "'.", vbExclamation, "Error"
			Exit Sub
		End If
		
		contentTag.WriteEx Empty
		
	End If
		
End Sub

<xatm_config_screens.Footer.ListBox:ListBox_OnChangedContent()>
Sub ListBox_OnChangedContent()
	
	' Initialize DocString property as timer helper
	DocString = "0"

End Sub

<xatm_config_screens.Footer.ListBox:ListBox_OnStartRunning()>
Sub ListBox_OnStartRunning()

	ListBox_OnChangedContent
	
End Sub

<xatm_config_screens.Footer.ListBox:ListBox_Timer()>
Sub ListBox_Timer()
	
	Const DELAY = 1
	
	If CInt(DocString) >= DELAY Then
		
		' Finish timer
		DocString = "-1"
		
		AddArrayContentToList
		
	Else
	
		DocString = CInt(DocString) + 1
		
	End If
	
End Sub


' The longest item the list box is asked to take.
Const MAX_ITEM_LENGTH = 512


' What the list box will accept: one line, and a bounded one.
'
' The engine writes one physical line per entry now, so this is a belt on
' top of braces. It stays because AddItem refuses an item carrying a line
' break outright, and the loop below stops at the first refusal - so one
' bad entry would not cost one row, it would hide every row after it.
Function SafeItem(v)

    Dim s
    s = ""

    On Error Resume Next
    s = CStr(v)
    On Error Goto 0

    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    If Len(s) > MAX_ITEM_LENGTH Then s = Left(s, MAX_ITEM_LENGTH - 3) & "..."

    SafeItem = s

End Function

Sub AddArrayContentToList()

    Dim contentTag
    Set contentTag = Nothing
    On Error Resume Next
    Set contentTag = Application.GetObject("xatm_config_data.ConsoleLogEngine").Item("Data").Item("Content")
    On Error Goto 0

    If contentTag Is Nothing Then
        MsgBox "No 'Console Log Engine' was bound to '" & Me.PathName & "'.", vbExclamation, "Error"
        Exit Sub
    End If

    If IsEmpty(contentTag.Value) Then
        Clear()
        Exit Sub
    End If

    If Not IsArray(contentTag.Value) Then
        Exit Sub
    End If

    Dim contentArr
    contentArr = contentTag.Value

    Dim arrSize
    arrSize = UBound(contentArr) + 1

    ' If ListCount >= arrSize, buffer is full and shifted — full redraw needed
    If ListCount >= arrSize Then
        Clear()
    End If

    Dim i
    For i = ListCount To UBound(contentArr)
    
    	AddItem SafeItem(contentArr(i)), 0
    	
        'If Application.GetObject("xatm_config_data.ConsoleLogEngine").AutoScrolling Then
            ListIndex = 0
        'End If
    Next
    
End Sub

<xatm_config_screens.Frame.Divisor.Direito.Superior:Superior_OnScreenChange()>
Sub Superior_OnScreenChange()
	
	Dim screenName
	screenName = SplitLink
	
	Dim pos
  	pos = InStr(screenName, "?")
  	If pos > 0 Then screenName = Left(screenName, pos - 1)
  	
  	pos = InStrRev(screenName, ".")
  	If pos > 0 Then screenName = Mid(screenName, pos + 1)
  	
	Application.GetObject("xatm_config_data.CurrentSreenName").WriteEx UCase(screenName)
		
End Sub

<xatm_config_screens.Frame.Divisor.Direito.Superior:Superior_OnStartRunning()>
Sub Superior_OnStartRunning()

	Superior_OnScreenChange

End Sub

<xatm_config_screens.Menu.btnClose:btnClose_Click()>
Sub btnClose_Click()
	
	Screen.Close(0)
	
End Sub

<xatm_config_screens.Menu.btnConfig:btnConfig_Click()>
Sub btnConfig_Click()
	
	Dim fr
	Set fr = Application.GetFrame("Superior")
	fr.OpenScreen "xatm_config_screens.Config?4?0", 0
		
End Sub

<xatm_config_screens.Menu.btnMinimize:btnMinimize_Click()>
Sub btnMinimize_Click()
	Me.Frame.MinimizeFrame()	
End Sub

<xatm_config_screens.Menu.btnRASEAT:btnRASEAT_Click()>
Sub btnRASEAT_Click()

	Dim autos
	Set autos = Application.GetObject("XATM_Data.Automation")

	' There is one of these in a station, so there is no list to order and
	' no instance to choose between - unlike the transfers, which get a
	' submenu each.
	Dim obj, target
	Set target = Nothing

	For Each obj In autos
		If TypeName(obj) = "xatm_RASEAT" Then
			Set target = obj
			Exit For
		End If
	Next

	If target Is Nothing Then
		MsgBox "No RASEAT automation found!"
		Exit Sub
	End If

	' --- the menu -------------------------------------------------------
	Dim menu
	menu = target.Name & "{Force reclosing||" & _
	       IIf(target.OperatorBlock, "*", "") & "Operator Block|Reset}"

	Dim lOption
	lOption = Application.SelectMenu(menu)
	If lOption <= 0 Then Exit Sub

	Select Case lOption

		Case 1

			' Asked for the way a transformer's lockout relay asks, so every
			' gate that Start keeps is kept for this too - enabled, not
			' already running, the three blocks, the preconditions, no other
			' automation in progress. A refusal is logged there and says why.
			'
			' Zero as the asker, because no relay did this. What the
			' sequence acts on is read in Step 0 from the field either way,
			' so a forced start on a healthy station finds nothing to
			' isolate and goes straight to the reclose.
			If MsgBox("Force a reclosing on " & target.Name & "?" & vbCrLf & vbCrLf & _
			          "The switchyard is operated for real.", _
			          vbYesNo + vbExclamation + vbDefaultButton2, "Confirm") = vbYes Then

				target.Item("Commands").Item("Start").WriteEx 0

			End If

		Case 2

			' Operator Block toggle
			target.OperatorBlock = Not target.OperatorBlock

		Case 3

			' Reset - clears the latched step failures, the general block
			' and the result of the last run.
			target.Item("Commands").Item("Reset").WriteEx True

	End Select
	
End Sub

<xatm_config_screens.Menu.btnSimulationMode:btnSimulationMode_Click()>
Sub btnSimulationMode_Click()
	
	Dim tag
	Set tag = Application.GetObject("xatm_config_data.SimulationMode")
	tag.WriteEx (Not CBool(tag.Value))
		
End Sub

<xatm_config_screens.Menu.btnSingleLineDiagram:btnSingleLineDiagram_Click()>
Sub btnSingleLineDiagram_Click()
	
	Dim fr
	Set fr = Application.GetFrame("Superior")
	fr.OpenScreen "xatm_config_screens.SingleLineDiagram?4?0", 0
		
End Sub

<xatm_config_screens.Menu.btnTA:btnTA_Click()>
Sub btnTA_Click()

	Dim autos
	Set autos = Application.GetObject("XATM_Data.Automation")

	' --- Collect every automatic transfer -----------------------------
	Dim obj, total
	total = 0
	For Each obj In autos
		If TypeName(obj) = "xatm_TA" Then total = total + 1
	Next

	If total = 0 Then
		MsgBox "No automatic transfer found!"
		Exit Sub
	End If

	Dim instances()
	ReDim instances(total - 1)
	Dim k
	k = 0
	For Each obj In autos
		If TypeName(obj) = "xatm_TA" Then
			Set instances(k) = obj
			k = k + 1
		End If
	Next

	' --- Order by the transformer each one answers for ----------------
	Dim i, j, tmp
	For i = 0 To total - 2
		For j = 0 To total - 2 - i
			If MenuKey(instances(j)) > MenuKey(instances(j + 1)) Then
				Set tmp = instances(j)
				Set instances(j) = instances(j + 1)
				Set instances(j + 1) = tmp
			End If
		Next
	Next

	' --- Build the menu, recording what each option does --------------
	'
	' Every submenu here is the same three entries, so the mapping could be
	' a division. It is recorded instead, the way the manual menu records
	' it: an instance bound to nothing gets a different submenu, and one
	' exception is all it takes for a division to point at the wrong
	' automation.
	'
	' Separators cost nothing: an empty entry takes no option number.
	Dim menu, actions, tr
	Set actions = CreateObject("Scripting.Dictionary")
	menu = ""

	For i = 0 To total - 1

		If i > 0 Then menu = menu & "|"

		Set tr = TransformerOf(instances(i))

		If tr Is Nothing Then

			' Bound to no transformer. Shown so that it is visibly there to
			' be fixed rather than quietly absent from a menu that looks
			' complete - a TA answers one particular transformer's trip,
			' and one bound to nothing answers nothing.
			Remember actions, i, "NONE"
			menu = menu & instances(i).Name & "{not configured}"

		Else

			menu = menu & tr.Name & "{"

			Remember actions, i, "TA"
			menu = menu & "TA||"

			Remember actions, i, "BLOCK"
			menu = menu & IIf(instances(i).OperatorBlock, "*", "") & "Operator Block|"

			Remember actions, i, "RESET"
			menu = menu & "Reset}"

		End If

	Next

	Dim lOption
	lOption = Application.SelectMenu(menu)
	If lOption <= 0 Then Exit Sub

	If Not actions.Exists(CStr(lOption)) Then
		MsgBox "The menu and the automation list disagree - nothing sent.", _
		       vbExclamation, "Error"
		Exit Sub
	End If

	' --- Do what the chosen option said -------------------------------
	Dim parts
	parts = Split(actions(CStr(lOption)), "|")

	Dim target
	Set target = instances(CLng(parts(0)))

	Select Case parts(1)

		Case "NONE"

			MsgBox target.Name & " is bound to no transformer, so there is nothing " & _
			       "it can be asked to run.", vbExclamation, "Not configured"

		Case "BLOCK"

			target.OperatorBlock = Not target.OperatorBlock

		Case "RESET"

			target.Item("Commands").Item("Reset").WriteEx True

		Case Else

			' Forced from here and asked for by nobody in service: a TA
			' answers a trip, through the transformer's own trigger. The
			' entry exists because somebody commissioning the station has
			' to be able to make it run.
			'
			' Written with two fields on purpose. Start reads the field for
			' the contingency, the way it does for a protection trip, so
			' there is nothing here for an operator to choose beyond which
			' transformer.
			If MsgBox("Force TA on " & target.Transformer.Name & "?", _
			          vbYesNo + vbQuestion, "Confirm") = vbYes Then
				target.Item("Commands").Item("Start").WriteEx "TA:" & target.Transformer.Id
			End If

	End Select

End Sub


' Records what the next option number is to do: which instance, and what to
' ask of it.
'
' Keyed by the count so far plus one, because SelectMenu numbers the entries
' it returns from one and passes over the empty ones. Called immediately
' before the label it belongs to is appended, which is what keeps the two in
' step. The key is written as text on both sides: a Dictionary keyed on a
' number answers Exists only for a lookup of the same subtype, and what
' SelectMenu hands back is not this script's to choose.
'
' No contingency argument, unlike the manual menu's copy. Nobody declares
' one for a TA - Start reads the field.
Sub Remember(actions, idx, kind)

	actions.Add CStr(actions.Count + 1), idx & "|" & kind

End Sub


' The transformer an instance answers for, or Nothing when it is bound to
' none. Its own copy, because one E3 object cannot call a procedure in
' another's scope.
Function TransformerOf(obj)

	Set TransformerOf = Nothing

	On Error Resume Next
	Set TransformerOf = obj.Transformer
	On Error Goto 0

End Function


' What the menu is ordered by: the transformer each one answers for, and
' anything bound to nothing after them.
'
' A string and not a number, so that the unbound ones sort last without
' having to stand for a transformer Id they do not have.
Function MenuKey(obj)

	Dim tr
	Set tr = TransformerOf(obj)

	If tr Is Nothing Then
		MenuKey = "2"
		Exit Function
	End If

	Dim id
	id = 0

	On Error Resume Next
	id = CLng(tr.Id)
	On Error Goto 0

	MenuKey = "1" & Right("00000" & id, 5)

End Function

Sub EndOfScope()
	
End Sub

<xatm_config_screens.Menu.btnTMTNM:btnTMTNM_Click()>
Sub btnTMTNM_Click()

	Dim autos
	Set autos = Application.GetObject("XATM_Data.Automation")

	' --- Collect every automation instance ----------------------------
	Dim obj, total
	total = 0
	For Each obj In autos
		If TypeName(obj) = "xatm_TMTNM" Then total = total + 1
	Next

	If total = 0 Then
		MsgBox "No automation found!"
		Exit Sub
	End If

	Dim instances()
	ReDim instances(total - 1)
	Dim k
	k = 0
	For Each obj In autos
		If TypeName(obj) = "xatm_TMTNM" Then
			Set instances(k) = obj
			k = k + 1
		End If
	Next

	' --- Order: transformers by Id, then the busbar pairs -------------
	Dim i, j, tmp
	For i = 0 To total - 2
		For j = 0 To total - 2 - i
			If MenuKey(instances(j)) > MenuKey(instances(j + 1)) Then
				Set tmp = instances(j)
				Set instances(j) = instances(j + 1)
				Set instances(j + 1) = tmp
			End If
		Next
	Next

	' --- Build the menu, recording what each option does --------------
	'
	' Recorded as it is built, rather than worked back out afterwards.
	' The mapping used to be a division, which needed every instance to
	' contribute the same number of options - true while they were all
	' transformer automations, and false the moment one of them moves a
	' busbar instead: that one offers two maneuvers and no contingencies
	' where its neighbours offer TM and NM for every transformer, plus TA.
	'
	' Building the label and registering the command in one pass is also
	' what keeps them honest. Two passes that agreed today would be two
	' things to keep in step, and the failure would be silent - a label
	' naming one maneuver and the command sending another.
	'
	' Separators cost nothing: an empty entry takes no option number, which
	' is why the blocks can be kept apart freely.
	Dim menu, actions
	Set actions = CreateObject("Scripting.Dictionary")
	menu = ""

	For i = 0 To total - 1

		If i > 0 Then menu = menu & "|"

		If BusbarPairOf(instances(i)) <> "" Then

			AddBusbarMenu instances, i, menu, actions

		ElseIf Not TransformerOf(instances(i)) Is Nothing Then

			AddTransformerMenu instances, total, i, menu, actions

		Else

			' Bound to no transformer and moving no busbar. Shown, so that
			' it is visibly there to be fixed rather than quietly absent
			' from a menu that looks complete, and with nothing under it
			' but the reason - there is nothing it could be asked to do.
			Remember actions, i, "NONE", 0
			menu = menu & instances(i).Name & "{not configured}"

		End If

	Next

	Dim lOption
	lOption = Application.SelectMenu(menu)
	If lOption <= 0 Then Exit Sub

	If Not actions.Exists(CStr(lOption)) Then
		MsgBox "The menu and the automation list disagree - nothing sent.", _
		       vbExclamation, "Error"
		Exit Sub
	End If

	' --- Do what the chosen option said -------------------------------
	Dim parts
	parts = Split(actions(CStr(lOption)), "|")

	Dim target
	Set target = instances(CLng(parts(0)))

	Dim kind, impedeId
	kind     = parts(1)
	impedeId = CLng(parts(2))

	Select Case kind

		Case "NONE"

			MsgBox target.Name & " is bound to no transformer and names no " & _
			       "busbar pair, so there is nothing it can be asked to run.", _
			       vbExclamation, "Not configured"

		Case "BLOCK"

			target.OperatorBlock = Not target.OperatorBlock

		Case "RESET"

			target.Item("Commands").Item("Reset").WriteEx True

		Case Else

			SendManeuver target, instances, total, kind, impedeId

	End Select

End Sub


' A TM or an NM, asked for the way the instance it is going to expects.
'
' A busbar automation is bound to nothing and takes no contingency, so its
' request names neither: the pair it moves is a property of its own, and
' Start reads it there. That is the same string StartMode writes when the
' command comes off a tag instead of this menu.
Sub SendManeuver(target, instances, total, mode, impedeId)

	Dim pairMoved
	pairMoved = BusbarPairOf(target)

	If pairMoved <> "" Then

		If MsgBox("Force " & mode & " " & pairMoved & "?", _
		          vbYesNo + vbQuestion, "Confirm") = vbYes Then
			target.Item("Commands").Item("Start").WriteEx mode & ":0:0"
		End If

		Exit Sub

	End If

	Dim tr
	Set tr = target.Transformer

	Dim question
	If impedeId = 0 Then
		question = "Force " & mode & " on " & tr.Name & "?"
	Else
		question = "Force " & mode & " on " & tr.Name & _
		           " with " & TransformerNameById(instances, total, impedeId) & " out of service?"
	End If

	If MsgBox(question, vbYesNo + vbQuestion, "Confirm") = vbYes Then
		target.Item("Commands").Item("Start").WriteEx _
			mode & ":" & tr.Id & ":" & impedeId
	End If

End Sub


' One manual automation's submenu: both maneuvers with every contingency,
' then the two things done to the automation itself.
'
' No TA. That one is a class of its own, and a button of its own.
Sub AddTransformerMenu(instances, total, idx, menu, actions)

	menu = menu & instances(idx).Transformer.Name & "{"

	menu = menu & ModeBlock("TM", instances, total, idx, actions) & "||"
	menu = menu & ModeBlock("NM", instances, total, idx, actions) & "||"


	Remember actions, idx, "BLOCK", 0
	menu = menu & IIf(instances(idx).OperatorBlock, "*", "") & "Operator Block|"

	Remember actions, idx, "RESET", 0
	menu = menu & "Reset}"

End Sub


' One busbar automation's submenu, titled with the pair it moves.
'
' Two maneuvers and no more. TM hands the busbar to the far side of the
' ring and NM brings it back; there is no contingency to offer, because the
' path does not vary, and no TA, because that one answers a trip and is
' never asked for by a person.
Sub AddBusbarMenu(instances, idx, menu, actions)

	menu = menu & BusbarPairOf(instances(idx)) & "{"

	Remember actions, idx, "TM", 0
	menu = menu & "TM|"

	Remember actions, idx, "NM", 0
	menu = menu & "NM||"

	Remember actions, idx, "BLOCK", 0
	menu = menu & IIf(instances(idx).OperatorBlock, "*", "") & "Operator Block|"

	Remember actions, idx, "RESET", 0
	menu = menu & "Reset}"

End Sub


' What one mode contributes to one transformer's submenu: the plain
' maneuver, then the same maneuver with each other transformer out.
'
' Labelled the way the spec labels them - "TM-TR3" on TR4's submenu is
' section 1.4.2.23's "TM TR4-TR3" - so the menu, the CommandStartTM<n00>
' tags and the spec all name the same maneuver.
Function ModeBlock(mode, instances, total, idx, actions)

	Remember actions, idx, mode, 0

	Dim out, n, other
	out = mode

	For n = 1 To total - 1

		Set other = OtherTransformer(instances, total, idx, n)

		If Not other Is Nothing Then
			Remember actions, idx, mode, other.Id
			out = out & "|" & mode & "-" & other.Name
		End If

	Next

	ModeBlock = out

End Function


' Records what the next option number is to do: which instance, what to
' ask of it, and which transformer it is to treat as out of service.
'
' Keyed by the count so far plus one, because SelectMenu numbers the
' entries it returns from one and passes over the empty ones. Called
' immediately before the label it belongs to is appended, which is what
' keeps the two in step.
'
' The key is written as text on both sides. A Dictionary keyed on a
' number answers Exists only for a lookup of the same subtype, and what
' SelectMenu hands back is not this script's to choose.
Sub Remember(actions, idx, kind, impedeId)

	actions.Add CStr(actions.Count + 1), idx & "|" & kind & "|" & impedeId

End Sub


' The nth transformer that is not this instance's own, counting from 1 in
' the ascending order the list was sorted into.
'
' One walk shared by the menu builder and the option mapper. Two walks
' that agreed today would be two things to keep in step, and the failure
' would be silent: a label naming one transformer and the command naming
' another.
'
' Busbar automations are passed over rather than counted. They are bound
' to no transformer, so there is nothing about them to declare out of
' service and nothing to offer as a contingency.
Function OtherTransformer(instances, total, idx, n)

	Set OtherTransformer = Nothing

	Dim i, seen, tr
	seen = 0

	For i = 0 To total - 1

		If i <> idx Then

			Set tr = TransformerOf(instances(i))

			If Not tr Is Nothing Then

				seen = seen + 1

				If seen = n Then
					Set OtherTransformer = tr
					Exit Function
				End If

			End If

		End If

	Next

End Function


' The name of the transformer carrying an Id, among the ones on this menu,
' and the Id itself when no instance is bound to it.
Function TransformerNameById(instances, total, id)

	TransformerNameById = "TR" & id

	Dim i, tr
	For i = 0 To total - 1

		Set tr = TransformerOf(instances(i))

		If Not tr Is Nothing Then
			If tr.Id = id Then
				TransformerNameById = tr.Name
				Exit Function
			End If
		End If

	Next

End Function


' The transformer an instance is bound to, or Nothing - which a busbar
' automation always is, and a half-configured one may be.
Function TransformerOf(obj)

	Set TransformerOf = Nothing

	On Error Resume Next
	Set TransformerOf = obj.Transformer
	On Error Goto 0

End Function


' The pair of busbars an instance moves, or "" for one that moves none -
' and for a library too old to have the property at all.
Function BusbarPairOf(obj)

	BusbarPairOf = ""

	On Error Resume Next
	BusbarPairOf = Trim(obj.BusbarPair & "")
	On Error Goto 0

End Function


' What the menu is ordered by: transformer automations first and by Id,
' then the busbar ones by the pair they move.
'
' A string and not a number, because a busbar automation has no Id to sort
' on - and because the two families should not interleave, which the
' leading digit settles.
Function MenuKey(obj)

	Dim pairMoved
	pairMoved = BusbarPairOf(obj)

	If pairMoved <> "" Then
		MenuKey = "2" & UCase(pairMoved)
		Exit Function
	End If

	Dim id
	id = 0

	Dim tr
	Set tr = TransformerOf(obj)

	If Not tr Is Nothing Then
		On Error Resume Next
		id = CLng(tr.Id)
		On Error Goto 0
	End If

	MenuKey = "1" & Right("00000" & id, 5)

End Function

Sub EndOfScope()	
	
End Sub

<xatm_config_screens.Menu.grpDemo.lblDemo:lblDemo_OnStartRunning()>
Sub lblDemo_OnStartRunning()
	
	' Blank in a runtime build, so the label simply is not there. Anything
	' else - a greyed word, a different colour - is something an operator
	' learns to stop seeing.
	Value = ""

	If IsDemoLibrary() Then Value = "DEMO"

End Sub


' Whether the library driving this station is the demo build.
'
' Asked of the equipment rather than of a singleton somewhere: every
' breaker and disconnector carries a Build, all of them come from the same
' library file, so the first one found answers for the rest and there is
' nothing to keep in agreement.
'
' No device carrying a Build means either a substation that has not been
' built yet or a library too old to have the class. Neither is a demo, and
' neither is this label's business to announce - the command gate is what
' actually withholds anything, and it fails closed on its own.
Function IsDemoLibrary()

	IsDemoLibrary = False

	Dim substation
	Set substation = Nothing

	On Error Resume Next
	Set substation = Application.GetObject("XATM_Data.Substation")
	On Error Goto 0

	If substation Is Nothing Then Exit Function

	Dim found
	found = Empty

	FindBuild substation, found

	If Not IsEmpty(found) Then IsDemoLibrary = CBool(found)

End Function


' Walks the substation until something answers, then stops.
'
' found stays Empty while nothing has, which is what lets the caller tell
' "no device carries a Build" from "a device carries one and it says
' runtime" - the two look alike at a glance and mean quite different
' things.
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

		' A folder, or a device from a library that has no Build - either
		' way, look inside it. A tag refuses to be walked, and the error
		' that raises is how the walk finds the leaves.
		On Error Resume Next
		FindBuild obj, found
		On Error Goto 0

	Next
		
End Sub

<xatm_config_screens.Menu.grpDemo.objArea:objArea_Click()>
Sub objArea_Click()

	If HELP_LANG = "pt-BR" Then
		MsgBox DemoTextPt(), vbInformation, "Versão de demonstração"
	Else
		MsgBox DemoTextEn(), vbInformation, "Demo build"
	End If

End Sub


' Which language the panel speaks. Written out again here because one E3
' object cannot read another's constants - the same copy the Config screen
' keeps for the property help.
Const HELP_LANG = "pt-BR"


' What a demo build withholds, said in full.
'
' Functions and not constants: VBScript will not build a Const out of an
' expression, so the line breaks between the paragraphs have to be joined
' where they can be - inside a body.
'
' Worth spelling out rather than shortening: somebody clicking this label
' is asking why the switchyard did not move, and the useful answer is three
' things - that simulation still works, that the sequence they just watched
' was real logic on simulated plant, and what operating for real would
' take. "This is a demo" leaves them with the same question.
Function DemoTextPt()

	DemoTextPt = "Esta cópia da biblioteca de automatismos é uma versão de demonstração." & vbCrLf & vbCrLf & _
	             "Os comandos alcançam apenas equipamentos em modo de simulação. Nada é enviado ao pátio: " & _
	             "um disjuntor simulado se move como se moveria em serviço, e uma sequência inteira pode ser " & _
	             "executada do início ao fim - mas nenhum equipamento real é operado." & vbCrLf & vbCrLf & _
	             "Operar equipamento real exige a versão de runtime, licenciada para executar somente em " & _
	             "hardkey que contenha essa licença."

End Function

Function DemoTextEn()

	DemoTextEn = "This copy of the automation library is a demo build." & vbCrLf & vbCrLf & _
	             "Commands reach simulated equipment only. Nothing is sent to the switchyard: a simulated " & _
	             "breaker moves as it would in service, and a whole sequence can be run end to end - but no " & _
	             "real equipment is operated." & vbCrLf & vbCrLf & _
	             "Operating real equipment needs the runtime build, licensed to run only on a hardkey " & _
	             "carrying that licence."

End Function


' A scope may not end on a Function - E3 takes the script without complaint
' and then behaves as though the last one were not there.
Sub Foo()

End Sub

<xatm_config_screens.Menu.lblVersion:lblVersion_Click()>
Sub lblVersion_Click()

	' What is installed, and then the notes if they are wanted.
	'
	' The version is read off the object rather than written down here: it
	' comes from a constant inside xatm_lib.lib and is published as a
	' property, so a screen keeping its own copy would be one more thing to
	' remember on a release. The demo flag is not put in the dialogue at all
	' - the blinking label on this same screen is what says that, and says
	' it without being asked.
	Dim versionObject
	Set versionObject = Nothing

	On Error Resume Next
	Set versionObject = Application.GetObject(VERSION_OBJECT)
	On Error Goto 0

	If versionObject Is Nothing Then

		MsgBox MissingText(), vbExclamation, DialogTitle()
		Exit Sub

	End If

	If MsgBox(AskText(VersionOf(versionObject)), vbInformation + vbYesNo, DialogTitle()) <> vbYes Then Exit Sub

	ShowReport versionObject

End Sub


' Where the one version object lives. One for the library and not one per
' device: what it publishes is true of the file rather than of any piece of
' equipment, which is the opposite of how the demo flag is carried.
Const VERSION_OBJECT = "xatm_config_data.Version"


' Which language the panel speaks. Written out again here because one E3
' object cannot read another's constants - the same copy the Config screen
' and the demo label each keep.
Const HELP_LANG = "pt-BR"


' What is installed, written out and put in front of whoever asked for it.
'
' Done here in the screen rather than in the library object, which is where
' it started. An E3 data server runs as a service: its %TEMP% is
' C:\Windows\Temp and not the operator's own, and a process in session 0
' has no desktop to put a Notepad window on - it starts and nobody ever
' sees it. The viewer runs as the person at the keyboard, so both problems
' go away by moving the work rather than by working around them.
'
' A dictionary rather than a built-up string: the report is a list of
' facts, and keeping it as one until the moment it is rendered means the
' order is the order it was filled in, and adding a line is one line.
Sub ShowReport(obj)

	Dim info
	Set info = CreateObject("Scripting.Dictionary")

	info.Add "Library",       "xatm_lib"
	info.Add "Version",       VersionOf(obj)
	info.Add "Edition",       EditionText()
	info.Add "Generated",     Now
	info.Add "Release notes", NotesOf(obj)

	Dim failed
	failed = ""

	Dim path
	path = WriteReport(info, failed)

	If path = "" Then
		MsgBox WriteFailedText(failed), vbExclamation, DialogTitle()
		Exit Sub
	End If

	failed = OpenFile(path)

	If failed <> "" Then
		MsgBox OpenFailedText(path, failed), vbExclamation, DialogTitle()
	End If

End Sub


' Renders the dictionary into %TEMP% and returns where it went - "" when it
' went nowhere, with why in the second argument.
Function WriteReport(info, ByRef failed)

	WriteReport = ""
	failed      = ""

	Dim fso
	Set fso = Nothing

	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	On Error Goto 0

	If fso Is Nothing Then
		failed = "Scripting.FileSystemObject"
		Exit Function
	End If

	Dim folder
	folder = ReportFolder(fso)

	If Trim(folder & "") = "" Then
		failed = "%TEMP%"
		Exit Function
	End If

	Dim path
	path = fso.BuildPath(folder, REPORT_FILE_NAME)

	On Error Resume Next
	Err.Clear

	Dim stream
	Set stream = fso.CreateTextFile(path, True)
	stream.Write Rendered(info)
	stream.Close

	If Err.Number <> 0 Then failed = Err.Description
	On Error Goto 0

	If failed <> "" Then Exit Function

	WriteReport = path

End Function


' The operator's own temporary folder, and the scripting object's idea of
' one only if %TEMP% will not answer.
'
' Read from this process on purpose. In the viewer that is the account of
' the person who logged in, which is the whole reason the report is written
' here: the data server's answer to the same question is C:\Windows\Temp,
' a folder they cannot so much as cd into.
Function ReportFolder(fso)

	ReportFolder = ""

	Dim shell
	Set shell = NewShell()

	Dim expanded
	expanded = ""

	If Not shell Is Nothing Then

		On Error Resume Next
		expanded = shell.ExpandEnvironmentStrings("%TEMP%")
		On Error Goto 0

	End If

	' An unset variable comes back as the name it was written as, so an
	' answer still carrying a per cent sign is not a folder.
	If InStr(expanded & "", "%") = 0 And Trim(expanded & "") <> "" Then

		If fso.FolderExists(expanded) Then
			ReportFolder = expanded
			Exit Function
		End If

	End If

	On Error Resume Next
	ReportFolder = fso.GetSpecialFolder(TEMPORARY_FOLDER).Path
	On Error Goto 0

End Function


' The dictionary as text: one fact per line, the keys padded so the values
' line up and the thing can be read down rather than across.
Function Rendered(info)

	Dim width, k
	width = 0

	For Each k In info.Keys
		If Len(k) > width Then width = Len(k)
	Next

	Dim out
	out = "xatm_lib" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf

	For Each k In info.Keys
		out = out & k & String(width - Len(k) + 2, " ") & info(k) & vbCrLf
	Next

	Rendered = out

End Function


' A WScript.Shell, or Nothing. Two things want one - the folder to write in
' and the editor to open it - and neither is worth failing the whole report
' over on its own.
Function NewShell()

	Set NewShell = Nothing

	On Error Resume Next
	Set NewShell = CreateObject("WScript.Shell")
	On Error Goto 0

End Function


' Puts the report on the screen. Returns "" when it opened, and why not
' when it did not.
'
' Notepad by name rather than handing the path to whatever the machine
' opens a .txt with: it is on every Windows there is and takes a path and
' nothing else.
Function OpenFile(path)

	OpenFile = ""

	Dim shell
	Set shell = NewShell()

	If shell Is Nothing Then
		OpenFile = "WScript.Shell"
		Exit Function
	End If

	On Error Resume Next
	Err.Clear
	shell.Run "notepad.exe """ & path & """", 1, False
	If Err.Number <> 0 Then OpenFile = Err.Description
	On Error Goto 0

End Function


' The version as the object publishes it, falling back to the three numbers
' when the written form is blank. Both are set at start from the same
' constant, so either one answers.
Function VersionOf(obj)

	Dim written
	written = ""

	On Error Resume Next
	written = Trim(obj.Version & "")
	On Error Goto 0

	If written <> "" Then
		VersionOf = written
		Exit Function
	End If

	Dim major, minor, patch
	major = 0
	minor = 0
	patch = 0

	On Error Resume Next
	major = CLng(obj.Major)
	minor = CLng(obj.Minor)
	patch = CLng(obj.Patch)
	On Error Goto 0

	VersionOf = major & "." & minor & "." & patch

End Function


Function NotesOf(obj)

	NotesOf = ""

	On Error Resume Next
	NotesOf = Trim(obj.ReleaseNotes & "")
	On Error Goto 0

	If NotesOf = "" Then NotesOf = "(none published)"

End Function


' Which build is driving this station, for the report to carry - and said
' as unknown rather than guessed at when nothing answers.
'
' Asked of the equipment: every breaker and disconnector carries an
' xatm_Build, all of them come from the same library file, so the first one
' found answers for the rest. The walk is written out again here because
' one E3 object cannot call a procedure in another's scope - the demo label
' on this same screen keeps its own copy for the same reason.
Function EditionText()

	EditionText = "unknown - no device has reported its build"

	Dim substation
	Set substation = Nothing

	On Error Resume Next
	Set substation = Application.GetObject(SUBSTATION)
	On Error Goto 0

	If substation Is Nothing Then Exit Function

	Dim found
	found = ""

	FindEdition substation, found

	If found <> "" Then EditionText = found

End Function

Sub FindEdition(folder, ByRef found)

	If found <> "" Then Exit Sub

	Dim obj
	For Each obj In folder

		If found <> "" Then Exit Sub

		Dim build
		Set build = Nothing

		On Error Resume Next
		Set build = obj.Item("Build")
		On Error Goto 0

		If Not build Is Nothing Then

			On Error Resume Next
			found = Trim(build.Edition & "")
			On Error Goto 0

			If found <> "" Then Exit Sub

		End If

		On Error Resume Next
		FindEdition obj, found
		On Error Goto 0

	Next

End Sub


Function DialogTitle()

	If HELP_LANG = "pt-BR" Then
		DialogTitle = "Versão da biblioteca"
	Else
		DialogTitle = "Library version"
	End If

End Function


' What the dialogue says, and the question it ends on.
'
' The notes themselves are not put in here. They are a page rather than a
' line, they grow with every release, and a MsgBox that has to be scrolled
' is a worse way to read them than the text file this offers instead.
Function AskText(versionText)

	If HELP_LANG = "pt-BR" Then

		AskText = "Biblioteca de automatismos xatm_lib, versão " & versionText & "." & vbCrLf & vbCrLf & _
		          "Deseja abrir as notas desta versão em um arquivo de texto?"

	Else

		AskText = "Automation library xatm_lib, version " & versionText & "." & vbCrLf & vbCrLf & _
		          "Open the release notes in a text file?"

	End If

End Function


' Said plainly rather than as a failed read. The object is created by hand,
' once, and a station that never had it done reads exactly like one where
' the name was mistyped - so the name is what the message carries.
Function MissingText()

	If HELP_LANG = "pt-BR" Then

		MissingText = "Não foi possível ler a versão da biblioteca." & vbCrLf & vbCrLf & _
		              "O objeto " & VERSION_OBJECT & " não foi encontrado no projeto."

	Else

		MissingText = "The library version could not be read." & vbCrLf & vbCrLf & _
		              "The object " & VERSION_OBJECT & " was not found in the project."

	End If

End Function


Function WriteFailedText(reason)

	If HELP_LANG = "pt-BR" Then

		WriteFailedText = "Não foi possível gravar o arquivo com as notas desta versão." & vbCrLf & vbCrLf & _
		                  reason

	Else

		WriteFailedText = "The file with the release notes could not be written." & vbCrLf & vbCrLf & _
		                  reason

	End If

End Function


' The path is given even so. The file is written and only the opening
' failed, so telling somebody where it is leaves them able to go and read
' it themselves.
Function OpenFailedText(path, reason)

	If HELP_LANG = "pt-BR" Then

		OpenFailedText = "As notas foram gravadas, mas o Bloco de Notas não abriu." & vbCrLf & vbCrLf & _
		                 path & vbCrLf & vbCrLf & reason

	Else

		OpenFailedText = "The notes were written, but Notepad would not open." & vbCrLf & vbCrLf & _
		                 path & vbCrLf & vbCrLf & reason

	End If

End Function


' Where the equipment lives, Scripting's own number for the temporary
' folder that ReportFolder falls back on, and what the report is called
' once it is written.
Const SUBSTATION        = "XATM_Data.Substation"
Const TEMPORARY_FOLDER  = 2
Const REPORT_FILE_NAME  = "xatm_lib.txt"


' A scope may not end on a Function - E3 takes the script without complaint
' and then behaves as though the last one were not there.
Sub EndOfScope()

End Sub

<xatm_config_screens.SingleLineDiagram:SingleLineDiagram_OnPreShow(Arg)>
Sub SingleLineDiagram_OnPreShow(Arg)
	
	Dim devices
	Set devices = BuildDeviceDictionary()
	
	Dim obj
	
	For Each obj In Screen
		
		If UCase(Left(TypeName(obj), 5)) = "XATM_" Then
			
			Dim id
			id = -1
			On Error Resume Next
			id = obj.Id
			On Error Goto 0
			
			If devices.Exists(id) Then
				
				Dim device
                Set device = devices(id)
                
                obj.Source = device.PathName
                
            End If
			
			
		End If
		
	Next
	
	' TODO: make dynamic
	'On Error Resume Next
	'Item("xatm_TMTNMStatus1").Source = Application.GetObject("XATM_Data.Automation.TMTNM1").PathName
	
End Sub

Function BuildDeviceDictionary()

    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")

    Dim substation
    Set substation = Application.GetObject("XATM_Data.Substation")

    ScanFolder substation, dict

    Set BuildDeviceDictionary = dict

End Function

Sub ScanFolder(ByVal folder, ByRef dict)

    Dim item
    For Each item In folder

        Select Case UCASE(TypeName(item))

            Case "XATM_BREAKER", "XATM_TRANSFORMER", "XATM_DISCONNECTOR"

                If dict.Exists(item.Id) Then
                    'WriteLog "BuildDeviceDictionary: duplicate Id " & item.Id & " at " & item.PathName
                Else
                    dict.Add item.Id, item
                End If

            Case Else

                ' Recurse into sub-folders; non-iterable items are silently skipped
                On Error Resume Next
                ScanFolder item, dict
                On Error GoTo 0

        End Select

    Next
    
End Sub

