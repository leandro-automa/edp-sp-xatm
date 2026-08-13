-----------------------
Documentação de Scripts
-----------------------
XATM_CONFIG (C:\ProjDev\edp_sp\xatm_config.prj)
Wed Aug 12 11:18:24 2026
-----------------------

<xatm_config_data.Catalog.XMLBuilderAfterDelay:XMLBuilderAfterDelay_Functions()>
Sub XMLBuilderAfterDelay_Functions()
End Sub

Sub ExportElipseCatalog()

	Dim outputPath
	outputPath = "C:\temp\elipse-catalog.xml"

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

	Application.Trace "Elipse catalog export started: " & outputPath
	
	WalkEnumerable doc, root, Application, "Application", 0
	
	'doc.save outputPath
	
	Parent.Item("XMLContent").WriteEx doc.xml
	
	Application.Trace "Elipse catalog export finished: " & outputPath

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


' Writes one property, and only when the document carries a value for it.
'
' A property with no value attribute is unset in the document, and the
' export writes no value for an IOTag or for a reading at all - so absent
' means leave it alone, not clear it.
Sub WriteProperty(obj, property, report, where)

	Dim a
	Set a = property.getAttributeNode("value")
	If a Is Nothing Then Exit Sub

	WriteValue obj, property.getAttribute("name"), property.getAttribute("type"), _
	           a.value, report, where

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


' What an object answers to across a rename - its Id where it has one,
' and its name where it has not. A BTC carries no Id, and SyncAutomation
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

	If IsEmpty(value) Or IsNull(value) Then
		property.removeAttribute "source"
	ElseIf CStr(value) = "" Then
		property.removeAttribute "source"
	Else
		property.setAttribute "source", CStr(value)
	End If

End Function


' Renames the object and rebuilds the path it is known by, the name being
' the last piece of it.
'
' TODO: a property somewhere else in the document may hold the old path -
' a BTC's Transformer does - and those are not followed yet.
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

	SetMember alarm, "DigitalLimit",             p.AlarmLimit(),      where, problem
	SetMember alarm, "DigitalMessageText",       p.AlarmActiveText(), where, problem
	SetMember alarm, "DigitalReturnMessageText", p.AlarmNormalText(), where, problem
	SetMember alarm, "DigitalSeverity",          p.AlarmSeverity,     where, problem
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

	IsDiscrete = (LCase(p.DataType & "") = "boolean") And p.IsInterfaced()

End Function


' A command the interface carries: an InternalTag property a manifest
' interfaces. RebuildInterface leaves these unlinked and points the
' XObject's property at them, so writing one is how a maneuver is asked
' for - and level 3 asks through the driver.
Function IsCommand(p)

	IsCommand = (LCase(p.DataType & "") = "internaltag") And p.IsInterfaced()

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
' one of the object's own - the BTC's command interface. Both export as a
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

<xatm_config_data.PropertiesHelper.xatm_BTC:xatm_BTC_OnStartRunning()>
Sub xatm_BTC_OnStartRunning()

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
	
	AddProperty bag, "Preconditions", "Boolean", True, _
		"Field conditions that have to hold before a sequence may start. Bound to an expression - True while the maneuver is permitted.", _
		"Condições de campo que devem valer antes de uma sequência partir. Vinculada a uma expressão - True enquanto a manobra é permitida."

	AddProperty bag, "OperatorBlock", "Boolean", False, _
		"Operator lock. Blocks the start until the operator releases it.", _
		"Bloqueio do operador. Impede a partida até que o operador libere."

	AddProperty bag, "GeneralBlock", "Boolean", False, _
		"General interlock. Blocks the start, and is latched by a step failure until Reset clears it.", _
		"Intertravamento geral. Impede a partida e é selado por uma falha de passo até que o Reset o apague."
		
	
	AddProperty bag, "AutomaticBlock", "Boolean", False, _
		"Field conditions that block the automation. Bound to an expression - True keeps a sequence from starting, alongside OperatorBlock and GeneralBlock.", _
		"Condições de campo que bloqueiam o automatismo. Vinculada a uma expressão - True impede a partida, junto com OperatorBlock e GeneralBlock."


	' --- the command interface -----------------------------------------

	AddProperty bag, "CommandReset", "InternalTag", Empty, _
		"Reset command. Clears the latched step failures and the general block, so that a sequence can be started again.", _
		"Comando de reset. Apaga as falhas seladas de passo e o bloqueio geral, para que uma sequência possa partir novamente."

	AddProperty bag, "CommandStartTM", "InternalTag", Empty, _
		"Start command for a manual transfer. Written by the operator's screen to ask for the maneuver.", _
		"Comando de partida da transferência manual. Escrito pela tela do operador para pedir a manobra."

	AddProperty bag, "CommandStartNM", "InternalTag", Empty, _
		"Start command for a manual normalisation. Written by the operator's screen to ask for the maneuver.", _
		"Comando de partida da normalização manual. Escrito pela tela do operador para pedir a manobra."

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
	' The four that are readings - the blocks and Preconditions - are not
	' saved: their expression is the configuration, and the reading is
	' whatever the switchyard was doing at the time.
	'
	' EXPOSE_INTERFACE is a different question from the rest, and answered
	' separately: not what the panel may do with a property, but whether
	' the Elipse application is given a tag of its own for it. What a
	' screen draws or acts on is interfaced; what only an engineer sets -
	' the transformer this instance drives - is not.
	SetExposure bag, "Enabled",        EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "Transformer",    EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "Preconditions",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE
	SetExposure bag, "OperatorBlock",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE + EXPOSE_INTERFACE
	SetExposure bag, "GeneralBlock",   EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE + EXPOSE_INTERFACE
	SetExposure bag, "AutomaticBlock", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE

	' Running and the six step latches are the automation talking about
	' itself. Neither belongs on the configuration panel - there is
	' nothing about them to configure - but both are what a screen draws,
	' so they are interfaced and nothing else.
	SetExposure bag, "Running", EXPOSE_INTERFACE

	For i = 1 To 6
		SetExposure bag, "StepExecutionFailed" & i, EXPOSE_INTERFACE
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
	SetExposure bag, "CommandReset",         EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE
	SetExposure bag, "CommandStartTM",       EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE
	SetExposure bag, "CommandStartNM",       EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE
	SetExposure bag, "CommandOperatorBlock", EXPOSE_VIEW + EXPOSE_SAVED + EXPOSE_INTERFACE


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
	SetAlarm bag, "AutomaticBlock", "BLOQUEIO AUTOMÁTICO", PAIR_BLOCKED,      SEV_MEDIUM
	SetAlarm bag, "Preconditions",  "PRECONDIÇÕES",        PAIR_PRECONDITION, SEV_MEDIUM
	SetAlarm bag, "Running",        "AUTOMATISMO",         PAIR_RUNNING,      SEV_LOW

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
Const PAIR_RUNNING      = "PARADO|EM OPERAÇÃO|1"

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

	' The digital state that raises the alarm - DigitalLimit. It travels
	' inside the pair rather than beside it, so the words and the state
	' they describe cannot be changed independently of one another.
	Public Function AlarmLimit()
		AlarmLimit = CLng("0" & PairPart(2))
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
	SetExposure bag, "CommandTimeout",       EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueCommandOpen",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "RawValueCommandClose", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED

	SetExposure bag, "CommandOpen",          EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOOpen",       EXPOSE_VIEW + EXPOSE_SAVED
	SetExposure bag, "CommandOpenAlt",       EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOOpenAlt",    EXPOSE_VIEW + EXPOSE_SAVED
	SetExposure bag, "CommandClose",         EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOClose",      EXPOSE_VIEW + EXPOSE_SAVED
	SetExposure bag, "CommandCloseAlt",      EXPOSE_VIEW + EXPOSE_FORCE + EXPOSE_SAVED
	SetExposure bag, "CommandSBOCloseAlt",   EXPOSE_VIEW + EXPOSE_SAVED
	
	
	' Defective is configured the way a transformer's relays are: an
	' expression is written for it on the panel, it can be forced for a
	' test, and it is never saved - the expression is the configuration
	' and the reading is whatever the switchyard was doing at the time.
	SetExposure bag, "Defective", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE + EXPOSE_INTERFACE

	' Position is shown and never touched. Reading it back on the panel is
	' how an engineer sees that the raw values above were configured the
	' right way round, which is worth a row of its own; forcing it would
	' only make the panel lie about the switchyard.
	SetExposure bag, "Position", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_INTERFACE

	' The two command failures are latches the automation sets and Reset
	' clears - the breaker's own StepExecutionFailed. Nothing about them
	' is configuration, so they are interfaced and nothing else.
	SetExposure bag, "CommandOpenFailed",  EXPOSE_INTERFACE
	SetExposure bag, "CommandCloseFailed", EXPOSE_INTERFACE


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
Const PAIR_RUNNING      = "FINALIZADO|EM ANDAMENTO|1"

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

	' The digital state that raises the alarm - DigitalLimit. It travels
	' inside the pair rather than beside it, so the words and the state
	' they describe cannot be changed independently of one another.
	Public Function AlarmLimit()
		AlarmLimit = CLng("0" & PairPart(2))
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
Const PAIR_RUNNING      = "PARADO|EM OPERAÇÃO|1"

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

	' The digital state that raises the alarm - DigitalLimit. It travels
	' inside the pair rather than beside it, so the words and the state
	' they describe cannot be changed independently of one another.
	Public Function AlarmLimit()
		AlarmLimit = CLng("0" & PairPart(2))
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

' Returns True if at least one position tag has good quality — breaker does NOT need sim mode.
Function AnyBreakerTagHealthy(breaker)

    If TagHealthy(breaker.PositionOpen)      Then AnyBreakerTagHealthy = True : Exit Function
    If TagHealthy(breaker.PositionClosed)    Then AnyBreakerTagHealthy = True : Exit Function
    If TagHealthy(breaker.PositionOpenAlt)   Then AnyBreakerTagHealthy = True : Exit Function
    If TagHealthy(breaker.PositionClosedAlt) Then AnyBreakerTagHealthy = True : Exit Function
    AnyBreakerTagHealthy = False

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
            Case "XATM_BREAKER"
    
                Dim shouldSim
                
                If simEnabled Then
                    shouldSim = Not AnyBreakerTagHealthy(obj)
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
Const ROW_LEFT_PX = 1024
Const ROW_TOP_PX  = 32
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
	' Not under Automation, though: SyncAutomation finds a BTC by the
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
' name order, a BTC having no Id to go by.
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

<xatm_config_screens.Config.btnApply:btnApply_Click()>
Sub btnApply_Click()

	Dim transformerType, busbarType
	transformerType = SelectedLayout("SelectLayoutTransformer")
	busbarType      = SelectedLayout("SelectLayoutBusbar")

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

	SyncFolder doc, TRANSFORMER_PATH, TransformerDevices(transformerType), removed, added, failed
	SyncFolder doc, BUSBAR_PATH, BusbarDevices(busbarType), removed, added, failed
	SyncAutomation doc, AutomationCount(transformerType), removed, added, failed

	SetLayoutTag doc, "Transformer", transformerType
	SetLayoutTag doc, "Busbar", busbarType

	contentTag.WriteEx DocumentText(doc)
	
	Application.GetObject("xatm_config_data.Config.UpdateTreeviewSignal").WriteEx True
	
	MsgBox ApplyReport(transformerType, busbarType, removed, added, failed), vbInformation, "Apply"

End Sub

Const CONFIG_DATA      = "xatm_config_data"
Const HELPER_FOLDER    = "PropertiesHelper"
Const AUTOMATION_PATH  = "/xatm-config/folder[@name='Automation']"
Const LAYOUT_PATH      = "/xatm-config/folder[@name='Automation']/folder[@name='Layout']"
Const TRANSFORMER_PATH = "/xatm-config/folder[@name='Substation']/folder[@name='Transformer']"
Const BUSBAR_PATH      = "/xatm-config/folder[@name='Substation']/folder[@name='Busbar']"

Const BTC_CLASS         = "xatm_BTC"
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

		' A class with no tag of its own falls back to the first two
		' letters of its name - XATM_Feeder is FE.
		Case Else                : TypePrefix = UCase(Left(Replace(LCase(className), "xatm_", ""), 2))

	End Select

End Function


' Transformer n is Id n00, the way the layouts hand the Ids out, so a
' BTC the layout adds knows which transformer it drives.
Function AutomationTransformerId(n)

	AutomationTransformerId = n * 100

End Function


' One BTC automation per transformer, so the transformer layout sets
' the count.
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
		failed = failed & vbCrLf & "  the " & FolderLabel(folderPath) & _
		         " folder is not in the document"
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


' Brings the Automation folder to the layout. The BTC instances carry no
' Id property, so they go by the number in the name - BTC1..BTCn stay,
' the rest are dropped, and a missing one is created bound to the
' transformer of the same number.
Sub SyncAutomation(doc, keepCount, removed, added, failed)

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
	Set nodes = folder.selectNodes("object[@type='" & BTC_CLASS & "']")

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

	Dim name, path, overrides
	For n = 1 To keepCount

		If Not kept.Exists(n) Then

			name = "BTC" & n

			' The transformer folder is brought to the layout first, so a
			' transformer this BTC drives is already there to bind to.
			path = TransformerPath(doc, AutomationTransformerId(n))

			If path = "" Then
				overrides = Array()
			Else
				overrides = Array("Transformer", path)
			End If

			Set node = NewObject(folder, BTC_CLASS, name, overrides)

			If node Is Nothing Then
				failed = failed & vbCrLf & "  " & name & " in Automation - no manifest for " & BTC_CLASS
			ElseIf path = "" Then
				added = added & vbCrLf & "  " & name & " in Automation, with no transformer to bind to"
			Else
				added = added & vbCrLf & "  " & name & " in Automation, bound to " & path
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
' one of the object's own - the BTC's command interface. Both export as a
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

Function ApplyReport(transformerType, busbarType, removed, added, failed)

	Dim text
	text = "Applied " & transformerType & " + " & busbarType & "."

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


' Trailing digits of a name - BTC12 is 12, a name without them is 0.
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

Sub Foo()	
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

Sub Foo()
	
	
End Sub

<xatm_config_screens.Config.btnAlarms:btnAlarms_Click()>
Sub btnAlarms_Click()

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


' E3 will not take a Function as the last thing in a scope. The script
' is refused with nothing useful said about why, so an empty Sub goes
' on the end to keep a Function from being last. The other scopes in
' this dump carry one for the same reason - they are not leftovers.
Sub EndOfScope()
End Sub

<xatm_config_screens.Config:Config_OnPreShow(Arg)>
Sub Config_OnPreShow(Arg)
	
	' ================================
	' LAYOUT
	' ================================
	
	Dim imageList, i
	
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

<xatm_config_screens.Config.btnDistribution:btnDistribution_Click()>
Sub btnDistribution_Click()

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

' E3 will not take a Function as the last thing in a scope. The script
' is refused with nothing useful said about why, so an empty Sub goes
' on the end to keep a Function from being last. The other scopes in
' this dump carry one for the same reason - they are not leftovers.
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

Function CatalogXmlText()

	Dim xmlText
	xmlText = ""

	On Error Resume Next
	Dim xmlTag
	Set xmlTag = Application.GetObject("domainbrowser_Data.Catalog.XMLContent")
	If Err.Number = 0 Then
		xmlText = CStr(xmlTag.Value)
	Else
		Err.Clear
	End If
	Set xmlTag = Nothing
	On Error GoTo 0

	If Len(xmlText) = 0 Then
		Dim doc
		Set doc = CreateObject("MSXML2.DOMDocument.6.0")
		doc.async = False
		doc.load "C:\temp\appbrowser\elipse-catalog.xml"
		If doc.parseError.errorCode = 0 Then
			xmlText = doc.xml
		End If
		Set doc = Nothing
	End If

	CatalogXmlText = xmlText

End Function

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

Function CatalogXmlText()

	Dim xmlText
	xmlText = ""

	On Error Resume Next
	Dim xmlTag
	Set xmlTag = Application.GetObject("domainbrowser_Data.Catalog.XMLContent")
	If Err.Number = 0 Then
		xmlText = CStr(xmlTag.Value)
	Else
		Err.Clear
	End If
	Set xmlTag = Nothing
	On Error GoTo 0

	If Len(xmlText) = 0 Then
		
		Dim doc
		Set doc = CreateObject("MSXML2.DOMDocument.6.0")
		doc.async = False
		
		Dim xmlContent
		xmlContent = Application.GetObject("xatm_config_data.Catalog.XMLContent").Value
		
		doc.loadXML xmlContent
		
		If doc.parseError.errorCode = 0 Then
			xmlText = doc.xml
		End If
		
		Set doc = Nothing
	End If

	CatalogXmlText = xmlText

End Function

Function TreeNodeText(xmlNode)

	Dim text
	text = xmlNode.getAttribute("name")

	If Len(text) = 0 Then
		text = xmlNode.getAttribute("path")
	End If
	If Len(text) = 0 Then
		text = "[object]"
	End If

	TreeNodeText = text

End Function

Function TreeNodeImageKey(xmlNode, imageList)

	Dim className
	className = xmlNode.getAttribute("type")

	Dim imageValue
	imageValue = ImageListValueForTag(imageList, className)

	If Len(CStr(imageValue)) = 0 Then
		imageValue = ImageListValueForTag(imageList, "default")
	End If

	If Len(CStr(imageValue)) > 0 Then
		TreeNodeImageKey = imageValue
	ElseIf Len(className) = 0 Then
		TreeNodeImageKey = "Object"
	Else
		TreeNodeImageKey = className
	End If

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

Function AddTreeNode(tree, parentKey, xmlNode, sequence, useImages)

	sequence = sequence + 1

	Dim key
	key = "n" & CStr(sequence)

	Dim text
	text = TreeNodeText(xmlNode)

	Dim imageKey
	imageKey = ""
	If useImages Then
		imageKey = TreeNodeImageKey(xmlNode, tree.ImageList)
	End If

	On Error Resume Next
	Dim treeNode
	Set treeNode = Nothing
	If useImages Then
		If Len(parentKey) = 0 Then
			Set treeNode = tree.Nodes.Add(, , key, text, imageKey, imageKey)
		Else
			Set treeNode = tree.Nodes.Add(parentKey, 4, key, text, imageKey, imageKey)
		End If
	End If
	If Err.Number <> 0 Then
		Application.Trace "TreeView image key failed: " & imageKey & " - " & Err.Description
		Err.Clear
	End If
	If treeNode Is Nothing Then
		If Len(parentKey) = 0 Then
			Set treeNode = tree.Nodes.Add(, , key, text)
		Else
			Set treeNode = tree.Nodes.Add(parentKey, 4, key, text)
		End If
	End If
	If Err.Number <> 0 Then
		Application.Trace "TreeView add failed: " & xmlNode.getAttribute("path") & " - " & Err.Description
		Err.Clear
		AddTreeNode = ""
		Exit Function
	End If
	On Error GoTo 0

	If useImages Then
		On Error Resume Next
		treeNode.Image = imageKey
		treeNode.SelectedImage = imageKey
		If Err.Number <> 0 Then
			Application.Trace "TreeView image assign failed: " & imageKey & " - " & Err.Description
			Err.Clear
		End If
		On Error GoTo 0
	End If

	On Error Resume Next
	treeNode.Tag = xmlNode.getAttribute("path")
	treeNode.ToolTipText = xmlNode.getAttribute("path")
	Err.Clear
	On Error GoTo 0

	Dim child
	For Each child In xmlNode.childNodes
		If child.nodeName = "Object" Then
			Dim childKey
			childKey = AddTreeNode(tree, key, child, sequence, useImages)
		End If
	Next

	If xmlNode.getAttribute("depth") = "1" Then
		On Error Resume Next
		treeNode.Expanded = True
		Err.Clear
		On Error GoTo 0
	End If

	AddTreeNode = key

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
	treeNode.Expanded = True
	Err.Clear
	On Error GoTo 0

	AddTreeRootNode = key

End Function

Sub FillTreeViewFromCatalogXml(tree, xmlText)

	On Error Resume Next
	tree.Nodes.Clear
	If Err.Number <> 0 Then
		Application.Trace "TreeView clear failed: " & Err.Description
		Err.Clear
	End If
	On Error GoTo 0

	If Len(xmlText) = 0 Then
		Application.Trace "Catalog XML is empty; TreeView was not filled."
		Exit Sub
	End If

	Dim doc
	Set doc = CreateObject("MSXML2.DOMDocument.6.0")
	doc.async = False
	doc.loadXML xmlText

	If doc.parseError.errorCode <> 0 Then
		Application.Trace "Catalog XML parse failed: " & doc.parseError.reason
		Exit Sub
	End If

	Dim sequence
	sequence = 0

	Dim useImages
	useImages = BindTreeImageList(tree)

	Dim parentKey
	parentKey = ""

	Dim domainName
	domainName = ElipseDomainName()
	If Len(domainName) > 0 Then
		parentKey = AddTreeRootNode(tree, domainName, sequence, useImages)
	End If

	Dim objectNode
	For Each objectNode In doc.documentElement.childNodes
		If objectNode.nodeName = "Object" Then
			Dim nodeKey
			nodeKey = AddTreeNode(tree, parentKey, objectNode, sequence, useImages)
		End If
	Next

	Application.Trace "TreeView filled from catalog XML with " & CStr(sequence) & " nodes."

End Sub

<xatm_config_screens.DomainBrowser:DomainBrowser_OnShow()>
Sub DomainBrowser_OnShow()
	FillTree	
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
        AddItem contentArr(i), 0
        'If Application.GetObject("xatm_config_data.ConsoleLogEngine").AutoScrolling Then
            ListIndex = 0
        'End If
    Next
    
End Sub

<xatm_config_screens.Menu.btnBTC:btnBTC_Click()>
Sub btnBTC_Click()

	Dim autos
	Set autos = Application.GetObject("XATM_Data.Automation")

	Dim modes
	modes = Array("TM", "NM", "TA")

	Dim nModes, nActions
	nModes   = UBound(modes) + 1     ' trigger actions
	nActions = nModes + 2            ' + Operator Block + Reset

	' --- Collect every BTC instance -----------------------------------
	Dim obj, total
	total = 0
	For Each obj In autos
		If TypeName(obj) = "xatm_BTC" Then total = total + 1
	Next

	If total = 0 Then
		MsgBox "No BTC automation found!"
		Exit Sub
	End If

	Dim btc()
	ReDim btc(total - 1)
	Dim k
	k = 0
	For Each obj In autos
		If TypeName(obj) = "xatm_BTC" Then
			Set btc(k) = obj
			k = k + 1
		End If
	Next

	' --- Order by transformer Id (stable menu order) ------------------
	Dim i, j, tmp
	For i = 0 To total - 2
		For j = 0 To total - 2 - i
			If btc(j).Transformer.Id > btc(j + 1).Transformer.Id Then
				Set tmp = btc(j)
				Set btc(j) = btc(j + 1)
				Set btc(j + 1) = tmp
			End If
		Next
	Next

	' --- Build the menu -----------------------------------------------
	Dim menu, tr
	menu = ""
	For i = 0 To total - 1
		Set tr = btc(i).Transformer
		If i > 0 Then menu = menu & "|"
		menu = menu & tr.Name & "{" & Join(modes, "|") & "||" & _
		       IIf(btc(i).OperatorBlock, "*", "") & "Operator Block|Reset}"
	Next

	Dim lOption
	lOption = Application.SelectMenu(menu)
	If lOption <= 0 Then Exit Sub

	' --- Map option -> (instance, action) -----------------------------
	Dim idx, action
	idx    = (lOption - 1) \ nActions
	action = ((lOption - 1) Mod nActions) + 1

	Dim target
	Set target = btc(idx)
	Set tr = target.Transformer

	If action <= nModes Then                       
		
		' TM / NM / TA
		Dim mode
		mode = modes(action - 1)
		If MsgBox("Force " & mode & " on " & tr.Name & "?", _
		          vbYesNo + vbQuestion, "Confirm") = vbYes Then
			target.Item("Commands").Item("Start").WriteEx mode & ":" & tr.Id
		End If

	ElseIf action = nModes + 1 Then
		
		' Operator Block toggle
		target.OperatorBlock = Not target.OperatorBlock

	Else
		
		' Reset
		target.Item("Commands").Item("Reset").WriteEx True

	End If
		
End Sub

<xatm_config_screens.Menu.btnConfig:btnConfig_Click()>
Sub btnConfig_Click()
	
	Dim fr
	Set fr = Application.GetFrame("Superior")
	fr.OpenScreen "xatm_config_screens.Config?4?0", 0
		
End Sub

<xatm_config_screens.Menu.btnSimulationMode:btnSimulationMode_Click()>
Sub btnSimulationMode_Click()
	
	Dim tag
	Set tag = Application.GetObject("xatm_config_data.SimulationMode")
	tag.WriteEx (Not CBool(tag.Value))
		
End Sub

<xatm_config_screens.Menu.btnSupervision:btnSupervision_Click()>
Sub btnSupervision_Click()
	
	Dim fr
	Set fr = Application.GetFrame("Superior")
	fr.OpenScreen "xatm_config_screens.SingleLineDiagram?4?0", 0
		
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
	On Error Resume Next
	Item("xatm_BTCStatus1").Source = Application.GetObject("XATM_Data.Automation.BTC1").PathName
	
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

