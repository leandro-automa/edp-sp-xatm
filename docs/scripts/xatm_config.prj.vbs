-----------------------
Documentação de Scripts
-----------------------
XATM_CONFIG (C:\ProjDev\edp_sp\xatm_config.prj)
Fri Jul 31 17:15:45 2026
-----------------------

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

	AddProperty bag, "OperatorBlock", "Boolean", False, _
		"Operator lock. Blocks the start until the operator releases it.", _
		"Bloqueio do operador. Impede a partida até que o operador libere."

	AddProperty bag, "GeneralBlock", "Boolean", False, _
		"General interlock. Blocks the start, and is latched by a step failure until Reset clears it.", _
		"Intertravamento geral. Impede a partida e é selado por uma falha de passo até que o Reset o apague."

	Dim i
	For i = 1 To 6

		AddProperty bag, "StepExecutionFailed" & i, "Boolean", Empty, _
			"Latched failure of step " & i & ". Set when the step does not execute and the automation goes to global lockout, cleared by Reset.", _
			"Falha selada do passo " & i & ". Marcada quando o passo não executa e o automatismo entra em bloqueio geral, apagada pelo Reset."

	Next

	Set Value = bag

End Sub

Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

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

	Set Value = bag

End Sub

Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

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

	Set Value = bag

End Sub

Class PropertyInfo

	Public Name
	Public DataType
	Public InitialValue
	Public Exposure
	Public HelpEn
	Public HelpPt

	Public Function Help(lang)

		If lang = "pt-BR" Then
			Help = HelpPt
		Else
			Help = HelpEn
		End If

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

<xatm_config_data.TagInterno1:TagInterno1_OnStartRunning()>
Sub TagInterno1_OnStartRunning()

Sub Foo()
End Sub

<xatm_config_data.XML.XMLBuilder:XMLBuilder_OnStartRunning()>
Sub XMLBuilder_OnStartRunning()
	
	Dim xmlContent
	xmlContent = ExportXml()

	Parent.Item("XMLContent").WriteEx xmlContent
	
	SaveXml "C:\ProjDev\edp_sp\docs\scripts\old\file.xml"
	
End Sub

Const XML_VERSION   = "1"
Const DATA_ROOT     = "XATM_Data"
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

	xml = xml & "<xatm-config version=""" & XML_VERSION & """" & _
	                    " exported=""" & IsoNow() & """" & _
	                    " root=""" & EscapeXml(root.Name) & """>" & vbCrLf

	xml = xml & ExportFolder(root, vbTab)

	xml = xml & "</xatm-config>" & vbCrLf

	ExportXml = xml

End Function


' Writes the document to disk as UTF-8.
Sub SaveXml(path)

	Dim stream
	Set stream = CreateObject("ADODB.Stream")

	stream.Type     = 2			' adTypeText
	stream.Charset  = "utf-8"
	stream.Open
	stream.WriteText ExportXml()
	stream.SaveToFile path, 2	' adSaveCreateOverWrite
	stream.Close

End Sub


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

		Dim value
		If IsLinkType(p.DataType) Then
			value = Empty
		Else
			value = ReadProperty(obj, p.Name)
		End If

		xml = xml & indent & vbTab & "<property name=""" & EscapeXml(p.Name) & """" & _
		                                  " type=""" & EscapeXml(p.DataType) & """" & _
		                                  ValueAttribute(value) & "/>" & vbCrLf

	Next

	xml = xml & indent & "</object>" & vbCrLf

	ExportObject = xml

End Function

' The manifest declared for a class, or Nothing when the class has
' none - which is how a folder tells itself apart from an object.
Function ManifestOf(className)

	Set ManifestOf = Nothing

	On Error Resume Next
	Set ManifestOf = Application.GetObject(CONFIG_DATA).Item(HELPER_FOLDER).Item(className).Value
	On Error Goto 0

End Function


' A property whose configuration is an association, not a value.
Function IsLinkType(dataType)

	IsLinkType = (LCase(dataType) = "iotag")

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

<xatm_config_screens.Config.btnApply:btnApply_Click()>
Sub btnApply_Click()

	Dim transformerType, busbarType
	transformerType = SelectedLayout("SelectLayoutTransformer")
	busbarType      = SelectedLayout("SelectLayoutBusbar")

	' If Not IsTransformerLayout(transformerType) Then
	' 	MsgBox "'" & transformerType & "' is not a transformer layout this screen can apply.", _
	' 	       vbExclamation, "Apply"
	' 	Exit Sub
	' End If

	' If Not IsBusbarLayout(busbarType) Then
	' 	MsgBox "'" & busbarType & "' is not a busbar layout this screen can apply.", _
	' 	       vbExclamation, "Apply"
	' 	Exit Sub
	' End If

	Dim contentTag
	Set contentTag = Nothing
	On Error Resume Next
	Set contentTag = Application.GetObject(CONFIG_DATA).Item("XML").Item("XMLContent")
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

	Dim removed, missing
	removed = ""
	missing = ""

	PruneById doc, TRANSFORMER_PATH, TransformerIds(transformerType), removed, missing
	PruneById doc, BUSBAR_PATH, BusbarIds(busbarType), removed, missing
	PruneAutomation doc, AutomationCount(transformerType), removed, missing

	SetLayoutTag doc, "Transformer", transformerType
	SetLayoutTag doc, "Busbar", busbarType

	contentTag.WriteEx DocumentText(doc)

	MsgBox ApplyReport(transformerType, busbarType, removed, missing), vbInformation, "Apply"

End Sub

Const CONFIG_DATA      = "xatm_config_data"
Const AUTOMATION_PATH  = "/xatm-config/folder[@name='Automation']"
Const LAYOUT_PATH      = "/xatm-config/folder[@name='Automation']/folder[@name='Layout']"
Const TRANSFORMER_PATH = "/xatm-config/folder[@name='Substation']/folder[@name='Transformer']"
Const BUSBAR_PATH      = "/xatm-config/folder[@name='Substation']/folder[@name='Busbar']"

Const NODE_TEXT = 3


' ------------------------------------------------------------
'  LAYOUTS
' ------------------------------------------------------------

' Every Id the transformer layout expects to find in the Transformer
' folder - the transformers themselves and their secondary breakers.
Function TransformerIds(layoutType)

	Select Case UCase(layoutType)

		Case "4TR4LV"
			TransformerIds = Array(100, 200, 300, 400, 120, 220, 320, 420)

		Case "2TR2LV"
			TransformerIds = Array(100, 200, 120, 220)

		Case Else
			TransformerIds = Array()

	End Select

End Function


' Every Id the busbar layout expects to find in the Busbar folder.
Function BusbarIds(layoutType)

	Select Case UCase(layoutType)

		Case "6BB6TIERING"
			BusbarIds = Array(700, 710, 720, 730, 740, 900)

		Case "2BB1TIE"
			BusbarIds = Array(700)

		Case Else
			BusbarIds = Array()

	End Select

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

' Drops every object in the folder whose Id the layout does not name,
' and reports the Ids the layout names that the document does not have.
Sub PruneById(doc, folderPath, ids, removed, missing)

	Dim wanted
	Set wanted = CreateObject("Scripting.Dictionary")

	Dim i
	For i = 0 To UBound(ids)
		wanted.Add CStr(ids(i)), False
	Next

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim nodes, node, id, n
	Set nodes = doc.selectNodes(folderPath & "/object")

	For n = 0 To nodes.length - 1

		Set node = nodes.item(n)
		id = PropertyValue(node, "Id")

		If wanted.Exists(id) Then
			wanted(id) = True
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

	Dim key
	For Each key In wanted.Keys
		If Not wanted(key) Then
			missing = missing & vbCrLf & "  Id " & key & " in " & FolderLabel(folderPath)
		End If
	Next

End Sub


' The BTC instances carry no Id property, so the layout keeps
' BTC1..BTCn by the number in the name and drops the rest.
Sub PruneAutomation(doc, keepCount, removed, missing)

	Dim kept
	Set kept = CreateObject("Scripting.Dictionary")

	Dim doomed
	Set doomed = CreateObject("Scripting.Dictionary")

	Dim nodes, node, n, num
	Set nodes = doc.selectNodes(AUTOMATION_PATH & "/object[@type='xatm_BTC']")

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

	For n = 1 To keepCount
		If Not kept.Exists(n) Then
			missing = missing & vbCrLf & "  BTC" & n & " in Automation"
		End If
	Next

End Sub


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
			If Trim(ws.text) = "" Then
				ws.parentNode.removeChild ws
			End If
		End If
	End If

	node.parentNode.removeChild node

End Sub


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

Function ApplyReport(transformerType, busbarType, removed, missing)

	Dim text
	text = "Applied " & transformerType & " + " & busbarType & "."

	If removed <> "" Then
		text = text & vbCrLf & vbCrLf & "Removed from the document:" & removed
	End If

	If missing <> "" Then
		text = text & vbCrLf & vbCrLf & _
		       "Expected by the layout and not in the document:" & missing
	End If

	If removed = "" And missing = "" Then
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

