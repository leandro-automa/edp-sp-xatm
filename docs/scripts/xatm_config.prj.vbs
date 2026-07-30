-----------------------
Documentação de Scripts
-----------------------
XATM_CONFIG (C:\ProjDev\edp_sp\xatm_config.prj)
Thu Jul 30 17:41:13 2026
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
	                    " type=""" & EscapeXml(TypeName(obj)) & """>" & vbCrLf

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

<xatm_config_screens.Config.CommandButton1:CommandButton1_Click()>
Sub CommandButton1_Click()
	Set o = Application.GetObject("xatm_config_data.PropertiesHelper.xatm_Breaker")	
	
	For Each key In o.value
		
		msgbox (o.value(key).HelpEn),,key
		exit for
		
	Next


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

