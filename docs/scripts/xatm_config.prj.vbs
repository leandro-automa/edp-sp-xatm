-----------------------
Documentação de Scripts
-----------------------
XATM_CONFIG (C:\ProjDev\edp_sp\xatm_config.prj)
Wed Aug  5 16:12:12 2026
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
	' whatever the switchyard was doing at the time. StepExecutionFailed1
	' to 6 are latches the automation sets and Reset clears, so they are
	' neither shown nor saved.
	SetExposure bag, "Enabled",        EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EDIT + EXPOSE_SAVED
	SetExposure bag, "Transformer",    EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_SAVED
	SetExposure bag, "Preconditions",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE
	SetExposure bag, "OperatorBlock",  EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE
	SetExposure bag, "GeneralBlock",   EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_FORCE
	SetExposure bag, "AutomaticBlock", EXPOSE_VIEW + EXPOSE_VALUE + EXPOSE_EXPRESSION + EXPOSE_FORCE

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

<xatm_config_data.XML.ImportXml:ImportXml_OnChangedValue()>
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
		Execute "Set obj." & name & " = gWriteObject"

	Else

		Execute "obj." & name & " = gWriteValue"

	End If

	If Err.Number <> 0 Then
		failed = Err.Description
		Err.Clear
	End If

	On Error Goto 0

	If failed <> "" Then
		report = report & vbCrLf & "  could not set " & name & " on " & where & " - " & failed
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

<xatm_config_data.XML.ImportXml:ImportXml_OnStartRunning()>
Sub ImportXml_OnStartRunning()

	DocString = ""

End Sub

<xatm_config_data.XML.SaveXML:SaveXML_OnChangedValue()>
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

<xatm_config_data.XML.SaveXML:SaveXML_OnStartRunning()>
Sub SaveXML_OnStartRunning()

	DocString = ""
		
End Sub

<xatm_config_data.XML.SetProperty:SetProperty_OnChangedValue()>
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

		Case KIND_NAME, KIND_PROPERTY

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

' What the command is asking to change: the object itself, or one of the
' properties declared on it.
Const KIND_NAME     = "name"
Const KIND_PROPERTY = "property"

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

<xatm_config_data.XML.SetProperty:SetProperty_OnStartRunning()>
Sub SetProperty_OnStartRunning()

	DocString = "-1"

End Sub

<xatm_config_data.XML.XMLBuilderAfterDelay:XMLBuilderAfterDelay_Functions()>
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
		'
		' A reading is left out altogether. Its expression is the setting; the
		' number beside it is whatever the switchyard happened to be doing,
		' and putting that in the document is how a step failure latched
		' during a test ends up looking like a configured value.
		Dim value, source
		source = SourceOf(obj, p.Name, p.DataType)

		If IsLinkType(p.DataType) Or Not p.IsSaved() Then
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

<xatm_config_data.XML.XMLBuilderAfterDelay:XMLBuilderAfterDelay_OnStartRunning()>
Sub XMLBuilderAfterDelay_OnStartRunning()

	Value = 3	' delay in seconds
	
End Sub

<xatm_config_data.XML.XMLBuilderAfterDelay:XMLBuilderAfterDelay_TickCountdown()>
Sub XMLBuilderAfterDelay_TickCountdown()
	
	If Value > 0 Then
			
		Value = Value - 1
	
	Else
		
		' stop
		Value = -1
		
		BuildXML
		
	End If
	
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
	content = Application.GetObject("xatm_config_data.XML.XMLContent").Value
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
Const ROW_TOP_PX  = 16
Const ROW_GAP_PX  = 0


' What the object's own name is shown as. It has no manifest entry, so
' the row has no help text for it and nothing declares its type.
Const NAME_PROPERTY = "Name"
Const NAME_TYPE     = "String"

' The folder whose objects are named by the layout rather than by hand.
Const AUTOMATION_FOLDER = "Automation"

' Which of the two an edit is asking to change, told to the row when it
' is built so it never has to work it out from the property name.
Const KIND_NAME     = "name"
Const KIND_PROPERTY = "property"

' What marks a key as an Id rather than a path.
Const ID_PREFIX     = "id:"

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


' One row per property of the selected object, stacked down the panel in
' the order the export wrote them - which is the order the manifest
' declares them in.
Sub BuildPropertyRows(objectNode, key)

	Dim objectType
	objectType = Attribute(objectNode, "type")

	Dim properties
	Set properties = objectNode.selectNodes("property")

	Dim i, property, row, y
	y = Himetric(ROW_TOP_PX)

	' The object's own name goes first. It is the object rather than
	' something declared on it, so it is neither a property of the document
	' nor an entry in the manifest, and it is built here instead of coming
	' out of the loop.
	'
	' Not under Automation, though: SyncAutomation finds a BTC by the
	' number on the end of its name, so the name is structure there and not
	' the operator's to change.
	If Not IsAutomation(objectNode) Then
		Set row = NewRow(KIND_NAME, key, objectType, NAME_PROPERTY, NAME_TYPE, _
		                 Attribute(objectNode, "name"), y)
		y = y + row.Height + Himetric(ROW_GAP_PX)
	End If

	For i = 0 To properties.length - 1

		Set property = properties.item(i)

		Set row = NewRow(KIND_PROPERTY, key, objectType, Attribute(property, "name"), _
		                 Attribute(property, "type"), Attribute(property, "value"), y)

		' A row comes out at the size it was drawn, so the next one goes
		' under whatever that turned out to be.
		y = y + row.Height + Himetric(ROW_GAP_PX)

	Next

End Sub


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
Function NewRow(kind, key, objectType, propertyName, propertyType, value, y)

	Dim row
	Set row = Screen.AddObject(ROW_CLASS, False)

	row.X = Himetric(ROW_LEFT_PX)
	row.Y = y
	
	row.Kind         = kind
	row.Source       = key
	row.ObjectType   = objectType
	row.PropertyName = propertyName
	row.PropertyType = propertyType
	row.Value        = value

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


' A property whose configuration is an association, not a value.
Function IsLinkType(dataType)

	IsLinkType = (LCase(dataType) = "iotag")

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

	' 1. The document into the project. Nothing else happens if it will
	' not go in - a file saved against a project that does not match it is
	' worse than no save at all.
	Dim importer
	Set importer = Application.GetObject(IMPORT_XML)
	importer.WriteEx True

	If importer.DocString <> EXIT_SUCCESS Then
		MsgBox "Nothing was saved. The document could not be written into the " & _
		       "project - the console says why.", vbCritical, "Save"
		Exit Sub
	End If

	' 2. The layout the screen is showing, and then the project file. The
	' objects the import just wrote are in the same container, so one Save
	' persists them too.
	Dim layoutFolder
	Set layoutFolder = Application.GetObject("XATM_Data.Automation.Layout")

	layoutFolder.Item("Transformer").WriteEx Screen.Item("SelectLayoutTransformer").Index
	layoutFolder.Item("Busbar").WriteEx Screen.Item("SelectLayoutBusbar").Index

	layoutFolder.Context("Container").Save()

	' 3. And the file.
	Dim saver
	Set saver = Application.GetObject(SAVE_XML)
	saver.WriteEx True

	If saver.DocString <> EXIT_SUCCESS Then
		MsgBox "The project was saved, but the XML file was not - the console " & _
		       "says why.", vbExclamation, "Save"
		Exit Sub
	End If

End Sub


Const IMPORT_XML   = "xatm_config_data.XML.ImportXml"
Const SAVE_XML     = "xatm_config_data.XML.SaveXML"
Const EXIT_SUCCESS = "EXIT_SUCCESS"

<xatm_config_screens.Config.btnUpdateTreeview:btnUpdateTreeview_Click()>
Sub btnUpdateTreeview_Click()
	
	' Builds the tree by hand, from the export as it stands in the tag. The
	' same thing Config_OnPreShow does on open and btnApply does after an
	' Apply - here it says out loud what went wrong instead of doing nothing.

	Dim tree
	Set tree = Screen.Item("TreeView")

	If tree Is Nothing Then
		MsgBox "No tree control named '" & TREE_CONTROL & "' was found on this screen.", _
		       vbExclamation, "Tree"
		Exit Sub
	End If

	Dim content
	content = Empty

	On Error Resume Next
	content = Application.GetObject("xatm_config_data.XML.XMLContent").Value
	On Error Goto 0

	If IsEmpty(content) Or IsNull(content) Then
		MsgBox "'XMLContent' is empty - run the XML export first.", vbExclamation, "Tree"
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

' What an engineer puts in the DocString of a folder or a device to keep
' it, and everything under it, off the tree.
Const NO_SHOW      = "$NoShow$"

' Fills the tree from the document: one node per folder and per object.
' The two folders at the top of the document - Automation and Substation -
' become the two roots. A property or a tag configures a node, it is not
' a node of its own.
Sub PopulateTree(tree, doc)

	tree.Nodes.Clear

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

	node.Text = NodeText(element)

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

