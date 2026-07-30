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