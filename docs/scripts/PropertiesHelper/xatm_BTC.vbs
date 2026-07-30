Const EXPOSE_NONE = 0

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
