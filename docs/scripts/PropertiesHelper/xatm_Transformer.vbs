Const EXPOSE_NONE = 0

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
