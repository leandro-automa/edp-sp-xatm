-----------------------
Documenta��o de Scripts
-----------------------
xatm_TransformerSymbol
Tue Jul  7 16:20:29 2026
-----------------------

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
	          "Reset||Cancel"

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

