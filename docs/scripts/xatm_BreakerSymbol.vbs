-----------------------
Documenta��o de Scripts
-----------------------
xatm_BreakerSymbol
Wed Jul  8 16:57:50 2026
-----------------------

<xatm_BreakerSymbol.Source:xatm_BreakerSymbol_OnSourceChanged()>
Sub xatm_BreakerSymbol_OnSourceChanged()
	
	Item("lblName").Value = Source.Name
	
	Links.CreateLink "Position", Source.Item("Data").Item("Position").PathName & ".Value"
	
	Links.CreateLink "MemorizedPosition", Source.Item("Data").Item("MemorizedPosition").PathName & ".Value"
	Links.CreateLink "MemorizedPositionTimer", Source.Item("Data").Item("MemorizedPosition").PathName & ".DocString"
	
	Links.CreateLink "CommunicationFailure", Source.Item("Data").Item("CommunicationFailure").PathName & ".Value"
	
	Links.CreateLink "SimulationModeEnabled", Source.Item("Data").Item("SimulationModeEnabled").PathName & ".Value"

	Links.CreateLink "SimulateCommandFailure", Source.Item("Data").Item("SimulateCommandFailure").PathName & ".Value"

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
		Case 1 : openCmd  = "!Open"   ' already open  � disable Open
		Case 2 : closeCmd = "!Close"  ' already closed � disable Close
	End Select

	Dim failTag
	Set failTag = source.Item("Data").Item("SimulateCommandFailure")

	Dim options
	options = "Command{" & openCmd & "|" & closeCmd & "}|" & _
	          IIf(CBool(failTag.Value), "*", "") & "Command Failure|" & _
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

