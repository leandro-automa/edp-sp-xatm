-----------------------
Documenta��o de Scripts
-----------------------
xatm_Transformer
Tue Jul  7 12:11:08 2026
-----------------------

<xatm_Transformer.Data.Timers.UndervoltageRelay:UndervoltageRelay_Counter()>
Sub UndervoltageRelay_Counter()
	
	If Value > 0 Then
		
		If Value <= 3 Then
			
			WriteLog "Undervoltage trigger is about to be activated in " & Value & "..."
			
		End If
			
		Value = Value - 1
		Exit Sub
	
	End If
	
	
	
	' ================
	' TRIGGER BTC
	' ================
	
	
	WriteLog "Undervoltage relay counter has reached zero - requesting TA."
	RequestBTC "TA"
	Value = -1

End Sub

<xatm_Transformer.Data.Timers.UndervoltageRelay:UndervoltageRelay_Functions()>
Sub UndervoltageRelay_Functions()
End Sub


Sub WriteLog(message)
	
	Dim consoleLogEngine
	Set consoleLogEngine = Nothing
	
	On Error Resume Next
	Set consoleLogEngine = Application.GetObject("xatm_config_data.ConsoleLogEngine")
	Application.Trace "[" & Parent.Parent.Parent.Name & "] - " & message
	On Error Goto 0
	
	If Not consoleLogEngine Is Nothing Then
		consoleLogEngine.WriteLine = "[" & Parent.Parent.Parent.Name & "] - " & message
	End If
		
End Sub

' Finds the xatm_BTC bound to this transformer and requests an automation
' by writing "MODE:Id" to its Commands.Start. The BTC arm routine applies
' every interlock (Enabled, Running, blocks, CR2 pre-condition) - this only
' requests; it does not gate.
Sub RequestBTC(mode)

	Dim obj, bound, tr
	Set bound = Nothing

	For Each obj In Application.GetObject("XATM_Data.Automation")
		If TypeName(obj) = "xatm_BTC" Then

			Set tr = Nothing
			On Error Resume Next
			Set tr = obj.Transformer          ' may be unbound
			On Error GoTo 0

			If Not tr Is Nothing Then
				If tr.Id = xatm_Transformer.Id Then
					Set bound = obj
					Exit For
				End If
			End If

		End If
	Next

	If bound Is Nothing Then
		WriteLog "No BTC bound to this transformer - " & mode & " request ignored."
		Exit Sub
	End If

	bound.Item("Commands").Item("Start").WriteEx mode & ":" & xatm_Transformer.Id
	WriteLog mode & " requested via " & bound.Name & "."

End Sub

<xatm_Transformer.Data.Timers.UndervoltageRelay:UndervoltageRelay_OnChangedValue()>
Sub UndervoltageRelay_OnChangedValue()
	
	If xatm_Transformer.UndervoltageRelay Then
		
		Dim delay
		
		If xatm_Transformer.UndervoltageDelay > 0 Then
			delay = xatm_Transformer.UndervoltageDelay
		Else
			delay = 25 ' default value
		End If
		
		WriteLog "Undervoltage timer initialized."
		Value = delay
	
	Else
		
		If Value <> -1 Then
			WriteLog "Undervoltage timer reset."
			Value = -1
		End If
		
	End If
	
End Sub

