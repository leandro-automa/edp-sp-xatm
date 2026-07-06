<autoTelas.autoTelasComuns.XATM_Menu.btnBTC:btnBTC_Click()>
' Operator menu to force a BTC automation on any transformer.
' Menu is built dynamically from every xatm_BTC in XATM_Data.Automation,
' so adding a transformer (or a mode) needs no change here.
'
' Per transformer:  T.Name{TM|NM|TA|Operator Block|Reset}
' Triggering writes "MODE:triggerId" to that instance's Commands.Start.
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
		menu = menu & tr.Name & "{" & Join(modes, "|") & "|" & _
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

	If action <= nModes Then                       ' TM / NM / TA
		Dim mode
		mode = modes(action - 1)
		If MsgBox("Force " & mode & " on " & tr.Name & "?", _
		          vbYesNo + vbQuestion, "Confirm") = vbYes Then
			target.Item("Commands").Item("Start").WriteEx mode & ":" & tr.Id
		End If

	ElseIf action = nModes + 1 Then                ' Operator Block toggle
		target.OperatorBlock = Not target.OperatorBlock

	Else                                            ' Reset
		target.Item("Commands").Item("Reset").WriteEx "Reset"
	End If

End Sub
