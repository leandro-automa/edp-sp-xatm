-----------------------
Documentação de Scripts
-----------------------
txtTct
Mon Jul  6 10:57:46 2026
-----------------------

<autoTelas.autoTelasComuns.XATM_Menu.txtTct:txtTct_Click()>
Sub txtTct_Click()
	Dim bFlag
	Dim lOption
	Dim objTR
	Dim sAux
	Dim objTCT1, objTCT2, objTCT3, objTCT4
	
	sAux = "{Disparo Liberação|Disparo Normalização|Bloqueio Operador|Reset}"

	sAuxTR1 = "!N/A" & sAux
	sAuxTR2 = "!N/A" & sAux
	sAuxTR3 = "!N/A" & sAux
	sAuxTR4 = "!N/A" & sAux
	bFlag = False
	
	For Each obj in Application.GetObject("autoDados.Automatismo")
		If TypeName(obj) = "XATM_TCT" Then
			Set objTR = obj.Transformador
			Select Case objTR.Id
				Case 100 
					' sTCT1 = objTR.Name
					Set objTCT1 = obj
					sAuxTR1 = objTR.Name & "{Disparo Liberação|Disparo Normalização|" & _
							IIf(obj.BloqueioOperador, "*", "") & "Bloqueio Operador|Reset}"
					bFlag = True
				Case 200
					' sTCT2 = objTR.Name
					Set objTCT2 = obj
					sAuxTR2 = objTR.Name & "{Disparo Liberação|Disparo Normalização|" & _
							IIf(obj.BloqueioOperador, "*", "") & "Bloqueio Operador|Reset}"
					bFlag = True
				Case 300
					' sTCT3 = objTR.Name
					Set objTCT3 = obj
					sAuxTR3 = objTR.Name & "{Disparo Liberação|Disparo Normalização|" & _
							IIf(obj.BloqueioOperador, "*", "") & "Bloqueio Operador|Reset}"
					bFlag = True
				Case 400
					' sTCT4 = objTR.Name
					Set objTCT4 = obj
					sAuxTR4 = objTR.Name & "{Disparo Liberação|Disparo Normalização|" & _
							IIf(obj.BloqueioOperador, "*", "") & "Bloqueio Operador|Reset}"
					bFlag = True
			End Select
		End If
	Next
	
	If ( bFlag ) Then
		lOption = Application.SelectMenu( sAuxTR1 & "|" & sAuxTR2  & "|" & sAuxTR3  & "|" & sAuxTR4)
		Select Case lOption
			Case 1 'Disparo Liberação TR1
				If MsgBox("Tem certeza que deseja forçar o disparo do automatismo TCT?",vbYesno+vbQuestion,"Confirma")=vbNo Then Exit Sub
				objTCT1.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
			Case 2 'Disparo Normalização TR1
				If MsgBox("Tem certeza que deseja forçar o disparo do automatismo TCT?",vbYesno+vbQuestion,"Confirma")=vbNo Then Exit Sub
				objTCT1.Item("DadosTCT").Item("PartidaNormalizaçãoTCT").WriteEx "Normalizar"
			Case 3 'Bloqueio Operador TR1
				objTCT1.BloqueioOperador = Not objTCT1.BloqueioOperador
			Case 4 'Reset TR1
				objTCT1.Item("DadosTCT").Item("Comandos").WriteEx "Reset"
			Case 5 'Disparo Liberação TR2
				If MsgBox("Tem certeza que deseja forçar o disparo do automatismo TCT?",vbYesno+vbQuestion,"Confirma")=vbNo Then Exit Sub
				objTCT2.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
			Case 6 'Disparo Normalização TR2
				If MsgBox("Tem certeza que deseja forçar o disparo do automatismo TCT?",vbYesno+vbQuestion,"Confirma")=vbNo Then Exit Sub
				objTCT2.Item("DadosTCT").Item("PartidaNormalizaçãoTCT").WriteEx "Normalizar"
			Case 7 'Bloqueio Operador TR2
				objTCT2.BloqueioOperador = Not objTCT2.BloqueioOperador
			Case 8 'Reset TR2
				objTCT2.Item("DadosTCT").Item("Comandos").WriteEx "Reset"
			Case 9 'Disparo Liberação TR3
				If MsgBox("Tem certeza que deseja forçar o disparo do automatismo TCT?",vbYesno+vbQuestion,"Confirma")=vbNo Then Exit Sub
				objTCT3.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
			Case 10 'Disparo Normalização TR3
				If MsgBox("Tem certeza que deseja forçar o disparo do automatismo TCT?",vbYesno+vbQuestion,"Confirma")=vbNo Then Exit Sub
				objTCT3.Item("DadosTCT").Item("PartidaNormalizaçãoTCT").WriteEx "Normalizar"
			Case 11 'Bloqueio Operador TR3
				objTCT3.BloqueioOperador = Not objTCT3.BloqueioOperador
			Case 12 'Reset TR3
				objTCT3.Item("DadosTCT").Item("Comandos").WriteEx "Reset"	
			Case 13 'Disparo Liberação TR4
				If MsgBox("Tem certeza que deseja forçar o disparo do automatismo TCT?",vbYesno+vbQuestion,"Confirma")=vbNo Then Exit Sub
				objTCT4.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
			Case 14 'Disparo Normalização TR4
				If MsgBox("Tem certeza que deseja forçar o disparo do automatismo TCT?",vbYesno+vbQuestion,"Confirma")=vbNo Then Exit Sub
				objTCT4.Item("DadosTCT").Item("PartidaNormalizaçãoTCT").WriteEx "Normalizar"
			Case 15 'Bloqueio Operador TR4
				objTCT4.BloqueioOperador = Not objTCT4.BloqueioOperador
			Case 16 'Reset TR4
				objTCT4.Item("DadosTCT").Item("Comandos").WriteEx "Reset"
		End Select 
	Else
		Msgbox "Automatismo ou transformador de referência não encontrado!!!"
	End If
End Sub

