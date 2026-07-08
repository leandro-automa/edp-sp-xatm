-----------------------
Documentação de Scripts
-----------------------
XATM_Disj_Red
Wed Jul  8 14:02:04 2026
-----------------------

<XATM_Disj_Red.Dados.CmdAbrirFechar:CmdAbrirFechar_CmdAbFe()>
Sub CmdAbrirFechar_CmdAbFe()
	Dim Cmd1
	Dim Cmd2
	Dim AbrirFlag
	Dim FecharFlag
	Dim CmdExecutadoFlag
	Dim ValorCmdAbrirSaida
	Dim ValorCmdFecharSaida
	Dim ValorCmd
	Dim FalhaComRele1
	Dim FalhaComRele2
	Dim Qualidade
	Dim CmdSBO1
	Dim CmdSBO2
	'############ VERIFICA TIPO DO COMANDO (ABRIR OU FECHAR) ###########

	ValorCmd = Value
	Qualidade = 192
	AbrirFlag = False
	FecharFlag = False
	
	If ValorCmd = 1 Then
		Set CmdSBO1 = XATM_Disj_Red.[1_CmdSBOAbrir]	
		Set CmdSBO2 = XATM_Disj_Red.[2_CmdSBOAbrir]
		Set Cmd1 = XATM_Disj_Red.[1_CmdAbrir_Saida]	
		Set Cmd2 = XATM_Disj_Red.[2_CmdAbrir_Saida]
		AbrirFlag = True
	else
		If ValorCmd = 2 Then
			Set CmdSBO1 = XATM_Disj_Red.[1_CmdSBOFechar]	
			Set CmdSBO2 = XATM_Disj_Red.[2_CmdSBOFechar]		
			Set Cmd1 = XATM_Disj_Red.[1_CmdFechar_Saida]	
			Set Cmd2 = XATM_Disj_Red.[2_CmdFechar_Saida]
			FecharFlag = True
		Else
			XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Valor do comando de entrada não confere"
			Exit Sub
		End If
	End If
	
	'############ VERIFICA FALHA DE COMUNICAÇÃO ###########
	If XATM_Disj_Red.HabilitaSimulador = False Then
		If CheckLink(XATM_Disj_Red.[1_PosAberto]) = False Then
			Set Pos = XATM_Disj_Red.[1_PosAberto]
			If Pos.Quality >= Qualidade Then
				FalhaComRele1 = False
			ElseIf CheckLink(XATM_Disj_Red.[1_PosFechado]) = False Then
				Set Pos = XATM_Disj_Red.[1_PosFechado]
				If Pos.Quality >= Qualidade Then
					FalhaComRele1 = False
				Else
					FalhaComRele1 = True
				End If
			Else
				FalhaComRele1 = True
			End If 
		ElseIf CheckLink(XATM_Disj_Red.[1_PosFechado]) = False Then
			Set Pos = XATM_Disj_Red.[1_PosFechado]
			If Pos.Quality >= Qualidade Then
				FalhaComRele1 = False
			Else
				FalhaComRele1 = True
			End If 
		Else
			FalhaComRele1 = True
		End If
		
		If CheckLink(XATM_Disj_Red.[2_PosAberto]) = False Then
			Set Pos = XATM_Disj_Red.[2_PosAberto]
			If Pos.Quality >= Qualidade Then
				FalhaComRele2 = False
			ElseIf CheckLink(XATM_Disj_Red.[2_PosFechado]) = False Then
				Set Pos = XATM_Disj_Red.[2_PosFechado]
				If Pos.Quality >= Qualidade Then
					FalhaComRele2 = False
				Else
					FalhaComRele2 = True
				End If
			Else
				FalhaComRele2 = True
			End If 
		ElseIf CheckLink(XATM_Disj_Red.[2_PosFechado]) = False Then
			Set Pos = XATM_Disj_Red.[2_PosFechado]
			If Pos.Quality >= Qualidade Then
				FalhaComRele2 = False
			Else
				FalhaComRele2 = True
			End If 
		Else
			FalhaComRele2 = True
		End If
	Else
		FalhaComRele1 = False
		FalhaComRele2 = False
	End If

	'############ DEFINIÇÃO DO DESTINO DO COMANDO ###########
	
	CmdExecutadoFlag = False
	ValorCmdAbrirSaida = XATM_Disj_Red.ValorCmdAbrir_Saida
	ValorCmdFecharSaida = XATM_Disj_Red.ValorCmdFechar_Saida
	
	Select Case XATM_Disj_Red.RelePrioritario
		Case 2
			If FalhaComRele2 = False Then
				If AbrirFlag = True Then
					If Parent.Item("Posição").Value = 2 Then
						If (XATM_Disj_Red.HabilitaSimulador = False) And (XATM_Disj_Red.Item("XATM_TipoApp").Tipo = "REAL") Then
							If Typename(CmdSBO2)="IOTag" Then
								CmdSBO2.WriteEx ValorCmdAbrirSaida
							End If
							Cmd2.WriteEx ValorCmdAbrirSaida
							CmdExecutadoFlag = True
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Abrir 2 - Valor = " & ValorCmdAbrirSaida
						Else
							If Parent.Item("FalhaCmd").Value = False Then
								Parent.Item("PosiçãoProvisória").WriteEx 1
							End If
							CmdExecutadoFlag = True
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Abrir - Valor = " & ValorCmdAbrirSaida & " via simulador"
						End If
					Else
						XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Comando Abrir 2 não executado - Intertravamento posição"
					End If
				End If
				If FecharFlag = True Then
					If Parent.Item("Posição").Value = 1 Then
						If (XATM_Disj_Red.HabilitaSimulador = False) And (XATM_Disj_Red.Item("XATM_TipoApp").Tipo = "REAL") Then
							If Typename(CmdSBO2)="IOTag" Then
								CmdSBO2.WriteEx ValorCmdFecharSaida
							End If						
							Cmd2.WriteEx ValorCmdFecharSaida
							CmdExecutadoFlag = True
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Fechar 2 - Valor = " & ValorCmdFecharSaida
						Else
							If Parent.Item("FalhaCmd").Value = False Then
								Parent.Item("PosiçãoProvisória").WriteEx 2
							End If
							CmdExecutadoFlag = True
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Fechar - Valor = " & ValorCmdFecharSaida & " via simulador"
						End If
					Else
						XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Comando Fechar 2 não executado - Intertravamento posição"
					End If
				End If
			Else
				If FalhaComRele1 = False Then
					If AbrirFlag = True Then
						If Parent.Item("Posição").Value = 2 Then
							If (XATM_Disj_Red.HabilitaSimulador = False) And (XATM_Disj_Red.Item("XATM_TipoApp").Tipo = "REAL") Then
								If Typename(CmdSBO1)="IOTag" Then
									CmdSBO1.WriteEx ValorCmdAbrirSaida
								End If								
								Cmd1.WriteEx ValorCmdAbrirSaida
								CmdExecutadoFlag = True
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Abrir 1 - Valor = " & ValorCmdAbrirSaida
							Else
								If Parent.Item("FalhaCmd").Value = False Then
									Parent.Item("PosiçãoProvisória").WriteEx 1
								End If
								CmdExecutadoFlag = True
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Abrir - Valor = " & ValorCmdAbrirSaida & " via simulador"
							End If
						Else
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Comando Abrir 1 não executado - Intertravamento posição"
						End If
					End If
					If FecharFlag = True Then
						If Parent.Item("Posição").Value = 1 Then
							If (XATM_Disj_Red.HabilitaSimulador = False) And (XATM_Disj_Red.Item("XATM_TipoApp").Tipo = "REAL") Then
								If Typename(CmdSBO1)="IOTag" Then
									CmdSBO1.WriteEx ValorCmdFecharSaida
								End If								
								Cmd1.WriteEx ValorCmdFecharSaida
								CmdExecutadoFlag = True
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Fechar 1 - Valor = " & ValorCmdFecharSaida
							Else
								If Parent.Item("FalhaCmd").Value = False Then
									Parent.Item("PosiçãoProvisória").WriteEx 2
								End If
								CmdExecutadoFlag = True
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Fechar - Valor = " & ValorCmdFecharSaida & " via simulador"
							End If
						Else
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Comando Fechar 1 não executado - Intertravamento posição"
						End If
					End If
				Else
					XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Nenhum relé disponivel para comando"
				End If
			End If
			
		Case Else
			If FalhaComRele1 = False Then
				If AbrirFlag = True Then
					If Parent.Item("Posição").Value = 2 Then
						If (XATM_Disj_Red.HabilitaSimulador = False) And (XATM_Disj_Red.Item("XATM_TipoApp").Tipo = "REAL") Then
							If Typename(CmdSBO1)="IOTag" Then
								CmdSBO1.WriteEx ValorCmdAbrirSaida
							End If							
							Cmd1.WriteEx ValorCmdAbrirSaida
							CmdExecutadoFlag = True
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Abrir 1 - Valor = " & ValorCmdAbrirSaida
						Else
							If Parent.Item("FalhaCmd").Value = False Then
								Parent.Item("PosiçãoProvisória").WriteEx 1
							End If
							CmdExecutadoFlag = True
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Abrir - Valor = " & ValorCmdAbrirSaida & " via simulador"
						End If
					Else
						XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Comando Abrir 1 não executado - Intertravamento posição"
					End If
				End If
				If FecharFlag = True Then
					If Parent.Item("Posição").Value = 1 Then
						If (XATM_Disj_Red.HabilitaSimulador = False) And (XATM_Disj_Red.Item("XATM_TipoApp").Tipo = "REAL") Then
							If Typename(CmdSBO1)="IOTag" Then
								CmdSBO1.WriteEx ValorCmdFecharSaida
							End If							
							Cmd1.WriteEx ValorCmdFecharSaida
							CmdExecutadoFlag = True
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Fechar 1 - Valor = " & ValorCmdFecharSaida
						Else
							If Parent.Item("FalhaCmd").Value = False Then
								Parent.Item("PosiçãoProvisória").WriteEx 2
							End If
							CmdExecutadoFlag = True
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Fechar - Valor = " & ValorCmdFecharSaida & " via simulador"
						End If
					Else
						XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Comando Fechar 1 não executado - Intertravamento posição"
					End If
				End If
			Else
				If FalhaComRele2 = False Then
					If AbrirFlag = True Then
						If Parent.Item("Posição").Value = 2 Then
							If (XATM_Disj_Red.HabilitaSimulador = False) And (XATM_Disj_Red.Item("XATM_TipoApp").Tipo = "REAL") Then
								If Typename(CmdSBO2)="IOTag" Then
									CmdSBO2.WriteEx ValorCmdAbrirSaida
								End If								
								Cmd2.WriteEx ValorCmdAbrirSaida
								CmdExecutadoFlag = True
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Abrir 2 - Valor = " & ValorCmdAbrirSaida
							Else
								If Parent.Item("FalhaCmd").Value = False Then
									Parent.Item("PosiçãoProvisória").WriteEx 1
								End If
								CmdExecutadoFlag = True
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Abrir - Valor = " & ValorCmdAbrirSaida & " via simulador"
							End If
						Else
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Comando Abrir 2 não executado - Intertravamento posição"
						End If
					End If
					If FecharFlag = True Then
						If Parent.Item("Posição").Value = 1 Then
							If (XATM_Disj_Red.HabilitaSimulador = False) And (XATM_Disj_Red.Item("XATM_TipoApp").Tipo = "REAL") Then
								If Typename(CmdSBO2)="IOTag" Then
									CmdSBO2.WriteEx ValorCmdFecharSaida
								End If								
								Cmd2.WriteEx ValorCmdFecharSaida
								CmdExecutadoFlag = True
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Fechar 2 - Valor = " & ValorCmdFecharSaida
							Else
								If Parent.Item("FalhaCmd").Value = False Then
									Parent.Item("PosiçãoProvisória").WriteEx 2
								End If
								CmdExecutadoFlag = True
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Saiu Comando Fechar - Valor = " & ValorCmdFecharSaida & " via simulador"
							End If
						Else
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Comando Fechar 2 não executado - Intertravamento posição"
						End If
					End If
				Else
					XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Nenhum relé disponivel para comando"
				End If
			End If
	End Select	
	If CmdExecutadoFlag = True Then
		If Parent.Item("CmdEmExecução").Value <> 2 Then
			Parent.Item("CmdEmExecução").WriteEx 2
		End If
	End If
End Sub

Function CheckLink(Sinal)
	Dim Aux
	On Error Resume Next
		Aux = Sinal.Value
		Aux = Sinal.Quality
		If Err.Number <> 0 Then
			Err.Clear
			CheckLink = True
		Else
			CheckLink = False
		End If
	On Error Goto 0
End Function

Sub Fim()
End Sub

<XATM_Disj_Red.Dados.DisparoTripBF:DisparoTripBF_Trip50BF()>
Sub DisparoTripBF_Trip50BF()
	If XATM_Disj_Red.Trip50BF.Quality >= 192 Then
		If (XATM_Disj_Red.Trip50BF.Value = 1) Or (XATM_Disj_Red.Trip50BF.Value = 2) Then
			For Each Atm in Application.GetObject("autoDados.Automatismo")
				Select Case TypeName(Atm)
					Case "XATM_RASEBF"
						If (Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value <> "GISBT") Then
							Atm.Item("Dados_RaseBF").Item("Disjuntor").WriteEx XATM_Disj_Red.Id
							Atm.Item("Dados_RaseBF").Item("DisparoRaseBF").WriteEx "DISJ"
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" &  XATM_Disj_Red.Name & "] - Atuação do Trip 50BF - Disparo RaseBF"
						End If
				End Select
			Next		
		End If
	Else
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" &  XATM_Disj_Red.Name & "] - Sinal de Trip 50BF com Qualidade Baixa"
	End If
End Sub

<XATM_Disj_Red.Dados.DisparoTripBF:[DisparoTripBF_ModoSimulação]()>
Sub [DisparoTripBF_ModoSimulação]()
	If XATM_Disj_Red.HabilitaSimulador = True Then
		If (Value = True) Then
			For Each Atm in Application.GetObject("autoDados.Automatismo")
				Select Case TypeName(Atm)
					Case "XATM_RASEBF"
						If (Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value <> "GISBT") Then
							Atm.Item("Dados_RaseBF").Item("Disjuntor").WriteEx XATM_Disj_Red.Id
							Atm.Item("Dados_RaseBF").Item("DisparoRaseBF").WriteEx "DISJ"
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" &  XATM_Disj_Red.Name & "] - Atuação do Trip 50BF - Disparo RaseBF"
						End If
				End Select
			Next		
		End If
	End If	
End Sub

<XATM_Disj_Red.Dados.FalhaCmd:FalhaCmd_OnStartRunning()>
Sub FalhaCmd_OnStartRunning()
	WriteEx(False)	
End Sub

<XATM_Disj_Red.Dados.Reset:Reset_Reset()>
Sub Reset_Reset()
	If Value >= 1 Then
		Parent.Item("TimerCmd").Value = -1
		Parent.Item("CmdEmExecução").Value = 0
		Value = 0
	End If
End Sub

<XATM_Disj_Red.Dados.ResetAutomatismos:ResetAutomatismos_Reset()>
Sub ResetAutomatismos_Reset()
	If Value = 1 Then
		Parent.Item("TimerCmd").Value = -1
		Parent.Item("CmdEmExecução").Value = 0
		XATM_Disj_Red.FalhaCmdAbrir = False
		XATM_Disj_Red.FalhaCmdFechar = False
		Parent.Item("PosiçãoMem").WriteEx Parent.Item("Posição").Value
		Parent.Item("SeloIAD").WriteEx False
		Value = 0
	Else
		If Value = 10 Then
			Parent.Item("TimerSF6").WriteEx -1
		End If
	End If
End Sub

<XATM_Disj_Red.Dados.SeloIAD:SeloIAD_Functions()>
Sub SeloIAD_Functions()

End Sub

Function BuscaAutomatismo(TipoAutomatismo, IdEquipamento, byRef Atm_Flag)
	Atm_Flag = False

	For Each Atm in Application.GetObject("autoDados.Automatismo")
		If TypeName(Atm) = TipoAutomatismo Then
			Select Case TipoAutomatismo
				Case "XATM_TCT"
					IdTransformador = -999
					On Error Resume Next
						IdTransformador = Atm.Transformador.Id
					On Error GoTo 0
					If IsNumeric(IdTransformador) And IsNumeric(IdEquipamento) Then
						If IdTransformador = IdEquipamento Then
							Atm_Flag = True
						End If	
					End If	
				Case "XATM_IAD","XATM_TCD"
					IdDisjuntor = -999
					On Error Resume Next
						IdDisjuntor = Atm.Disjuntor.Id
					On Error GoTo 0
					If IsNumeric(IdDisjuntor) And IsNumeric(IdEquipamento) Then
						If IdDisjuntor = IdEquipamento Then
							Atm_Flag = True
						End If
					End If	
				Case Else
					Atm_Flag = True
			End Select		
		End If
		If Atm_Flag Then
			Exit For
		End If
	Next
	
	If Atm_Flag Then
		Set BuscaAutomatismo = Atm
	Else
		Set BuscaAutomatismo = Nothing
	End If
End Function



Sub Fim()	
End Sub

<XATM_Disj_Red.Dados.SeloIAD:SeloIAD_PartidaIAD()>
Sub SeloIAD_PartidaIAD()
	
	Dim ObjRase
	Dim ObjRaseD
	Dim ObjIAD
	Dim SF62
	Dim RaseFlag
	Dim RaseDFlag
	Dim IADFlag
	Dim DJAux
	
	Set SF62 = Parent.Item("SF62")
	RaseFlag = False
	RaseDFlag = False
	IADFlag = False

	If (SF62.Value = True) Or (SF62.Value = 1) Then
		If (XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 1) Or (((XATM_Disj_Red.Id = 22) Or (XATM_Disj_Red.Id = 23)) And (Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value <> "4BR2DJ")) Or (XATM_Disj_Red.Id = 24) Or (XATM_Disj_Red.Id = 25) Or (XATM_Disj_Red.Id = 26) Or (XATM_Disj_Red.Id = 40) Then
			Set ObjRase = BuscaAutomatismo("XATM_RASET",-1,RaseFlag)
			Set ObjRaseD = BuscaAutomatismo("XATM_RASED",-1,RaseDFlag)
			Set ObjIAD = BuscaAutomatismo("XATM_IAD",XATM_Disj_Red.Id,IADFlag)
			
			If (Application.GetObject("autoDados.Automatismo.Aux.TipoEntrada").Value = "1DJGBY") _
					Or (Application.GetObject("autoDados.Automatismo.Aux.TipoEntrada").Value = "2SD1DJGBY") _
					Or (Application.GetObject("autoDados.Automatismo.Aux.TipoEntrada").Value = "1DJBY2DJ") _
					Or ((Application.GetObject("autoDados.Automatismo.Aux.TipoEntrada").Value = "3DJBP") _
						And XATM_Disj_Red.Id = 1) Then
				If IADFlag = True Then
					If Parent.Item("Posição").Value = 1 Then
						'$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
						ObjIAD.Item("DadosIAD").Item("PartidaIAD").WriteEx "IAD"
						XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Partida do Automatismo " & ObjIAD.Name
						Value = False
					End If
				End If
				Exit Sub
			End If
			
			If RaseFlag = True Then
				If (ObjRase.AutomatismoEmAndamento = False) Or (ObjRase.Enable = False) Then
					'If (ObjRase.Item("Timers").Item("BloqueioAutomatico").Output = True) Or (ObjRase.BloqueioOperador = True) Or (ObjRase.Item("Timers").Item("PreCondicoes").Output = False) Or (ObjRase.BloqueioGeral = True) Then
						If IADFlag = True Then
							If Parent.Item("Posição").Value = 1 Then
								'$CausaDisparo = 1o est. SF6 atuado por 10s, 2o est. SF6 atuado e RaseT com algum bloqueio
								ObjIAD.Item("DadosIAD").Item("PartidaIAD").WriteEx "IAD"
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Partida do Automatismo " & ObjIAD.Name
								Value = False
							End If
						End If
					'End If
				Else
					If IADFlag = True Then
						If Parent.Item("Posição").Value = 1 Then
							'$CausaDisparo = 1o est. SF6 atuado por 10s, 2o est. SF6 atuado e RaseT em andamento
							ObjIAD.Item("DadosIAD").Item("PartidaIAD").WriteEx "IAD"
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Partida do Automatismo " & ObjIAD.Name
							Value = False
						End If
					End If
				End If
			End If
			If RaseDFlag = True Then
				If (ObjRaseD.AutomatismoEmAndamento = False) Or (ObjRaseD.Enable = False) Then
					'If (ObjRaseD.Item("Timers").Item("BloqueioAutomatico").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.Item("Timers").Item("PreCondicoes").Output = False) Or (ObjRaseD.BloqueioGeral = True) Then
						If IADFlag = True Then
							If Parent.Item("Posição").Value = 1 Then
								'$CausaDisparo = 1o est. SF6 atuado por 10s, 2o est. SF6 atuado e RaseD com algum bloqueio
								ObjIAD.Item("DadosIAD").Item("PartidaIAD").WriteEx "IAD"
								XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Partida do Automatismo " & ObjIAD.Name
								Value = False
							End If
						End If
					'End If
				Else
					If IADFlag = True Then
						If Parent.Item("Posição").Value = 1 Then
							'$CausaDisparo = 1o est. SF6 atuado por 10s, 2o est. SF6 atuado e RaseD em andamento
							ObjIAD.Item("DadosIAD").Item("PartidaIAD").WriteEx "IAD"
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Partida do Automatismo " & ObjIAD.Name
							Value = False
						End If
					End If
				End If
			End If
		End If
	Else
		Value = False
	End If
End Sub


Function VerificaLink(Sinal)
	Dim Aux
	On Error Resume Next
		Aux = Sinal.Id
		If Err.Number <> 0 Then
			Err.Clear
			VerificaLink = True
		Else
			VerificaLink = False
		End If
	On Error Goto 0
End Function


Sub Fim()	
	
End Sub

<XATM_Disj_Red.Dados.SeloIAD:SeloIAD_Valor()>
Sub SeloIAD_Valor()
	If Value =  True Then
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Atuado Selo IAD - Automatismo prédisposto a iniciar"
	Else
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Desatuado Selo IAD"
	End If	
End Sub

<XATM_Disj_Red.Dados.SinaisSF6.SF61_1:SF61_1_Disparo()>
Sub SF61_1_Disparo()
	If Quality >= 192 Then
		If XATM_Disj_Red.SF61_Rele1.Value = XATM_Disj_Red.ValorSF6Atuado Then
			Parent.Parent.Item("SF61").WriteEx 1
		Else
			Parent.Parent.Item("SF61").WriteEx 0
		End If
	Else
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - SF6 1º estágio não executado (Relé 1) - Baixa qualidade do sinal."
	End If	
End Sub

<XATM_Disj_Red.Dados.SinaisSF6.SF61_2:SF61_2_Disparo()>
Sub SF61_2_Disparo()
	If Quality >= 192 Then
		If XATM_Disj_Red.SF61_Rele2.Value = XATM_Disj_Red.ValorSF6Atuado Then
			Parent.Parent.Item("SF61").WriteEx 1
		Else
			Parent.Parent.Item("SF61").WriteEx 0
		End If
	Else
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - SF6 1º estágio não executado (Relé 2) - Baixa qualidade do sinal."
	End If	
End Sub

<XATM_Disj_Red.Dados.SinaisSF6.SF62_1:SF62_1_Disparo()>
Sub SF62_1_Disparo()
	If Quality >= 192 Then
		If XATM_Disj_Red.SF62_Rele1.Value = XATM_Disj_Red.ValorSF6Atuado Then
			Parent.Parent.Item("SF62").WriteEx 1
		Else
			Parent.Parent.Item("SF62").WriteEx 0
		End If
	Else
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - SF6 2º estágio não executado (Relé 1) - Baixa qualidade do sinal."
	End If		
End Sub

<XATM_Disj_Red.Dados.SinaisSF6.SF62_2:SF62_2_Disparo()>
Sub SF62_2_Disparo()
	If Quality >= 192 Then
		If XATM_Disj_Red.SF62_Rele2.Value = XATM_Disj_Red.ValorSF6Atuado Then
			Parent.Parent.Item("SF62").WriteEx 1
		Else
			Parent.Parent.Item("SF62").WriteEx 0
		End If
	Else
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - SF6 2º estágio não executado (Relé 2) - Baixa qualidade do sinal."
	End If	
End Sub

<XATM_Disj_Red.Dados.TimerCmd:TimerCmd_Timer()>
Sub TimerCmd_Timer()
	Dim TipoCmd
	
	TipoCmd = Parent.Item("CmdAbrirFechar").Value
	If TipoCmd = 1 Then
		If Parent.Item("Posição").Value = 1 Then
			Value = -1
			Parent.Item("CmdEmExecução").WriteEx 3
			Exit Sub
		ElseIf (Value = CByte(XATM_Disj_Red.TimeOutCmd /2)) Then
			Parent.Item("CmdAbrirFechar").WriteEx 1
			XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Reenvio do comando de abrir"
		End If
	Else
		If TipoCmd = 2 Then
			If Parent.Item("Posição").Value = 2 Then
				Value = -1
				Parent.Item("CmdEmExecução").WriteEx 3
				Exit Sub
			ElseIf (Value = CByte(XATM_Disj_Red.TimeOutCmd /2)) Then
				Parent.Item("CmdAbrirFechar").WriteEx 2
				XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Reenvio do comando de fechar"
			End If
		Else
			XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Temporizador - Tipo do comando inválido"
			Value = -1
			Exit Sub
		End If
	End if

	If Value >= XATM_Disj_Red.TimeOutCmd Then
		If Value >= (3 * XATM_Disj_Red.TimeOutCmd) Then
			Parent.Item("CmdEmExecução").WriteEx 0
			Value = -1
		Else
			If Parent.Item("CmdEmExecução").Value <> 1 Then
				Parent.Item("CmdEmExecução").WriteEx 1
			End If
		End If
		Select Case TipoCmd
			Case 1
				XATM_Disj_Red.FalhaCmdAbrir = True
			Case 2
				XATM_Disj_Red.FalhaCmdFechar = True
		End Select
		Value = -1
	End If
	
	If (Value >= 0) Then
		Value = Value + 1
	End If	
End Sub

<XATM_Disj_Red.Dados.TimerCmdIndividual:TimerCmdIndividual_Delay()>
Sub TimerCmdIndividual_Delay()

	If Value > 1 Then
		
		Value = -1
		
		If Parent.Item("Posição").Value	= 1 Then
			
			ReleaseTransformerMastership
			
		End If
		
		Exit Sub
			
	End If
	
	Value = Value + 1
	
End Sub

Sub ReleaseTransformerMastership()
	' Envio de comando de individual para os transformadores pelo evento de abertura do disjuntor
	
	Dim breakerId
	breakerId = XATM_Disj_Red.Id

	' ============================
	' Interligadores -> id na faixa 700
	' e geral de baixa -> id termina com 20
	' ============================
	If (breakerId \ 100 = 7) Or (breakerId Mod 100 = 20) Or (breakerId Mod 100 = 10) Or breakerId = 920 Then
		SendIndividualCommand
	End If

End Sub

Sub SendIndividualCommand()
	
	Dim transformerType, busbarType
	transformerType = Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
	busbarType = Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value
	
	Dim layoutKey
	layoutKey = transformerType & "_" & busbarType
	
	Dim targetTransformersId
	Set targetTransformersId = CreateObject("Scripting.Dictionary")

	Dim bayTransformerId
    bayTransformerId = (XATM_Disj_Red.Id \ 100) * 100

	Dim breaker

	' ===========================
	' Identificação dos transformadores que podem receber o comando de individual
	' ===========================
	Select Case layoutKey
	
		Case "3TRSC_2BR", "3TRSCBT_2BR"
			
			' Paralelismo entre TR02 e TR03
			targetTransformersId.Add 200, Empty 
			targetTransformersId.Add 300, Empty
		
		Case "3TRSC_3BR3DJ", "3TRSCBT_3BR3DJ"
			
			Select Case XATM_Disj_Red.Id
				Case 700

					targetTransformersId.Add 100, Empty 
					targetTransformersId.Add 200, Empty

				Case 710

					targetTransformersId.Add 200, Empty 
					targetTransformersId.Add 300, Empty

				Case 750

					targetTransformersId.Add 300, Empty 
					targetTransformersId.Add 100, Empty

				Case Else
					
					targetTransformersId.Add bayTransformerId, Empty
					
					For Each breaker In Application.GetObject("autoDados.[Subestação].Barra13_8kV")
						
						If TypeName(breaker) = "XATM_Disj_Red" Then
							
							If breaker.Item("Dados").Item("Posição").Value = 2 Then
								
								If bayTransformerId = 100 And breaker.Id = 700 Then             
									
									' ----> T1|T2 <----
									If Not targetTransformersId.Exists(200) Then targetTransformersId.Add 200, Empty
									
								ElseIf bayTransformerId = 100 And breaker.Id = 750 Then         
									
									' ----> T1|T3 <----
									If Not targetTransformersId.Exists(300) Then targetTransformersId.Add 300, Empty
									
								ElseIf bayTransformerId = 200 And breaker.Id = 700 Then         
									
									' ----> T2|T1 <----
									If Not targetTransformersId.Exists(100) Then targetTransformersId.Add 100, Empty
									
								ElseIf bayTransformerId = 200 And breaker.Id = 710 Then         
									
									' ----> T2|T3 <----
									If Not targetTransformersId.Exists(300) Then targetTransformersId.Add 300, Empty
									
								ElseIf bayTransformerId = 300 And breaker.Id = 750 Then         
									
									' ----> T3|T1 <----
									If Not targetTransformersId.Exists(100) Then targetTransformersId.Add 100, Empty
									
								ElseIf bayTransformerId = 300 And breaker.Id = 710 Then         
									
									' ----> T3|T2 <----
									If Not targetTransformersId.Exists(200) Then targetTransformersId.Add 200, Empty
								
								End If
							End If
						
						End If
						
					Next
			End Select

		Case "3TRDJATSCBTACBT_2SINGLEBB"

			' ===========================
			' Sempre adiciona o transformador do próprio vão. O DJ920 pertence ao TR200.
			' ===========================
			
			Select Case bayTransformerId

				Case 100
					
					targetTransformersId.Add bayTransformerId, Empty
					
					' Paralelismo T1 <-> T2
					For Each breaker In Application.GetObject("autoDados.[Subestação].Transformador")

						If TypeName(breaker) = "XATM_Disj_Red" Then

							If breaker.Id = 220 And breaker.Item("Dados").Item("Posição").Value = 2 Then
								targetTransformersId.Add 200, Empty
								Exit For
							End If

						End If

					Next

				Case 200
					
					targetTransformersId.Add bayTransformerId, Empty
					
					For Each breaker In Application.GetObject("autoDados.[Subestação].Transformador")

						If TypeName(breaker) = "XATM_Disj_Red" Then

							If breaker.Item("Dados").Item("Posição").Value = 2 Then

								Select Case breaker.Id

									Case 120
										If Not targetTransformersId.Exists(100) Then
											targetTransformersId.Add 100, Empty
										End If

									Case 920
										If Not targetTransformersId.Exists(300) Then
											
											' O geral de baixa ID=220 do T200 nunca deve impactar no T300 neste layout
											If XATM_Disj_Red.Id <> 220 Then
												targetTransformersId.Add 300, Empty
											End If
											
										End If

								End Select

							End If

						End If

					Next
					
				Case 900
					
					' Disjuntor de baixa (ID=920) que conecta o T200 a barra do T300 
					
					targetTransformersId.Add 300, Empty
					
					For Each breaker In Application.GetObject("autoDados.[Subestação].Transformador")

						If TypeName(breaker) = "XATM_Disj_Red" Then

							If breaker.Id = 220 And breaker.Item("Dados").Item("Posição").Value = 1 Then
						
								' Só é seguro enviar comando de Indidividual para o T200 caso o GBT esteja aberto.
								' Caso contrário, é possível que estejamos lidando com um evento de abertura de 
								' abertura do DJ920 e o T100/T200 estejam operando em paralelo.
								
								targetTransformersId.Add 200, Empty			
				
							End If
						
						End If
						
					Next
					
				Case 300
					
					targetTransformersId.Add bayTransformerId, Empty
					
					' Paralelismo T3 <-> T2
					For Each breaker In Application.GetObject("autoDados.[Subestação].Transformador")

						If TypeName(breaker) = "XATM_Disj_Red" Then

							If breaker.Id = 920 And breaker.Item("Dados").Item("Posição").Value = 2 Then
								targetTransformersId.Add 200, Empty
								Exit For
							End If

						End If

					Next
					
			End Select

		Case Else
			
			' ===========================
			' DEFAULT
			' ===========================

			If XATM_Disj_Red.Id = 700 Then
				
				targetTransformersId.Add 100, Empty
				targetTransformersId.Add 200, Empty

			ElseIf XATM_Disj_Red.Id = 720 Then
				
				targetTransformersId.Add 300, Empty
				targetTransformersId.Add 400, Empty

			Else

				Select Case bayTransformerId
					Case 100, 200
						targetTransformersId.Add 100, Empty
						targetTransformersId.Add 200, Empty 
					Case 300, 400
						targetTransformersId.Add 300, Empty
						targetTransformersId.Add 400, Empty 
				End Select

			End If
		
	End Select

	' ===========================
	' Envio do comando de individual para o transformador em mestre (se aplicável)
	' ===========================
	Dim transformer
	For Each transformer In Application.GetObject("autoDados.[Subestação].Transformador")
	
		If TypeName(transformer) = "XATM_Trafo" Then
		
			If targetTransformersId.Exists(transformer.Id) Then
				
				If transformer.TrafoEmMestre Then
				
					If transformer.HabilitaSimulador Then
					   
						transformer.TrafoEmMestre = False
						XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Enviado via simulador comando de individual para [" & transformer.Name & "]"

					Else

						Dim commandTag
						Set commandTag = Nothing
						On Error Resume Next
						Set commandTag = transformer.CmdIndividual
						On Error GoTo 0

						If commandTag Is Nothing Then
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Comando de individual não encontrado para [" & transformer.Name & "]"
							Exit Sub
						End If

						commandTag.WriteEx transformer.ValorCmdIndividual
						FinalizePulseIfNeeded transformer, commandTag, transformer.ValorCmdIndividual						
						XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Enviado comando de individual para [" & transformer.Name & "]"

					End If

				End If

			End If

		End If
	
	Next

End Sub

Sub FinalizePulseIfNeeded(transformer, commandTag, commandValue)

	If transformer.PulsarCmdMestreIndividual = True Then
		
		Select Case commandValue
			
			Case 1	: commandTag.WriteEx 0
			Case 2	: commandTag.WriteEx 1
			Case 65	:commandTag.WriteEx 129
			
		End Select
		
	End If
		
End Sub

<XATM_Disj_Red.Dados.TimerCmdIndividual:TimerCmdIndividual_WhenOpen()>
Sub TimerCmdIndividual_WhenOpen()
	
	' Inicia o delay
	Value = 0
		
End Sub

<XATM_Disj_Red.Dados.TimerPos:TimerPos_Timer()>
Sub TimerPos_Timer()
	Dim DisjPos
	
	DisjPos = Parent.Item("Posição").Value
	
	If Value = 0 Then
		If (DisjPos = 1) Or (DisjPos = 2) Then
			Parent.Item("PosiçãoMem").Value = DisjPos
			Value = -1
		End If
	Else
		Value = Value - 1
	End If	
	If Parent.Item("PosiçãoMem").Value = DisjPos Then
		Value = -1
	End If
End Sub

<XATM_Disj_Red.Dados.TimerSF6:TimerSF6_Disparos()>
Sub TimerSF6_Disparos()
    Dim TR
    Dim DJ
    Dim TCT_TR1
    Dim TCT_TR2
    Dim TCT_TR3
    Dim TCT_TR4
    Dim TCT_Flag
    Dim TCD
    Dim TCD_Flag
    Dim Secc1
    Dim Secc2
    Dim SeccBP
    
    TCT_Flag = False
    TCD_Flag = False
    Secc1 = 0
    Secc2 = 0
    SeccBP = 0
    
    Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
		Case "4TRDJSC", "4TRSC", "4TR2D2JSC"
			Select Case XATM_Disj_Red.Id
				Case 24, 120
					Set TCT_TR1 = BuscaAutomatismo("XATM_TCT",100,TCT_Flag)
				Case 25, 220
					Set TCT_TR2 = BuscaAutomatismo("XATM_TCT",200,TCT_Flag)
				Case 26, 320
					Set TCT_TR3 = BuscaAutomatismo("XATM_TCT",300,TCT_Flag)
				Case 27, 420
					Set TCT_TR4 = BuscaAutomatismo("XATM_TCT",400,TCT_Flag)
			End Select
		Case "3TR3DJ"
			Select Case XATM_Disj_Red.Id
				Case 420, 120
					Set TCT_TR1 = BuscaAutomatismo("XATM_TCT",100,TCT_Flag)
				Case 24, 220
					Set TCT_TR2 = BuscaAutomatismo("XATM_TCT",200,TCT_Flag)
				Case 25, 320
					Set TCT_TR3 = BuscaAutomatismo("XATM_TCT",300,TCT_Flag)
			End Select
		Case "3TRSC", "3TRSCBT"
			If Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "2BR" Then
				Select Case XATM_Disj_Red.Id
					Case 220
						Set TCT_TR1 = BuscaAutomatismo("XATM_TCT",200,TCT_Flag)
					Case 320
						Set TCT_TR2 = BuscaAutomatismo("XATM_TCT",300,TCT_Flag)
				End Select
			Else
				Select Case XATM_Disj_Red.Id
					Case 110, 120
						Set TCT_TR1 = BuscaAutomatismo("XATM_TCT",100,TCT_Flag)
					Case 210, 220
						Set TCT_TR2 = BuscaAutomatismo("XATM_TCT",200,TCT_Flag)
					Case 310, 320
						Set TCT_TR3 = BuscaAutomatismo("XATM_TCT",300,TCT_Flag)
					Case 410, 420
						Set TCT_TR4 = BuscaAutomatismo("XATM_TCT",400,TCT_Flag)
				End Select
			End If
		Case Else
			Select Case XATM_Disj_Red.Id
                Case 110, 120
                    Set TCT_TR1 = BuscaAutomatismo("XATM_TCT",100,TCT_Flag)
                Case 210, 220
                    Set TCT_TR2 = BuscaAutomatismo("XATM_TCT",200,TCT_Flag)
                Case 310, 320
                    Set TCT_TR3 = BuscaAutomatismo("XATM_TCT",300,TCT_Flag)
                Case 410, 420
                    Set TCT_TR4 = BuscaAutomatismo("XATM_TCT",400,TCT_Flag)
            End Select
	End Select

    Set TCD = BuscaAutomatismo("XATM_TCD",XATM_Disj_Red.Id,TCD_Flag)
    Set TCT = BuscaAutomatismo("XATM_TCT",(XATM_Disj_Red.Id \ 100) * 100, TCT_Flag)
    
    Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoEntrada").Value
        Case "2SD1DJGBY"
            If TCD_Flag = True Then
                For Each Equip in Application.GetObject("autoDados.[Subestação].Entrada")
                    Select Case XATM_Disj_Red.Id
                        Case 0
                            If Equip.Id = 7 Then
                                If TypeName(Equip) = "XATM_Secc" Then
                                    Secc1 = Equip.Item("Dados").Item("Posição").Value
                                Else
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Equipamento de ID 7 com Tipo inválido - Bloqueio do Automatismo"
                                    Main_BloqueioGeral
                                End If
                            End If
                            If Equip.Id = 8 Then
                                If TypeName(Equip) = "XATM_Secc" Then
                                    Secc2 = Equip.Item("Dados").Item("Posição").Value
                                Else
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Equipamento de ID 8 com Tipo inválido - Bloqueio do Automatismo"
                                    Main_BloqueioGeral
                                End If
                            End If
                            If Equip.Id = 9 Then
                                If TypeName(Equip) = "XATM_Secc" Then
                                    SeccBP = Equip.Item("Dados").Item("Posição").Value
                                Else
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Equipamento de ID 9 com Tipo inválido - Bloqueio do Automatismo"
                                    Main_BloqueioGeral
                                End If
                            End If  
                    End Select
                Next
                If Parent.Item("Posição").Value = 2 Then
                    If (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 2) Then
                        '$CausaDisparo = 1o est. SF6 atuado por 60s
                        TCD.Item("Dados").Item("PartidaLiberacaoTCD").WriteEx "Liberar"
                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD - SF6 1º atuado por mais de 60s"    
                    Else
                        If (Secc1 = 2) And (Secc2 = 2) And (SeccBP = 1) Then
                            '$CausaDisparo = 1o est. SF6 atuado por 60s
                            TCD.Item("Dados").Item("PartidaLiberacaoTCD").WriteEx "Liberar"
                            XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD - SF6 1º atuado por mais de 60s"
                        Else
                            XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - TCD não disparada - Seccionadora do disjuntor aberta"
                        End If
                    End If  
                Else
                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD Liberação não executado - Disjuntor inicialmente aberto"
                End If
            Else
                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo TCD não encontrado - SF6 1º atuado por mais de 60s"
            End If
        Case "2SECC"
            If Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "2TRDJ" Then
                If TCT_Flag = True Then
                    If Parent.Item("Posição").Value = 2 Then
                        Select Case XATM_Disj_Red.Id
                            Case 110, 120, 210, 220
                                '$CausaDisparo = 1o est. SF6 atuado por 60s
                                TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                        End Select
                    Else
                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação não executado - Disjuntor inicialmente aberto"
                    End If
                End If
            End If
        Case "2DJG", "1DJBY2DJ", "2DJGIB2", "2CHDJ", "2DJ2SCAT"
            If ((Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "4TRDJSC") _
                    Or (Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "4TRSC") _
                    Or (Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "2TRDJ") _
                    Or (Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "3TR3DJ") _
                    Or (Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "3TRSC") _
                    Or (Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "3TRSCBT") _
                    Or (Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "4TR2D2JSC")) Then
                If TCD_Flag = True Then
                    If (Parent.Item("Posição").Value = 2) And ((XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10)) Then
                            '$CausaDisparo = 1o est. SF6 atuado por 60s
                            TCD.Item("Dados").Item("PartidaLiberacaoTCD").WriteEx "Liberar"
                            XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD - SF6 1º atuado por mais de 60s"
                    Else
                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD Liberação não executado - Disjuntor inicialmente aberto"
                    End If
                Else
                    ' XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo TCD não encontrado - SF6 1º atuado por mais de 60s"
                End If
            End If

            If TCT_Flag = True Then
                If Parent.Item("Posição").Value = 2 Then
                    Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
                        Case "3TR3DJ"
                            Select Case XATM_Disj_Red.Id
                                Case 420
                                    '$CausaDisparo = 1o est. SF6 atuado por 60s
                                    TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                Case 24
                                    '$CausaDisparo = 1o est. SF6 atuado por 60s
                                    TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                Case 25
                                    '$CausaDisparo = 1o est. SF6 atuado por 60s
                                    TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                            End Select
                        Case "3TRSC", "3TRSCBT"
                            If Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "2BR" Then
                                Select Case XATM_Disj_Red.Id
                                    Case 220
                                        '$CausaDisparo = 1o est. SF6 atuado por 60s
                                        TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                    Case 320
                                        '$CausaDisparo = 1o est. SF6 atuado por 60s
                                        TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                End Select
                            Else
                                Select Case XATM_Disj_Red.Id
                                    Case 120
                                        '$CausaDisparo = 1o est. SF6 atuado por 60s
                                        TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                    Case 220
                                        '$CausaDisparo = 1o est. SF6 atuado por 60s
                                        TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                    Case 320
                                        '$CausaDisparo = 1o est. SF6 atuado por 60s
                                        TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                    Case 420
                                        '$CausaDisparo = 1o est. SF6 atuado por 60s
                                        TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                End Select
                            End If
                        Case Else
                            Select Case XATM_Disj_Red.Id
                                Case 24, 320, 120, 110, 310
                                    '$CausaDisparo = 1o est. SF6 atuado por 60s
                                    TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                Case 25, 420, 220, 210, 410
                                    '$CausaDisparo = 1o est. SF6 atuado por 60s
                                    TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                Case 26
                                    '$CausaDisparo = 1o est. SF6 atuado por 60s
                                    TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                                Case 27
                                    '$CausaDisparo = 1o est. SF6 atuado por 60s
                                    TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                            End Select
                    End Select
                
                Else
                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação não executado - Disjuntor inicialmente aberto"
                End If
                'End If
            Else
                If (TCD_Flag = True) And (Application.GetObject("autoDados.Automatismo.Aux.TipoEntrada").Value <> "2CHDJ") Then
                    If Parent.Item("Posição").Value = 2 Then
                        '$CausaDisparo = 1o est. SF6 atuado por 60s
                        TCD.Item("Dados").Item("PartidaLiberacaoTCD").WriteEx "Liberar"
                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD - SF6 1º atuado por mais de 60s"
                    Else
                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD Liberação não executado - Disjuntor inicialmente aberto"
                    End If
                Else
                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo TCD não encontrado - SF6 1º atuado por mais de 60s"
                End If
            End If
        Case "2DJGIBGIS"
            If (Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "2TRDJ") And ((XATM_Disj_Red.Id = 1020) Or (XATM_Disj_Red.Id = 220) Or (XATM_Disj_Red.Id = 320) Or (XATM_Disj_Red.Id = 420)) Then
                If TCT_Flag = True Then
                    If Parent.Item("Posição").Value = 2 Then
                        Select Case XATM_Disj_Red.Id
                            Case 320, 120
                                '$CausaDisparo = 1o est. SF6 atuado por 60s
                                TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                            Case 420, 220
                                '$CausaDisparo = 1o est. SF6 atuado por 60s
                                TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
                                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
                        End Select
                    Else
                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação não executado - Disjuntor inicialmente aberto"
                    End If
                End If
            Else
                If TCD_Flag = True Then
                    If Parent.Item("Posição").Value = 2 Then
                            '$CausaDisparo = 1o est. SF6 atuado por 60s
                            TCD.Item("Dados").Item("PartidaLiberacaoTCD").WriteEx "Liberar"
                            XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD - SF6 1º atuado por mais de 60s"
                    Else
                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD Liberação não executado - Disjuntor inicialmente aberto"
                    End If
                Else
                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo TCD não encontrado - SF6 1º atuado por mais de 60s"
                End If
            End If
        Case "2DJGIB", "2DJG2IB"
            If TCD_Flag = True Then
                For Each Equip in Application.GetObject("autoDados.[Subestação].Entrada")
                    Select Case XATM_Disj_Red.Id
                        Case 1
                            If Equip.Id = 7 Then
                                If TypeName(Equip) = "XATM_Secc" Then
                                    Secc1 = Equip.Item("Dados").Item("Posição").Value
                                Else
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Equipamento de ID 7 com Tipo inválido - Bloqueio do Automatismo"
                                    Main_BloqueioGeral
                                End If
                            End If
                            If Equip.Id = 8 Then
                                If TypeName(Equip) = "XATM_Secc" Then
                                    Secc2 = Equip.Item("Dados").Item("Posição").Value
                                Else
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Equipamento de ID 7 com Tipo inválido - Bloqueio do Automatismo"
                                    Main_BloqueioGeral
                                End If
                            End If
                        Case 0
                            If Equip.Id = 9 Then
                                If TypeName(Equip) = "XATM_Secc" Then
                                    Secc1 = Equip.Item("Dados").Item("Posição").Value
                                Else
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Equipamento de ID 9 com Tipo inválido - Bloqueio do Automatismo"
                                    Main_BloqueioGeral
                                End If
                            End If
                            If Equip.Id = 10 Then
                                If TypeName(Equip) = "XATM_Secc" Then
                                    Secc2 = Equip.Item("Dados").Item("Posição").Value
                                Else
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Equipamento de ID 9 com Tipo inválido - Bloqueio do Automatismo"
                                    Main_BloqueioGeral
                                End If
                            End If
                    End Select
                Next
                If Parent.Item("Posição").Value = 2 Then
                    If (Secc1 = 2) And (Secc2 = 2) Then
                        '$CausaDisparo = 1o est. SF6 atuado por 60s
                        TCD.Item("Dados").Item("PartidaLiberacaoTCD").WriteEx "Liberar"
                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD - SF6 1º atuado por mais de 60s"
                    Else
                        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - TCD não disparada - Seccionadora do disjuntor aberta"
                    End If
                Else
                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD Liberação não executado - Disjuntor inicialmente aberto"
                End If
            Else
                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo TCD não encontrado - SF6 1º atuado por mais de 60s"
            End If
		Case "3DJBP"
            If TCD_Flag = True Then
                If (Parent.Item("Posição").Value = 2) And ((XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 1)) Then
                    '$CausaDisparo = 1o est. SF6 atuado por 60s
                    TCD.Item("Dados").Item("PartidaLiberacaoTCD").WriteEx "Liberar"
                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD - SF6 1º atuado por mais de 60s"
                Else
                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da TCD Liberação não executado - Disjuntor inicialmente aberto"
                End If
            Else
                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo TCD não encontrado - SF6 1º atuado por mais de 60s"
            End If
		Case Else
			Select Case XATM_Disj_Red.Id
				Case 120, 220, 320, 420, 110, 210, 310, 410
					If TCT_Flag = True Then
						If Parent.Item("Posição").Value = 2 Then
							'$CausaDisparo = 1o est. SF6 atuado por 60s
							TCT.Item("DadosTCT").Item("PartidaLiberaçãoTCT").WriteEx "Liberar"
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação " & TCT.Name & " - SF6 1º atuado por mais de 60s"
						Else
							XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Disparo da Liberação não executado - Disjuntor inicialmente aberto"
						End If
					End If
			End Select
	End Select	
End Sub

<XATM_Disj_Red.Dados.TimerSF6:TimerSF6_SF62()>
Sub TimerSF6_SF62()
    Dim ObjIAD
    Dim ObjRaseT
    Dim ObjRaseD
    Dim IADFlag
    Dim RaseTFlag
    Dim RaseDFlag
    Dim DJAux
    Dim Disparo

    IADFlag = False
    RaseTFlag = False
    RaseDFlag = False
    Disparo = False
    
    If (Parent.Item("SF61").Value = 1) Then
        If (Parent.Item("SF62").Value = 1) Then
            Set ObjRaseT = BuscaAutomatismo("XATM_RASET",-1,RaseTFlag)
            Set ObjRaseD = BuscaAutomatismo("XATM_RASED",-1,RaseDFlag)
            Set ObjIAD = BuscaAutomatismo("XATM_IAD",XATM_Disj_Red.Id,IADFlag)
            
            If Value >= 5 Then
                ' Disparo RASED (ou IAD)
                Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoEntrada").Value
                    Case "1DJGBY"
                        If Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "2TRDJ" Then
                            If Parent.Item("PosiçãoMem").Value = 2 Then
                                If RaseTFlag = True Then
                                    Select Case XATM_Disj_Red.Id
                                        Case 0
                                            If IADFlag = True Then              
                                                Parent.Item("SeloIAD").WriteEx True
                                                Value = -1
                                            End If
                                        Case 110 
                                            If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 15
                                                Disparo = True
                                            End If
                                            If (IADFlag = True) And ((ObjRaseT.Item("Timers").Item("PreCondicoesCompTR1").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR1").Output = True) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then           
                                                Parent.Item("SeloIAD").WriteEx True
                                                Value = -1
                                            End If
                                        Case 210
                                            If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 16
                                                Disparo = True
                                            End If
                                            If (IADFlag = True) And ((ObjRaseT.Item("Timers").Item("PreCondicoesCompTR2").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR2").Output = True) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then       
                                                Parent.Item("SeloIAD").WriteEx True
                                                Value = -1
                                            End If
                                    End Select
                                Else
                                    XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo RASET não encontrado - Disparado por SF6 2º não executado"
                                    Parent.Item("SeloIAD").WriteEx True
                                    Value = -1
                                End If
                            Else
                                Parent.Item("SeloIAD").WriteEx True
                                Value = -1
                            End If
                        Else
                            If IADFlag = True Then              
                                Parent.Item("SeloIAD").WriteEx True
                                Value = -1
                            End If
                        End If
                    Case "2SD1DJGBY"
                        If Parent.Item("PosiçãoMem").Value = 2 Then
                            If RaseDFlag = True Then
                                Select Case XATM_Disj_Red.Id
                                    Case 0  
                                        If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                            '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                            ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                            Disparo = True
                                        End If
                                        If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj0").Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj0").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then               
                                            Parent.Item("SeloIAD").WriteEx True
                                            Value = -1
                                        End If
                                    Case 1  
                                        If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                            '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                            ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                            Disparo = True
                                        End If
                                        If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj1").Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj1").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then           
                                            Parent.Item("SeloIAD").WriteEx True
                                            Value = -1
                                        End If
    
                                    Case 2  
                                        If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                            '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                            ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                            Disparo = True
                                        End If
                                        If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj22").Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj22").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then         
                                            Parent.Item("SeloIAD").WriteEx True
                                            Value = -1
                                        End If
                                End Select
                            Else
                                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo RASED não encontrado - Disparado por SF6 2º não executado"
                                Parent.Item("SeloIAD").WriteEx True
                                Value = -1
                            End If
                        Else
                            Parent.Item("SeloIAD").WriteEx True
                            Value = -1
                        End If
                    Case "2DJG", "1DJBY2DJ", "2DJGIB2", "2DJGIBGIS", "2CHDJ", "2DJ2SCAT"
                        If Parent.Item("PosiçãoMem").Value = 2 Then
                            If RaseDFlag = True Then
                                Select Case XATM_Disj_Red.Id
                                    Case 0
                                        If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                            '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                            ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                            Disparo = True
                                        End If
                                        If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj0").Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj0").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then               
                                            Parent.Item("SeloIAD").WriteEx True
                                        End If
                                        Value = -1
                                    Case 10
                                        If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                            '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                            ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                            Disparo = True
                                        End If
                                        If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj1").Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj1").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then           
                                            Parent.Item("SeloIAD").WriteEx True
                                        End If
                                        Value = -1
                                    Case 1
                                        If IADFlag = True Then              
                                            Parent.Item("SeloIAD").WriteEx True
                                            Value = -1
                                        End If
                                    Case 120, 110
                                        Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
                                            Case "3TRSCBP", "2TRSCLT"
                                                If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                                    '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                    ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                                    Disparo = True
                                                End If
                                                If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj22").Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj22").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then         
                                                    Parent.Item("SeloIAD").WriteEx True
                                                    Value = -1
                                                End If
                                            Case "2TRDJ"
                                                'If (Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "GISBT") Then
                                                    If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                        '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                        ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 15
                                                        Disparo = True
                                                    End If
                                                'End If
                                            Case Else
                                                If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                    '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                    ObjRaseT.Item("Dados").Item("DisparoPorSF6").WriteEx XATM_Disj_Red.Id
                                                    ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 15
                                                    Disparo = True
                                                End If
                                        End Select
                                    Case 220, 210
                                        Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
                                            Case "3TR3DJ"
                                                If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                    '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                    ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 15
                                                    Disparo = True
                                                End If
                                                If (IADFlag = True) And ((ObjRaseT.Item("Timers").Item("PreCondicoesCompTR2").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR2").Output = True) Or (ObjRaseT.Item("Timers").Item("PreCondicoesCompTR3").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR3").Output = True) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then          
                                                    Parent.Item("SeloIAD").WriteEx True
                                                    Value = -1
                                                End If
                                            Case Else
                                            'If Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "GISBT" Then
                                                If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                    '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                    ObjRaseT.Item("Dados").Item("DisparoPorSF6").WriteEx XATM_Disj_Red.Id
                                                    ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 16
                                                    Disparo = True
                                                End If
                                            'End If
                                            End Select
                                    Case 24
                                        Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
                                            Case "4TRDJSC", "4TRSC"
                                                If ((Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "4BR2DJ") Or (Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "4BR1DJ")) Then
                                                    If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                        '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                        ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 15
                                                        Disparo = True
                                                    End If
                                                    If (IADFlag = True) And ((ObjRaseT.Item("Timers").Item("PreCondicoesCompTR1").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR1").Output = True) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then           
                                                        Parent.Item("SeloIAD").WriteEx True
                                                        Value = -1
                                                    End If
                                                End If
                                            Case "3TR3DJ"
                                                If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                    '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                    ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 16
                                                    Disparo = True
                                                End If
                                                If (IADFlag = True) And ((ObjRaseT.Item("Timers").Item("PreCondicoesCompTR1").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR1").Output = True) Or (ObjRaseT.Item("Timers").Item("PreCondicoesCompTR3").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR3").Output = True) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then          
                                                    Parent.Item("SeloIAD").WriteEx True
                                                    Value = -1
                                                End If
                                        End Select
                                    Case 25
                                        Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
                                            Case "4TRDJSC", "4TRSC"
                                                If ((Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "4BR2DJ") Or (Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "4BR1DJ")) Then
                                                    If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                        '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                        ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 16
                                                        Disparo = True
                                                    End If
                                                    If (IADFlag = True) And ((ObjRaseT.Item("Timers").Item("PreCondicoesCompTR2").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR2").Output = True) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then           
                                                        Parent.Item("SeloIAD").WriteEx True
                                                        Value = -1
                                                    End If
                                                End If
                                            Case "3TR3DJ"
                                                If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                    '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                    ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 17
                                                    Disparo = True
                                                End If
                                                If (IADFlag = True) And ((ObjRaseT.Item("Timers").Item("PreCondicoesCompTR1").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR1").Output = True) Or (ObjRaseT.Item("Timers").Item("PreCondicoesCompTR2").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR2").Output = True) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then          
                                                    Parent.Item("SeloIAD").WriteEx True
                                                    Value = -1
                                                End If
                                        End Select
                                    Case 26
                                        Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
                                            Case "4TRDJSC", "4TRSC", "4TR2D2JSC" 
                                                If ((Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "4BR2DJ") Or (Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "4BR1DJ")) Then
                                                    If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                        '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                        ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 17
                                                        Disparo = True
                                                    End If
                                                    If (IADFlag = True) And ((ObjRaseT.Item("Timers").Item("PreCondicoesCompTR3").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR3").Output = True) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then           
                                                        Parent.Item("SeloIAD").WriteEx True
                                                        Value = -1
                                                    End If
                                                End If
                                        End Select
                                    Case 27
                                        Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
                                            Case "4TR2D2JSC" 
                                                If ((Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "4BR2DJ") Or (Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "4BR1DJ")) Then
                                                    If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                        '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                        ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 18
                                                        Disparo = True
                                                    End If
                                                    If (IADFlag = True) And ((ObjRaseT.Item("Timers").Item("PreCondicoesTR4").Output = False) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then            
                                                        Parent.Item("SeloIAD").WriteEx True
                                                        Value = -1
                                                    End If
                                                End If
                                        End Select
                                    Case 40
                                        If Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "GIS2BR1DJ" Then
                                            If RaseTFlag = True Then
                                                If (ObjRaseT.AutomatismoEmAndamento = False) Then
                                                    '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                    ObjRaseT.Item("Dados").Item("DisparoPorSF6").WriteEx XATM_Disj_Red.Id
                                                    ObjRaseT.Item("Dados").Item("DisparoRaseCompleta").WriteEx 40
                                                    Disparo = True
                                                End If
                                                If (IADFlag = True) And ((ObjRaseT.Item("MaquinaEstados").Item("Main").Value >= 9) Or (ObjRaseT.Item("Timers").Item("PreCondicoesCompTR1").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR1").Output = True) Or (ObjRaseT.Item("Timers").Item("PreCondicoesCompTR2").Output = False) Or (ObjRaseT.Item("Timers").Item("BloqAutomaticoCompTR2").Output = True) Or (ObjRaseT.BloqueioOperador = True) Or (ObjRaseT.BloqueioGeral = True)) Then         
                                                    Parent.Item("SeloIAD").WriteEx True
                                                    Value = -1
                                                End If
                                            Else
                                                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo RASET não encontrado - Disparado por SF6 2º não executado"
                                                Parent.Item("SeloIAD").WriteEx True
                                                Value = -1
                                            End If
                                        End If
                                End Select
                            Else
                                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo RASED não encontrado - Disparado por SF6 2º não executado"
                                Parent.Item("SeloIAD").WriteEx True
                                Value = -1
                            End If
                        Else
                            Parent.Item("SeloIAD").WriteEx True
                            Value = -1
                        End If
                    Case "2DJGIB", "2DJG2IB", "2DJGBY2BR"
                        If Parent.Item("PosiçãoMem").Value = 2 Then
                            If RaseDFlag = True Then
                                Select Case XATM_Disj_Red.Id
                                    Case 0  
                                        If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                            '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                            ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                            Disparo = True
                                        End If
                                        If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj0").Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj0").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then
                                            Parent.Item("SeloIAD").WriteEx True
                                            Value = -1
                                        End If
                                    Case 1  
                                        If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                            '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                            ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                            Disparo = True
                                        End If
                                        If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj1").Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj1").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then
                                            Parent.Item("SeloIAD").WriteEx True
                                            Value = -1
                                        End If
                                    Case 110
                                        If Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "3TRSCBP" Then  
                                            If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                                '$CausaDisparo = 1o est. SF6 atuado por 10s e 2o est. SF6 atuado
                                                ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                                Disparo = True
                                            End If
                                            If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj22").Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj22").Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then             
                                                Parent.Item("SeloIAD").WriteEx True
                                                Value = -1
                                            End If
                                        End If
                                End Select
                            Else
                                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Automatismo RASED não encontrado - Disparado por SF6 2º não executado"
                                Parent.Item("SeloIAD").WriteEx True
                                Value = -1
                            End If
                        Else
                            Parent.Item("SeloIAD").WriteEx True
                            Value = -1
                        End If                        
                    Case "3DJBP"
                        If Parent.Item("PosiçãoMem").Value = 2 Then
                            If RaseDFlag = True Then
                                Select Case XATM_Disj_Red.Id
                                    Case 0, 1, 10
                                        If (ObjRaseD.AutomatismoEmAndamento = False) Then
                                            ' $CausaDisparo = 1o est. SF6 atuado por 5s e 2o est. SF6 atuado
                                            ObjRaseD.Item("Dados").Item("DisparoRaseD").WriteEx "ID_" & XATM_Disj_Red.Id
                                            Disparo = True
                                        End If
                                        If (IADFlag = True) And ((ObjRaseD.Item("Timers").Item("PreCondicoesDisj" & XATM_Disj_Red.Id).Output = False) Or (ObjRaseD.Item("Timers").Item("BloqueioAutomaticoDisj" & XATM_Disj_Red.Id).Output = True) Or (ObjRaseD.BloqueioOperador = True) Or (ObjRaseD.BloqueioGeral = True)) Then
                                            Parent.Item("SeloIAD").WriteEx True
                                        End If
                                        Value = -1
                                    Case Else
                                        ' Não se aplica
                                End Select
                            End If
                        End If
                End Select

                ' Disparo RASET
                If Disparo = False Then
                    If RaseTFlag Then
                        If Parent.Item("PosiçãoMem").Value = 2 Then
                            If Not ObjRaseT.AutomatismoEmAndamento Then
                                ' Verifica se trata-se de um disjuntor de AT do transformador (ou seja, final 10, e.g.: 110, 210, 310, ...)
                                Select Case CInt(Right(XATM_Disj_Red.Id, 2))
                                    Case 10
                                    	ObjRaseT.Item("Dados").Item("DisparoPorSF6").WriteEx True
                                        ObjRaseT.Item("Dados").Item("ModoAutomatismo").WriteEx 0
                                        ObjRaseT.Item("Dados").Item("Disparo").WriteEx (XATM_Disj_Red.Id \ 100) * 100
                                        Disparo = True
                                        Value = -1
                                End Select
                            End If
                        End If
                    End If
                End If
            Else
                Value = -1
                XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Tempo entre a atuação do SF6 1º e 2º estágio menor que 5s - Automatismos não disparados"
            End If  
        End If
    End If
    If Value >= 55 Then
        Value = -1
        XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Timer SF6 Encerrado por tempo excedido"
    End If	
End Sub

<XATM_Disj_Red.Dados.TimerSF6:TimerSF6_Start()>
Sub TimerSF6_Start()
	Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoEntrada").Value
		Case "1DJGBY", "2SD1DJGBY", "1DJBY2DJ"
			If (XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 2) Or (XATM_Disj_Red.Id = 110) Or (XATM_Disj_Red.Id = 210) Then
				IniciaTimer
			End If
		Case "2SECC"
			If Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "2TRDJ" Then
				If (XATM_Disj_Red.Id = 110) Or (XATM_Disj_Red.Id = 210) Then
					IniciaTimer
				End If
			End If
		Case "2DJG", "2DJGIB2", "2CHDJ", "2DJ2SCAT"
			If ((Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "4TRDJSC") Or (Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "4TRSC") Or (Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "4TR2D2JSC")) Then
				If (XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 24) Or (XATM_Disj_Red.Id = 25) Or (XATM_Disj_Red.Id = 26) Or (XATM_Disj_Red.Id = 27) Then
					IniciaTimer
				End If
			Else
				If Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "2TRDJ" Then
					If Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "GIS2BR1DJ" Then
						If ((XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 110) Or (XATM_Disj_Red.Id = 210) Or (XATM_Disj_Red.Id = 20) Or (XATM_Disj_Red.Id = 21)) Then
							IniciaTimer
						End If
					Else
						If (XATM_Disj_Red.Id = 110) Or (XATM_Disj_Red.Id = 210) Then
							IniciaTimer
						End If
					End If
				Else
					If Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value = "3TR3DJ" Then
						If (XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 210) Or (XATM_Disj_Red.Id = 24) Or (XATM_Disj_Red.Id = 25) Then
							IniciaTimer
						End If
					Else
						If Application.GetObject("autoDados.Automatismo.Aux.TipoBarra").Value = "GIS2BR1DJ" Then
							If ((XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 20) Or (XATM_Disj_Red.Id = 21) Or (XATM_Disj_Red.Id = 40)) Then
								IniciaTimer
							End If
						Else
							If (XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Then
								IniciaTimer
							End If
						End If
					End If
				End If
			End If
		Case "2DJGIBGIS"
			If (XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 20) Or (XATM_Disj_Red.Id = 21) Or (XATM_Disj_Red.Id = 110) Or (XATM_Disj_Red.Id = 210) Then
				IniciaTimer
			End If
		Case "2DJGIB", "2DJG2IB", "2DJGBY2BR"
			If (XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Then
				IniciaTimer
			End If
		Case "3DJBP"
            If ((XATM_Disj_Red.Id = 0) Or (XATM_Disj_Red.Id = 10) Or (XATM_Disj_Red.Id = 1)) Then
				IniciaTimer
			End If
	End Select
	Select Case Application.GetObject("autoDados.Automatismo.Aux.TipoTrafo").Value
		Case "3TRSCBP", "2TRSCLT"
			If XATM_Disj_Red.Id = 110 Then
				IniciaTimer
			End If
		Case "3TRSC", "3TRSCBT"
			Select Case XATM_Disj_Red.Id
				Case 120, 220, 320
					IniciaTimer
			End Select
		Case Else
			Select Case XATM_Disj_Red.Id
				Case 120, 220, 320, 420, 110, 210, 310, 410
					IniciaTimer
			End Select
	End Select
End Sub

Sub IniciaTimer()
	If (Parent.Item("SF61").Value = 1) Then
		Value = 0
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Timer SF6 Iniciado"
	else
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Timer SF6 Encerrado"
		Value = -1
	End If	
End Sub

<XATM_Disj_Red.Dados.TimerSF6:TimerSF6_Timer()>
Sub TimerSF6_Timer()
	If (Parent.Item("SF62").Value = 0) Then
		Value = Value + 1
	End If
End Sub

<XATM_Disj_Red.Dados.TimerSF6:[TimerSF6_Funções]()>
Sub [TimerSF6_Funções]()

End Sub

Function CheckLink(Sinal)
	Dim Aux
	On Error Resume Next
		Aux = Sinal.Value
		Aux = Sinal.Quality
		If Err.Number <> 0 Then
			Err.Clear
			CheckLink = True
		Else
			CheckLink = False
		End If
	On Error Goto 0
End Function

Function CheckLibLink(Sinal)
	Dim Aux
	On Error Resume Next
		Aux = Sinal.Id
		If Err.Number <> 0 Then
			Err.Clear
			CheckLibLink = True
		Else
			CheckLibLink = False
		End If
	On Error Goto 0
End Function

Function BuscaAutomatismo(TipoAutomatismo, IdEquipamento, byRef Atm_Flag)
	Atm_Flag = False

	For Each Atm in Application.GetObject("autoDados.Automatismo")
		If TypeName(Atm) = TipoAutomatismo Then
			Select Case TipoAutomatismo
				Case "XATM_TCT"
					IdTransformador = -999
					On Error Resume Next
						IdTransformador = Atm.Transformador.Id
					On Error GoTo 0
					If IsNumeric(IdTransformador) And IsNumeric(IdEquipamento) Then
						If IdTransformador = IdEquipamento Then
							Atm_Flag = True
						End If	
					End If	
				Case "XATM_IAD","XATM_TCD"
					IdDisjuntor = -999
					On Error Resume Next
						IdDisjuntor = Atm.Disjuntor.Id
					On Error GoTo 0
					If IsNumeric(IdDisjuntor) And IsNumeric(IdEquipamento) Then
						If IdDisjuntor = IdEquipamento Then
							Atm_Flag = True
						End If
					End If	
				Case Else
					Atm_Flag = True
			End Select		
		End If
		If Atm_Flag Then
			Exit For
		End If
	Next
	
	If Atm_Flag Then
		Set BuscaAutomatismo = Atm
	Else
		Set BuscaAutomatismo = Nothing
	End If
End Function

Sub Lixo()
	
End Sub

<XATM_Disj_Red.Dados.[CmdEmExecução]:[CmdEmExecução_CmdExec]()>
Sub [CmdEmExecução_CmdExec]()
	If (Me.Value = 2) Then
		' Comando em execucao
		Parent.Item("TimerCmd").Value = 0
	ElseIf (Me.Value = 1) Then 
		' Falha na execucao do comando
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & _
				"] - Falha na Execução do comando"
	ElseIf (Me.Value = 3) Then
		' Fim execucao do comando
		'Me.Value = 0
	End If	
End Sub

<XATM_Disj_Red.Dados.[PosiçãoMem]:[PosiçãoMem_InicioTimer]()>
Sub [PosiçãoMem_InicioTimer]()
	If Value = 0 Then
		Value = Parent.Item("Posição").Value
	Else
		Parent.Item("TimerPos").Value = 2 * 60 '2 min
	End If
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_CalcPos]()>
Sub [PosiçãoProvisória_CalcPos]()
End Sub

Sub CalculaPos()
	Dim PosAb1
	Dim LinkAb1Flag
	Dim PosFe1
	Dim LinkFe1Flag
	Dim PosAb2
	Dim LinkAb2Flag
	Dim PosFe2
	Dim LinkFe2Flag
	Dim LinkInvalido
	Dim Qualidade
		
	Set PosAb1 = XATM_Disj_Red.[1_PosAberto]
	Set PosFe1 = XATM_Disj_Red.[1_PosFechado]
	Set PosAb2 = XATM_Disj_Red.[2_PosAberto]
	Set PosFe2 = XATM_Disj_Red.[2_PosFechado]
	
	Qualidade = 192
	LinkAb1Flag = CheckLink(PosAb1)
	LinkFe1Flag = CheckLink(PosFe1)
	LinkAb2Flag = CheckLink(PosAb2)
	LinkFe2Flag = CheckLink(PosFe2)
	LinkInvalido = False
	
	If XATM_Disj_Red.PontoDuploPos = True Then '######### PONTO DUPLO ############
		Select Case XATM_Disj_Red.RelePrioritario
			Case 2
				If LinkFe2Flag = False Then
					If PosFe2.Quality >= Qualidade Then
						CalcPosConsistDuplo(PosFe2.Value)
					Else
						If LinkAb2Flag = False Then
							If PosAb2.Quality >= Qualidade Then
								CalcPosConsistDuplo(PosAb2.Value)
							Else
								If LinkFe1Flag = False Then
									If PosFe1.Quality >= Qualidade Then
										CalcPosConsistDuplo(PosFe1.Value)
									Else
										If LinkAb1Flag = False Then
											If PosAb1.Quality >= Qualidade Then
												CalcPosConsistDuplo(PosAb1.Value)
											Else
												XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
											End If
										Else
											LinkInvalido = True
										End If
									End If
								Else
									If LinkAb1Flag = False Then
										If PosAb1.Quality >= Qualidade Then
											CalcPosConsistDuplo(PosAb1.Value)
										Else
											XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
										End If
									Else
										LinkInvalido = True
									End If
								End If
							End If
						Else
							If LinkFe1Flag = False Then
								If PosFe1.Quality >= Qualidade Then
									CalcPosConsistDuplo(PosFe1.Value)
								Else
									If LinkAb1Flag = False Then
										If PosAb1.Quality >= Qualidade Then
											CalcPosConsistDuplo(PosAb1.Value)
										Else
											XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
										End If
									Else
										LinkInvalido = True
									End If
								End If
							Else
								If LinkAb1Flag = False Then
									If PosAb1.Quality >= Qualidade Then
										CalcPosConsistDuplo(PosAb1.Value)
									Else
										XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
									End If
								Else
									LinkInvalido = True
								End If
							End If
						End If
					End If
				Else
					If LinkAb2Flag = False Then
						If PosAb2.Quality >= Qualidade Then
							CalcPosConsistDuplo(PosAb2.Value)
						Else
							If LinkFe1Flag = False Then
								If PosFe1.Quality >= Qualidade Then
									CalcPosConsistDuplo(PosFe1.Value)
								Else
									If LinkAb1Flag = False Then
										If PosAb1.Quality >= Qualidade Then
											CalcPosConsistDuplo(PosAb1.Value)
										Else
											XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
										End If
									Else
										LinkInvalido = True
									End If
								End If
							Else
								If LinkAb1Flag = False Then
									If PosAb1.Quality >= Qualidade Then
										CalcPosConsistDuplo(PosAb1.Value)
									Else
										XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
									End If
								Else
									LinkInvalido = True
								End If
							End If
						End If
					Else
						If LinkFe1Flag = False Then
							If PosFe1.Quality >= Qualidade Then
								CalcPosConsistDuplo(PosFe1.Value)
							Else
								If LinkAb1Flag = False Then
									If PosAb1.Quality >= Qualidade Then
										CalcPosConsistDuplo(PosAb1.Value)
									Else
										XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
									End If
								Else
									LinkInvalido = True
								End If
							End If
						Else
							If LinkAb1Flag = False Then
								If PosAb1.Quality >= Qualidade Then
									CalcPosConsistDuplo(PosAb1.Value)
								Else
									XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
								End If
							Else
								LinkInvalido = True
							End If
						End If
					End If
				End If
			Case Else
				If LinkFe1Flag = False Then
					If PosFe1.Quality >= Qualidade Then
						CalcPosConsistDuplo(PosFe1.Value)
					Else
						If LinkAb1Flag = False Then
							If PosAb1.Quality >= Qualidade Then
								CalcPosConsistDuplo(PosAb1.Value)
							Else
								If LinkFe2Flag = False Then
									If PosFe2.Quality >= Qualidade Then
										CalcPosConsistDuplo(PosFe2.Value)
									Else
										If LinkAb2Flag = False Then
											If PosAb2.Quality >= Qualidade Then
												CalcPosConsistDuplo(PosAb2.Value)
											Else
												XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
											End If
										Else
											LinkInvalido = True
										End If
									End If
								Else
									If LinkAb2Flag = False Then
										If PosAb2.Quality >= Qualidade Then
											CalcPosConsistDuplo(PosAb2.Value)
										Else
											XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
										End If
									Else
										LinkInvalido = True
									End If
								End If
							End If
						Else
							If LinkFe2Flag = False Then
								If PosFe2.Quality >= Qualidade Then
									CalcPosConsistDuplo(PosFe2.Value)
								Else
									If LinkAb2Flag = False Then
										If PosAb2.Quality >= Qualidade Then
											CalcPosConsistDuplo(PosAb2.Value)
										Else
											XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
										End If
									Else
										LinkInvalido = True
									End If
								End If
							Else
								If LinkAb2Flag = False Then
									If PosAb2.Quality >= Qualidade Then
										CalcPosConsistDuplo(PosAb2.Value)
									Else
										XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
									End If
								Else
									LinkInvalido = True
								End If
							End If
						End If
					End If
				Else
					If LinkAb1Flag = False Then
						If PosAb1.Quality >= Qualidade Then
							CalcPosConsistDuplo(PosAb1.Value)
						Else
							If LinkFe2Flag = False Then
								If PosFe2.Quality >= Qualidade Then
									CalcPosConsistDuplo(PosFe2.Value)
								Else
									If LinkAb2Flag = False Then
										If PosAb2.Quality >= Qualidade Then
											CalcPosConsistDuplo(PosAb2.Value)
										Else
											XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
										End If
									Else
										LinkInvalido = True
									End If
								End If
							Else
								If LinkAb2Flag = False Then
									If PosAb2.Quality >= Qualidade Then
										CalcPosConsistDuplo(PosAb2.Value)
									Else
										XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
									End If
								Else
									LinkInvalido = True
								End If
							End If
						End If
					Else
						If LinkFe2Flag = False Then
							If PosFe2.Quality >= Qualidade Then
								CalcPosConsistDuplo(PosFe2.Value)
							Else
								If LinkAb2Flag = False Then
									If PosAb2.Quality >= Qualidade Then
										CalcPosConsistDuplo(PosAb2.Value)
									Else
										XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
									End If
								Else
									LinkInvalido = True
								End If
							End If
						Else
							If LinkAb2Flag = False Then
								If PosAb2.Quality >= Qualidade Then
									CalcPosConsistDuplo(PosAb2.Value)
								Else
									XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Sinal de Posição com Qualidade Baixa"
								End If
							Else
								LinkInvalido = True
							End If
						End If
					End If
				End If
		End Select
	Else '######### PONTO SIMPLES ############
		If (LinkAb1Flag = False And LinkFe1Flag = False) And (PosFe1.Quality >= Qualidade And PosAb1.Quality >= Qualidade) Then
			CalcPosConsistSimples PosAb1.Value, True, PosFe1.Value, True
		ElseIf (LinkAb2Flag = False And LinkFe2Flag = False) And (PosFe2.Quality >= Qualidade And PosAb2.Quality >= Qualidade) Then
			CalcPosConsistSimples PosAb2.Value, True, PosFe2.Value, True
		Else
			WriteEx 0
			XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Associações de posição, inválidas ou nula ou com baixa qualidade"
		End If
	End If
	If LinkInvalido = True Then
		WriteEx 0
		XATM_Disj_Red.Item("XATM_LogDJ").Texto = "[" & XATM_Disj_Red.Name & "] - Associações de posição, inválidas ou nula"
	End If

End Sub

Function CheckLink(Sinal)
	Dim Aux
	On Error Resume Next
		Aux = Sinal.Value
		Aux = Sinal.Quality
		If Err.Number <> 0 Then
			Err.Clear
			CheckLink = True
		Else
			CheckLink = False
		End If
	On Error Goto 0
End Function

Sub CalcPosConsistSimples(PosAb, PosAbFlag, PosFe, PosFeFlag)
	If PosAbFlag = True And PosFeFlag = True Then
		If PosAb = 0 And PosFe = 1 Then 
			WriteEx 2
		ElseIf PosAb = 1 And PosFe = 0 Then 
			WriteEx 1
		Else
			WriteEx 0
		End If
	Else
		If PosAbFlag = True Then
			If PosAb = 1 Then 
				WriteEx 1
			Else
				WriteEx 2
			End If
		ElseIf PosFeFlag = True Then
			If PosFe = 1 Then 
				WriteEx 2
			Else
				WriteEx 1
			End If
		Else
			WriteEx 0
		End If
	End If
End Sub

Sub CalcPosConsistDuplo(Pos)
	Select Case Pos
		Case XATM_Disj_Red.ValorPosAberto
			WriteEx 1
		Case XATM_Disj_Red.ValorPosFechado
			WriteEx 2
		Case Else
			WriteEx 0
	End Select
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_ModoSimulacao]()>
Sub [PosiçãoProvisória_ModoSimulacao]()
	If XATM_Disj_Red.HabilitaSimulador = True Then
		WriteEx 2
		Parent.Item("Posição").WriteEx 2
	Else
		WriteEx 0
		Parent.Item("Posição").WriteEx 0
		CalculaPos()
	End If	
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_OnStartRunning]()>
Sub [PosiçãoProvisória_OnStartRunning]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	Else
		WriteEx 2
		Parent.Item("PosiçãoMem").Value = Value
	End If
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_PosAb1Quality]()>
Sub [PosiçãoProvisória_PosAb1Quality]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	End If		
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_PosAb1]()>
Sub [PosiçãoProvisória_PosAb1]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	End If	
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_PosAb2Quality]()>
Sub [PosiçãoProvisória_PosAb2Quality]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	End If		
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_PosAb2]()>
Sub [PosiçãoProvisória_PosAb2]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	End If	
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_PosFe1Quality]()>
Sub [PosiçãoProvisória_PosFe1Quality]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	End If		
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_PosFe1]()>
Sub [PosiçãoProvisória_PosFe1]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	End If	
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_PosFe2Quality]()>
Sub [PosiçãoProvisória_PosFe2Quality]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	End If		
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_PosFe2]()>
Sub [PosiçãoProvisória_PosFe2]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	End If	
End Sub

<XATM_Disj_Red.Dados.[PosiçãoProvisória]:[PosiçãoProvisória_Prioridade]()>
Sub [PosiçãoProvisória_Prioridade]()
	If XATM_Disj_Red.HabilitaSimulador = False Then
		CalculaPos()
	End If
End Sub

<XATM_Disj_Red.Dados.[Posição]:[Posição_ConfirmClosedPositionAfterDelay]()>
Sub [Posição_ConfirmClosedPositionAfterDelay]()
	If CInt(DocString) >= XATM_Disj_Red.AtrasoPosFeConsist Then
		WriteEx Parent.Item("PosiçãoProvisória").Value
		DocString = "-1"
	Else
		DocString = CStr(CInt(DocString) + 1)
	End If
End Sub

<XATM_Disj_Red.Dados.[Posição]:[Posição_OnProvisionalPositionChanged]()>
Sub [Posição_OnProvisionalPositionChanged]()
	If Parent.Item("PosiçãoProvisória").Value = 2 And XATM_Disj_Red.AtrasoPosFeConsist > 0 Then
		' Inicializa temporizador, verificar evento de usuario "ConfirmClosedPositionAfterDelay"
		DocString = "0"
	Else
		DocString = "-1"
		WriteEx Parent.Item("PosiçãoProvisória").Value
	End If
End Sub

<XATM_Disj_Red.Dados.[Posição]:[Posição_OnStartRunning]()>
Sub [Posição_OnStartRunning]()
	[Posição_OnProvisionalPositionChanged]	
End Sub

<XATM_Disj_Red.Dados.[Posição]:[Posição_WatchDogPosition]()>
Sub [Posição_WatchDogPosition]()
	WriteEx Parent.Item("PosiçãoProvisória").Value
End Sub

