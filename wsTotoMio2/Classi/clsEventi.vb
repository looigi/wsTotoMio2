Imports System.Diagnostics.Eventing.Reader
Imports Microsoft.SqlServer

Public Class clsEventi
	Private Structure StrutturaGiocatore
		Dim idUtente As Integer
		Dim NickName As String
		Dim Punti As Integer
		Dim RisultatiEsatti As Integer
		Dim RisCasaTot As Integer
		Dim RisFuoriTot As Integer
		Dim Segni As Integer
		Dim SommaGoal As Integer
		Dim DifferenzaGoal As Integer
		Dim Giocate As Integer
	End Structure

	Public Function GestioneEventi(Mp As String, idAnno As Integer, idGiornata As Integer, idEvento As Integer,
								   QuantiGiocatori As Integer, Importanza As Integer, InizioGiornata As Integer,
								   Conn As Object, Connessione As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim Rec As Object

		Sql = "Select * From Eventi Where idEvento=" & idEvento
		Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Not Rec.Eof Then
				Dim Evento As String = Rec("Descrizione").Value
				Rec.Close

				Dim Dettaglio As String = ""
				Dim Dettaglio2 As String = ""
				Dim Cosa As String = ""

				If Evento.ToUpper.Contains("CREAZIONE ") Then
					Cosa = "CREAZIONE"
					Dettaglio = Mid(Evento, 10, Evento.Length).Trim()
				End If

				If Evento.ToUpper.Contains("CHIUSURA ") Then
					Cosa = "CHIUSURA"
					Dettaglio = Mid(Evento, 9, Evento.Length).Trim()
					Dim c() As String = Dettaglio.Split(" ")
					Dettaglio2 = c(0) ' FINALE / SEMIFINALE / GIRONI
					Dettaglio = Dettaglio.Replace(Dettaglio2, "").Trim ' NOME TROFEO
				End If

				If Evento.ToUpper.Contains("PARTITA ") Then
					Cosa = "PARTITA"
					Dettaglio = Mid(Evento, 8, Evento.Length).Trim()
					Dettaglio2 = Mid(Dettaglio, Dettaglio.Length - 3, Dettaglio.Length).Trim()
					Dettaglio = Dettaglio.Replace(Dettaglio2, "").Trim()
				End If

				If Evento.ToUpper.Contains("SEMIFINALE ") And Not Evento.ToUpper.Contains("CHIUSURA") Then
					Cosa = "SEMIFINALE"
					Dettaglio = Mid(Evento, 11, Evento.Length).Trim() ' NOME TORNEO
				End If

				If Evento.ToUpper.Contains("FINALE ") And Not Evento.ToUpper.Contains("SEMIFINALE ") And Not Evento.ToUpper.Contains("CHIUSURA") Then
					Cosa = "FINALE"
					Dettaglio = Mid(Evento, 7, Evento.Length).Trim() ' NOME TORNEO
				End If

				Select Case Cosa
					Case "CREAZIONE"
						Ritorno = CreazioneCoppa(Mp, idAnno, idGiornata, QuantiGiocatori, Importanza, InizioGiornata, Conn, Connessione, Dettaglio)
					Case "PARTITA"
						Ritorno = Partita(Mp, idAnno, idGiornata, idEvento, Conn, Connessione)
					Case "SEMIFINALE"
					Case "FINALE"
					Case "CHIUSURA"
				End Select
			End If
		End If

		Return Ritorno
	End Function

	Private Function PrendeGiocatori(Mp As String, idAnno As Integer, idGiornata As Integer, Conn As Object, Connessione As String) As List(Of StrutturaGiocatore)
		Dim Ritorno As New List(Of StrutturaGiocatore)
		Dim Giocatori As String = RitornaClassificaGenerale(Mp, idAnno, idGiornata, Conn, Connessione)
		Dim Righe() As String = Giocatori.Split("§")
		For Each R As String In Righe
			If R <> "" Then
				Dim Campi() As String = R.Split(";")
				Dim s As New StrutturaGiocatore
				s.idUtente = Val(Campi(0))
				s.NickName = Campi(1)
				s.Punti = Val(Campi(2))
				s.RisultatiEsatti = Val(Campi(3))
				s.RisCasaTot = Val(Campi(4))
				s.RisFuoriTot = Val(Campi(5))
				s.Segni = Val(Campi(6))
				s.SommaGoal = Val(Campi(7))
				s.DifferenzaGoal = Val(Campi(8))
				s.Giocate = Val(Campi(9))
				Ritorno.Add(s)
			End If
		Next
		Return Ritorno
	End Function

	Private Function Partita(Mp As String, idAnno As Integer, idGiornata As Integer, idEvento As Integer, Conn As Object, Connessione As String) As String
		Dim Ritorno As String = "OK"
		Dim Sql As String = "SELECT A.*, B.NickName As Casa, C.NickName As Fuori, D.Descrizione As Evento, E.Punti As Punti1, F.Punti As Punti2, " &
			"E.SegniPresi As SegniPresi1, F.SegniPresi As SegniPresi2, E.RisultatiEsatti As RisEsatti1, F.RisultatiEsatti As RisEsatti2, " &
			"E.RisultatiCasaTot As RisCasa1, F.RisultatiCasaTot As RisCasa2, E.RisultatiFuoriTot As RisFuori1, F.RisultatiFuoriTot As RisFuori2, " &
			"E.SommeGoal As SommeGoal1, F.SommeGoal As SommeGoal2, E.DifferenzeGoal As DiffGoal1, F.DifferenzeGoal As DiffGoal2, " &
			"G.Pronostico As Risultato1, H.Pronostico As Risultato2 " &
			"FROM EventiPartite A " &
			"Left Join Utenti B On A.idAnno = B.idAnno And A.idGiocatore1 = B.idUtente " &
			"Left Join Utenti C On A.idAnno = B.idAnno And A.idGiocatore2 = C.idUtente " &
			"Left Join Eventi D On A.idEvento = D.idEvento And A.idGiornata = D.InizioGiornata " &
			"Left Join Risultati E On A.idAnno = E.idAnno And A.idGiornata = E.idConcorso And A.idGiocatore1 = E.idUtente " &
			"Left Join Risultati F On A.idAnno = F.idAnno And A.idGiornata = F.idConcorso And A.idGiocatore2 = F.idUtente " &
			"Left Join Pronostici G On A.idAnno = G.idAnno And A.idGiornata = G.idConcorso And A.idGiocatore1 = G.idUtente And A.idPartita = G.idPartita " &
			"Left Join Pronostici H On A.idAnno = H.idAnno And A.idGiornata = H.idConcorso And A.idGiocatore2 = H.idUtente And A.idPartita = H.idPartita " &
			"WHERE A.idAnno = " & idAnno & " And A.idEvento = " & idEvento & " And idGiornata = " & idGiornata & " " &
			"Order By idPartita"
		Dim Rec As Object = CreaRecordset(Mp, Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessuna giornata di andata rilevata"
			Else
				Do Until Rec.Eof
					Dim Vincente As Integer = -1
					Dim idPartita As Integer = Rec("idPartita").Value
					Dim Risultato1 As String = Rec("Risultato1").Value
					Dim Risultato2 As String = Rec("Risultato2").Value

					Dim Punti1 As Integer = Rec("Punti1").Value
					Dim Punti2 As Integer = Rec("Punti2").Value
					If Punti1 > Punti2 Then
						Vincente = 1
					Else
						If Punti1 < Punti2 Then
							Vincente = 2
						Else
							Dim Segni1 As Integer = Rec("SegniPresi1").Value
							Dim Segni2 As Integer = Rec("SegniPresi2").Value

							If Segni1 > Segni2 Then
								Vincente = 1
							Else
								If Segni2 < Segni2 Then
									Vincente = 2
								Else
									Dim RisEsa1 As Integer = Rec("RisEsatti1").Value
									Dim RisEsa2 As Integer = Rec("RisEsatti2").Value

									If RisEsa1 > RisEsa2 Then
										Vincente = 1
									Else
										If RisEsa1 < RisEsa2 Then
											Vincente = 2
										Else
											Dim RisFuori1 As Integer = Rec("RisFuori1").Value
											Dim RisFuori2 As Integer = Rec("RisFuori2").Value

											If RisFuori1 > RisFuori2 Then
												Vincente = 1
											Else
												If RisFuori1 < RisFuori2 Then
													Vincente = 2
												Else
													Dim RisCasa1 As Integer = Rec("RisCasa1").Value
													Dim RisCasa2 As Integer = Rec("RisCasa2").Value

													If RisCasa1 > RisCasa2 Then
														Vincente = 1
													Else
														If RisCasa1 < RisCasa2 Then
															Vincente = 2
														Else
															Dim SommeGoal1 As Integer = Rec("SommeGoal1").Value
															Dim SommeGoal2 As Integer = Rec("SommeGoal2").Value

															If SommeGoal1 > SommeGoal2 Then
																Vincente = 1
															Else
																If SommeGoal1 < SommeGoal2 Then
																	Vincente = 2
																Else
																	Dim DiffGoal1 As Integer = Rec("DiffGoal1").Value
																	Dim DiffGoal2 As Integer = Rec("DiffGoal2").Value

																	If DiffGoal1 > DiffGoal2 Then
																		Vincente = 1
																	Else
																		If DiffGoal1 < DiffGoal2 Then
																			Vincente = 2
																		Else
																			Dim Random As Integer = CInt(Int((5 * Rnd())))
																			If Random > 2 Then
																				Vincente = 1
																			Else
																				Vincente = 2
																			End If
																		End If
																	End If
																End If
															End If
														End If
													End If
												End If
											End If
										End If
									End If
								End If
							End If
						End If
					End If

					Sql = "Update EventiPartite Set " &
						"idVincente=" & Vincente & ", Risultato1='" & Risultato1 & "', Risultato2='" & Risultato2 & "' " &
						"Where idAnno=" & idAnno & " And idEvento=" & idEvento & " And idGiornata=" & idGiornata & " And idPartita=" & idPartita
					Dim Rit As String = Conn.EsegueSql(Mp, Sql, Connessione, False)
					If Rit.Contains("ERROR") Then
						Ritorno = Rit
						Exit Do
					End If

					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	Private Function RitornaGiocatoriScelti(QuantiGiocatori As Integer, Importanza As Integer, Classifica As List(Of StrutturaGiocatore)) As List(Of StrutturaGiocatore)
		Dim Scelti As New List(Of StrutturaGiocatore)
		Dim Quanti As Integer = Math.Abs(QuantiGiocatori - 1)

		If QuantiGiocatori > 0 Then
			Dim Inizio As Integer = -1
			Dim Fine As Integer = -1

			Select Case Importanza
				Case 1
					Inizio = 1
					Fine = Quanti + 1
				Case 2
					Inizio = CInt((Classifica.Count - 1) / 3)
					Fine = Inizio + Quanti + 1
				Case 3
					Inizio = (Classifica.Count - 1) / 2
					Fine = Inizio + Quanti + 1
				Case 4
					Inizio = (Classifica.Count - 1) / 1.3
					Fine = Inizio + Quanti + 1
			End Select

			While Fine > Classifica.Count - 1
				Fine -= 1
				Inizio -= 1
			End While

			For i As Integer = Inizio To Fine
				Scelti.Add(Classifica.Item(i))
			Next
		Else
			For i As Integer = (Classifica.Count - 1) To (Classifica.Count - 1 - Quanti) Step -1
				Scelti.Add(Classifica.Item(i))
			Next
		End If

		Return Scelti
	End Function

	Private Function CreazioneCoppa(Mp As String, idAnno As Integer, idGiornata As Integer, QuantiGiocatori As Integer, Importanza As Integer,
											InizioGiornata As Integer, Conn As Object, Connessione As String, Torneo As String) As String
		Dim Ritorno As String = ""
		Dim Classifica As List(Of StrutturaGiocatore) = PrendeGiocatori(Mp, idAnno, idGiornata, Conn, Connessione)
		Dim QuantiGiocatoriPresenti As Integer = Classifica.Count - 1
		If QuantiGiocatoriPresenti > QuantiGiocatori Then
			Dim Scelti As List(Of StrutturaGiocatore) = RitornaGiocatoriScelti(QuantiGiocatori, Importanza, Classifica)

			Dim GiornateAndata As New List(Of Integer)
			Dim idEventiAndata As New List(Of Integer)
			Dim GiornateRitorno As New List(Of Integer)
			Dim idEventiRitorno As New List(Of Integer)

			Dim Sql As String = "Select * From Eventi Where Upper(Descrizione) Like '%" & Torneo.ToUpper & " A%' Order By InizioGiornata"
			Dim Rec As Object = CreaRecordset(Mp, Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Rec.Eof Then
					Ritorno = "ERROR: Nessuna giornata di andata rilevata"
				Else
					Do Until Rec.Eof
						GiornateAndata.Add(Rec("InizioGiornata").Value)
						idEventiAndata.Add(Rec("idEvento").Value)

						Rec.MoveNext
					Loop
					Rec.Close

					Sql = "Select * From Eventi Where Upper(Descrizione) Like '%" & Torneo.ToUpper & " R%' Order By InizioGiornata"
					Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = "ERROR: Nessuna giornata di ritorno rilevata"
						Else
							Do Until Rec.Eof
								GiornateRitorno.Add(Rec("InizioGiornata").Value)
								idEventiRitorno.Add(Rec("idEvento").Value)

								Rec.MoveNext
							Loop
							Rec.Close

							Dim GiornataSemiFinale As Integer = -1
							Dim idEventoSemifinale As Integer = -1

							Sql = "Select * From Eventi Where Upper(Descrizione) Like '%SEMIFINALE " & Torneo.ToUpper & "%' Order By InizioGiornata"
							Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ritorno = "ERROR: Nessuna semifinale rilevata"
								Else
									Do Until Rec.Eof
										GiornataSemiFinale = Rec("InizioGiornata").Value
										idEventoSemifinale = Rec("idEvento").Value

										Rec.MoveNext
									Loop
									Rec.Close

									Dim GiornataFinale As Integer = -1
									Dim idEventoFinale As Integer = -1

									Sql = "Select * From Eventi Where Upper(Descrizione) Like '%FINALE " & Torneo.ToUpper & "%' And UPPER(Descrizione) Not Like '%SEMIFINALE " & Torneo.ToUpper & "%' Order By InizioGiornata"
									Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										If Rec.Eof Then
											Ritorno = "ERROR: Nessuna finale rilevata"
										Else
											Do Until Rec.Eof
												GiornataFinale = Rec("InizioGiornata").Value
												idEventoFinale = Rec("idEvento").Value

												Rec.MoveNext
											Loop
											Rec.Close

											For i As Integer = 0 To QuantiGiocatori - 2
												'Dim Connessione2 As String = RitornaPercorso(Mp, 5)
												'Dim Conn2 As Object = New clsGestioneDB(TipoServer)

												Sql = "Select * From ScontriDiretti Where NumeroSquadre=" & QuantiGiocatori & " And Giornata=" & i + 1 & " Order By Progressivo"
												Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													If Rec.Eof Then
														Ritorno = "ERROR: Nessuna giornata rilevata"
													Else
														Dim Prima As New List(Of Integer)
														Dim Seconda As New List(Of Integer)

														Do Until Rec.Eof
															Prima.Add(Rec("Squadra1").Value)
															Seconda.Add(Rec("Squadra2").Value)

															Rec.MoveNext
														Loop
														Rec.Close

														Dim GiornataAndata As Integer = GiornateAndata.Item(i)
														Dim idEventoAndata As Integer = idEventiAndata.Item(i)
														Dim GiornataRitorno As Integer = GiornateRitorno.Item(i)
														Dim idEventoRitorno As Integer = idEventiRitorno.Item(i)

														Dim Conta As Integer = 0
														Dim idPartita As Integer = 1
														For Each p As Integer In Prima
															Dim Squadra1 As Integer = Scelti.Item(Prima(Conta) - 1).idUtente
															Dim Squadra2 As Integer = Scelti.Item(Seconda(Conta) - 1).idUtente

															Sql = "Insert Into EventiPartite Values (" &
																" " & idAnno & ", " &
																" " & idEventoAndata & ", " &
																" " & GiornataAndata & ", " &
																" " & idPartita & ", " &
																" " & Squadra1 & ", " &
																"'', " &
																" " & Squadra2 & ", " &
																"'', " &
																"-1 " &
																")"
															Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
															If Ritorno.Contains("ERROR") And Not Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
																Exit For
															Else
																Sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & GiornataAndata & ", " & idEventoAndata & ")"
																Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
																If Ritorno.Contains("ERROR") And Not Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
																	Exit For
																Else
																	Sql = "Insert Into EventiPartite Values (" &
																		" " & idAnno & ", " &
																		" " & idEventoRitorno & ", " &
																		" " & GiornataRitorno & ", " &
																		" " & idPartita & ", " &
																		" " & Squadra2 & ", " &
																		"'', " &
																		" " & Squadra1 & ", " &
																		"'', " &
																		"-1 " &
																		")"
																	Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
																	If Ritorno.Contains("ERROR") And Not Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
																		Exit For
																	Else
																		Sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & GiornataRitorno & ", " & idEventoRitorno & ")"
																		Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
																		If Ritorno.Contains("ERROR") And Not Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
																			Exit For
																		End If
																	End If
																End If
															End If

															'Rec.MoveNext
															Conta += 1
															idPartita += 1
														Next
														'Rec.Close
													End If
												End If
											Next

											If Not Ritorno.Contains("ERROR") Or Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
												If idEventoSemifinale <> -1 Then
													Sql = "Insert Into EventiPartite Values (" &
															" " & idAnno & ", " &
															" " & idEventoSemifinale & ", " &
															" " & GiornataSemiFinale & ", " &
															"1, " &
															"-1, " &
															"'', " &
															"-1, " &
															"'', " &
															"-1 " &
															")"
													Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
													If Not Ritorno.Contains("ERROR") And Not Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
														Sql = "Insert Into EventiPartite Values (" &
															" " & idAnno & ", " &
															" " & idEventoSemifinale & ", " &
															" " & GiornataSemiFinale & ", " &
															"2, " &
															"-1, " &
															"'', " &
															"-1, " &
															"'', " &
															"-1 " &
															")"
														Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
														If Not Ritorno.Contains("ERROR") And Not Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
															Sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & GiornataSemiFinale & ", " & idEventoSemifinale & ")"
															Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
														End If
													End If
												End If
											End If

											If Not Ritorno.Contains("ERROR") Or Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
												If idEventoFinale <> -1 Then
													Sql = "Insert Into EventiPartite Values (" &
															" " & idAnno & ", " &
															" " & idEventoFinale & ", " &
															" " & GiornataFinale & ", " &
															"1, " &
															"-1, " &
															"'', " &
															"-1, " &
															"'', " &
															"-1 " &
															")"
													Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
													If Not Ritorno.Contains("ERROR") And Not Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
														Sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & GiornataFinale & ", " & idEventoFinale & ")"
														Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
													End If
												End If
											End If

											If Not Ritorno.Contains("ERROR") Or Ritorno.ToUpper.Contains("DUPLICATE ENTRY") Then
												Ritorno = "OK"
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		Else
			Ritorno = "ERROR: Poche squadre per creare la coppa"
		End If

		Return Ritorno
	End Function

End Class
