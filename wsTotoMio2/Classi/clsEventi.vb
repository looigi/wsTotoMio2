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

				If Evento.ToUpper.Contains("PARTITA ") Then
					Cosa = "PARTITA"
					Dettaglio = Mid(Evento, 8, Evento.Length).Trim()
					Dettaglio2 = Mid(Dettaglio, Dettaglio.Length - 3, Dettaglio.Length).Trim()
					Dettaglio = Dettaglio.Replace(Dettaglio2, "").Trim()
				End If

				If Evento.ToUpper.Contains("SEMIFINALE ") Then
					Cosa = "SEMIFINALE"
					Dettaglio = Mid(Evento, 11, Evento.Length).Trim()
				End If

				If Evento.ToUpper.Contains("FINALE ") And Not Evento.ToUpper.Contains("SEMIFINALE") Then
					Cosa = "FINALE"
					Dettaglio = Mid(Evento, 7, Evento.Length).Trim()
				End If

				Select Case Cosa
					Case "CREAZIONE"
						Ritorno = CreazioneCoppa(Mp, idAnno, idGiornata, QuantiGiocatori, Importanza, InizioGiornata, Conn, Connessione, Dettaglio)
					Case "PARTITA"
					Case "SEMIFINALE"
					Case "FINALE"
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
