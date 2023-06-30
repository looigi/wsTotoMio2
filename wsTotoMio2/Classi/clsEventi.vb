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

				Select Case Evento
					Case "Creazione Rotolo di Coppa"
						Ritorno = CreazioneCoppa(Mp, idAnno, idGiornata, QuantiGiocatori, Importanza, InizioGiornata, Conn, Connessione, "Rotolo di Coppa")
					Case "Creazione Coppa Coppata"
						Ritorno = CreazioneCoppa(Mp, idAnno, idGiornata, QuantiGiocatori, Importanza, InizioGiornata, Conn, Connessione, "Coppa Coppata")
					Case "Creazione Coppa TotoMio"
						Ritorno = CreazioneCoppa(Mp, idAnno, idGiornata, QuantiGiocatori, Importanza, InizioGiornata, Conn, Connessione, "Coppa TotoMio")
					Case "Creazione Coppa dei Settimini"
						Ritorno = CreazioneCoppa(Mp, idAnno, idGiornata, QuantiGiocatori, Importanza, InizioGiornata, Conn, Connessione, "Coppa dei Settimini")
					Case "Creazione Pippettero"
						Ritorno = CreazioneCoppa(Mp, idAnno, idGiornata, QuantiGiocatori, Importanza, Conn, InizioGiornata, Connessione, "Pippettero")
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
				Ritorno.Add(s)
			End If
		Next
		Return Ritorno
	End Function

	Private Function CreazioneCoppa(Mp As String, idAnno As Integer, idGiornata As Integer, QuantiGiocatori As Integer, Importanza As Integer,
											InizioGiornata As Integer, Conn As Object, Connessione As String, Torneo As String) As String
		Dim Ritorno As String = ""
		Dim Classifica As List(Of StrutturaGiocatore) = PrendeGiocatori(Mp, idAnno, idGiornata, Conn, Connessione)
		Dim QuantiGiocatoriPresenti As Integer = Classifica.Count - 1
		If QuantiGiocatoriPresenti > QuantiGiocatori Then
			Dim Scelti As New List(Of StrutturaGiocatore)
			For i As Integer = 0 To QuantiGiocatori - 1
				Scelti.Add(Classifica.Item(i))
			Next

			Dim GiornateAndata As New List(Of Integer)
			Dim idEventiAndata As New List(Of Integer)
			Dim GiornateRitorno As New List(Of Integer)
			Dim idEventiRitorno As New List(Of Integer)

			Dim Sql As String = "Select * From Eventi Where Upper(Descrizione) Like '%" & Torneo.ToUpper & "%' And UPPER(Descrizione) Like '% A%' Order By InizioGiornata"
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

					Sql = "Select * From Eventi Where Upper(Descrizione) Like '%" & Torneo.ToUpper & "%' And UPPER(Descrizione) Like '% R%' Order By InizioGiornata"
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

							Sql = "Select * From Eventi Where Upper(Descrizione) Like '%" & Torneo.ToUpper & "%' And UPPER(Descrizione) Like '%SEMIFINALE%' Order By InizioGiornata"
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

									Sql = "Select * From Eventi Where Upper(Descrizione) Like '%" & Torneo.ToUpper & "%' And UPPER(Descrizione) Like '%FINALE%' And UPPER(Descrizione) Not Like '%SEMIFINALE%' Order By InizioGiornata"
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

											For i As Integer = 1 To QuantiGiocatori - 1
												Sql = "Select * From ScontriDiretti Where NumeroSquadre=" & QuantiGiocatori & " And Giornata=" & i & " Order By Progressivo"
												Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													If Rec.Eof Then
														Ritorno = "ERROR: Nessuna giornata rilevata"
													Else
														Dim GiornataAndata As Integer = GiornateAndata.Item(i)
														Dim idEventoAndata As Integer = idEventiAndata.Item(i)
														Dim GiornataRitorno As Integer = GiornateRitorno.Item(i)
														Dim idEventoRitorno As Integer = idEventiRitorno.Item(i)

														Do Until Rec.Eof
															Dim Squadra1 As Integer = Scelti.Item(Rec("Squadra1").Value).idUtente
															Dim Squadra2 As Integer = Scelti.Item(Rec("Squadra2").Value).idUtente

															Sql = "Insert Into EventiPartite Values (" &
																" " & idAnno & ", " &
																" " & idEventoAndata & ", " &
																" " & GiornataAndata & ", " &
																" " & i & ", " &
																" " & Squadra1 & ", " &
																"'', " &
																" " & Squadra2 & ", " &
																"'', " &
																"-1 " &
																")"
															Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
															If Not Ritorno.Contains("ERROR") Then
																Sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & GiornataAndata & ", " & idEventoAndata & ")"
																Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
																If Ritorno.Contains("ERROR") Then
																	Exit For
																Else
																	Sql = "Insert Into EventiPartite Values (" &
																		" " & idAnno & ", " &
																		" " & idEventoRitorno & ", " &
																		" " & GiornataRitorno & ", " &
																		" " & i & ", " &
																		" " & Squadra2 & ", " &
																		"'', " &
																		" " & Squadra1 & ", " &
																		"'', " &
																		"-1 " &
																		")"
																	Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
																	If Ritorno.Contains("ERROR") Then
																		Exit For
																	Else
																		Sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & idGiornata & ", " & idEventoRitorno & ")"
																		Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
																		If Ritorno.Contains("ERROR") Then
																			Exit For
																		End If
																	End If
																End If
															End If
															Rec.MoveNext
														Loop
														Rec.Close
													End If
												End If
											Next

											If Not Ritorno.Contains("ERROR") Then
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
													If Not Ritorno.Contains("ERROR") Then
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
														If Not Ritorno.Contains("ERROR") Then
															Sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & GiornataSemiFinale & ", " & idEventoSemifinale & ")"
															Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
														End If
													End If
												End If
											End If

											If Not Ritorno.Contains("ERROR") Then
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
													If Not Ritorno.Contains("ERROR") Then
														Sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & GiornataFinale & ", " & idEventoFinale & ")"
														Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
													End If
												End If
											End If

											If Not Ritorno.Contains("ERROR") Then
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
