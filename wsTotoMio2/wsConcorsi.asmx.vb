Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looConcorsiTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsConcorsi
	Inherits System.Web.Services.WebService

	'<WebMethod()>
	'Public Function AggiungeConcorso(idAnno As String, Dati As String) As String
	'	Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
	'	Dim Conn As Object = New clsGestioneDB(TipoServer)
	'	Dim Ritorno As String = ""
	'	Dim sql As String = ""

	'	sql = "Start transaction"
	'	Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
	'	If Not Ritorno.Contains("ERROR:") Then
	'		sql = "Select Coalesce(Max(idGiornata)+1,1) From Concorsi Where idAnno=" & idAnno
	'		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
	'		If TypeOf (Rec) Is String Then
	'			Ritorno = Rec
	'		Else
	'			Dim idGiornata As String = Rec(0).Value

	'			Dim Dati2() As String = Dati.Split("§")
	'			For Each D As String In Dati2
	'				Dim D2() As String = D.Split(";")

	'				' idAnno	idGiornata	idPartita	Prima	Seconda	Risultato	Segno
	'				sql = "Insert Into Concorsi Values (" &
	'					" " & idAnno & ", " &
	'					" " & idGiornata & ", " &
	'					" " & D2(0) & ", " &
	'					"'" & SistemaStringaPerDB(D2(1)) & "', " &
	'					"'" & SistemaStringaPerDB(D2(2)) & "', " &
	'					"'" & SistemaStringaPerDB(D2(3)) & "', " &
	'					"'' " &
	'					")"
	'				Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
	'				If Ritorno.Contains(StringaErrore) Then
	'					Exit For
	'				End If
	'			Next
	'		End If

	'		If Ritorno = "OK" Then
	'			sql = "commit"
	'			Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione)
	'		Else
	'			sql = "rollback"
	'			Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione)
	'		End If
	'	End If

	'	Return Ritorno
	'End Function

	<WebMethod()>
	Public Function ApreConcorso(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Dim sql As String = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
		If Not Ritorno.Contains("ERROR:") Then
			sql = "Select * From ModalitaConcorso Where Descrizione='Aperto'"
			Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Rec.Eof Then
					Ritorno = "ERROR: Nessuna tipologia rilevata"
				Else
					Dim idModalita As String = Rec("idModalitaConcorso").Value
					Dim Descrizione As String = Rec("Descrizione").Value
					Rec.Close

					sql = "Update Globale Set idModalitaConcorso=" & idModalita & ", idGiornata = idGiornata + 1 Where idAnno=" & idAnno
					Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					If Not Ritorno.Contains("ERROR") Then
						Ritorno = idModalita & ";" & Descrizione

						sql = "Select * From Globale Where idAnno=" & idAnno
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = "ERROR: Nessuna giornata rilevata"
							Else
								Dim idGiornata As Integer = Rec("idGiornata").Value
								Rec.Close

								sql = "Select * From Eventi Where InizioGiornata=" & idGiornata
								Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									sql = "Delete From EventiCalendario Where idAnno=" & idAnno & " And idGiornata=" & idGiornata
									Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
									If Not Ritorno.Contains("ERROR") Then
										If Rec.Eof Then
										Else
											Do Until Rec.Eof
												sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & idGiornata & ", " & Rec("idEvento").Value & ")"
												Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
												If Ritorno.Contains("ERROR") Then
													Exit Do
												End If

												Rec.MoveNext
											Loop
											Rec.Close
										End If
									End If

									If Not Ritorno.Contains("ERROR") Then
										sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & idGiornata & ", 1)"
										Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
									End If
								End If
							End If
						End If
					End If
				End If
			End If

			If Ritorno = "OK" Then
				sql = "commit"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			Else
				sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ChiudeConcorso(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Dim sql As String = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
		If Not Ritorno.Contains("ERROR:") Then
			sql = "Select * From ModalitaConcorso Where Descrizione='Controllato'"
			Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Rec.Eof Then
					Ritorno = "ERROR: Nessuna tipologia rilevata"
				Else
					Dim idModalita As String = Rec("idModalitaConcorso").Value
					Dim Descrizione As String = Rec("Descrizione").Value
					Rec.Close

					sql = "Update Globale Set idModalitaConcorso=" & idModalita & " Where idAnno=" & idAnno
					Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					If Not Ritorno.Contains("ERROR") Then
						Ritorno = idModalita & ";" & Descrizione
					Else
						sql = "Select * From Globale Where idAnno=" & idAnno
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = "ERROR: Nessuna giornata rilevata"
							Else
								Dim idGiornata As Integer = Rec("idGiornata").Value
								Rec.Close

								sql = "Select * From EventiCalendario Where idAnno=" & idAnno & " And idGiornata=" & idGiornata & " And idEvento<>1"
								Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Not Rec.Eof Then
										Dim ev As New clsEventi

										Do Until Rec.Eof
											Ritorno = ev.GestioneEventi(Server.MapPath("."), idAnno, idGiornata, Rec("idEvento").Value, Conn, Connessione)

											Rec.MoveNext
										Loop
										Rec.CLose
									End If
								End If
							End If
						End If
					End If
				End If
			End If

			If Ritorno = "OK" Then
				sql = "commit"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			Else
				sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function impostaConcorsoPerControllo(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From ModalitaConcorso Where Descrizione='Da Controllare'"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessuna tipologia rilevata"
			Else
				Dim idModalita As String = Rec("idModalitaConcorso").Value
				Dim Descrizione As String = Rec("Descrizione").Value
				Rec.Close

				sql = "Update Globale Set idModalitaConcorso=" & idModalita & " Where idAnno=" & idAnno
				Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
				If Not Ritorno.Contains("ERROR") Then
					Ritorno = idModalita & ";" & Descrizione
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaConcorso(idAnno As String, idConcorso As String, Dati As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""

		sql = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
		If Not Ritorno.Contains("ERROR:") Then
			sql = "Delete From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
			Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			If Not Ritorno.Contains("ERROR:") Then
				Dim Dati2() As String = Dati.Split("§")
				For Each D As String In Dati2
					If D <> "" Then
						Dim D2() As String = D.Split(";")

						' idAnno	idGiornata	idPartita	Prima	Seconda	Risultato	Segno
						sql = "Insert Into Concorsi Values (" &
							" " & idAnno & ", " &
							" " & idConcorso & ", " &
							" " & D2(0) & ", " &
							"'" & SistemaStringaPerDB(D2(1)) & "', " &
							"'" & SistemaStringaPerDB(D2(2)) & "', " &
							"'" & SistemaStringaPerDB(D2(3)) & "', " &
							"'" & SistemaStringaPerDB(D2(4)) & "' " &
							")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Ritorno.Contains(StringaErrore) Then
							Exit For
						End If
					End If
				Next
			End If

			If Ritorno = "OK" Then
				sql = "commit"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			Else
				sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornoConcorso(idAnno As String, idConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""

		sql = "Select * From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " Order By idPartita"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun dato presente"
			Else
				Do Until Rec.Eof
					Ritorno &= Rec("idPartita").Value & ";" & SistemaStringaPerRitorno(Rec("Prima").Value) & ";" &
						SistemaStringaPerRitorno(Rec("Seconda").Value) & ";" & Rec("Risultato").Value & ";" &
						Rec("Segno").Value & "§"
					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function controllaConcorso(idAnno As String, idUtente As String, ModalitaConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From Globale Where idAnno=" & idAnno
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun anno rilevato"
			Else
				Dim idGiornata As String = Rec("idGiornata").Value
				Rec.Close

				sql = "Select * From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " Order By idPartita"
				Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = "ERROR: Nessun concorso rilevato"
					Else
						Dim Partite As New List(Of String)

						Do Until Rec.Eof
							Partite.Add(Rec("idPartita").Value & ";" & SistemaStringaPerRitorno(Rec("Prima").Value) & ";" &
										SistemaStringaPerRitorno(Rec("Seconda").Value) & ";" &
										Rec("Risultato").Value & ";" & Rec("Segno").Value)

							Rec.MoveNext
						Loop
						Rec.Close

						sql = "Select A.NickName, B.idTipologia, B.Descrizione From Utenti A " &
							"Left Join UtentiTipologie B On A.idTipologia = B.idTipologia " &
							"Where idAnno=" & idAnno & " And idUtente=" & idUtente
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = "ERROR: Nessun utente rilevato"
							Else
								Dim idTipologia As Integer = Rec("idTipologia").Value
								Dim NickName As String = Rec("NickName").Value
								Dim Tipologia As String = Rec("Descrizione").Value
								Rec.Close

								If idTipologia = 0 Or ModalitaConcorso = "Controllato" Then
									' Controllo per amministratore
									sql = "Select * From Utenti Where idAnno=" & idAnno & " And Eliminato='N'"
									Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										If Rec.Eof Then
											Ritorno = "ERROR: Nessun utente rilevato"
										Else
											Dim idUtenti As New List(Of String)
											Dim NickNames As New List(Of String)

											Do Until Rec.Eof
												idUtenti.Add(Rec("idUtente").Value)
												NickNames.Add(Rec("NickName").Value)

												Rec.MoveNext
											Loop
											Rec.Close

											Dim q As Integer = 0
											For Each id As String In idUtenti
												Dim NN As String = NickNames.Item(q)
												q += 1
												sql = "Select * From Pronostici Where idAnno=" & idAnno & " And idUtente=" & id & " And idConcorso=" & idGiornata & " Order By idPartita"
												Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													If Rec.Eof Then
														Dim Controllo As String = ControllaPunti(idAnno, id, idGiornata, NN,
																								 Partite, New List(Of String), Conn, Connessione, Server.MapPath("."),
																								 ModalitaConcorso)
														Ritorno &= Controllo & "%"
													Else
														Dim Pronostici As New List(Of String)

														Do Until Rec.Eof
															Pronostici.Add(Rec("idPartita").Value & ";" & Rec("Pronostico").Value & ";" & Rec("Segno").Value)

															Rec.MoveNext
														Loop
														Rec.Close

														Dim Controllo As String = ControllaPunti(idAnno, id, idGiornata, NN,
																								 Partite, Pronostici, Conn, Connessione, Server.MapPath("."),
																								 ModalitaConcorso)
														Ritorno &= Controllo & "%"
													End If
												End If
											Next
										End If
									End If
								Else
									' Controllo per utente
									sql = "Select * From Pronostici Where idAnno=" & idAnno & " And idUtente=" & idUtente & " And idConcorso=" & idGiornata & " Order By idPartita"
									Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										If Rec.Eof Then
											Dim Controllo As String = ControllaPunti(idAnno, idUtente, idGiornata, NickName,
																					Partite, New List(Of String), Conn, Connessione, Server.MapPath("."),
																					ModalitaConcorso)
											Ritorno &= Controllo & "%"
										Else
											Dim Pronostici As New List(Of String)

											Do Until Rec.Eof
												Pronostici.Add(Rec("idPartita").Value & ";" & Rec("Risultato").Value & ";" & Rec("Segno").Value)

												Rec.MoveNext
											Loop
											Rec.Close

											Dim Controllo As String = ControllaPunti(idAnno, idUtente, idGiornata, NickName,
																					 Partite, Pronostici, Conn, Connessione, Server.MapPath("."),
																					 ModalitaConcorso)
											Ritorno &= Controllo & "%"
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If

		' 1;28|1;Pippa;Pippetta;1-2;2;1-1;X;3§%
		' IdUtente;PuntiTotali|idPartita;Squadra1;Squadra2;Risultato;Segno;Pronostico;PronosticoSegno;PuntiPartita§%
		Return Ritorno
	End Function

End Class