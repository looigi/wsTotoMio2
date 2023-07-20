Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports wsTotoMio2.clsRecordset

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
	Public Function ApreConcorso(idAnno As String, Scadenza As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim idModalita As String = ""
		Dim Descrizione As String = ""

		Dim sql As String = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
		If Not Ritorno.Contains("ERROR:") Then
			sql = "Select * From ModalitaConcorso Where Descrizione='Aperto'"
			Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Rec.Eof Then
					Ritorno = "ERROR: Nessuna tipologia rilevata"
				Else
					idModalita = Rec("idModalitaConcorso").Value
					Descrizione = Rec("Descrizione").Value
					Rec.Close

					sql = "Update Globale Set idModalitaConcorso=" & idModalita & ", idGiornata = idGiornata + 1, Scadenza='" & Scadenza & "' Where idAnno=" & idAnno
					Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					If Not Ritorno.Contains("ERROR") Then
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
												If Rec("idEvento").Value <> 1 Then
													sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & idGiornata & ", " & Rec("idEvento").Value & ")"
													Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
													If Rit.Contains("ERROR") Then
														Ritorno = Rit
														Exit Do
													End If
												End If

												Rec.MoveNext
											Loop
											Rec.Close
										End If
									End If

									If Not Ritorno.Contains("ERROR") Then
										sql = "Insert Into EventiCalendario Values (" & idAnno & ", " & idGiornata & ", 1)"
										Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
										If Rit.Contains("ERROR") Then
											Ritorno = Rit
										End If
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
				If idModalita <> "" Then
					Ritorno = idModalita & ";" & Descrizione
				End If
			Else
				sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMaxGiornataCoppa(idAnno As String, idCoppa As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Select Coalesce(Max(idGiornataVirtuale), 1) As idGiornata From EventiPartite Where idEvento In (Select idEvento From Eventi Where idCoppa = " & idCoppa & ") " &
			"And Risultato1 <> '' And Risultato1 Is Not Null And Risultato2 <> '' And Risultato2 Is Not Null And idAnno = " & idAnno
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessuna giornata rilevata"
			Else
				Ritorno = Rec("idGiornata").Value

				Rec.Close

				Sql = "Select Coalesce(Max(idGiornataVirtuale), 1) As idGiornata From EventiPartite Where idEvento In " &
					"(Select idEvento From Eventi A Left Join EventiTipologie B On A.idTipologia = B.idTipologia " &
					"Where idCoppa = " & idCoppa & " And B.Descrizione <> 'Semifinale' And B.Descrizione <> 'Finale') " &
					"And idAnno = " & idAnno
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = "ERROR: Nessuna giornata rilevata"
					Else
						Ritorno &= ";" & Rec("idGiornata").Value

						Rec.Close
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaVincitori(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim idModalita As String = ""
		Dim Sql As String = "Select * From EventiNomi Where Attiva='S'"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessuna coppa rilevata"
			Else
				Dim idCoppe As New List(Of Integer)
				Dim Descrizione As New List(Of String)
				Dim Finale As New List(Of Boolean)
				Dim Giocatori As New List(Of Integer)

				Do Until Rec.Eof
					Giocatori.Add(Rec("QuantiGiocatori").Value)
					idCoppe.Add(Rec("idCoppa").Value)
					Descrizione.Add(Rec("Descrizione").Value)
					Finale.Add(Rec("Finale").Value = "S")

					Rec.MoveNext
				Loop
				Rec.Close

				' CAMPIONATO
				Dim ev As New clsEventi
				Dim Classifica As List(Of clsEventi.StrutturaGiocatore) = ev.PrendeGiocatori(Server.MapPath("."), idAnno, 38, Conn, Connessione)
				Ritorno &= "Campione di TotoMIO;" & Classifica.Item(0).NickName & "§"
				Ritorno &= "Vice Campione;" & Classifica.Item(1).NickName & "§"
				Ritorno &= "Terzo;" & Classifica.Item(2).NickName & "§"
				Ritorno &= "Cucchiara di legno;" & Classifica.Item(Classifica.Count - 1).NickName & "§"

				' COPPE
				Dim Conta As Integer = 0

				For Each idCoppa As Integer In idCoppe
					If Finale.Item(Conta) Then
						Sql = "SELECT A.idPartita, D.NickName As Giocatore1, E.NickName As Giocatore2, A.Risultato1, A.Risultato2, Coalesce(A.idVincente, -99) As idVincente FROM EventiPartite A " &
								"Left Join Eventi B On A.idEvento = B.idEvento " &
								"Left Join EventiTipologie C On B.idTipologia = C.idTipologia " &
								"Left Join Utenti D On A.idGiocatore1 = D.idUtente And A.idAnno = D.idAnno " &
								"Left Join Utenti E On A.idGiocatore2 = E.idUtente And A.idAnno = E.idAnno " &
								"Left Join Risultati F On F.idAnno = A.idAnno And F.idConcorso = A.idGiornata And F.idUtente = A.idGiocatore1 " &
								"Left Join Risultati G On G.idAnno = A.idAnno And G.idConcorso = A.idGiornata And G.idUtente = A.idGiocatore2 " &
								"Where A.idAnno = " & idAnno & " And B.idCoppa = " & idCoppa & " And C.Descrizione = 'Finale'"
						Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = "ERROR: Nessuna finale rilevata"
							Else
								If Rec("idVincente").Value = -99 Or Rec("idVincente").Value = -1 Then
									Ritorno &= Descrizione.Item(Conta) & ";Non giocata finale§"
								Else
									If Rec("idVincente").Value = 1 Then
										Ritorno &= Descrizione.Item(Conta) & ";" & Rec("Giocatore1").Value & "§"
									Else
										Ritorno &= Descrizione.Item(Conta) & ";" & Rec("Giocatore2").Value & "§"
									End If
								End If
								Rec.Close
							End If
						End If
					Else
						Dim idGiornata As Integer = (Giocatori.Item(Conta) - 1) * 2
						Dim ClassificaTorneo As String = ev.CalcolaClassificaTorneo(Server.MapPath("."), idAnno, idGiornata, idCoppa, True, Conn, Connessione)
						If Not ClassificaTorneo.Contains("ERROR") Then
							Dim r() As String = ClassificaTorneo.Split(";")

							Ritorno &= Descrizione.Item(Conta) & ";" & r(1) & "§"
						Else
							Ritorno &= Descrizione.Item(Conta) & ";Non ancora creata§"
						End If
					End If

					Conta += 1
				Next
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ChiudeConcorso(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim idModalita As String = ""
		Dim Descrizione As String = ""

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
					idModalita = Rec("idModalitaConcorso").Value
					Descrizione = Rec("Descrizione").Value
					Rec.Close

					sql = "Update Globale Set idModalitaConcorso=" & idModalita & ", Scadenza='' Where idAnno=" & idAnno
					Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					If Ritorno.Contains("ERROR") Then
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

								sql = "Select A.*, B.Descrizione As Tipologia, C.Descrizione As Torneo, C.QuantiGiocatori, C.Importanza, " &
									"A.InizioGiornata, C.Descrizione As Torneo, B.Dettaglio " &
									"From Eventi A " &
									"Left Join EventiTipologie B On A.idTipologia = B.idTipologia " &
									"Left Join EventiNomi C On A.idCoppa = C.idCoppa " &
									"Where InizioGiornata=" & idGiornata & " Order By idEvento" ' idAnno=" & idAnno & " And idGiornata=" & idGiornata & " And idEvento<>1"
								Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Not Rec.Eof Then
										Dim ev As New clsEventi

										Do Until Rec.Eof
											If Rec("idEvento").Value <> 1 Then
												Ritorno = ev.GestioneEventi(Server.MapPath("."), idAnno, idGiornata, Rec("idEvento").Value,
																		Rec("QuantiGiocatori").Value, Rec("Importanza").Value,
																		Rec("InizioGiornata").Value, Rec("Tipologia").Value, Rec("Torneo").Value,
																		Rec("Dettaglio").Value, Rec("idCoppa").Value, Conn, Connessione)
												If Ritorno.Contains("ERROR") Then
													Exit Do
												End If
											End If

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
				If idModalita <> "" Then
					Ritorno = idModalita & ";" & Descrizione
				End If
			Else
				sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaClassificaCoppe(idAnno As String, idGiornata As String, Torneo As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim e As New clsEventi
		Dim Ritorno As String = e.CalcolaClassificaTorneo(Server.MapPath("."), idAnno, idGiornata, Torneo, False, Conn, Connessione)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ritornaNomiCoppe() As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From EventiNomi Order By idCoppa"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessuna coppa rilevata"
			Else
				Do Until Rec.Eof
					Ritorno &= Rec("idCoppa").Value & ";" & SistemaStringaPerRitorno(Rec("Descrizione").Value) & ";" & Rec("SemiFinale").Value & ";" & Rec("Finale").Value & "§"

					Rec.MoveNext
				Loop
				Rec.Close
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

				sql = "Update Globale Set idModalitaConcorso=" & idModalita & ", Scadenza='' Where idAnno=" & idAnno
				Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
				If Not Ritorno.Contains("ERROR") Then
					Ritorno = idModalita & ";" & Descrizione
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function LeggePartitaJolly(idAnno As String, idConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From PartiteJolly Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessuna partita jolly rilevata"
			Else
				Ritorno = Rec("idPartita").Value
				Rec.Close
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
						Else
							Ritorno = CreaPartitaJolly(Server.MapPath("."), idAnno, idConcorso, Conn, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Exit For
							End If
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
		Dim idGiornata As String = ""

		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun anno rilevato"
			Else
				idGiornata = Rec("idGiornata").Value
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

						Dim PartitaJolly As Integer = -1

						sql = "Select Coalesce(idPartita, -1) As idPartita From PartiteJolly Where idAnno=" & idAnno & " And idConcorso=" & idGiornata
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							PartitaJolly = Rec("idPartita").Value
							Rec.Close
						End If

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
																								 ModalitaConcorso, PartitaJolly)
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
																								 ModalitaConcorso, PartitaJolly)
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
																					ModalitaConcorso, PartitaJolly)
											Ritorno &= Controllo & "%"
										Else
											Dim Pronostici As New List(Of String)

											Do Until Rec.Eof
												Pronostici.Add(Rec("idPartita").Value & ";" & Rec("Pronostico").Value & ";" & Rec("Segno").Value)

												Rec.MoveNext
											Loop
											Rec.Close

											Dim Controllo As String = ControllaPunti(idAnno, idUtente, idGiornata, NickName,
																					 Partite, Pronostici, Conn, Connessione, Server.MapPath("."),
																					 ModalitaConcorso, PartitaJolly)
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
		' IdUtente;PuntiTotali|idPartita;Squadra1;Squadra2;Risultato;Segno;Pronostico;PronosticoSegno;PuntiPartita;Jolly§%

		' Aggiorna primi ultimi
		Dim Classifica As String = RitornaClassificaGenerale(Server.MapPath("."), idAnno, idGiornata, Conn, Connessione, True)
		If Classifica <> "" Then
			Dim c() As String = Classifica.Split("§")
			Dim PrimaRiga() As String = c(0).Split(";")
			Dim idUltimo As Integer = PrimaRiga(0)
			Dim UltimaRiga() As String = c(c.Count - 2).Split(";")
			Dim idPrimo As Integer = UltimaRiga(0)
			Dim Ritorno2 As String = "OK"

			sql = "Update RisultatiAltro Set Vittorie = 0, Ultimo = 0 Where idAnno=" & idAnno & " And idConcorso=" & idGiornata
			Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)

			sql = "Select * From RisultatiAltro Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente=" & idPrimo
			Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno2 = Rec
			Else
				If Rec.Eof Then
					sql = "Insert Into RisultatiAltro Values (" & idAnno & ", " & idGiornata & ", " & idPrimo & ", 1, 0, 0)"
					Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
				Else
					sql = "Update RisultatiAltro Set Vittorie = 1 Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente = " & idPrimo
					Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
				End If
			End If

			sql = "Select * From RisultatiAltro Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente=" & idUltimo
			Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno2 = Rec
			Else
				If Rec.Eof Then
					sql = "Insert Into RisultatiAltro Values (" & idAnno & ", " & idGiornata & ", " & idUltimo & ", 0, 1, 0)"
					Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
				Else
					sql = "Update RisultatiAltro Set Ultimo = 1 Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente = " & idUltimo
					Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class