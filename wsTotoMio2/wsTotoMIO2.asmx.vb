Imports System.ComponentModel
Imports System.Diagnostics.Eventing.Reader
Imports System.Net.Mail
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports ADODB

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsTotoMIO2
	Inherits System.Web.Services.WebService

	Private Structure StrutturaCoppe
		Dim idCoppa As Integer
		Dim Descrizione As String
		Dim Importanza As Integer
		Dim QuantiGiocatori As Integer
		Dim Percentuale As Integer
		Dim Semifinale As Boolean
		Dim Finale As Boolean
	End Structure

	<WebMethod()>
	Public Function Test() As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From Utenti"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)

		Do Until Rec.eof
			Ritorno &= Rec(2).Value & " " & Rec(3).Value & ";"

			Rec.movenext
		Loop
		Rec.Close

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function TestMail(Operazione As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = "*"

		Dim Testo As String = ""
		Testo = "Prova<br />"
		Testo &= "Per entrare nel sito e vedere il resto: <a href=""" & IndirizzoSito & """>Click QUI</a>"
		InvaMailATutti(Server.MapPath("."), 1, "TotoMIO:Mail di prova", Testo, Conn, Connessione, "Controllo")

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function TesteCreazioneCoppa() As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = "*"

		Dim c As New clsEventi
		Ritorno = c.CreazioneCoppa(Server.MapPath("."), 1, 5, 6, 1, 1, Conn, Connessione)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaDatiGenerali(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Select *, C.Descrizione As ModalitaConcorso From Globale A " &
			"Left Join Anni B On A.idAnno = B.idAnno " &
			"Left Join ModalitaConcorso C On C.idModalitaConcorso = A.idModalitaConcorso " &
			"Where A.idAnno=" & idAnno
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun valore ritornato"
			Else
				Dim idGiornata As Integer = Rec("idGiornata").Value

				Ritorno = Rec("idGiornata").Value & ";" &
						Rec("idModalitaConcorso").Value & ";" &
						Rec("ModalitaConcorso").Value & ";" &
						Rec("Scadenza").Value & ";" &
						"|"
				Rec.Close

				Sql = "Select * From Anni Order By idAnno Desc"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof
						Ritorno &= Rec("idAnno").Value & ";" & Rec("Descrizione").Value & "§"

						Rec.MoveNext
					Loop
					Rec.Close

					Ritorno &= "|"
					Sql = "Select Distinct C.Descrizione As Tipologia, C.Dettaglio, D.Descrizione As Torneo, E.idGiornataVirtuale From EventiCalendario A " &
						"Left Join Eventi B On A.idEvento = B.idEvento " &
						"Left Join EventiTipologie C On B.idTipologia = C.idTipologia " &
						"Left Join EventiNomi D On B.idCoppa = D.idCoppa " &
						"Left Join EventiPartite E On A.idAnno = E.idAnno And A.idEvento = B.idEvento And E.idGiornata = A.idGiornata And E.idPartita = 1 " &
						"Where A.idAnno = " & idAnno & " And A.idGiornata = " & idGiornata & " And B.idEvento Is Not null"
					Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Do Until Rec.Eof
							If Rec("Tipologia").Value = "Partita" Then
								Ritorno &= Rec("Tipologia").Value & " " & Rec("idGiornataVirtuale").Value & " " & Rec("Dettaglio").Value & " " & Rec("Torneo").Value & "§"
							Else
								Ritorno &= Rec("Tipologia").Value & " " & Rec("Dettaglio").Value & " " & Rec("Torneo").Value & "§"
							End If

							Rec.MoveNext
						Loop
						Rec.Close
					End If
				End If
			End If
		End If
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AggiornaDatiGenerali(idAnno As String, idGiornata As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Update Globale Set " &
			"idGiornata=" & idGiornata & " " &
			"Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
		If Not Ritorno.Contains(StringaErrore) Then
			Ritorno = "*"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AvanzaGiornata(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Dim Sql As String = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
		If Not Ritorno.Contains("Error: ") Then
			Sql = "Update Globale Set " &
				"idGiornata=idGiornata+1 " &
				"Where idAnno=" & idAnno
			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
			If Not Ritorno.Contains(StringaErrore) Then
				Sql = "Select Coalesce(Count(*),0) From Utenti Where idAnno=" & idAnno & " And Eliminato='N'"
				Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Dim QuantiGiocatori As Integer = Rec(0).Value
					Rec.Close

					If QuantiGiocatori = 0 Then
						Ritorno = "ERROR: Nessun giocatore presente"
					Else
						Ritorno = AggiungeGiornataDiCampionatoAEventi(Conn, Connessione, idAnno)
						If Not Ritorno.Contains("ERROR") Then
							Ritorno = RitornaEventiGiornata(idAnno)
						End If
					End If
				End If
			End If

			If Not Ritorno.Contains("Error") Then
				Sql = "commit"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			Else
				Sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaEventiGiornata(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Dim idGiornata As String = RitornaGiornata(Server.MapPath("."), Conn, Connessione, idAnno)
		If Not idGiornata.Contains("ERROR") Then
			Dim Sql As String = "Select A.idAnno, A.idGiornata, A.idEvento, A.idTipologia, B.Descrizione As Evento, C.Descrizione As Tipologia From " &
				"EventiCalendario As A " &
				"Left Join Eventi As B On A.idEvento = B.idEvento " &
				"Left Join EventiTipologie As C On A.idTipologia = C.idTipologia " &
				"Where A.idAnno=" & idAnno & " And A.idGiornata=" & idGiornata
			Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Not Rec.Eof Then
					Do Until Rec.Eof
						Ritorno &= Rec("idAnno").Value & ";"
						Ritorno &= Rec("idGiornata").Value & ";"
						Ritorno &= Rec("idEvento").Value & ";"
						Ritorno &= Rec("idTipologia").Value & ";"
						Ritorno &= SistemaStringaPerRitorno(Rec("Evento").Value) & ";"
						Ritorno &= SistemaStringaPerRitorno(Rec("Tipologia").Value) & "§"

						Rec.MoveNext
					Loop
				Else
					Ritorno = "Nessun evento in programma per la giornata " & idGiornata
				End If
				Rec.Close
			End If
		Else
			Ritorno = idGiornata
		End If

		Return Ritorno
	End Function

	Public Function AggiungeGiornataDiCampionatoAEventi(Conn As Object, Connessione As String, idAnno As String) As String
		Dim Ritorno As String = ""
		Dim idGiornata As String = RitornaGiornata(Server.MapPath("."), Conn, Connessione, idAnno)
		If Not idGiornata.Contains("ERROR") Then
			' Inserimento Partita di campionato
			Dim Sql As String = "Insert Into EventiCalendario Values (" &
				" " & idAnno & ", " &
				" " & idGiornata & ", " &
				"1, " &
				"1 " &
				")"
			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
			If Not Ritorno.Contains(StringaErrore) Then
				' Gestire tutti gli altri eventi giornata per giornata
				Select Case idGiornata
					Case 10
				End Select

				Ritorno = "*"
			End If
		Else
			Ritorno = idGiornata
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function TestRandom() As String
		Dim Ritorno As String = ""
		For i As Integer = 1 To 10
			Ritorno &= GetRandom(1, 10) & ";"
		Next
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function PuliziaDatiDebug(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Dim Sql As String = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
		If Not Ritorno.Contains("Error:") Then

			Sql = "Delete From Concorsi Where idAnno=" & idAnno
			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
			If Not Ritorno.Contains(StringaErrore) Then

				Sql = "Delete From EventiCalendario Where idAnno=" & idAnno
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
				If Not Ritorno.Contains(StringaErrore) Then

					Sql = "Delete From EventiPartite Where idAnno=" & idAnno
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
					If Not Ritorno.Contains(StringaErrore) Then

						Sql = "Update Globale Set idGiornata=0, idModalitaConcorso=0, Scadenza='' Where idAnno=" & idAnno
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
						If Not Ritorno.Contains(StringaErrore) Then

							Sql = "Delete From Pronostici Where idAnno=" & idAnno
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
							If Not Ritorno.Contains(StringaErrore) Then

								Sql = "Delete From Bilancio Where idAnno=" & idAnno
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
								If Not Ritorno.Contains(StringaErrore) Then

									Sql = "Delete From Risultati Where idAnno=" & idAnno
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
									If Not Ritorno.Contains(StringaErrore) Then

										Sql = "Delete From Utenti Where idAnno=" & idAnno & " And Cognome Like '%Utente Cognome%'"
										Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
										If Not Ritorno.Contains(StringaErrore) Then

											Sql = "Delete From RisultatiAltro Where idAnno=" & idAnno
											Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
											If Not Ritorno.Contains(StringaErrore) Then

												Sql = "Delete From PartiteJolly Where idAnno=" & idAnno
												Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
												If Not Ritorno.Contains(StringaErrore) Then

													Sql = "Delete From SquadreRandom Where idAnno=" & idAnno
													Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
													If Not Ritorno.Contains(StringaErrore) Then
														If Not Ritorno.Contains(StringaErrore) Then

															Sql = "Delete From PartiteScelte Where idAnno=" & idAnno
															Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
															If Not Ritorno.Contains(StringaErrore) Then

																Sql = "Delete From Chat Where idAnno=" & idAnno
																Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
																If Not Ritorno.Contains(StringaErrore) Then

																	Sql = "Update PremioPerFinto Set Importo = 0 Where idAnno=" & idAnno
																	Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
																	If Not Ritorno.Contains(StringaErrore) Then
																		Ritorno = CreaEventi()
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

			If Not Ritorno.Contains("Error") Then
				Sql = "commit"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			Else
				Sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaEventi() As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""

		sql = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione)
		If Not Ritorno.Contains("Error:") Then
			sql = "Delete From Eventi"
			Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			If Not Ritorno.Contains(StringaErrore) Then
				sql = "Select * From EventiNomi Where Attiva = 'S' Order By idCoppa"
				Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = "ERROR: Nessun evento nome rilevato"
					Else
						Dim Coppe As New List(Of StrutturaCoppe)
						Dim Inizio As New List(Of Integer)

						Do Until Rec.Eof
							Dim s As New StrutturaCoppe
							s.idCoppa = Rec("idCoppa").Value
							s.Descrizione = Rec("Descrizione").Value
							s.Importanza = Rec("Importanza").Value
							s.Percentuale = Rec("Percentuale").Value
							s.QuantiGiocatori = Rec("QuantiGiocatori").Value
							s.Semifinale = IIf(Rec("Semifinale").value = "S", True, False)
							s.Finale = IIf(Rec("Finale").value = "S", True, False)
							Coppe.Add(s)

							Rec.MoveNext
						Loop
						Rec.Close

						Dim QuanteCoppe As Integer = Coppe.Count
						Dim GiornateTotali As Integer = 38
						Dim PartiteMax As Integer = 0

						For Each s As StrutturaCoppe In Coppe
							Dim Quanti As Integer = s.QuantiGiocatori
							Dim Partite As Integer = (Quanti - 1) * 2
							If s.Semifinale Then
								Partite += 2
							End If
							If s.Finale Then
								Partite += 1
							End If
							If Partite > PartiteMax Then
								PartiteMax = Partite
							End If
							'PartiteTotali += (Quanti - 1) * 2
							'If s.Semifinale Then
							'	PartiteTotali += 2
							'End If
							'If s.Finale Then
							'	PartiteTotali += 2
							'End If

							'' Creazione
							'PartiteTotali += 1
						Next
						If PartiteMax > GiornateTotali Then
							Ritorno = "ERROR: Troppe partite da calcolare: " & PartiteMax & "/" & GiornateTotali
						Else
							Dim InizioGiornata As Integer = 5
							Dim FineGiornata As Integer = GiornateTotali - QuanteCoppe - 1
							Dim Salto As New List(Of Integer)
							'Dim InizioPerCoppa As Integer = InizioGiornata

							For i As Integer = 1 To 9
								Inizio.Add(-1)
								Salto.Add(-1)
							Next

							For Each s As StrutturaCoppe In Coppe
								Dim QuantiGiocatori As Integer = s.QuantiGiocatori
								Dim QuantoSalto As Integer = Int((FineGiornata - InizioGiornata) / ((QuantiGiocatori - 1) * 2))

								Inizio(s.idCoppa) = InizioGiornata + (s.Importanza - 1)
								Salto(s.idCoppa) = QuantoSalto
								'InizioPerCoppa += 1
							Next

							Dim idEvento As Integer = 2
							sql = "Select * From EventiTipologie Order By idTipologia"
							Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ritorno = "ERROR: Nessun evento tipologia rilevato"
								Else
									Do Until Rec.Eof
										Dim idEventoTipologia As Integer = Rec("idTipologia").Value
										Dim Descrizione As String = Rec("Descrizione").Value

										For Each s As StrutturaCoppe In Coppe
											Dim idCoppa As Integer = s.idCoppa
											Dim Ok As Boolean = True
											Dim QuantePartite As Integer = 1

											'Select * FROM `Eventi` As A 
											'Left Join EventiTipologie B On A.idTipologia = B.idTipologia
											'Left Join EventiNomi C On A.idCoppa = C.idCoppa
											'Where C.Descrizione Like '%2%'

											If Descrizione = "Chiusura" Then
												Inizio(idCoppa) -= Salto(idCoppa)
											End If

											If Descrizione = "Finale" Then
												Inizio(idCoppa) = GiornateTotali - (s.Importanza - 1)
											End If

											If Descrizione = "Semifinale" And s.Semifinale = False Then
												Ok = False
											Else
												If Descrizione = "Semifinale" And s.Semifinale = True Then
													Inizio(idCoppa) = GiornateTotali - QuanteCoppe - 1
												End If
											End If

											If Descrizione = "Finale" And s.Finale = False Then
												Ok = False
											End If

											If Ok Then
												If Descrizione.Contains("Partita") Then
													QuantePartite = (s.QuantiGiocatori - 1)
												Else
													QuantePartite = 1
												End If

												For i As Integer = 1 To QuantePartite
													sql = "Insert Into Eventi Values (" &
														" " & idEvento & ", " &
														" " & idCoppa & ", " &
														" " & idEventoTipologia & ", " &
														" " & Inizio(s.idCoppa) & " " &
														")"
													Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
													If Ritorno.Contains(StringaErrore) Then
														Exit For
													End If

													If Descrizione <> "Creazione" Then
														Inizio(s.idCoppa) += Salto(idCoppa)
													Else
														Inizio(s.idCoppa) += 1
													End If
												Next
												idEvento += 1

												If Ritorno.Contains(StringaErrore) Then
													Exit For
												End If
											End If
										Next

										If Ritorno.Contains(StringaErrore) Then
											Exit Do
										End If

										Rec.MoveNext
									Loop
									Rec.Close
								End If
							End If
						End If
					End If
				End If
			End If

			If Not Ritorno.Contains("Error") Then
				sql = "commit"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione)
			Else
				sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function CreaDatiDiDebug(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""
		Dim NomiSquadre() As String = {"Lazio", "Roma", "Juventus", "Torino", "Milan", "Inter", "Napoli", "Sampdoria", "Genoa", "Acireale", "Afragolese",
			"Castelverde", "Real Madrid", "Tiburtino III", "Padova", "Parma", "Palermo", "Cerebrolesi", "Ignobili", "Fetenti", "Sgarrupati", "Smutandati",
			"Tedeschi", "Svizzeri", "Babbioni", "Gabibbi", "Bounty", "Davidini", "Angeliconi", "Micaeleschi", "Testoni", "Arsenal", "Ravenna",
			"Reggiana", "Reggina", "Atalanta", "Bari", "Cesena", "Cagliari", "Ternana"}

		Dim Quanti As Integer = 10

		Ritorno = PuliziaDatiDebug(idAnno)
		If Not Ritorno.Contains("ERROR:") Then
			sql = "Start transaction"
			Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			If Not Ritorno.Contains("ERROR") Then
				For i As Integer = 1 To Quanti
					sql = "Insert Into Utenti Values (" &
						" " & idAnno & ", " &
						" " & i + 1 & ", " &
						"'Utente " & i & "', " &
						"'Utente Cognome " & i & "', " &
						"'Utente Nome " & i & "', " &
						"'utente', " &
						"'utente" & i & "@utente.it', " &
						"1, " &
						"'N' " &
						")"
					Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					If Ritorno.Contains(StringaErrore) Then
						Exit For
					Else
						Dim gi As New GestioneImmagini
						gi.CreaAvatar(Server.MapPath("."), idAnno, i + 1, "Utente " & i, "Nome", "Cognome")
					End If
				Next

				If Not Ritorno.Contains("ERROR") Then
					sql = "Select Count(*) From Utenti Where Eliminato = 'N'"
					Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = "ERROR: Nessun utente rilevato"
						Else
							Dim QuantiGiocatori As Integer = Rec(0).Value
							Rec.Close

							' Creazione concorsi
							sql = "Delete From Concorsi Where idAnno = " & idAnno
							Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
							If Not Ritorno.Contains(StringaErrore) Then
								For i = 1 To 38
									For k = 1 To 10
										Dim Ris1 As Integer = GetRandom(0, 4)
										Dim Ris2 As Integer = GetRandom(0, 4)
										Dim Risultato As String = Ris1 & "-" & Ris2
										Dim Segno As String = ""
										If Ris1 > Ris2 Then
											Segno = "1"
										Else
											If Ris1 < Ris2 Then
												Segno = "2"
											Else
												Segno = "X"
											End If
										End If
										Dim x As Integer = GetRandom(0, NomiSquadre.Count - 1)
										Dim Casa As String = NomiSquadre(x)
										x = GetRandom(0, NomiSquadre.Count - 1)
										Dim Fuori As String = NomiSquadre(x)
										sql = "Insert Into Concorsi Values (" &
											" " & idAnno & ", " &
											" " & i & ", " &
											" " & k & ", " &
											"'" & Casa & "', " &
											"'" & Fuori & "', " &
											"'" & Risultato & "', " &
											"'" & Segno & "' " &
											")"
										Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
										If Ritorno.Contains(StringaErrore) Then
											Exit For
										End If
									Next k

									If Ritorno.Contains(StringaErrore) Then
										Exit For
									Else
										'Ritorno = CreaPartitaJolly(Server.MapPath("."), idAnno, i, Conn, Connessione)
										'If Ritorno.Contains(StringaErrore) Then
										'	Exit For
										'End If
									End If
								Next i
							End If

							If Not Ritorno.Contains(StringaErrore) Then
								' Creazione Pronostici
								sql = "Delete From Pronostici Where idAnno = " & idAnno
								Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
								If Not Ritorno.Contains(StringaErrore) Then
									Dim QuantePartite As Integer = 10

									For i = 1 To 38
										For z = 1 To QuantiGiocatori
											For k = 1 To QuantePartite
												Dim Ris1 As Integer = GetRandom(0, 4)
												Dim Ris2 As Integer = GetRandom(0, 4)
												Dim Risultato As String = Ris1 & "-" & Ris2
												Dim Segno As String = ""
												If Ris1 > Ris2 Then
													Segno = "1"
												Else
													If Ris1 < Ris2 Then
														Segno = "2"
													Else
														Segno = "X"
													End If
												End If
												sql = "Insert Into Pronostici Values (" &
													" " & idAnno & ", " &
													" " & z & ", " &
													" " & i & ", " &
													" " & k & ", " &
													"'" & Risultato & "', " &
													"'" & Segno & "' " &
													")"
												Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
												If Ritorno.Contains(StringaErrore) Then
													Exit For
												End If
											Next k

											If Ritorno.Contains(StringaErrore) Then
												Exit For
											Else
												sql = "Delete From PartiteScelte Where idAnno=" & idAnno & " And idConcorso=" & i & " And idUtente=" & z
												Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
												If Not Ritorno.Contains("ERROR:") Then
													Dim idPartitaScelta As Integer = GetRandom(1, QuantePartite)

													sql = "Insert Into PartiteScelte Values (" & idAnno & ", " & i & ", " & z & ", " & idPartitaScelta & ")"
													Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
												Else
													Exit For
												End If
											End If
										Next z
										If Ritorno.Contains("ERROR") Then
											Exit For
										End If
									Next i
								End If

								If Not Ritorno.Contains("ERROR") Then
									sql = "Update Globale Set idGiornata=0, idModalitaConcorso=0, Scadenza='' Where idAnno=" & idAnno
									Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
									If Not Ritorno.Contains(StringaErrore) Then
										Ritorno = "OK"
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
	Public Function TestCreazioneImmagine() As String
		Dim gi As New GestioneImmagini
		gi.CreaAvatar(Server.MapPath("."), 1, 1, "Catadiottro", "Pippeto", "Polpetti")
		Return "OK"
	End Function

End Class