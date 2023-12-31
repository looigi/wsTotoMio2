﻿Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Linq.Expressions
Imports System.Net.Security
Imports System.Runtime.CompilerServices
Imports System.Security.Policy
Imports System.Threading
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Windows.Forms
Imports wsTotoMio2.clsRecordset

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looConcorsiTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsConcorsi
	Inherits System.Web.Services.WebService

	Public Structure StruttPronostico
		Dim Pronostico As String
		Dim Quante As Integer
	End Structure

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
		Dim idGiornata As Integer

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
								idGiornata = Rec("idGiornata").Value
								Rec.Close

								sql = "Select A.idEvento, Coalesce(B.Descrizione, '') As NomeEvento, Coalesce(C.Descrizione, '') As Tipologia, Coalesce(C.Dettaglio, '') As Dettaglio From Eventi A " &
									"Left Join EventiNomi B On A.idCoppa = B.idCoppa " &
									"Left Join EventiTipologie C On A.idTipologia = C.idTipologia " &
									"Where InizioGiornata=" & idGiornata
								Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									Dim Eventi As String = ""

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

													Eventi &= Rec("NomeEvento").Value & " " & Rec("Tipologia").Value & " " & Rec("Dettaglio").Value & "<br />"
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
										Else
											Dim Testo As String = ""
											Testo = "E' stato aperto il concorso TotoMIO numero " & idGiornata & ".<br />"
											Testo &= "Eventi della giornata:<br /><br />"
											Testo &= "Partita di campionato<br /><style=""font-weight: bold;"">"
											Testo &= Eventi
											Testo &= "</style><br />Chiusura concorso: <style=""font-weight: bold;"">" & Scadenza & "</style><br />"
											Testo &= "Per partecipare: <a href=""" & IndirizzoSito & """>Click QUI</a>"
											InvaMailATutti(Server.MapPath("."), idAnno, "TotoMIO: Apertura concorso " & idGiornata, Testo, Conn, Connessione, "Apertura")
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If

			If Ritorno = "OK" Then
				GestisceTorneo23(Server.MapPath("."), idAnno, idGiornata, Conn, Connessione)
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

				Dim Totale As Single

				Sql = "Select Coalesce(Sum(Importo),0) As Totale From Bilancio Where idAnno=" & idAnno & " And idMovimento=1 And Eliminato='N'"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
					Else
						Totale = Rec("Totale").Value
					End If
					Rec.Close
				End If

				Sql = "Select Coalesce(Sum(Importo),0) As Totale From Bilancio Where idAnno=" & idAnno & " And idMovimento<>1 And Eliminato='N'"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
					Else
						Totale -= Rec("Totale").Value
					End If
					Rec.Close
				End If

				Dim Percentuale As New List(Of Integer)
				For i As Integer = 1 To 10
					Percentuale.Add(0)
				Next
				Percentuale.Add(0)
				Dim TotPerc As Integer = 0

				Sql = "Select * From EventiNomi Where Attiva='S' Order By idCoppa"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof
						Percentuale.Item(Rec("idCoppa").Value) = Rec("Percentuale").Value
						TotPerc += Rec("Percentuale").Value

						Rec.MoveNext
					Loop
					Rec.Close
				End If

				Sql = "Select * From EventiNomi Where Descrizione='23 Aiutame Te' Order By idCoppa"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Percentuale.Item(2) = Rec("Percentuale").Value
					TotPerc += Rec("Percentuale").Value
					Rec.Close
				End If
				Percentuale.Item(0) = 100 - TotPerc

				Ritorno = ""
				Dim c As Integer = 0
				For Each p As Integer In Percentuale
					Dim Vincita As Single = (CInt(Totale * p) / 100)
					Ritorno &= c & ";" & p & ";" & Totale & ";" & Vincita & "§"
					c += 1
				Next
				Ritorno &= "|"

				' CAMPIONATO
				Dim ev As New clsEventi
				Dim Classifica As List(Of clsEventi.StrutturaGiocatore) = ev.PrendeGiocatori(Server.MapPath("."), idAnno, 38, Conn, Connessione)
				Ritorno &= "Campione di TotoMIO;" & Classifica.Item(0).idUtente & ";" & Classifica.Item(0).NickName & ";0§"
				Ritorno &= "Secondo;" & Classifica.Item(1).idUtente & ";" & Classifica.Item(1).NickName & ";999§"
				Ritorno &= "Terzo;" & Classifica.Item(2).idUtente & ";" & Classifica.Item(2).NickName & ";999§"
				Ritorno &= "Cucchiarella de legno;" & Classifica.Item(Classifica.Count - 1).idUtente & ";" & Classifica.Item(Classifica.Count - 1).NickName & ";999§"

				' COPPE
				Dim Conta As Integer = 0

				For Each idCoppa As Integer In idCoppe
					If Finale.Item(Conta) Then
						Sql = "SELECT A.idPartita, D.NickName As Giocatore1, E.NickName As Giocatore2, A.Risultato1, A.Risultato2, Coalesce(A.idVincente, -99) As idVincente " &
								"FROM EventiPartite A " &
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
								Ritorno &= Descrizione.Item(Conta) & ";-1;Torneo non ancora creato;" & idCoppa & "§"
							Else
								If Rec("idVincente").Value = -99 Or Rec("idVincente").Value = -1 Then
									Ritorno &= Descrizione.Item(Conta) & ";-1;Finale non giocata;" & idCoppa & "§"
								Else
									If Rec("idVincente").Value = 1 Then
										Ritorno &= Descrizione.Item(Conta) & ";" & Rec("idVincente").Value & ";" & Rec("Giocatore1").Value & ";" & idCoppa & "§"
									Else
										Ritorno &= Descrizione.Item(Conta) & ";" & Rec("idVincente").Value & ";" & Rec("Giocatore2").Value & ";" & idCoppa & "§"
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

							Ritorno &= Descrizione.Item(Conta) & ";" & r(0) & ";" & r(1) & ";" & idCoppa & "§"
						Else
							Ritorno &= Descrizione.Item(Conta) & ";-1;Torneo non ancora creato;" & idCoppa & "§"
						End If
					End If

					Conta += 1
				Next

				' 23 Aiutame Te
				Sql = "Select * From (" &
					"SELECT A.idUtente, B.NickName, Sum(A.Punti) As Punti FROM SquadreRandom As A " &
					"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
					"Where A.idAnno = " & idAnno & " " &
					"Group By A.idUtente, B.NickName " &
					"Union ALL " &
					"Select idUtente, NickName, 0 As Punti From Utenti " &
					"Where idUtente Not In (Select idUtente From SquadreRandom Where idAnno = " & idAnno & ") " &
					"And idAnno = " & idAnno & " " &
					") As A " &
					"Group By idUtente, NickName " &
					"Order By 3 Desc, 2 "
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						' Ritorno = "ERROR: Nessuna coppa rilevata"
					Else
						Ritorno &= "23 Aiutame Te;" & Rec("idUtente").Value & ";" & Rec("NickName").Value & ";2§"
						Rec.Close
					End If
				End If

				Sql = "Select A.idUtente, NickName, Sum(Punti) As Totale From ( " &
					"Select A.idAnno, A.idUtente, ((Sum(A.Punti) / Count(*)) + (Sum(SegniPresi) * 7) + (Sum(RisultatiEsatti) * 5) + (Sum(RisultatiCasaTot) * 3) + (Sum(RisultatiFuoriTot) * 3) + " &
					"(Sum(SommeGoal) * 3) + (Sum(DifferenzeGoal) * 3) + (Sum(Jolly) * 11) + (Sum(C.Punti) * 11) + (Sum(A.PuntiPartitaScelta) * 11)) / Count(*) As Punti " &
					"From Risultati A " &
					"Left Join RisultatiAltro B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
					"Left Join SquadreRandom C On A.idAnno = C.idAnno And A.idUtente = C.idUtente " &
					"Where A.idAnno = " & idAnno & " " &
					"Group By A.idUtente " &
					"Union All " &
					"Select A.idAnno, A.idUtente, (Count(*) * 3) * 15 As Punti From Utenti A " &
					"Left Join EventiPartite B On A.idAnno = B.idAnno And A.idUtente = B.idGiocatore1 And B.idVincente = 1 " &
					"Where A.idAnno = " & idAnno & " " &
					"Group By A.idUtente " &
					"Union ALL " &
					"Select A.idAnno, A.idUtente, (Count(*)) * 7 As Punti From Utenti A " &
					"Left Join EventiPartite B On A.idAnno = B.idAnno And A.idUtente = B.idGiocatore1 And B.idVincente = 0 " &
					"Where A.idAnno = " & idAnno & " " &
					"Group By A.idUtente " &
					"Union All " &
					"Select A.idAnno, A.idUtente, (Count(*) * 3) * 15 As Punti From Utenti A " &
					"Left Join EventiPartite B On A.idAnno = B.idAnno And A.idUtente = B.idGiocatore2 And B.idVincente = 2 " &
					"Where A.idAnno = " & idAnno & " " &
					"Group By A.idUtente " &
					"Union ALL " &
					"Select A.idAnno, A.idUtente, (Count(*)) * 7 As Punti From Utenti A " &
					"Left Join EventiPartite B On A.idAnno = B.idAnno And A.idUtente = B.idGiocatore2 And B.idVincente = 0 " &
					"Where A.idAnno = " & idAnno & " " &
					"Group By A.idUtente " &
					") As A " &
					"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
					"Group By A.idAnno, A.idUtente " &
					"Order By Punti Desc, NickName"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						' Ritorno = "ERROR: Nessuna coppa rilevata"
					Else
						Ritorno &= "Campione dei campioni (che non vince niente) con punti media " & Rec("Totale").Value & ";" & Rec("idUtente").Value & ";" & Rec("NickName").Value & "§"
						Rec.MoveLast
						Ritorno &= "Pippone dei pipponi (che non perde niente) con punti media " & Rec("Totale").Value & ";" & Rec("idUtente").Value & ";" & Rec("NickName").Value & "§"
						Rec.Close

						Ritorno &= "|"
						Sql = "SELECT A.idUtente, B.NickName, Sum(Importo) As Totale FROM Bilancio As A " &
							"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
							"Where A.idAnno = " & idAnno & " And (A.idMovimento = 3 Or A.idMovimento = 4) " &
							"Group By A.idUtente, B.NickName"
						Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								' Ritorno = "ERROR: Nessuna coppa rilevata"
							Else
								Do Until Rec.Eof
									Ritorno &= Rec("idUtente").Value & ";" & Rec("NickName").Value & ";" & Rec("Totale").Value & "§"

									Rec.MoveNext
								Loop
								Rec.Close
							End If
						End If
					End If
				End If

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
		Dim idGiornata As Integer

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
								idGiornata = Rec("idGiornata").Value
								Rec.Close

								Dim CreazioneCoppa As String = ""
								Dim GiocatePartite As String = ""

								sql = "Select A.*, B.Descrizione As Tipologia, C.Descrizione As Torneo, C.QuantiGiocatori, C.Importanza, " &
									"A.InizioGiornata, C.Descrizione As Torneo, B.Dettaglio, D.idGiornataVirtuale As Giornata " &
									"From Eventi A " &
									"Left Join EventiTipologie B On A.idTipologia = B.idTipologia " &
									"Left Join EventiNomi C On A.idCoppa = C.idCoppa " &
									"Left Join EventiPartite D On A.idEvento = D.idEvento And D.idGiornata = A.InizioGiornata And D.idAnno = " & idAnno & " " &
									"Where A.InizioGiornata=" & idGiornata & " Order By A.idEvento" ' idAnno=" & idAnno & " And idGiornata=" & idGiornata & " And idEvento<>1"
								Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Not Rec.Eof Then
										Dim ev As New clsEventi

										Do Until Rec.Eof
											If Rec("idEvento").Value <> 1 Then
												SceltiPerCreazione = ""
												GiocataPartita = ""

												Ritorno = ev.GestioneEventi(Server.MapPath("."), idAnno, idGiornata, Rec("idEvento").Value,
																		Rec("QuantiGiocatori").Value, Rec("Importanza").Value,
																		Rec("InizioGiornata").Value, Rec("Tipologia").Value, Rec("Torneo").Value,
																		Rec("Dettaglio").Value, Rec("idCoppa").Value, Conn, Connessione)
												If Ritorno.Contains("ERROR") Then
													Exit Do
												Else
													If SceltiPerCreazione <> "" Then
														' Creazione coppa
														CreazioneCoppa &= "<hr /><span style=""font-weight: bold;"">Creazione Coppa '" & Rec("Torneo").Value & "'. Partecipanti:</span><br /><br />"
														Dim s() As String = SceltiPerCreazione.Split(";")
														For Each ss As String In s
															If ss <> "" Then
																CreazioneCoppa &= ss & "<br />"
															End If
														Next
														CreazioneCoppa &= "<br />"
													End If
													If GiocataPartita <> "" Then
														' Giocata partita
														Dim Tabella As String = "<hr /><span style=""font-weight: bold;"">Giocata partita di Coppa '" & Rec("Torneo").Value & "': " & Rec("Dettaglio").Value & " Giornata " & Rec("Giornata").Value & ".</span><br />"
														Tabella &= "<table style=""border: 1px solid #999;"">"
														Tabella &= "<tr>"
														Tabella &= "<th>Casa</th>"
														Tabella &= "<th>Fuori</th>"
														Tabella &= "<th>Risultato 1</th>"
														Tabella &= "<th>Risultato 2</th>"
														Tabella &= "<th>Vincente</th>"
														Tabella &= "</tr>"
														Dim righe() As String = GiocataPartita.Split("§")
														For Each r As String In righe
															If r <> "" Then
																Dim campi() As String = r.Split(";")
																Dim Vincente As String = campi(4)
																If Vincente = "1" Then
																	Vincente = campi(0)
																Else
																	If Vincente = "2" Then
																		Vincente = campi(1)
																	Else
																		Vincente = "Pareggio"
																	End If
																End If
																Tabella &= "<tr>"
																Tabella &= "<td>" & campi(0) & "</td>"
																Tabella &= "<td>" & campi(1) & "</td>"
																Tabella &= "<td>" & campi(2) & "</td>"
																Tabella &= "<td>" & campi(3) & "</td>"
																Tabella &= "<td>" & Vincente & "</td>"
																Tabella &= "</tr>"
															End If
														Next
														Tabella &= "</table><br />"
														GiocatePartite &= Tabella
													End If
												End If
											End If

											Rec.MoveNext
										Loop
										Rec.CLose
									End If
								End If

								Dim TestoRis As String = ""

								sql = "Select * From Concorsi A " &
									"Where A.idAnno=" & idAnno & " And A.idConcorso=" & idGiornata & " Order By A.idPartita"
								Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
									'Ritorno = Rec
								Else
									Dim Partite As New List(Of String)
									TestoRis = "<hr /><span style=""font-weight: bold;"">Risultati Concorso</span><br /><table style=""border: 1px solid #999;"">"
									TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
									TestoRis &= "<th>Casa</th>"
									TestoRis &= "<th>Fuori</th>"
									TestoRis &= "<th>Risultato</th>"
									TestoRis &= "<th>Segno</th>"
									TestoRis &= "</tr>"
									Do Until Rec.Eof
										TestoRis &= "<tr>"
										TestoRis &= "<td>" & Rec("Prima").Value & "</td>"
										TestoRis &= "<td>" & Rec("Seconda").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("Risultato").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("Segno").Value & "</td>"
										TestoRis &= "</tr>"
										Partite.Add(Rec("Prima").Value & ";" & Rec("Seconda").Value & ";" & Rec("Segno").Value)
										Rec.MoveNext
									Loop
									Rec.Close
									TestoRis &= "</table><br />"

									Dim Sorprese As List(Of StrutturaSorprese) = PrendeSorprese(Server.MapPath("."), Conn, Connessione, idAnno, idGiornata)
									TestoRis &= "<hr /><span style=""font-weight: bold;"">Risultati a sorpresa del concorso</span><br /><table style=""border: 1px solid #999;"">"
									TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
									TestoRis &= "<th>Casa</th>"
									TestoRis &= "<th>Fuori</th>"
									TestoRis &= "<th>Segno Pronosticato</th>"
									TestoRis &= "<th>Percentuale</th>"
									TestoRis &= "<th>Segno Finale</<th>"
									TestoRis &= "<th>Realizzata</<th>"
									TestoRis &= "</tr>"
									For Each s As StrutturaSorprese In Sorprese
										Dim Risultato As String = ""
										For Each p As String In Partite
											Dim pp() As String = p.Split(";")
											If s.Casa.Trim.ToUpper = pp(0).Trim.ToUpper And s.Fuori.Trim.ToUpper = pp(1).Trim.ToUpper Then
												Risultato = pp(2)
												Exit For
											End If
										Next
										TestoRis &= "<tr>"
										TestoRis &= "<td>" & s.Casa & "</td>"
										TestoRis &= "<td>" & s.Fuori & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & s.Segno & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & s.Percentuale & "%</td>"
										TestoRis &= "<td style=""text-align: center"">" & Risultato & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & IIf(s.Segno = Risultato, "*", "-") & "</td>"
										TestoRis &= "</tr>"
									Next
									TestoRis &= "</table><br />"
								End If

								sql = "Select A.*, B.NickName, C.Jolly From Risultati A " &
									"Left Join Utenti B On A.idUtente = B.idUtente " &
									"Left Join RisultatiAltro C On A.idAnno = C.idAnno And A.idUtente = C.idUtente And A.idConcorso = C.idConcorso " &
									"Where A.idAnno=" & idAnno & " And A.idConcorso=" & idGiornata & " Order By A.Punti Desc"
								Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
									'Ritorno = Rec
								Else
									TestoRis &= "<hr /><span style=""font-weight: bold;"">Risultati Campionato</span><br /><table style=""width: 100%; border: 1px solid #999;"">"
									TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
									TestoRis &= "<th>Utente</th>"
									TestoRis &= "<th>Punti</th>"
									TestoRis &= "<th>Segni Presi</th>"
									TestoRis &= "<th>Ris. Esatti</th>"
									TestoRis &= "<th>Ris. Casa</th>"
									TestoRis &= "<th>Ris.Fuori</th>"
									TestoRis &= "<th>Somma Goal</th>"
									TestoRis &= "<th>Diff. Goal</th>"
									TestoRis &= "<th>Jolly</th>"
									TestoRis &= "<th>P.P.Sc.</th>"
									TestoRis &= "<th>Sorp.</th>"
									TestoRis &= "</tr>"
									Do Until Rec.Eof
										TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
										TestoRis &= "<td>" & Rec("NickName").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("Punti").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("SegniPresi").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("RisultatiEsatti").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("RisultatiCasaTot").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("RisultatiFuoriTot").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("SommeGoal").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("DifferenzeGoal").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("Jolly").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("PuntiPartitaScelta").Value & "</td>"
										TestoRis &= "<td style=""text-align: center"">" & Rec("PuntiSorpresa").Value & "</td>"
										TestoRis &= "<tr>"

										Rec.MoveNext
									Loop
									TestoRis &= " </table>"
									Rec.Close
								End If

								TestoRis &= "<br />"
								TestoRis &= CreazioneCoppa
								TestoRis &= GiocatePartite

								'sql = "Select C.idCoppa, C.Descrizione As NomeCoppa, D.Descrizione As Tipologia FROM EventiPartite A " &
								'	"Left Join Eventi B On A.idEvento = B.idEvento " &
								'	"Left Join EventiNomi C On B.idCoppa = C.idCoppa " &
								'	"Left Join EventiTipologie D On B.idTipologia = D.idTipologia " &
								'	"Where idAnno = " & idAnno & " And idGiornataVirtuale = " & idGiornata & " And D.Descrizione <> 'Creazione' " &
								'	"Group By C.idCoppa, C.Descrizione, D.Descrizione"
								'Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								'If TypeOf (Rec) Is String Then
								'	'Ritorno = Rec
								'Else
								'	Dim idCoppe As New List(Of Integer)
								'	Dim Torneo As New List(Of String)

								'	Do Until Rec.Eof
								'		idCoppe.Add(Rec("idCoppa").Value)
								'		Torneo.Add(Rec("NomeCoppa").Value)

								'		Rec.MoveNext
								'	Loop
								'	Rec.Close

								'	Dim n As Integer = 0
								'	For Each id As Integer In idCoppe
								'		TestoRis &= "Torneo: " & Torneo.Item(n) & " <br />"
								'		TestoRis &= "<table style=""width: 100%; border: 1px solid #999;"">"
								'		TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
								'		TestoRis &= "<th>Casa</th>"
								'		TestoRis &= "<th>Fuori</th>"
								'		TestoRis &= "<th>Risultato 1</th>"
								'		TestoRis &= "<th>Risultato 2</th>"
								'		TestoRis &= "<th>Vincente</th>"
								'		TestoRis &= "</tr>"
								'		sql = "Select A.*, B.NickName As Casa, C.NickName As Fuori FROM EventiPartite As A " &
								'			"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore1 = B.idUtente " &
								'			"Left Join Utenti As C On A.idAnno = C.idAnno And A.idGiocatore2 = C.idUtente " &
								'			"Where A.idAnno = " & idAnno & " And A.idGiornataVirtuale = " & idGiornata & " And " &
								'			"A.idEvento In (Select idEvento From Eventi As E Where idCoppa = " & id & " And A.idGiornata = E.InizioGiornata) " &
								'			"Order By idPartita"
								'		Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								'		If TypeOf (Rec) Is String Then
								'			'Ritorno = Rec
								'		Else
								'			Do Until Rec.Eof
								'				TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
								'				TestoRis &= "<td>" & Rec("Casa").Value & "</td>"
								'				TestoRis &= "<td>" & Rec("Fuori").Value & "</td>"
								'				TestoRis &= "<td>" & Rec("Risultato1").Value & "</td>"
								'				TestoRis &= "<td>" & Rec("Risultato2").Value & "</td>"
								'				If Rec("idVincente").Value = "1" Then
								'					TestoRis &= "<td>" & Rec("Casa").Value & "</td>"
								'				Else
								'					If Rec("idVincente").Value = "2" Then
								'						TestoRis &= "<td>" & Rec("Casa").Value & "</td>"
								'					Else
								'						If Rec("idVincente").Value = "-1" Then
								'						Else
								'							TestoRis &= "<td>Pareggio</td>"
								'						End If
								'					End If
								'				End If
								'				TestoRis &= "</tr>"

								'				Rec.MoveNext
								'			Loop
								'			Rec.Close
								'		End If
								'		TestoRis &= "</table><br /><br />"
								'		n += 1
								'	Next
								'End If

								sql = "Select * From SquadreRandom A " &
									"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
									"Where A.idAnno=" & idAnno & " And A.idConcorso=" & idGiornata & " Order By Punti Desc"
								Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
									'Ritorno = Rec
								Else
									TestoRis &= "<hr />"
									TestoRis &= "<span style=""font-weight: bold;"">23 Aiutame Te</span><br />"
									TestoRis &= "<table style=""border: 1px solid #999;"">"
									TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
									TestoRis &= "<th>Utente</th>"
									TestoRis &= "<th>Squadra</th>"
									TestoRis &= "<th>Punti</th>"
									TestoRis &= "</tr>"
									Do Until Rec.Eof
										If Rec("NickName").Value <> "" Then
											TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
											TestoRis &= "<td>" & Rec("NickName").Value & "</td>"
											TestoRis &= "<td>" & Rec("Squadra").Value & "</td>"
											TestoRis &= "<td style=""text-align: center;"">" & Rec("Punti").Value & "</td>"
											TestoRis &= "</tr>"
										End If

										Rec.MoveNext
									Loop
									Rec.Close
									TestoRis &= "</table><br />"
								End If

								Dim Testo As String = ""
								Testo = "E' stato controllato il concorso TotoMIO numero " & idGiornata & ".<br />"
								Testo &= "<br />" & TestoRis & "<br />"
								Testo &= "Per entrare nel sito e vedere il resto: <a href=""" & IndirizzoSito & """>Click QUI</a>"
								InvaMailATutti(Server.MapPath("."), idAnno, "TotoMIO: Controllo concorso " & idGiornata, Testo, Conn, Connessione, "Controllo")
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
	Public Function RitornaNomiCoppe() As String
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
	Public Function RitornaInadempienti(idAnno As String, idGiornata As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Rec As Object

		Dim sql As String = "SELECT Distinct idUtente, NickName FROM Utenti A " &
								"Where idAnno = " & idAnno & " And idUtente Not In (Select idUtente From Pronostici Where idAnno = " & idAnno & " And idConcorso = " & idGiornata & ") " &
								"And idTipologia<>2 " &
								"Order By NickName"
		Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			'Ritorno = Rec
		Else
			Do Until Rec.Eof
				Ritorno &= Rec("NickName").Value & "§"

				Rec.MoveNext
			Loop
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaColonnaUtente(idAnno As String, idUtente As String, idGiornata As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select A.idPartita, B.Prima, B.Seconda, A.Pronostico, A.Segno From Pronostici As A " &
			"Left Join Concorsi B On A.idAnno = B.idAnno And A.idConcorso = B.idConcorso And A.idPartita = B.idPartita " &
			"Where A.idAnno = " & idAnno & " And idUtente = " & idUtente & " And A.idConcorso = " & idGiornata & " " &
			"Order By A.idPartita"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessuna colonna rilevata"
			Else
				Do Until Rec.Eof
					Ritorno &= Rec("idPartita").Value & ";" & SistemaStringaPerRitorno(Rec("Prima").Value) & ";" &
						SistemaStringaPerRitorno(Rec("Seconda").Value) & ";" & Rec("Pronostico").Value & ";" &
						Rec("Segno").Value & "§"

					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ImpostaConcorsoPerControllo(idAnno As String) As String
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

				sql = "Select * From Globale Where idAnno=" & idAnno
				Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = "ERROR: problemi nella lettura della giornata"
					Else
						Dim idGiornata As String = Rec("idGiornata").Value
						Rec.Close

						sql = "Update Globale Set idModalitaConcorso=" & idModalita & ", Scadenza='' Where idAnno=" & idAnno
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Not Ritorno.Contains("ERROR") Then
							Ritorno = idModalita & ";" & Descrizione

							Dim Assenti As String = ""

							sql = "SELECT Distinct idUtente, NickName FROM Utenti A " &
								"Where idAnno = " & idAnno & " And idUtente Not In (Select idUtente From Pronostici Where idAnno = " & idAnno & " And idConcorso = " & idGiornata & ") " &
								"And A.idTipologia<>2"
							Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
							If TypeOf (Rec) Is String Then
								'Ritorno = Rec
							Else
								Do Until Rec.Eof
									Assenti &= Rec("NickName").Value & " <br />"

									Rec.MoveNext
								Loop
								Rec.Close
							End If

							Dim Random As String = ""

							sql = "Select * From SquadreRandom A " &
								"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
								"Where A.idAnno=" & idAnno & " And A.idConcorso=" & idGiornata & " And B.idTipologia<>2"
							Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
							If TypeOf (Rec) Is String Then
								'Ritorno = Rec
							Else
								Do Until Rec.eof
									If Rec("NickName").Value <> "" Then
										Random &= Rec("NickName").Value & ": " & Rec("Squadra").Value & "<br />"
									End If

									Rec.MoveNext
								Loop
								Rec.Close
							End If

							' Crea colonna utente finto se esistente
							Dim Rit As String = CreaColonnaUtenteFinto(Server.MapPath("."), idAnno, idGiornata, Conn, Connessione)
							If Rit.Contains(StringaErrore) Then
								Ritorno = Rit
							Else
								Dim Testo As String = ""
								Testo = "E' stato chiuso il concorso TotoMIO numero " & idGiornata & ".<br />"
								Testo &= "Non sarà più possibile giocare la schedina<br /><br />"
								If Assenti <> "" Then
									Testo &= "Non adempienti:<br />" & Assenti & "<br /><br />"
								End If
								If Random <> "" Then
									Testo &= "Squadre assegnate per 23 Aiutame Te:<br />" & Random & "<br /><br />"
								End If
								Testo &= "Per entrare nel sito: <a href=""" & IndirizzoSito & """>Click QUI</a>"
								InvaMailATutti(Server.MapPath("."), idAnno, "TotoMIO: Chiusura concorso " & idGiornata, Testo, Conn, Connessione, "Chiusura")
							End If
						End If
					End If
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
							"'" & SistemaStringaPerDB(D2(4)) & "', " &
							"'" & SistemaStringaPerDB(D2(5)) & "' " &
							")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Ritorno.Contains(StringaErrore) Then
							Exit For
						End If
					End If
				Next

				If Not Ritorno.Contains(StringaErrore) Then
					Ritorno = CreaPartitaJolly(Server.MapPath("."), idAnno, idConcorso, Conn, Connessione)
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
						Rec("Segno").Value & ";" & Rec("Sospesa").Value & "§"
					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ControllaConcorso(idAnno As String, idUtente As String, ModalitaConcorso As String, SoloControllo As String, InviaMailSoloControllo As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From Globale Where idAnno=" & idAnno
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		Dim idGiornata As String = ""
		Dim SquadreCasa As New List(Of String)
		Dim SquadreFuori As New List(Of String)
		Dim Risultati As New List(Of String)
		Dim Segni As New List(Of String)

		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun anno rilevato"
			Else
				idGiornata = Rec("idGiornata").Value
				Rec.Close

				If SoloControllo <> "SI" Then
					sql = "Update RisultatiAltro Set Vittorie = 0, Ultimo = 0 Where idAnno=" & idAnno & " And idConcorso=" & idGiornata
					Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
				End If

				Ritorno = ""

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
							If Rec("Sospesa").Value = "S" Then
								Partite.Add(Rec("idPartita").Value & ";" & SistemaStringaPerRitorno(Rec("Prima").Value) & ";" &
									SistemaStringaPerRitorno(Rec("Seconda").Value) & ";" &
									";" & Rec("Segno").Value & ";" & Rec("Sospesa").Value)
							Else
								If Rec("Risultato").Value = "" Or Rec("Segno").Value = "" Then
									Ritorno = "ERROR: Risultato della partita " & Rec("idPartita").Value & " vuoto"
									Exit Do
								End If

								SquadreCasa.Add(Rec("Prima").Value)
								SquadreFuori.Add(Rec("Seconda").Value)
								Risultati.Add(Rec("Risultato").Value)
								Segni.Add(Rec("Segno").Value)

								Partite.Add(Rec("idPartita").Value & ";" & SistemaStringaPerRitorno(Rec("Prima").Value) & ";" &
									SistemaStringaPerRitorno(Rec("Seconda").Value) & ";" &
									Rec("Risultato").Value & ";" & Rec("Segno").Value & ";" & Rec("Sospesa").Value)
							End If

							Rec.MoveNext
						Loop
						Rec.Close

						Dim PartitaJolly As Integer = -1

						If Not Ritorno.Contains(StringaErrore) Then
							sql = "Select Coalesce(idPartita, -1) As idPartita From PartiteJolly Where idAnno=" & idAnno & " And idConcorso=" & idGiornata
							Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Not Rec.Eof Then
									PartitaJolly = Rec("idPartita").Value
								Else
									Ritorno = "ERROR: Nessuna partita jolly rilevata"
								End If
								Rec.Close
							End If
						End If

						If Not Ritorno.Contains("ERROR") Then
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

									Dim Sorprese As List(Of StrutturaSorprese) = PrendeSorprese(Server.MapPath("."), Conn, Connessione, idAnno, idGiornata)

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
													Dim idPartitaScelta As Integer = -1

													sql = "Select Coalesce(idPartita, -1) As idPartita From PartiteScelte Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente=" & id
													Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
													If TypeOf (Rec) Is String Then
														Ritorno = Rec
													Else
														If Not Rec.Eof Then
															idPartitaScelta = Rec("idPartita").Value
														Else
															idPartitaScelta = GetRandom(1, 10)

															sql = "Insert Into PartiteScelte Values (" & idAnno & ", " & idGiornata & ", " & id & ", " & idPartitaScelta & ")"
															Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
														End If
														Rec.Close
													End If

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
																							 ModalitaConcorso, PartitaJolly, idPartitaScelta, Sorprese,
																				 SoloControllo)
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
																							 ModalitaConcorso, PartitaJolly, idPartitaScelta, Sorprese,
																				 SoloControllo)
															Ritorno &= Controllo & "%"
														End If
													End If
												Next
											End If
										End If
									Else
										Dim idPartitaScelta As Integer = -1

										sql = "Select Coalesce(idPartita, -1) As idPartita From PartiteScelte Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente=" & idUtente
										Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
										If TypeOf (Rec) Is String Then
											Ritorno = Rec
										Else
											If Not Rec.Eof Then
												idPartitaScelta = Rec("idPartita").Value
											Else
												idPartitaScelta = GetRandom(1, 10)

												sql = "Insert Into PartiteScelte Values (" & idAnno & ", " & idGiornata & ", " & idUtente & ", " & idPartitaScelta & ")"
												Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
											End If
											Rec.Close
										End If

										' Controllo per utente
										sql = "Select * From Pronostici Where idAnno=" & idAnno & " And idUtente=" & idUtente & " And idConcorso=" & idGiornata & " Order By idPartita"
										Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
										If TypeOf (Rec) Is String Then
											Ritorno = Rec
										Else
											If Rec.Eof Then
												Dim Controllo As String = ControllaPunti(idAnno, idUtente, idGiornata, NickName,
																				Partite, New List(Of String), Conn, Connessione, Server.MapPath("."),
																				ModalitaConcorso, PartitaJolly, idPartitaScelta, Sorprese,
																				 SoloControllo)
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
																				 ModalitaConcorso, PartitaJolly, idPartitaScelta, Sorprese,
																				 SoloControllo)
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
		End If

		If Not Ritorno.Contains("ERROR") Then
			' Controlla squadre random
			sql = "Select * From SquadreRandom Where idAnno=" & idAnno & " And idConcorso=" & idGiornata
			Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Rec.Eof Then
				Else
					Do Until Rec.Eof
						Dim idUtente23 As Integer = Rec("idUtente").Value
						sql = "Select * From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And (Prima='" & SistemaStringaPerDB(Rec("Squadra").Value) & "' Or Seconda='" & SistemaStringaPerDB(Rec("Squadra").Value) & "')"
						Dim Rec2 As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec2) Is String Then
							Ritorno = Rec2
						Else
							Dim Punti As Integer = 0
							Dim Casa As Boolean = True
							If Rec("Squadra").Value = Rec2("Seconda").Value Then
								Casa = False
							End If
							Dim Risultato As String = Rec2("Risultato").Value
							Dim Segno As String = Rec2("Segno").Value
							Dim idPartita As Integer = Rec2("idPartita").Value
							Rec2.Close

							sql = "Select * From Pronostici Where idAnno=" & idAnno & " And idUtente=" & idUtente23 & " And idConcorso=" & idGiornata & " And idPartita=" & idPartita
							Rec2 = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
							If TypeOf (Rec2) Is String Then
								Ritorno = Rec2
							Else
								If Not Rec2.Eof Then
									Dim Pronostico As String = Rec2("Pronostico").Value
									Dim SegnPron As String = Rec2("Segno").Value
									Rec2.Close

									If Risultato = Pronostico Then
										Punti += 7
									End If
									If Segno = SegnPron Then
										Punti += 5
									End If

									If Risultato <> "" And Risultato.Contains("-") Then
										Dim r() As String = Risultato.Split("-")
										Dim p() As String = Pronostico.Split("-")

										Dim r1 As Integer = r(0)
										Dim r2 As Integer = r(1)

										Dim p1 As Integer = p(0)
										Dim p2 As Integer = p(1)

										If r1 = p1 Or r2 = p2 Then
											Punti += 3
										End If
										If Math.Abs(r1 - r2) = Math.Abs(p1 - p2) Then
											Punti += 1
										End If
										If Math.Abs(r1 + r2) = Math.Abs(p1 + p2) Then
											Punti += 1
										End If

										If Not Casa Then
											Punti *= 1.75
											Punti = CInt(Punti)
										End If

										If SoloControllo <> "SI" Then
											sql = "Update SquadreRandom Set Punti=" & Punti & " Where " &
												"idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente=" & idUtente23
											Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
										End If
									End If

									' Ritorno &= idUtente23 & ";" & Punti & "§"
								End If
							End If
						End If

						Rec.MoveNext
					Loop
					Rec.CLose
				End If
			End If
		End If

		' 1;28|1;Pippa;Pippetta;1-2;2;1-1;X;3§%
		' IdUtente;PuntiTotali|idPartita;Squadra1;Squadra2;Risultato;Segno;Pronostico;PronosticoSegno;PuntiPartita;Jolly;PuntiPartitaScelta§%
		' idUtente23;Punti§

		If Not Ritorno.Contains("ERROR") Then
			' Aggiorna primi ultimi
			Dim Classifica As String = RitornaClassificaGenerale(Server.MapPath("."), idAnno, idGiornata, Conn, Connessione, True, "S", SoloControllo)

			If Classifica <> "" Then
				Dim c() As String = Classifica.Split("§")
				Dim PrimaRiga() As String = c(0).Split(";")
				Dim idUltimo As Integer = PrimaRiga(0)
				Dim UltimaRiga() As String = c(c.Count - 2).Split(";")
				Dim idPrimo As Integer = UltimaRiga(0)
				Dim Ritorno2 As String = "OK"
				Dim idUtenteFinto As Integer = -1

				sql = "Select idUtente From Utenti Where idAnno=" & idAnno & " And idTipologia=2"
				Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno2 = Rec
				Else
					idUtenteFinto = Rec("idUtente").Value
					Rec.Close
				End If
				Dim Premio As Integer = 0

				sql = "Select * From PremioPerFinto Where idAnno=" & idAnno
				Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno2 = Rec
				Else
					If Rec.Eof Then
						Premio = 0
						sql = "Insert Into PremioPerFinto Values (1, 0)"
						Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					Else
						Premio = Rec("Importo").Value
					End If
					Rec.Close
				End If

				If idUltimo <> idUtenteFinto Then
					sql = "Select Coalesce(Max(Progressivo)+1, 1) As Progressivo From Bilancio Where idAnno=" & idAnno
					Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno2 = Rec
					Else
						Dim Progressivo As String = Rec("Progressivo").Value
						Rec.Close

						If SoloControllo <> "SI" Then
							sql = "Delete From Bilancio Where idAnno=" & idAnno & " And Note = 'Vittoria TotoMIO Concorso N° " & idGiornata & "'"
							Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)

							Premio += 1
							Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year
							sql = "Insert Into Bilancio Values (" &
								" " & idAnno & ", " &
								" " & idUltimo & ", " &
								" " & Progressivo & ", " &
								"3, " &
								" " & Premio & ", " &
								"'" & Datella & "', " &
								"'Vittoria TotoMIO Concorso N° " & idGiornata & "', " &
								"'N', " &
								"1 " &
								")"
							Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)

							sql = "Update PremioPerFinto Set Importo = 0 Where idAnno = " & idAnno
							Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						End If
					End If
				Else
					If SoloControllo <> "SI" Then
						sql = "Update PremioPerFinto Set Importo = Importo + 1 Where idAnno = " & idAnno
						Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					End If
				End If

				'If idPrimo <> idUtenteFinto Then

				'End If

				If SoloControllo <> "SI" Then
					sql = "Select * From RisultatiAltro Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente=" & idPrimo
					Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							sql = "Insert Into RisultatiAltro Values (" & idAnno & ", " & idGiornata & ", " & idPrimo & ", 1, 0, 0)"
							Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						Else
							sql = "Update RisultatiAltro Set Ultimo = 1 Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente = " & idPrimo
							Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						End If
					End If
				End If

				If SoloControllo <> "SI" Then
					sql = "Select * From RisultatiAltro Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente=" & idUltimo
					Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							sql = "Insert Into RisultatiAltro Values (" & idAnno & ", " & idGiornata & ", " & idUltimo & ", 0, 1, 0)"
							Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						Else
							sql = "Update RisultatiAltro Set Vittorie = 1 Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente = " & idUltimo
							Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						End If
					End If
				End If
			End If
		End If

		If SoloControllo = "SI" And InviaMailSoloControllo = "SI" Then
			Dim Nomi As New List(Of String)
			Dim Punti As New List(Of Integer)
			Dim Splittone() As String = Ritorno.Split("%")
			For Each s As String In Splittone
				If s <> "" Then
					Dim Campi() As String = s.Split("|")
					If Campi.Length > 0 Then
						Dim Campi2() As String = Campi(0).Split(";")
						Dim NickName As String = Campi2(1)
						Dim Punti2 As Integer = Campi2(2)
						Nomi.Add(NickName)
						Punti.Add(Punti2)
					End If
				End If
			Next

			For i As Integer = 0 To Nomi.Count - 1
				For k As Integer = i + 1 To Nomi.Count - 1
					If Punti.Item(i) < Punti.Item(k) Then
						Dim Appo As String = Nomi.Item(i)
						Nomi.Item(i) = Nomi.Item(k)
						Nomi.Item(k) = Appo
						Dim Appo2 As Integer = Punti.Item(i)
						Punti.Item(i) = Punti.Item(k)
						Punti.Item(k) = Appo2
					End If
				Next
			Next

			Dim Testo As String = ""
			Testo = "Andamento concorso TotoMIO numero " & idGiornata & ".<br /><br />"

			Testo &= "Risultati<hr /><table>"
			Testo &= "<tr>"
			Testo &= "<th>Casa</th>"
			Testo &= "<th>Fuori</th>"
			Testo &= "<th>Risultato</th>"
			Testo &= "<th>Segno</th>"
			Testo &= "</tr>"
			Dim c2 As Integer = 0
			For Each s As String In SquadreCasa
				Testo &= "<tr>"
				Testo &= "<td>" & s & "</td>"
				Testo &= "<td>" & SquadreFuori.Item(c2) & "</td>"
				Testo &= "<td style=""text-align: center"">" & Risultati.Item(c2) & "</td>"
				Testo &= "<td style=""text-align: center"">" & Segni.Item(c2) & "</td>"
				Testo &= "</tr>"
				c2 += 1
			Next
			Testo &= "</table><br /><hr />Punteggi<hr />"

			Testo &= "<table>"
			Testo &= "<tr>"
			Testo &= "<th></th>"
			Testo &= "<th>Utente</th>"
			Testo &= "<th>Punti</th>"
			Testo &= "</tr>"
			Dim c3 As Integer = 0
			Dim Posiz As Integer = 1
			Dim VecchioPunteggio As Integer = Punti.Item(0)
			For Each s As String In Nomi
				Testo &= "<tr>"
				Testo &= "<td style=""text-align: center"">" & Posiz & "</td>"
				Testo &= "<td>" & s & "</td>"
				Testo &= "<td style=""text-align: center"">" & Punti.Item(c3) & "</td>"
				Testo &= "</tr>"
				If VecchioPunteggio <> Punti.Item(c3) Then
					Posiz += 1
					VecchioPunteggio = Punti.Item(c3)
				End If
				c3 += 1
			Next
			Testo &= "</table><br /><br />"

			Testo &= "Per entrare nel sito: <a href=""" & IndirizzoSito & """>Click QUI</a>"
			InvaMailATutti(Server.MapPath("."), idAnno, "TotoMIO: Andamento concorso " & idGiornata, Testo, Conn, Connessione, "Chiusura")
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaSquadre23(idAnno As String, idConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From SquadreRandom Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Do Until Rec.Eof
				Ritorno &= Rec("idUtente").Value & ";" & Rec("Punti").Value & ";" & SistemaStringaPerRitorno(Rec("Squadra").Value) & "§"

				Rec.MoveNext
			Loop
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaClassifica23(idAnno As String, idConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "SELECT A.idUtente, B.NickName, Sum(A.Punti) As Punti FROM SquadreRandom As A " &
			"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
			"Where A.idAnno = " & idAnno & " And A.idConcorso <= " & idConcorso & " And B.idTipologia <> 2 " &
			"Group By A.idUtente, B.NickName " &
			"Order By 2 Desc "
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Do Until Rec.Eof
				Ritorno &= Rec("idUtente").Value & ";" & Rec("NickName").Value & ";" & Rec("Punti").Value & "§"

				Rec.MoveNext
			Loop
			Rec.Close

			Ritorno &= "|"

			sql = "SELECT A.idUtente, B.NickName, A.Squadra, A.Punti As Punti FROM SquadreRandom As A " &
				"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
				"Where A.idAnno = " & idAnno & " And A.idConcorso = " & idConcorso & " And B.idTipologia <> 2 " &
				"Group By A.idUtente, B.NickName"
			Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				Do Until Rec.Eof
					Ritorno &= Rec("idUtente").Value & ";" & Rec("NickName").Value & ";" & SistemaStringaPerRitorno(Rec("Squadra").Value) & ";" & Rec("Punti").Value & "§"

					Rec.MoveNext
				Loop
				Rec.Close

			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function StatistichePartite(idAnno As String, idConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Select idPartita, Prima, Seconda, Segno, Quanti, TotalePartite, Round((Quanti / TotalePartite) * 100) As Media From ( " &
			"Select *, (Select Count(*) From Pronostici Where idAnno= " & idAnno & " And idConcorso = " & idConcorso & " And idPartita=1) As TotalePartite From ( " &
			"Select A.idPartita, B.Prima, B.Seconda, A.Segno, Count(*) As Quanti From Pronostici As A " &
			"Left Join Concorsi B On A.idAnno = B.idAnno And A.idConcorso = B.idConcorso And A.idPartita = B.idPartita " &
			"Where A.idAnno = " & idAnno & " And A.idConcorso=" & idConcorso & " " &
			"Group By A.idPartita, A.Segno " &
			") As A " &
			") As B"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Do Until Rec.Eof
				Dim idPartita As String = Rec("idPartita").Value
				Dim Casa As String = Rec("Prima").Value
				Dim Fuori As String = Rec("Seconda").Value
				Dim Segno As String = Rec("Segno").Value
				Dim Quanti As String = Rec("Quanti").Value
				Dim Percentuale As String = Rec("Media").Value
				Sql = "Select Pronostico From Pronostici Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idPartita=" & idPartita
				Dim Rec2 As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec2) Is String Then
					Ritorno = Rec2
				Else
					Dim RisultatoPiuGiocato As New List(Of StruttPronostico)
					Dim GoalCasaPiuGiocato As New List(Of StruttPronostico)
					Dim GoalFuoriPiuGiocato As New List(Of StruttPronostico)

					Do Until Rec2.Eof
						Dim Pronostico As String = Rec2("Pronostico").Value
						Dim Ok As Boolean = False
						Dim C As Integer = 0
						For Each r As StruttPronostico In RisultatoPiuGiocato
							If RisultatoPiuGiocato.Item(C).Pronostico = Pronostico Then
								Dim g As StruttPronostico = RisultatoPiuGiocato.Item(C)
								g.Quante += 1
								RisultatoPiuGiocato.Item(C) = g
								Ok = True
								Exit For
							End If
							C += 1
						Next
						If Not Ok Then
							Dim p2 As New StruttPronostico
							p2.Pronostico = Pronostico
							p2.Quante = 1

							RisultatoPiuGiocato.Add(p2)
						End If

						Dim p() As String = Pronostico.Split("-")
						Ok = False
						Dim RisCasa As Integer = p(0)
						C = 0
						For Each r As StruttPronostico In GoalCasaPiuGiocato
							If Val(GoalCasaPiuGiocato.Item(C).Pronostico) = RisCasa Then
								Dim g As StruttPronostico = GoalCasaPiuGiocato.Item(C)
								g.Quante += 1
								GoalCasaPiuGiocato.Item(C) = g
								Ok = True
								Exit For
							End If
							C += 1
						Next
						If Not Ok Then
							Dim p2 As New StruttPronostico
							p2.Pronostico = RisCasa
							p2.Quante = 1

							GoalCasaPiuGiocato.Add(p2)
						End If

						Dim RisFuori As Integer = p(1)
						C = 0
						Ok = False
						For Each r As StruttPronostico In GoalFuoriPiuGiocato
							If Val(GoalFuoriPiuGiocato.Item(C).Pronostico) = RisCasa Then
								Dim g As StruttPronostico = GoalFuoriPiuGiocato.Item(C)
								g.Quante += 1
								GoalFuoriPiuGiocato.Item(C) = g
								Ok = True
								Exit For
							End If
							C += 1
						Next
						If Not Ok Then
							Dim p2 As New StruttPronostico
							p2.Pronostico = RisFuori
							p2.Quante = 1

							GoalFuoriPiuGiocato.Add(p2)
						End If

						Rec2.MoveNext
					Loop
					Rec2.Close

					For i As Integer = 0 To RisultatoPiuGiocato.Count - 1
						For k As Integer = i + 1 To RisultatoPiuGiocato.Count - 1
							If RisultatoPiuGiocato.Item(i).Quante < RisultatoPiuGiocato.Item(k).Quante Then
								Dim App As StruttPronostico = RisultatoPiuGiocato.Item(i)
								RisultatoPiuGiocato.Item(i) = RisultatoPiuGiocato.Item(k)
								RisultatoPiuGiocato.Item(k) = App
							End If
						Next
					Next

					Dim RisPiuGiocato As String = RisultatoPiuGiocato.Item(0).Pronostico & ";" & RisultatoPiuGiocato.Item(0).Quante
					Dim u As Integer = RisultatoPiuGiocato.Count - 1
					Dim RisMenoGiocato As String = RisultatoPiuGiocato.Item(u).Pronostico & ";" & RisultatoPiuGiocato.Item(u).Quante

					For i As Integer = 0 To GoalCasaPiuGiocato.Count - 1
						For k As Integer = i + 1 To GoalCasaPiuGiocato.Count - 1
							If GoalCasaPiuGiocato.Item(i).Quante < GoalCasaPiuGiocato.Item(k).Quante Then
								Dim App As StruttPronostico = GoalCasaPiuGiocato.Item(i)
								GoalCasaPiuGiocato.Item(i) = GoalCasaPiuGiocato.Item(k)
								GoalCasaPiuGiocato.Item(k) = App
							End If
						Next
					Next

					Dim GoalPiuGiocatoCasa As String = GoalCasaPiuGiocato.Item(0).Pronostico & ";" & GoalCasaPiuGiocato.Item(0).Quante
					u = GoalCasaPiuGiocato.Count - 1
					Dim GoalCasaMenoGiocato As String = GoalCasaPiuGiocato.Item(u).Pronostico & ";" & GoalCasaPiuGiocato.Item(u).Quante

					For i As Integer = 0 To GoalFuoriPiuGiocato.Count - 1
						For k As Integer = i + 1 To GoalFuoriPiuGiocato.Count - 1
							If GoalFuoriPiuGiocato.Item(i).Quante < GoalFuoriPiuGiocato.Item(k).Quante Then
								Dim App As StruttPronostico = GoalFuoriPiuGiocato.Item(i)
								GoalFuoriPiuGiocato.Item(i) = GoalFuoriPiuGiocato.Item(k)
								GoalFuoriPiuGiocato.Item(k) = App
							End If
						Next
					Next

					Dim GoalPiuGiocatoFuori As String = GoalFuoriPiuGiocato.Item(0).Pronostico & ";" & GoalFuoriPiuGiocato.Item(0).Quante
					u = GoalFuoriPiuGiocato.Count - 1
					Dim GoalFuoriMenoGiocato As String = GoalFuoriPiuGiocato.Item(u).Pronostico & ";" & GoalFuoriPiuGiocato.Item(u).Quante

					Ritorno &= idPartita & ";" & Casa & ";" & Fuori & ";" & Segno & ";" & Quanti & ";" & Percentuale & ";" &
						RisPiuGiocato & ";" & RisMenoGiocato & ";" & GoalPiuGiocatoCasa & ";" & GoalCasaMenoGiocato & ";" &
						GoalPiuGiocatoFuori & ";" & GoalFuoriMenoGiocato & "§"
				End If

				Rec.MoveNext
			Loop
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SistemaPronostici(idAnno As String, idConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "SELECT * FROM Pronostici Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And (Segno = '' Or Segno Is Null)"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Dim Sqls As New List(Of String)

			Do Until Rec.Eof
				Dim Risultato As String = Rec("Pronostico").Value
				Dim c() As String = Risultato.Split("-")
				Dim Casa As Integer = Val(c(0))
				Dim Fuori As Integer = Val(c(1))
				Dim idPartita As String = Rec("idPartita").Value
				Dim idUtente As String = Rec("idUtente").Value
				Dim Segno As String = ""
				If Casa > Fuori Then
					Segno = "1"
				Else
					If Casa < Fuori Then
						Segno = "2"
					Else
						Segno = "X"
					End If
				End If
				Sql = "Update Pronostici Set Segno = '" & Segno & "' Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente & " And idPartita=" & idPartita
				Sqls.Add(Sql)

				Rec.MoveNext
			Loop
			Rec.Close

			For Each s As String In Sqls
				Ritorno = Conn.EsegueSql(Server.MapPath("."), s, Connessione, False)
			Next
		End If

		If Ritorno = "" Then
			Ritorno = "OK"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaStatistiche(idAnno As String, idGiornata As String, Casa As String, Fuori As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Dim Punti1 As Integer = 0
		Dim PuntiCasa1 As Integer = 0
		Dim PuntiFuori1 As Integer = 0
		Dim Vittorie1 As Integer = 0
		Dim Pareggi1 As Integer = 0
		Dim Sconfitte1 As Integer = 0
		Dim VittorieCasa1 As Integer = 0
		Dim PareggiCasa1 As Integer = 0
		Dim SconfitteCasa1 As Integer = 0
		Dim VittorieFuori1 As Integer = 0
		Dim PareggiFuori1 As Integer = 0
		Dim SconfitteFuori1 As Integer = 0
		Dim GoalFatti1 As Integer = 0
		Dim GoalSubiti1 As Integer = 0
		Dim GoalFattiCasa1 As Integer = 0
		Dim GoalFattiFuori1 As Integer = 0
		Dim GoalSubitiCasa1 As Integer = 0
		Dim GoalSubitiFuori1 As Integer = 0
		Dim Giocate1 As Integer = 0
		Dim GiocateCasa1 As Integer = 0
		Dim GiocateFuori1 As Integer = 0

		Dim Punti2 As Integer = 0
		Dim PuntiCasa2 As Integer = 0
		Dim PuntiFuori2 As Integer = 0
		Dim Vittorie2 As Integer = 0
		Dim Pareggi2 As Integer = 0
		Dim Sconfitte2 As Integer = 0
		Dim VittorieCasa2 As Integer = 0
		Dim PareggiCasa2 As Integer = 0
		Dim SconfitteCasa2 As Integer = 0
		Dim VittorieFuori2 As Integer = 0
		Dim PareggiFuori2 As Integer = 0
		Dim SconfitteFuori2 As Integer = 0
		Dim GoalFatti2 As Integer = 0
		Dim GoalSubiti2 As Integer = 0
		Dim GoalFattiCasa2 As Integer = 0
		Dim GoalFattiFuori2 As Integer = 0
		Dim GoalSubitiCasa2 As Integer = 0
		Dim GoalSubitiFuori2 As Integer = 0
		Dim Giocate2 As Integer = 0
		Dim GiocateCasa2 As Integer = 0
		Dim GiocateFuori2 As Integer = 0

		Dim Partite1 As New List(Of String)
		Dim Partite2 As New List(Of String)
		Dim Sql As String = "SELECT * FROM Concorsi Where idAnno = " & idAnno & " And idConcorso < " & idGiornata & " And (Prima = '" & Casa & "' Or Seconda = '" & Casa & "') And Sospesa = 'N' Order By idPartita Desc Limit 10"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				' Ritorno = "ERROR: Nessuna statistica rilevata per la squadra 1"
			Else
				Do Until Rec.Eof
					Partite1.Add(Rec("idConcorso").Value & ";" & Rec("idPartita").Value & ";" & Rec("Prima").Value & ";" & Rec("Seconda").Value & ";" & Rec("Risultato").Value & ";" & Rec("Segno").Value)
					Dim Prima As String = Rec("Prima").Value

					Dim Ris() As String = Rec("Risultato").Value.split("-")
					Dim GoalFatti As Integer = Ris(0)
					Dim GoalSubiti As Integer = Ris(1)

					Dim Segno As String = Rec("Segno").Value

					Giocate1 += 1

					If Casa = Prima Then
						GoalFatti1 += GoalFatti
						GoalSubiti1 += GoalSubiti
						GoalFattiCasa1 += GoalFatti
						GoalSubitiCasa1 += GoalSubiti
						GiocateCasa1 += 1

						Select Case Segno
							Case "1"
								Punti1 += 3
								PuntiCasa1 += 3
								Vittorie1 += 1
								VittorieCasa1 += 1
							Case "X"
								Punti1 += 1
								PuntiCasa1 += 1
								Pareggi1 += 1
								PareggiCasa1 += 1
							Case "2"
								Sconfitte1 += 1
								SconfitteCasa1 += 1
						End Select
					Else
						GiocateFuori1 += 1

						GoalFatti1 += GoalSubiti
						GoalSubiti1 += GoalFatti
						GoalFattiFuori1 += GoalSubiti
						GoalSubitiFuori1 += GoalFatti

						Select Case Segno
							Case "2"
								Punti1 += 3
								PuntiFuori1 += 3
								Vittorie1 += 1
								VittorieFuori1 += 1
							Case "X"
								Punti1 += 1
								PuntiFuori1 += 1
								Pareggi1 += 1
								PareggiFuori1 += 1
							Case "1"
								Sconfitte1 += 1
								SconfitteFuori1 += 1
						End Select
					End If

					Rec.MoveNext
				Loop
				Rec.Close

				Sql = "SELECT * FROM Concorsi Where idAnno = " & idAnno & " And idConcorso < " & idGiornata & " And (Prima = '" & Fuori & "' Or Seconda = '" & Fuori & "') And Sospesa = 'N' Order By idPartita Desc Limit 10"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						' Ritorno = "ERROR: Nessuna statistica rilevata per la squadra 1"
					Else
						Do Until Rec.Eof
							Partite2.Add(Rec("idConcorso").Value & ";" & Rec("idPartita").Value & ";" & Rec("Prima").Value & ";" & Rec("Seconda").Value & ";" & Rec("Risultato").Value & ";" & Rec("Segno").Value)

							Dim Prima As String = Rec("Prima").Value

							Dim Ris() As String = Rec("Risultato").Value.split("-")
							Dim GoalFatti As Integer = Ris(0)
							Dim GoalSubiti As Integer = Ris(1)

							Dim Segno As String = Rec("Segno").Value

							Giocate2 += 1

							If Casa <> Prima Then
								GiocateCasa2 += 1
								GoalFatti2 += GoalFatti
								GoalSubiti2 += GoalSubiti
								GoalFattiCasa2 += GoalFatti
								GoalSubitiCasa2 += GoalSubiti

								Select Case Segno
									Case "1"
										Punti2 += 3
										PuntiCasa2 += 3
										Vittorie2 += 1
										VittorieCasa2 += 1
									Case "X"
										Punti2 += 1
										PuntiCasa2 += 1
										Pareggi2 += 1
										PareggiCasa2 += 1
									Case "2"
										Sconfitte2 += 1
										SconfitteCasa2 += 1
								End Select
							Else
								GiocateFuori2 += 1
								GoalFatti2 += GoalSubiti
								GoalSubiti2 += GoalFatti
								GoalFattiFuori2 += GoalSubiti
								GoalSubitiFuori2 += GoalFatti

								Select Case Segno
									Case "2"
										Punti2 += 3
										PuntiFuori2 += 3
										Vittorie2 += 1
										VittorieFuori2 += 1
									Case "X"
										Punti2 += 1
										PuntiFuori2 += 1
										Pareggi2 += 1
										PareggiFuori2 += 1
									Case "1"
										Sconfitte2 += 1
										SconfitteFuori2 += 1
								End Select
							End If

							Rec.MoveNext
						Loop
						Rec.Close

						Ritorno = ""
						' Lista partite 1
						For Each p As String In Partite1
							Ritorno &= p & "§"
						Next

						Ritorno &= "|"
						' Lista partite 2
						For Each p As String In Partite2
							Ritorno &= p & "§"
						Next

						Ritorno &= "|"
						' Statistiche 1
						Ritorno &= Punti1 & ";"
						Ritorno &= PuntiCasa1 & ";"
						Ritorno &= PuntiFuori1 & ";"
						Ritorno &= Vittorie1 & ";"
						Ritorno &= Pareggi1 & ";"
						Ritorno &= Sconfitte1 & ";"
						Ritorno &= VittorieCasa1 & ";"
						Ritorno &= PareggiCasa1 & ";"
						Ritorno &= SconfitteCasa1 & ";"
						Ritorno &= VittorieFuori1 & ";"
						Ritorno &= PareggiFuori1 & ";"
						Ritorno &= SconfitteFuori1 & ";"
						Ritorno &= GoalFatti1 & ";"
						Ritorno &= GoalSubiti1 & ";"
						Ritorno &= GoalFattiCasa1 & ";"
						Ritorno &= GoalFattiFuori1 & ";"
						Ritorno &= GoalSubitiCasa1 & ";"
						Ritorno &= GoalSubitiFuori1 & ";"
						Ritorno &= Giocate1 & ";"
						Ritorno &= GiocateCasa1 & ";"
						Ritorno &= giocatefuori1

						Ritorno &= "|"
						' Statistiche 2
						Ritorno &= Punti2 & ";"
						Ritorno &= PuntiCasa2 & ";"
						Ritorno &= PuntiFuori2 & ";"
						Ritorno &= Vittorie2 & ";"
						Ritorno &= Pareggi2 & ";"
						Ritorno &= Sconfitte2 & ";"
						Ritorno &= VittorieCasa2 & ";"
						Ritorno &= PareggiCasa2 & ";"
						Ritorno &= SconfitteCasa2 & ";"
						Ritorno &= VittorieFuori2 & ";"
						Ritorno &= PareggiFuori2 & ";"
						Ritorno &= SconfitteFuori2 & ";"
						Ritorno &= GoalFatti2 & ";"
						Ritorno &= GoalSubiti2 & ";"
						Ritorno &= GoalFattiCasa2 & ";"
						Ritorno &= GoalFattiFuori2 & ";"
						Ritorno &= GoalSubitiCasa2 & ";"
						Ritorno &= GoalSubitiFuori2 & ";"
						Ritorno &= Giocate2 & ";"
						Ritorno &= GiocateCasa2 & ";"
						Ritorno &= GiocateFuori2

					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class