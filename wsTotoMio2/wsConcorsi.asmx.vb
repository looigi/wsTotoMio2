Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Linq.Expressions
Imports System.Runtime.CompilerServices
Imports System.Security.Policy
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
											Testo &= "Per partecipare: <a href=" & IndirizzoSito & """>Click QUI</a>"
											InvaMailATutti(Server.MapPath("."), idAnno, "TotoMIO: Apertura concorso " & idGiornata, Testo, Conn, Connessione)
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

				' CAMPIONATO
				Dim ev As New clsEventi
				Dim Classifica As List(Of clsEventi.StrutturaGiocatore) = ev.PrendeGiocatori(Server.MapPath("."), idAnno, 38, Conn, Connessione)
				Ritorno &= "Campione di TotoMIO;" & Classifica.Item(0).idUtente & ";" & Classifica.Item(0).NickName & "§"
				Ritorno &= "Vice Campione;" & Classifica.Item(0).idUtente & ";" & Classifica.Item(1).NickName & "§"
				Ritorno &= "Terzo;" & Classifica.Item(0).idUtente & ";" & Classifica.Item(2).NickName & "§"
				Ritorno &= "Cucchiarella;" & Classifica.Item(0).idUtente & ";" & Classifica.Item(Classifica.Count - 1).NickName & "§"

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
								'Ritorno = "ERROR: Nessuna finale rilevata"
							Else
								If Rec("idVincente").Value = -99 Or Rec("idVincente").Value = -1 Then
									Ritorno &= Descrizione.Item(Conta) & ";-1;Non giocata finale§"
								Else
									If Rec("idVincente").Value = 1 Then
										Ritorno &= Descrizione.Item(Conta) & ";" & Rec("idVincente").Value & ";" & Rec("Giocatore1").Value & "§"
									Else
										Ritorno &= Descrizione.Item(Conta) & ";" & Rec("idVincente").Value & ";" & Rec("Giocatore2").Value & "§"
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

							Ritorno &= Descrizione.Item(Conta) & ";" & r(0) & ";" & r(1) & "§"
						Else
							Ritorno &= Descrizione.Item(Conta) & ";-1;Non ancora creata§"
						End If
					End If

					Conta += 1
				Next

				' 23 Aiutame Te
				Sql = "SELECT A.idUtente, B.NickName, Sum(A.Punti) As Punti FROM SquadreRandom As A " &
					"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
					"Where A.idAnno = " & idAnno & " " &
					"Group By A.idUtente, B.NickName " &
					"Order By 2 Desc "
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						' Ritorno = "ERROR: Nessuna coppa rilevata"
					Else
						Ritorno &= "23 Aiutame Te;" & Rec("idUtente").Value & ";" & Rec("NickName").Value & "§"
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
					"Order By Punti Desc"
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

										Dim TestoRis As String = ""

										sql = "Select A.*, B.NickName, C.Jolly From Risultati A " &
											"Left Join Utenti B On A.idUtente = B.idUtente " &
											"Left Join RisultatiAltro C On A.idAnno = C.idAnno And A.idUtente = C.idUtente And A.idConcorso = C.idConcorso " &
											"Where A.idAnno=" & idAnno & " And A.idConcorso=" & idGiornata
										Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
										If TypeOf (Rec) Is String Then
											'Ritorno = Rec
										Else
											TestoRis = "Risultati Campionato<br /><table style=""width: 100%; border: 1px solid #999;"">"
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
												TestoRis &= "<tr>"

												Rec.MoveNext
											Loop
											TestoRis &= " </table>"
											Rec.Close
										End If

										'sql = "Select A.*, B.NickName, C.Jolly From Risultati A " &
										'	"Left Join Utenti B On A.idUtente = B.idUtente " &
										'	"Left Join RisultatiAltro C On A.idAnno = C.idAnno And A.idUtente = C.idUtente And A.idConcorso = C.idConcorso " &
										'	"Where A.idAnno=" & idAnno & " And A.idConcorso=" & idGiornata
										sql = "Select C.idCoppa, C.Descrizione As NomeCoppa, D.Descrizione As Tipologia FROM EventiPartite A " &
											"Left Join Eventi B On A.idEvento = B.idEvento " &
											"Left Join EventiNomi C On B.idCoppa = C.idCoppa " &
											"Left Join EventiTipologie D On B.idTipologia = D.idTipologia " &
											"Where idAnno = " & idAnno & " And idGiornataVirtuale = " & idGiornata & " And D.Descrizione <> 'Creazione' " &
											"Group By C.idCoppa, C.Descrizione, D.Descrizione"
										Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
										If TypeOf (Rec) Is String Then
											'Ritorno = Rec
										Else
											Dim idCoppe As New List(Of Integer)
											Dim Torneo As New List(Of String)

											Do Until Rec.Eof
												idCoppe.Add(Rec("idCoppa").Value)
												Torneo.Add(Rec("NomeCoppa").Value)

												Rec.MoveNext
											Loop
											Rec.Close

											TestoRis &= " <br /><br />"
											Dim n As Integer = 0
											For Each id As Integer In idCoppe
												TestoRis &= "Torneo: " & Torneo.Item(n) & " <br />"
												TestoRis &= "<table style=""width: 100%; border: 1px solid #999;"">"
												TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
												TestoRis &= "<th>Casa</th>"
												TestoRis &= "<th>Fuori</th>"
												TestoRis &= "<th>Risultato 1</th>"
												TestoRis &= "<th>Risultato 2</th>"
												TestoRis &= "<th>Vincente</th>"
												TestoRis &= "</tr>"
												sql = "Select A.*, B.NickName As Casa, C.NickName As Fuori FROM EventiPartite As A " &
													"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore1 = B.idUtente " &
													"Left Join Utenti As C On A.idAnno = C.idAnno And A.idGiocatore2 = C.idUtente " &
													"Where A.idAnno = " & idAnno & " And A.idGiornataVirtuale = " & idGiornata & " And " &
													"A.idEvento In (Select idEvento From Eventi As E Where idCoppa = " & id & " And A.idGiornata = E.InizioGiornata) " &
													"Order By idPartita"
												Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
												If TypeOf (Rec) Is String Then
													'Ritorno = Rec
												Else
													Do Until Rec.Eof
														TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
														TestoRis &= "<td>" & Rec("Casa").Value & "</td>"
														TestoRis &= "<td>" & Rec("Fuori").Value & "</td>"
														TestoRis &= "<td>" & Rec("Risultato1").Value & "</td>"
														TestoRis &= "<td>" & Rec("Risultato2").Value & "</td>"
														If Rec("idVincente").Value = "1" Then
															TestoRis &= "<td>" & Rec("Casa").Value & "</td>"
														Else
															If Rec("idVincente").Value = "2" Then
																TestoRis &= "<td>" & Rec("Casa").Value & "</td>"
															Else
																If Rec("idVincente").Value = "-1" Then
																Else
																	TestoRis &= "<td>Pareggio</td>"
																End If
															End If
														End If
														TestoRis &= "</tr>"

														Rec.MoveNext
													Loop
													Rec.Close
												End If
												TestoRis &= "</table><br /><br />"
												n += 1
											Next
										End If

										sql = "Select * From SquadreRandom A " &
											"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
											"Where A.idAnno=" & idAnno & " And A.idConcorso=" & idGiornata
										Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
										If TypeOf (Rec) Is String Then
											'Ritorno = Rec
										Else
											TestoRis &= " <br />"
											TestoRis &= "23 Aiutame Te<br />"
											TestoRis &= "<table style=""width: 100%; border: 1px solid #999;"">"
											TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
											TestoRis &= "<th>Utente</th>"
											TestoRis &= "<th>Squadra</th>"
											TestoRis &= "<th>Punti</th>"
											TestoRis &= "</tr>"
											Do Until Rec.Eof
												TestoRis &= "<tr style=""border-bottom: 1px solid #999"">"
												TestoRis &= "<td>" & Rec("NickName").Value & "</td>"
												TestoRis &= "<td>" & Rec("Squadra").Value & "</td>"
												TestoRis &= "<td style=""text-align: center;"">" & Rec("Punti").Value & "</td>"
												TestoRis &= "</tr>"

												Rec.MoveNext
											Loop
											Rec.Close
											TestoRis &= "</table><br />"
										End If

										Dim Testo As String = ""
										Testo = "E' stato controllato il concorso TotoMIO numero " & idGiornata & ".<br />"
										Testo &= "<br />" & TestoRis & "<br />"
										Testo &= "Per entrare nel sito e vedere il resto: <a href=" & IndirizzoSito & """>Click QUI</a>"
										InvaMailATutti(Server.MapPath("."), idAnno, "TotoMIO: Controllo concorso " & idGiornata, Testo, Conn, Connessione)
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
								"Where idAnno = " & idAnno & " And idUtente Not In (Select idUtente From Pronostici Where idAnno = " & idAnno & " And idConcorso = " & idGiornata & ")"
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

							Dim Testo As String = ""
							Testo = "E' stato chiuso il concorso TotoMIO numero " & idGiornata & ".<br />"
							Testo &= "Non sarà più possibile giocare la schedina<br /><br />"
							If Assenti <> "" Then
								Testo &= "Non adempienti:<br /><br />" & Assenti & " <br /><br />"
							End If
							Testo &= "Per entrare nel sito: <a href=" & IndirizzoSito & """>Click QUI</a>"
							InvaMailATutti(Server.MapPath("."), idAnno, "TotoMIO: Chiusura concorso " & idGiornata, Testo, Conn, Connessione)
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
	Public Function ControllaConcorso(idAnno As String, idUtente As String, ModalitaConcorso As String) As String
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

				sql = "Update RisultatiAltro Set Vittorie = 0, Ultimo = 0 Where idAnno=" & idAnno & " And idConcorso=" & idGiornata
				Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
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
							If Not Rec.Eof Then
								PartitaJolly = Rec("idPartita").Value
							Else
								Ritorno = "ERROR: Nessuna partita jolly rilevata"
							End If
							Rec.Close
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
																								 ModalitaConcorso, PartitaJolly, idPartitaScelta)
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
																								 ModalitaConcorso, PartitaJolly, idPartitaScelta)
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
																					ModalitaConcorso, PartitaJolly, idPartitaScelta)
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
																					 ModalitaConcorso, PartitaJolly, idPartitaScelta)
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

									sql = "Update SquadreRandom Set Punti=" & Punti & " Where " &
										"idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente=" & idUtente23
									Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)

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
			Dim Classifica As String = RitornaClassificaGenerale(Server.MapPath("."), idAnno, idGiornata, Conn, Connessione, True)
			If Classifica <> "" Then
				Dim c() As String = Classifica.Split("§")
				Dim PrimaRiga() As String = c(0).Split(";")
				Dim idUltimo As Integer = PrimaRiga(0)
				Dim UltimaRiga() As String = c(c.Count - 2).Split(";")
				Dim idPrimo As Integer = UltimaRiga(0)
				Dim Ritorno2 As String = "OK"

				sql = "Select * From RisultatiAltro Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente=" & idPrimo
				Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno2 = Rec
				Else
					If Rec.Eof Then
						sql = "Insert Into RisultatiAltro Values (" & idAnno & ", " & idGiornata & ", " & idPrimo & ", 1, 0, 0)"
						Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					Else
						sql = "Update RisultatiAltro Set Ultimo = 1 Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente = " & idPrimo
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
						sql = "Update RisultatiAltro Set Vittorie = 1 Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idUtente = " & idUltimo
						Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					End If
				End If
			End If
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
			"Where A.idAnno = " & idAnno & " And A.idConcorso <= " & idConcorso & " " &
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
				"Where A.idAnno = " & idAnno & " And A.idConcorso = " & idConcorso & " " &
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
End Class