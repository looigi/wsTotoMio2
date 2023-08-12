Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Windows.Forms
Imports wsTotoMio2.clsRecordset

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looUtentiTotoMio2.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsUtenti
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function AggiungeUtente(idAnno As String, NickName As String, Cognome As String, Nome As String,
								   Password As String, Mail As String, idTipologia As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		If Not ControllaValiditaMail(Mail) Then
			Ritorno = "ERROR: Mail non valida"
		Else
			If Cognome = "" Or Cognome.Length < 3 Then
				Return "ERROR: Cognome non valido"
			End If
			If Nome = "" Or Nome.Length < 3 Then
				Return "ERROR: Nome non valido"
			End If
			If Password = "" Or Password.Length < 3 Then
				Return "ERROR: Password non valida"
			End If

			Dim sql As String = "Select * From Utenti Where idAnno=" & idAnno & " And NickName='" & SistemaStringaPerDB(NickName) & "'"
			Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Not Rec.Eof Then
					Ritorno = "ERROR: NickName già presente"
				Else
					sql = "Select Coalesce(Max(idUtente) + 1, 1) From Utenti Where idAnno=" & idAnno
					Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Dim idUtente As String = Rec(0).Value

						sql = "Insert Into Utenti Values (" &
							" " & idAnno & ", " &
							" " & idUtente & ", " &
							"'" & SistemaStringaPerDB(NickName) & "', " &
							"'" & SistemaStringaPerDB(Cognome) & "', " &
							"'" & SistemaStringaPerDB(Nome) & "', " &
							"'" & SistemaStringaPerDB(Password) & "', " &
							"'" & SistemaStringaPerDB(Mail) & "', " &
							" " & idTipologia & ", " &
							"'N'" &
							")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Not Ritorno.Contains(StringaErrore) Then
							sql = "Insert Into UtentiMail Values (" &
								" " & idAnno & ", " &
								" " & idUtente & ", " &
								"'S', " &
								"'S', " &
								"'S', " &
								"'S', " &
								"'S', " &
								"'S' " &
								")"
							Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
							If Not Ritorno.Contains(StringaErrore) Then
								Ritorno = idUtente

								Dim gi As New GestioneImmagini
								gi.CreaAvatar(Server.MapPath("."), idAnno, idUtente, NickName, Nome, Cognome)

								Dim Testo As String = "Nuovo utente registrato:<br /><br /><style=""font-weight: bold;"">" & NickName & "</style><br />" &
									"<style=""font-weight: bold;"">" & Nome & " " & Cognome & "</style>"
								Testo &= "<br /><br />Per accedere: <a href=""" & IndirizzoSito & """>Click QUI</a>"

								Dim m As New mail(Server.MapPath("."))

								sql = "Select * From Utenti Where idAnno=" & idAnno & " And Eliminato='N' And idTipologia=0"
								Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									Dim Mails As New List(Of String)
									Dim mmm As String = ""
									Mails.Add(Mail)
									mmm &= Mail & ";"
									Do Until Rec.Eof
										If Not mmm.Contains(Rec("Mail").Value & ";") Then
											Mails.Add(Rec("Mail").Value)
											mmm &= Rec("Mail").Value
										End If

										Rec.MoveNext
									Loop
									Rec.Close
									For Each mm As String In Mails
										m.SendEmail(Server.MapPath("."), mm, "TotoMIO: Registrazione nuovo utente", Testo, {})
									Next
								End If
							End If
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaUtente(idAnno As String, idUtente As String, NickName As String, Cognome As String, Nome As String,
								   Password As String, Mail As String, idTipologia As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		If Not ControllaValiditaMail(Mail) Then
			Ritorno = "ERROR: Mail non valida"
		Else
			If Cognome = "" Or Cognome.Length < 3 Then
				Return "ERROR: Cognome non valido"
			End If
			If Nome = "" Or Nome.Length < 3 Then
				Return "ERROR: Nome non valido"
			End If
			If Password = "" Or Password.Length < 3 Then
				Return "ERROR: Password non valida"
			End If

			Dim sql As String = "Update Utenti Set " &
				"NickName='" & SistemaStringaPerDB(NickName) & "', " &
				"Cognome='" & SistemaStringaPerDB(Cognome) & "', " &
				"Nome='" & SistemaStringaPerDB(Nome) & "', " &
				"Password='" & SistemaStringaPerDB(Password) & "', " &
				"Mail='" & SistemaStringaPerDB(Mail) & "', " &
				"idTipologia=" & idTipologia & " " &
				"Where idAnno=" & idAnno & " And idUtente=" & idUtente
			Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			If Not Ritorno.Contains(StringaErrore) Then
				Ritorno = "*"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaUtente(idAnno As String, idUtente As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Update Utenti Set Eliminato = 'S' Where idAnno=" & idAnno & " And idUtente=" & idUtente
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
		If Not Ritorno.Contains(StringaErrore) Then
			Ritorno = "*"
		End If
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaUtentePerLogin(idAnno As String, NickName As String, Password As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Select * From Utenti As A " &
			"Left Join UtentiTipologie As B On A.idTipologia = B.idTipologia " &
			"Where idAnno=" & idAnno & " And Eliminato='N' And NickName='" & SistemaStringaPerDB(NickName) & "'"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun utente rilevato"
			Else
				If Rec("Password").Value <> Password Then
					Ritorno = "ERROR: Password non valida"
				Else
					'Do Until Rec.Eof
					Ritorno &= SistemaStringaPerRitorno(Rec("idUtente").Value) & ";"
					Ritorno &= SistemaStringaPerRitorno(Rec("NickName").Value) & ";"
					Ritorno &= SistemaStringaPerRitorno(Rec("Cognome").Value) & ";"
					Ritorno &= SistemaStringaPerRitorno(Rec("Nome").Value) & ";"
					Ritorno &= SistemaStringaPerRitorno(Rec("Password").Value) & ";"
					Ritorno &= SistemaStringaPerRitorno(Rec("Mail").Value) & ";"
					Ritorno &= SistemaStringaPerRitorno(Rec("idTipologia").Value) & ";"
					Ritorno &= SistemaStringaPerRitorno(Rec("Descrizione").Value) & ";"

					'	Rec.MoveNext
					'Loop
				End If
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaPronosticoUtente(idAnno As String, idUtente As String, idConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From Pronostici Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente & " Order By idPartita"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun utente rilevato"
			Else
				Dim Quante As Integer = 0

				Do Until Rec.Eof
					Ritorno &= Rec("idPartita").Value & ";" & Rec("Pronostico").Value & ";" & Rec("Segno").Value & "§"
					Quante += 1

					Rec.MoveNext
				Loop
				Rec.Close

				sql = "Select * From PartiteScelte Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
				Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Dim idPartitaScelta As Integer = -1

					If Rec.Eof Then
						sql = "Delete From PartiteScelte Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
						Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Not Ritorno2.Contains("ERROR:") Then
							idPartitaScelta = GetRandom(1, Quante)

							sql = "Insert Into PartiteScelte Values (" & idAnno & ", " & idConcorso & ", " & idUtente & ", " & idPartitaScelta & ")"
							Ritorno2 = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
							If Not Ritorno2.Contains("ERROR:") Then
							End If
						End If
					Else
						idPartitaScelta = Rec("idPartita").Value
					End If

					Ritorno &= "|" & idPartitaScelta
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaClassifica(idAnno As String, idConcorso As String, MostraFinto As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Return RitornaClassificaGenerale(Server.MapPath("."), idAnno, idConcorso, Conn, Connessione, False, MostraFinto)
	End Function

	<WebMethod()>
	Public Function RitornaTuttiGliUtenti(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = ""

		Sql = "Select * From Utenti Where idAnno=" & idAnno & " And Eliminato='N' Order By idUtente"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: nessun movimento rilevato"
			Else
				Do Until Rec.Eof
					Ritorno &= Rec("idUtente").Value & ";" & SistemaStringaPerRitorno(Rec("NickName").Value) & "§"

					Rec.MoveNext
				Loop
				Rec.CLose
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaPronosticoUtente(idAnno As String, idUtente As String, idConcorso As String, Dati As String, idPartitaScelta As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""

		sql = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
		If Not Ritorno.Contains("ERROR:") Then
			sql = "Delete From Pronostici Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
			Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			If Not Ritorno.Contains("ERROR:") Then
				Dim Righe() As String = Dati.Split("§")

				For Each r As String In Righe
					If r <> "" Then
						Dim Campi() As String = r.Split(";")
						sql = "Insert Into Pronostici Values (" &
							" " & idAnno & ", " &
							" " & idUtente & ", " &
							" " & idConcorso & ", " &
							" " & Campi(0) & ", " &
							"'" & Campi(1) & "', " &
							"'" & Campi(2) & "' " &
							")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Ritorno.Contains("ERROR:") Then
							Exit For
						Else
							Ritorno = "OK"
						End If
					End If
				Next

				If Not Ritorno.Contains("ERROR:") Then
					sql = "Delete From PartiteScelte Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
					Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					If Not Ritorno.Contains("ERROR:") Then
						sql = "Insert Into PartiteScelte Values (" & idAnno & ", " & idConcorso & ", " & idUtente & ", " & idPartitaScelta & ")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
					End If
				End If
			End If

			If Ritorno = "OK" Then
				Dim Risultati As String = ""

				sql = "SELECT * FROM Pronostici As A " &
					"Left Join Concorsi B On A.idAnno = B.idAnno And A.idConcorso = B.idConcorso And A.idPartita = B.idPartita " &
					"Where A.idConcorso = " & idConcorso & " And A.idAnno = " & idAnno & " And A.idUtente =  " & idUtente & " " &
					"Order By A.idPartita"
				Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					'Ritorno = Rec
				Else
					Risultati &= "<tr style=""border-bottom: 1px solid #999"">"
					Risultati &= "<th>N°</th>"
					Risultati &= "<th>Casa</th>"
					Risultati &= "<th>Fuori</th>"
					Risultati &= "<th>Pronostico</th>"
					Risultati &= "<th>Segno</th>"
					Risultati &= "</tr>"
					Do Until Rec.Eof
						Risultati &= "<tr style=""border-bottom: 1px solid #999"">"
						Risultati &= "<td>" & Rec("idPartita").Value & "</td>"
						Risultati &= "<td>" & Rec("Prima").Value & "</td>"
						Risultati &= "<td>" & Rec("Seconda").Value & "</td>"
						Risultati &= "<td style=""text-align: center"">" & Rec("Pronostico").Value & "</td>"
						Risultati &= "<td style=""text-align: center"">" & Rec("Segno").Value & "</td>"
						Risultati &= "</tr>"
						Rec.MoveNext
					Loop
					Rec.Close
					Risultati &= "</table><br />"
					Risultati &= "Partita scelta: " & idPartitaScelta & "<br />"
				End If

				Dim EMail As String = ""
				Dim NickName As String = ""

				sql = "Select * From Utenti A " &
					"Left Join UtentiMails B On A.idAnno=B.idAnno And A.idUtente=B.idUtente " &
					"Where A.idAnno=" & idAnno & " And A.idUtente=" & idUtente & " And Giocata='S'"
				Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					'Ritorno = Rec
				Else
					EMail = Rec("Mail").Value
					NickName = Rec("NickName").Value
					Rec.Close
				End If

				Dim Mails As New List(Of String)

				If EMail <> "" Then
					Dim Testo As String = ""
					Testo = "E' stata giocata la colonna da " & NickName.ToUpper & " per il concorso TotoMIO numero " & idConcorso & ".<br />"
					Testo &= "<br />" & Risultati & "<br />"
					Testo &= "Per entrare nel sito e vedere il resto: <a href=""" & IndirizzoSito & """>Click QUI</a>"

					Dim m As New mail(Server.MapPath("."))

					sql = "Select * From Utenti A " &
						"Left Join UtentiMails B On A.idAnno=B.idAnno And A.idUtente=B.idUtente " &
						"Where A.idAnno=" & idAnno & " And Eliminato='N' And idTipologia=0 And Giocata='S'"
					Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						Dim mmm As String = ""
						Mails.Add(EMail)
						mmm &= EMail & ";"
						Do Until Rec.Eof
							If Not mmm.Contains(Rec("Mail").Value & ";") Then
								Mails.Add(Rec("Mail").Value)
								mmm &= Rec("Mail").Value
							End If

							Rec.MoveNext
						Loop
						Rec.Close
					End If

					For Each mm As String In Mails
						m.SendEmail(Server.MapPath("."), mm, "TotoMIO: Colonna utente per concorso numero " & idConcorso, Testo, {})
					Next
				End If

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
	Public Function CreaImmagineStandard(idAnno As Integer, idUtente As Integer) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From Utenti Where idAnno=" & idAnno & " And idUtente=" & idUtente
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun utente rilevato"
			Else
				Dim NickName As String = Rec("NickName").Value
				Dim Cognome As String = Rec("Cognome").Value
				Dim Nome As String = Rec("Nome").Value
				Rec.Close

				Dim gi As New GestioneImmagini
				gi.CreaAvatar(Server.MapPath("."), idAnno, idUtente, NickName, Nome, Cognome)
				Ritorno = "OK"
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaUtentiMails(idAnno As String, idUtente As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Select * From UtentiMails Where idAnno=" & idAnno & " And idUtente=" & idUtente
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun utente rilevato"
			Else
				Ritorno = Rec("Apertura").Value & ";" & Rec("Reminder").Value & ";" &
					Rec("Controllo").Value & ";" & Rec("Chiusura").Value & ";" &
					Rec("Chat").Value & ";" & Rec("Giocata").Value
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ImpostaUtentiMails(idAnno As String, idUtente As String, Apertura As String, Reminder As String,
									   Controllo As String, Chiusura As String, Chat As String, Giocata As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Update UtentiMails Set " &
			"Apertura='" & Apertura & "', " &
			"Reminder='" & Reminder & "', " &
			"Controllo='" & Controllo & "', " &
			"Chiusura='" & Chiusura & "', " &
			"Chat='" & Chat & "', " &
			"Giocata='" & Giocata & "' " &
			"Where idAnno=" & idAnno & " And idUtente=" & idUtente
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function Statistiche(idAnno As Integer) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Ritorno1 As String = ElaboraStatistiche(idAnno, Conn, Connessione)
		Dim Ritorno2 As String = ElaboraStatistiche("", Conn, Connessione)
		If Not Ritorno1.Contains(StringaErrore) And Not Ritorno2.Contains(StringaErrore) Then
			Ritorno = "{"
			Ritorno &= "Annuale: " & Ritorno1 & ","
			Ritorno &= "Storico: " & Ritorno2
			Ritorno &= "}"
		End If
		Return Ritorno
	End Function

	Private Function ElaboraStatistiche(idAnno As String, Conn As Object, Connessione As String) As String
		Dim Ritorno As String = ""
		Dim Altro As String = ""
		Dim Altro2 As String = ""
		If idAnno <> "" Then
			Altro = "Where A.idAnno = " & idAnno & " "
			Altro2 = "A.idAnno = " & idAnno & " And"
		End If

		Dim sql As String = "SELECT A.idUtente, B.NickName, Coalesce(Avg(Punti), 0) As MediaPunti, Coalesce(Avg(SegniPresi), 0) As MediaSegni, " &
			"Coalesce(Avg(RisultatiEsatti), 0) As MediaRisEsatti, Coalesce(Avg(RisultatiCasaTot), 0) As MediaRisCasa, " &
			"Coalesce(Avg(RisultatiFuoriTot), 0) As MediaRisFuori, Coalesce(Avg(SommeGoal), 0) As MediaSomma, " &
			"Coalesce(Avg(DifferenzeGoal), 0) As MediaDiff, Coalesce(Avg(PuntiPartitaScelta), 0) As MediaPuntiPS,  " &
			"Coalesce(Sum(Punti), 0) As SommaPunti, Coalesce(Sum(SegniPresi), 0) As SommaSegni, " &
			"Coalesce(Sum(RisultatiEsatti), 0) As SommaRisEsatti, Coalesce(Sum(RisultatiCasaTot), 0) As SommaRisCasa, " &
			"Coalesce(Sum(RisultatiFuoriTot), 0) As SommaRisFuori, Coalesce(Sum(SommeGoal), 0) As SommaSomma, " &
			"Coalesce(Sum(DifferenzeGoal), 0) As SommaDiff, Coalesce(Sum(PuntiPartitaScelta), 0) As SommaPuntiPS " &
			"FROM Risultati As A " &
			"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
			Altro & " " &
			"Group By A.idUtente, B.NickName " &
			"Order By A.idUtente"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			'If Rec.Eof Then
			'	Ritorno = "ERROR: Nessun utente rilevato"
			'Else
			Dim StatisticheRisultati As String = "["
			Do Until Rec.Eof
				StatisticheRisultati &= "{"
				StatisticheRisultati &= "idUtente: " & Rec("idUtente").Value & ", "
				StatisticheRisultati &= "NickName: " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
				StatisticheRisultati &= "SommaPunti: " & SistemaNumeroDaDB(Rec("SommaPunti").Value, False) & ", "
				StatisticheRisultati &= "SommaSegni: " & SistemaNumeroDaDB(Rec("SommaSegni").Value, False) & ", "
				StatisticheRisultati &= "SommaRisEsatti: " & SistemaNumeroDaDB(Rec("SommaRisEsatti").Value, False) & ", "
				StatisticheRisultati &= "SommaRisCasa: " & SistemaNumeroDaDB(Rec("SommaRisCasa").Value, False) & ", "
				StatisticheRisultati &= "SommaRisFuori: " & SistemaNumeroDaDB(Rec("SommaRisFuori").Value, False) & ", "
				StatisticheRisultati &= "SommaSomma: " & SistemaNumeroDaDB(Rec("SommaSomma").Value, False) & ", "
				StatisticheRisultati &= "SommaDiff: " & SistemaNumeroDaDB(Rec("SommaDiff").Value, False) & ", "
				StatisticheRisultati &= "SommaPuntiPS: " & SistemaNumeroDaDB(Rec("SommaPuntiPS").Value, False) & ", "
				StatisticheRisultati &= "MediaPunti: " & SistemaNumeroDaDB(Rec("MediaPunti").Value, True) & ", "
				StatisticheRisultati &= "MediaSegni: " & SistemaNumeroDaDB(Rec("MediaSegni").Value, True) & ", "
				StatisticheRisultati &= "MediaRisEsatti: " & SistemaNumeroDaDB(Rec("MediaRisEsatti").Value, True) & ", "
				StatisticheRisultati &= "MediaRisCasa: " & SistemaNumeroDaDB(Rec("MediaRisCasa").Value, True) & ", "
				StatisticheRisultati &= "MediaRisFuori: " & SistemaNumeroDaDB(Rec("MediaRisFuori").Value, True) & ", "
				StatisticheRisultati &= "MediaSomma: " & SistemaNumeroDaDB(Rec("MediaSomma").Value, True) & ", "
				StatisticheRisultati &= "MediaDiff: " & SistemaNumeroDaDB(Rec("MediaDiff").Value, True) & ", "
				StatisticheRisultati &= "MediaPuntiPS: " & SistemaNumeroDaDB(Rec("MediaPuntiPS").Value, True) & " "
				StatisticheRisultati &= "}, "

				Rec.MoveNext
			Loop
			If StatisticheRisultati <> "[" Then
				StatisticheRisultati = Mid(StatisticheRisultati, 1, StatisticheRisultati.Length - 2)
			End If
			StatisticheRisultati &= "]"
			Rec.Close

			sql = "SELECT A.idUtente, B.NickName, Coalesce(Avg(Vittorie), 0) As MediaVittorie,  " &
				"Coalesce(Avg(Ultimo), 0) As MediaUltimo, Coalesce(Avg(Jolly), 0) As MediaJolly, " &
				"Coalesce(Sum(Vittorie), 0) As SommaVittorie, " &
				"Coalesce(Sum(Ultimo), 0) As SommaUltimo, Coalesce(Sum(Jolly), 0) As SommaJolly " &
				"From RisultatiAltro As A " &
				"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
				Altro & " " &
				"Group By A.idUtente, B.idUtente " &
				"Order By A.idUtente"
			Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				'If Rec.Eof Then
				'	Ritorno = "ERROR: Nessun utente rilevato"
				'Else
				Dim StatisticheRisultatiA As String = "["
				Do Until Rec.Eof
					StatisticheRisultatiA &= "{"
					StatisticheRisultatiA &= "idUtente: " & Rec("idUtente").Value & ", "
					StatisticheRisultatiA &= "NickName: " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
					StatisticheRisultatiA &= "SommaVittorie: " & SistemaNumeroDaDB(Rec("SommaVittorie").Value, False) & ", "
					StatisticheRisultatiA &= "SommaUltimo: " & SistemaNumeroDaDB(Rec("SommaUltimo").Value, False) & ", "
					StatisticheRisultatiA &= "SommaJolly: " & SistemaNumeroDaDB(Rec("SommaJolly").Value, False) & ", "
					StatisticheRisultatiA &= "MediaVittorie: " & SistemaNumeroDaDB(Rec("MediaVittorie").Value, True) & ", "
					StatisticheRisultatiA &= "MediaUltimo: " & SistemaNumeroDaDB(Rec("MediaUltimo").Value, True) & ", "
					StatisticheRisultatiA &= "MediaJolly: " & SistemaNumeroDaDB(Rec("MediaJolly").Value, True) & " "
					StatisticheRisultatiA &= "}, "

					Rec.MoveNext
				Loop
				If StatisticheRisultatiA <> "[" Then
					StatisticheRisultatiA = Mid(StatisticheRisultatiA, 1, StatisticheRisultatiA.Length - 2)
				End If
				StatisticheRisultatiA &= "]"
				Rec.Close

				sql = "Select idUtente, NickName, Sum(Vinte) As Vinte, Sum(Pareggiate) As Pareggiate, Sum(Perse) As Perse, (Sum(Vinte) + Sum(Pareggiate) + Sum(Perse)) As Giocate From (" &
					"SELECT B.idUtente, B.NickName, Coalesce(Count(*), 0) As Vinte, 0 As Pareggiate, 0 As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore1 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 1 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"SELECT B.idUtente, B.NickName, Coalesce(Count(*), 0) As Vinte, 0 As Pareggiate, 0 As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore2 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 2 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"Select B.idUtente, B.NickName, 0 As Vinte, Coalesce(Count(*), 0) As Pareggiate, 0 As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore1 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 0 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"SELECT B.idUtente, B.NickName, 0 As Vinte, Coalesce(Count(*), 0) As Pareggiate, 0 As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore2 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 0 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"SELECT B.idUtente, B.NickName, 0 As Vinte, 0 As Pareggiate, Coalesce(Count(*), 0) As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore1 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 2 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"SELECT B.idUtente, B.NickName, 0 As Vinte, 0 As Pareggiate, Coalesce(Count(*), 0) As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore2 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 1 " &
					"Group By B.idUtente, B.NickName " &
					") As A " &
					"Group By idUtente, NickName " &
					"Order By 3 Desc, 2 Desc"
				Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					'If Rec.Eof Then
					'	Ritorno = "ERROR: Nessun utente rilevato"
					'Else
					Dim StatisticheScontriDiretti As String = "["
					Do Until Rec.Eof
						StatisticheScontriDiretti &= "{"
						StatisticheScontriDiretti &= "idUtente: " & Rec("idUtente").Value & ", "
						StatisticheScontriDiretti &= "NickName: " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
						StatisticheScontriDiretti &= "Vinte: " & SistemaNumeroDaDB(Rec("Vinte").Value, False) & ", "
						StatisticheScontriDiretti &= "Pareggiate: " & SistemaNumeroDaDB(Rec("Pareggiate").Value, False) & ", "
						StatisticheScontriDiretti &= "Perse: " & SistemaNumeroDaDB(Rec("Perse").Value, False) & ", "
						StatisticheScontriDiretti &= "Giocate: " & SistemaNumeroDaDB(Rec("Giocate").Value, False) & ", "
						StatisticheScontriDiretti &= "MediaVinte: " & SistemaNumeroDaDB(Rec("Vinte").Value / Rec("Giocate").Value, False) & ", "
						StatisticheScontriDiretti &= "MediaPareggiate: " & SistemaNumeroDaDB(Rec("Pareggiate").Value / Rec("Giocate").Value, False) & ", "
						StatisticheScontriDiretti &= "MediaPerse: " & SistemaNumeroDaDB(Rec("Perse").Value / Rec("Giocate").Value, False) & ", "
						StatisticheScontriDiretti &= "}, "

						Rec.MoveNext
					Loop
					If StatisticheScontriDiretti <> "[" Then
						StatisticheScontriDiretti = Mid(StatisticheScontriDiretti, 1, StatisticheScontriDiretti.Length - 2)
					End If
					StatisticheScontriDiretti &= "]"
					Rec.Close

					Dim idGiornata As String = ""
					If idAnno <> "" Then
						sql = "Select * From Globale Where idAnno=" & idAnno
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							idGiornata = Rec("idGiornata").Value
							Rec.Close
						End If
					Else
						idGiornata = "Nessuna"
					End If

					Dim QuantiAnni As String = "1"
					If idAnno = "" Then
						sql = "Select Coalesce(Count(*), 0) As Quanti From Anni"
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							QuantiAnni = Rec("Quanti").Value
							Rec.Close
						End If
					End If

					sql = "Select idUtente, NickName, Sum(Entrate) As SommaEntrate, Sum(Uscite) As SommaUscite, Sum(Vincita) As SommaVincita, " &
						"(Sum(Entrate) + Sum(Vincita)) - Sum(Uscite) As SommaBilancio From ( " &
						"SELECT A.idUtente, B.NickName, Sum(Importo) As Entrate, 0 As Uscite, 0 As Vincita FROM Bilancio As A " &
						"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
						"Where " & Altro2 & " idMovimento = 1 And A.Eliminato = 'N' " &
						"Group By A.idUtente, B.NickName " &
						"Union ALL " &
						"SELECT A.idUtente, B.NickName, 0 Entrate, Sum(Importo) As Uscite, 0 As Vincita FROM Bilancio As A " &
						"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
						"Where " & Altro2 & " idMovimento = 2 And A.Eliminato = 'N' " &
						"Group By A.idUtente, B.NickName " &
						"Union ALL " &
						"SELECT A.idUtente, B.NickName, 0 Entrate, 0 As Uscite, Sum(Importo) As Vincita FROM Bilancio As A " &
						"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
						"Where " & Altro2 & " idMovimento = 3 And A.Eliminato = 'N' " &
						"Group By A.idUtente, B.NickName " &
						") AS A  " &
						"Group By idUtente, NickName " &
						"Order By 6 Desc"
					Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						'If Rec.Eof Then
						'	Ritorno = "ERROR: Nessun utente rilevato"
						'Else
						Dim StatisticheBilancio As String = "["
						Do Until Rec.Eof
							StatisticheBilancio &= "{"
							StatisticheBilancio &= "idUtente: " & Rec("idUtente").Value & ", "
							StatisticheBilancio &= "NickName: " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
							StatisticheBilancio &= "Entrate: " & SistemaNumeroDaDB(Rec("SommaEntrate").Value, True) & ", "
							StatisticheBilancio &= "Uscite: " & SistemaNumeroDaDB(Rec("SommaUscite").Value, True) & ", "
							StatisticheBilancio &= "Vincite: " & SistemaNumeroDaDB(Rec("SommaVincite").Value, True) & ", "
							StatisticheBilancio &= "Bilancio: " & SistemaNumeroDaDB(Rec("SommaBilancio").Value, True) & " "
							StatisticheBilancio &= "}, "

							Rec.MoveNext
						Loop
						If StatisticheBilancio <> "[" Then
							StatisticheBilancio = Mid(StatisticheBilancio, 1, StatisticheBilancio.Length - 2)
						End If
						StatisticheBilancio &= "]"
						Rec.Close

						Dim Anno As String = ""
						If idAnno <> "" Then
							Anno = "1"
						Else
							Anno = "Tutti"
						End If

						Dim Quanti As String = "0"
						sql = "SELECT Coalesce(Count(*), 0) As Quanti FROM Concorsi " &
							IIf(idAnno <> "", "Where idAnno = " & idAnno & " And idPartita=1 Group By idAnno", "Where idPartita=1")
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Not Rec.Eof Then
								Quanti = Rec("Quanti").Value
							End If
							Rec.Close

							sql = "SELECT A.idUtente, NickName, Coalesce(Count(*), 0) As Quanti FROM Pronostici As A " &
								"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
								"Where " & Altro2 & " A.idPartita = 1 " &
								"Group By A.idUtente, NickName"
							Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								'If Rec.Eof Then
								'	Ritorno = "ERROR: Nessun utente rilevato"
								'Else
								Dim StatistichePronostici As String = "["
								Do Until Rec.Eof
									StatistichePronostici &= "{"
									StatistichePronostici &= "idUtente: " & Rec("idUtente").Value & ", "
									StatistichePronostici &= "NickName: " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
									StatistichePronostici &= "Giocate: " & SistemaNumeroDaDB(Rec("Quanti").Value, False) & " "
									StatistichePronostici &= "}, "

									Rec.MoveNext
								Loop
								If StatistichePronostici <> "[" Then
									StatistichePronostici = Mid(StatistichePronostici, 1, StatistichePronostici.Length - 2)
								End If
								StatistichePronostici &= "]"
								Rec.Close

								Dim SquadrePrese As String = GeneraSquadrePrese(Server.MapPath("."), idAnno, Conn, Connessione)

								Ritorno &= "{"
								Ritorno &= "Anno: " & Chr(34) & Anno & Chr(34) & ", "
								Ritorno &= "Anni: " & QuantiAnni & ", "
								Ritorno &= "Giornata: " & Chr(34) & idGiornata & Chr(34) & ", "
								Ritorno &= "ConcorsiAperti: " & Quanti & ", "
								Ritorno &= "Risultati: " & StatisticheRisultati & ", "
								Ritorno &= "RisultatiAltro: " & StatisticheRisultatiA & ", "
								Ritorno &= "ScontriDiretti: " & StatisticheScontriDiretti & ", "
								Ritorno &= "Pronostici: " & StatistichePronostici & ", "
								Ritorno &= "SquadrePrese: " & SquadrePrese & ", "
								Ritorno &= "Bilancio: " & StatisticheBilancio
								Ritorno &= "}"
							End If
						End If

						'End If
					End If
				End If
				'End If
			End If
			'End If
		End If

		Return Ritorno
	End Function
End Class