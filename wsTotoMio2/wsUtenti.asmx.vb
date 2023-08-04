Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looUtentiTotoMio2.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
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
	Public Function RitornaClassifica(idAnno As String, idConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Return RitornaClassificaGenerale(Server.MapPath("."), idAnno, idConcorso, Conn, Connessione, False)
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

End Class