Imports Microsoft.SqlServer
Imports wsTotoMio2.clsRecordset

Module mdlGenerale
	Public Structure strutturaMail
		Dim Destinatario As String
		Dim Oggetto As String
		Dim newBody As String
		Dim Allegato() As String
	End Structure
	Public listaMails As New List(Of strutturaMail)
	Public timerMails As Timers.Timer = Nothing
	Public path1 As String = ""
	Public pathMail As String = ""
	Public effettuaLogMail As Boolean = True
	Public nomeFileLogmail As String = ""
	Public StringaErrore As String = "ERROR: "
	Public TipoServer As String = "MARIADB"
	Public CorpoMail As String = ""
	Public IndirizzoSito As String = "http://looigi.ddns.net:1080/"

	Public Function SistemaStringaPerDB(Stringa As String) As String
		Dim Ritorno As String = Stringa

		Ritorno = Ritorno.Replace("'", "''")
		Ritorno = Ritorno.Replace("*PV*", ";")
		Ritorno = Ritorno.Replace("*SS*", "§")

		Return Ritorno
	End Function

	Public Function SistemaStringaPerRitorno(Stringa As String) As String
		Dim Ritorno As String = Stringa

		Ritorno = Ritorno.Replace(";", "*PV*")
		Ritorno = Ritorno.Replace("§", "*SS*")

		Return Ritorno
	End Function

	Public Function SistemaStringaPerRitorno2(Stringa As String) As String
		Dim Ritorno As String = Stringa

		Ritorno = Ritorno.Replace("*PV*", ";")
		Ritorno = Ritorno.Replace("*SS*", "§")

		Return Ritorno
	End Function

	Public Function ConverteNome(Stringa As String) As String
		Dim sStringa As String = Stringa
		sStringa = sStringa.Replace("***AND***", "&")
		sStringa = sStringa.Replace("***PI***", "?")
		sStringa = sStringa.Replace("***BS***", "/")
		sStringa = sStringa.Replace("***BD***", "\")
		sStringa = sStringa.Replace("***PERC***", "%")

		Return sStringa
	End Function

	Public Sub ScriveLog(MP As String, Squadra As String, NomeFile As String, Cosa As String)
		If Not effettuaLogMail Then
			Return
		End If

		If Squadra = "" Then
			Squadra = "NessunaSquadra"
		End If

		Dim gf As New GestioneFilesDirectory

		'If nomeFileLogMail = "" Then
		Dim PathLog As String = MP & "\Logs"
		Dim nomeFileLog As String = PathLog & "\LogsIO.txt"
		gf.CreaDirectoryDaPercorso(nomeFileLog)
		'End If

		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

		gf.ApreFileDiTestoPerScrittura(nomeFileLog)
		gf.ScriveTestoSuFileAperto(Datella & " " & Cosa)
		gf.ChiudeFileDiTestoDopoScrittura()

		gf = Nothing
	End Sub

	Public Function SistemaPercorso(pathPassato As String) As String
		Dim pp As String = pathPassato

		pp = pp.Replace(vbCrLf, "").Trim()
		If Strings.Right(pp, 1) = "\" Or Strings.Right(pp, 1) = "/" Then
			pp = Mid(pp, 1, pp.Length - 1)
		End If

		Return pp
	End Function

	Public Function RitornaPercorso(Path As String, Quale As Integer) As String
		Dim gf As New GestioneFilesDirectory
		Dim tutto As String = gf.LeggeFileIntero(Path & "\PathDB.txt").Replace(vbCrLf, "")
		Dim righe() As String = tutto.Split(";")
		Dim ritorno As String = righe(Quale - 1)
		ritorno = Mid(ritorno, ritorno.IndexOf("=") + 2, ritorno.Length)
		gf = Nothing
		Return ritorno
	End Function

	Public Function CreaRecordset(Mp As String, Conn As Object, Sql As String, Connessione As String) As Object
		Dim rec As Object
		rec = Conn.LeggeQuery(Mp, Sql, Connessione)
		If rec Is Nothing Then
			rec = "ERROR: query non valida -> " & Sql
		End If

		Return rec
	End Function

	Public Function ControllaValiditaMail(Mail As String) As Boolean
		If Not Mail.Contains("@") And Not Mail.Contains(".") Then
			Return False
		Else
			Dim c() As String = Mail.Split("@")

			If c(0).Length < 3 Then
				Return False
			Else
				Dim c2() As String = c(1).Split(".")

				If c2(0).Length < 3 Then
					Return False
				Else
					If c2(1).Length > 3 Then
						Return False
					Else
						Return True
					End If
				End If
			End If
		End If
	End Function

	Public Function RitornaGiornata(MP As String, Conn As Object, Connessione As String, idAnno As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String = "Select * From Globale Where idAnno=" & idAnno
		Dim Rec As Object = CreaRecordset(MP, Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Dim idGiornata As String = Rec("idGiornata").Value
			Rec.Close
			Ritorno = idGiornata
		End If

		Return Ritorno
	End Function

	Public Function ControllaPunti(idAnno As String, idUtente As String, idConcorso As String, NickName As String, Partite As List(Of String), Pronostici As List(Of String),
								   Conn As Object, Connessione As String, MP As String, Modalita As String, PartitaJolly As Integer, idPartitaScelta As Integer) As String
		Dim PuntiTotali As Integer = 0
		Dim Ritorno As String = ""
		Dim SegniPresi As Integer = 0
		Dim RisultatoEsatto As Integer = 0
		Dim RisultatoCasaTot As Integer = 0
		Dim RisultatoFuoriTot As Integer = 0
		Dim SommaGoal As Integer = 0
		Dim DifferenzaGoal As Integer = 0
		Dim PuntiPartitaScelta2 As Integer = 0
		Dim Jolly2 As Integer = 0

		Jolly2 = 0
		For Each Partita As String In Partite
			'Partite.Add(Rec("idPartita").Value & ";" & SistemaStringaPerRitorno(Rec("Prima").Value) & ";" &
			'							SistemaStringaPerRitorno(Rec("Seconda").Value) & ";" &
			'							Rec("Risultato").Value & ";" & Rec("Segno").Value)
			Dim Campi2() As String = Partita.Split(";")
			Dim idPartita2 As Integer = Campi2(0)
			Dim Punti As Integer = 0
			Dim Squadra1 As String = Campi2(1)
			Dim Squadra2 As String = Campi2(2)
			Dim Risultato As String = Campi2(3)
			Dim r2() As String = Risultato.Split("-")
			Dim RisultatoCasa As Integer = r2(0)
			Dim RisultatoFuori As Integer = r2(1)
			Dim RisultatoSegno As String = Campi2(4)
			Dim PartitaTrovata As Boolean = False

			Dim SegnoPreso As Integer = 0
			Dim RisultatoEsattoPartita As Integer = 0
			Dim RisultatoCasaPartita As Integer = 0
			Dim RisultatoFuoriPartita As Integer = 0
			Dim SommaGoalPartita As Integer = 0
			Dim DifferenzaGoalPartita As Integer = 0

			For Each Pronostico As String In Pronostici
				Dim Jolly As Integer = 0
				Dim PuntiPartitaScelta As Integer = 0
				'Pronostici.Add(Rec("idPartita").Value & ";" & Rec("Risultato").Value & ";" & Rec("Segno").Value)
				Dim Campi() As String = Pronostico.Split(";")
				Dim idPartita As Integer = Campi(0)

				If idPartita = idPartita2 Then
					PartitaTrovata = True
					Dim Pronostico2 As String = Campi(1)
					Dim r() As String = Pronostico2.Split("-")
					Dim PronosticoCasa As Integer = r(0)
					Dim PronosticoFuori As Integer = r(1)
					Dim PronosticoSegno As String = Campi(2)

					Punti += 1
					If RisultatoSegno = PronosticoSegno Then
						If idPartita = idPartitaScelta Then
							PuntiPartitaScelta += 5
						End If
						If idPartita = PartitaJolly Then
							Jolly += 3
						End If
						Punti += 5
						SegniPresi += 1
						SegnoPreso = 1
					End If

					If PronosticoCasa = RisultatoCasa And PronosticoFuori = RisultatoFuori Then
						If idPartita = idPartitaScelta Then
							PuntiPartitaScelta += 8
						End If
						If idPartita = PartitaJolly Then
							Jolly += 8
						End If
						Punti += 10
						RisultatoEsatto += 1
						RisultatoCasaTot += 1
						RisultatoFuoriTot += 1
						RisultatoEsattoPartita = 1
						RisultatoCasaPartita = 1
						RisultatoFuoriPartita = 1
					Else
						If PronosticoCasa = RisultatoCasa And PronosticoCasa <> RisultatoFuori Then
							If idPartita = idPartitaScelta Then
								PuntiPartitaScelta += 4
							End If
							If idPartita = PartitaJolly Then
								Jolly += 2
							End If
							Punti += 3
							RisultatoCasaTot += 1
							RisultatoCasaPartita = 1
						Else
							If PronosticoCasa <> RisultatoCasa And PronosticoFuori = RisultatoFuori Then
								If idPartita = idPartitaScelta Then
									PuntiPartitaScelta += 4
								End If
								If idPartita = PartitaJolly Then
									Jolly += 2
								End If
								Punti += 3
								RisultatoFuoriTot += 1
								RisultatoFuoriPartita = 1
							End If
						End If
					End If

					If PronosticoCasa + PronosticoFuori = RisultatoCasa + RisultatoFuori Then
						If idPartita = idPartitaScelta Then
							PuntiPartitaScelta += 2
						End If
						If idPartita = PartitaJolly Then
							Jolly += 1
						End If
						Punti += 2
						SommaGoal += 1
						SommaGoalPartita = 1
					End If

					If Math.Abs(PronosticoCasa - PronosticoFuori) = Math.Abs(RisultatoCasa - RisultatoFuori) Then
						If idPartita = idPartitaScelta Then
							PuntiPartitaScelta += 2
						End If
						If idPartita = PartitaJolly Then
							Jolly += 1
						End If
						Punti += 2
						DifferenzaGoal += 1
						DifferenzaGoalPartita = 1
					End If

					Punti += Jolly
					Punti += PuntiPartitaScelta

					Ritorno &= idPartita & ";" & Squadra1 & ";" & Squadra2 & ";" & Risultato & ";" & RisultatoSegno & ";" & Pronostico2 & ";" & PronosticoSegno & ";" &
						SegnoPreso & ";" & RisultatoEsattoPartita & ";" & RisultatoCasaPartita & ";" & RisultatoFuoriPartita & ";" & SommaGoalPartita & ";" &
						DifferenzaGoalPartita & ";"
					Ritorno &= Punti & ";" & Jolly & ";" & PuntiPartitaScelta & "§"

					PuntiTotali += Punti

					Jolly2 += Jolly
					PuntiPartitaScelta2 += PuntiPartitaScelta
				End If
			Next

			If Not PartitaTrovata Then
				Ritorno &= idPartita2 & ";" & Squadra1 & ";" & Squadra2 & ";" & Risultato & ";" & RisultatoSegno & ";;;;;;;;;0;§"
			End If
		Next
		Ritorno = idUtente & ";" & SistemaStringaPerRitorno(NickName) & ";" & PuntiTotali & "|" & Ritorno

		If Modalita <> "Controllato" Then
			Dim Sql As String = "Delete From Risultati Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
			Dim Ritorno2 As String = Conn.EsegueSql(MP, Sql, Connessione, False)

			Sql = "Insert Into Risultati Values(" & idAnno & ", " & idConcorso & ", " & idUtente & ", " & PuntiTotali & "," &
				" " & SegniPresi & ", " & RisultatoEsatto & ", " & RisultatoCasaTot & ", " & RisultatoFuoriTot & "," &
				" " & SommaGoal & ", " & DifferenzaGoal & ", " & PuntiPartitaScelta2 & ")"
			Ritorno2 = Conn.EsegueSql(MP, Sql, Connessione, False)
			If Not Ritorno2.Contains("ERROR") Then
				Sql = "Select * From RisultatiAltro Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
				Dim Rec As Object = CreaRecordset(MP, Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Sql = "Insert Into RisultatiAltro Values (" & idAnno & ", " & idConcorso & ", " & idUtente & ", 0, 0, " & Jolly2 & ")"
						Ritorno2 = Conn.EsegueSql(MP, Sql, Connessione, False)
					Else
						Sql = "Update RisultatiAltro Set Jolly = " & Jolly2 & " Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
						Ritorno2 = Conn.EsegueSql(MP, Sql, Connessione, False)
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	Public Function RitornaClassificaGenerale(Mp As String, idAnno As Integer, idGiornata As Integer, Conn As Object, Connessione As String, SoloUnaGiornata As Boolean) As String
		Dim Ritorno As String = ""
		Dim Confronto As String = "<="
		If SoloUnaGiornata Then
			Confronto = "="
		End If
		Dim sql As String = "Select * From (" &
			"SELECT A.idUtente, NickName, Sum(Punti) As Punti, Sum(RisultatiEsatti) As RisultatiEsatti, " &
			"Sum(RisultatiCasaTot) As RisCasaTot, Sum(RisultatiFuoriTot) As RisFuoriTot, " &
			"Sum(SegniPresi) As Segni, Sum(SommeGoal) As SommaGoal, Sum(DifferenzeGoal) As DifferenzeGoal, " &
			"(SELECT Count(*) FROM Pronostici Where idUtente = A.idUtente And idPartita = 1 And idConcorso " & Confronto & " A.idConcorso) As Giocate, " &
			"Coalesce(Sum(C.Vittorie),0) As Vittorie, Coalesce(Sum(C.Ultimo),0) As Ultimo, Coalesce(Sum(C.Jolly), 0) As Jolly, " &
			"Coalesce(Sum(A.PuntiPartitaScelta), 0) As PuntiPartitaScelta " &
			"FROM Risultati A Left Join Utenti B On A.idUtente = B.idUtente And A.idAnno = B.idAnno " &
			"Left Join RisultatiAltro C On A.idAnno = C.idAnno And A.idConcorso = C.idConcorso And A.idUtente = C.idUtente " &
			"Where A.idAnno=" & idAnno & " And A.idConcorso " & Confronto & " " & idGiornata & " " &
			"Group By A.idUtente, NickName " &
			"Union ALL " &
			"Select idUtente, NickName, 0 As Punti, 0 As RisultatiEsatti, " &
			"0 As RisCasaTot, 0 As RisFuoriTot, " &
			"0 As Segni, 0 As SommaGoal, 0 As DifferenzeGoal, 0 As Giocate, " &
			"0 As Vittorie,0 As Ultimo, 0 As Jolly, 0 As PuntiPartitaScelta " &
			"From Utenti Where idUtente Not In (Select idUtente From Risultati) " &
			") As A " &
			"Order By 3 Desc, 4 Desc, 7 Desc, 5 Desc, 6 Desc, 8 Desc, 9 Desc, 10 Desc, 12 Desc, 13, 2"
		Dim Rec As Object = CreaRecordset(Mp, Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun utente rilevato"
			Else
				Do Until Rec.Eof
					Ritorno &= Rec("idUtente").Value & ";" & SistemaStringaPerRitorno(Rec("NickName").Value) & ";" & Rec("Punti").Value & ";" &
						Rec("RisultatiEsatti").Value & ";" & Rec("RisCasaTot").Value & ";" & Rec("RisFuoriTot").Value & ";" &
						Rec("Segni").Value & ";" & Rec("SommaGoal").Value & ";" & Rec("DifferenzeGoal").Value & ";" & Rec("Giocate").Value & ";" &
						Rec("Vittorie").Value & ";" & Rec("Ultimo").Value & ";" & Rec("Jolly").Value & ";" & Rec("PuntiPartitaScelta").Value &
						"§"

					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

	Public Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
		' by making Generator static, we preserve the same instance '
		' (i.e., do not create new instances with the same seed over and over) '
		' between calls '
		Static Generator As System.Random = New System.Random()
		Return Generator.Next(Min, Max)
	End Function

	Public Function CreaPartitaJolly(Mp As String, idAnno As Integer, idConcorso As Integer, Conn As Object, Connessione As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String = "Select Coalesce(Count(*),0) From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
		Dim Rec As Object = CreaRecordset(Mp, Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun concorso rilevato"
			Else
				Dim Quante As Integer = Rec(0).Value
				Rec.Close

				Dim x As Integer = GetRandom(1, Quante)

				Sql = "Select * From PartiteJolly Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
				Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Sql = "Insert Into PartiteJolly Values (" &
							" " & idAnno & ", " &
							" " & idConcorso & ", " &
							" " & x & " " &
							")"
						Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
						If Ritorno.Contains(StringaErrore) Then
						End If
					Else
						Rec.Close
					End If
				End If

				'Sql = "Delete From PartiteJolly Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
				'Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
				'If Not Ritorno.Contains(StringaErrore) Then
				'End If
			End If
		End If

		Return Ritorno
	End Function

	Public Function GestisceTorneo23(Mp As String, idAnno As Integer, idConcorso As Integer, Conn As Object, Connessione As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String = "Delete From SquadreRandom Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
		Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
		If Ritorno.Contains(StringaErrore) Then
			Return Ritorno
		End If

		Sql = "Select* From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
		Dim Rec As Object = CreaRecordset(Mp, Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				'Ritorno = "ERROR: Nessun concorso rilevato"
			Else
				Dim Squadre As New List(Of String)

				Do Until Rec.eof
					Squadre.Add(Rec("Prima").Value)
					Squadre.Add(Rec("Seconda").Value)

					Rec.MoveNext
				Loop
				Rec.Close

				Sql = "Select * From Utenti Where idAnno=" & idAnno & " And Eliminato='N'"
				Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						'Ritorno = "ERROR: Nessun concorso rilevato"
					Else
						Dim idGiocatore As New List(Of Integer)

						Do Until Rec.Eof
							idGiocatore.Add(Rec("idUtente").Value)

							Rec.MoveNext
						Loop
						Rec.Close

						Dim Quante As Integer = Squadre.Count - 1

						For Each id As Integer In idGiocatore
							Dim x As Integer = GetRandom(1, Quante)
							Dim Squadra As String = Squadre.Item(x)

							Sql = "Insert Into SquadreRandom Values (" &
								" " & idAnno & ", " &
								" " & idConcorso & ", " &
								" " & id & ", " &
								"'" & SistemaStringaPerDB(squadra) & "', " &
								"0 " &
								")"
							Ritorno = Conn.EsegueSql(Mp, Sql, Connessione, False)
							If Ritorno.Contains(StringaErrore) Then
								Exit For
							End If
						Next
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	Public Sub InvaMailATutti(Mp As String, idAnno As String, Oggetto As String, Testo As String, Conn As Object, Connessione As String, Operazione As String)
		Dim Ritorno As String = ""
		Dim Sql As String = "Select * From Utenti Where idAnno=" & idAnno & " And Eliminato = 'N'"
		If Operazione <> "" Then
			Sql = "Select * From Utenti A " &
				"Left Join UtentiMails B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
				"Where A.idAnno = " & idAnno & " And Eliminato = 'N' And " & Operazione & "='S' "
		Else
			Sql = "Select * From Utenti A " &
				"Left Join UtentiMails B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
				"Where A.idAnno = " & idAnno & " And Eliminato = 'N'"
		End If
		Dim Rec As Object = CreaRecordset(Mp, Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = ""
			Else
				Do Until Rec.Eof
					If Not Ritorno.Contains(Rec("Mail").Value & ";") Then
						Ritorno &= Rec("Mail").Value & ";"
					End If

					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		If Ritorno <> "" Then
			Dim m As New mail(Mp)
			Dim lm() As String = Ritorno.Split(";")
			For Each mm As String In lm
				m.SendEmail(Mp, mm, Oggetto, Testo, {})
			Next
		End If

	End Sub
End Module
