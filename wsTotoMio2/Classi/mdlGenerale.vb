﻿Imports System.Drawing.Printing
Imports System.Web.Configuration
Imports Microsoft.SqlServer
Imports wsTotoMio2.clsRecordset

Module mdlGenerale
	Public Structure strutturaMail
		Dim Destinatario As String
		Dim Oggetto As String
		Dim newBody As String
		Dim Allegato() As String
	End Structure

	Public Structure SquadrePrese
		Dim Squadra As String
		Dim Quante As Integer
		Dim Totale As Integer
	End Structure

	Public Structure StrutturaSorprese
		Dim idPartita As Integer
		Dim Casa As String
		Dim Fuori As String
		Dim Segno As String
		Dim Percentuale As Integer
	End Structure

	Public Structure StrutturaPosizioni
		Dim idUtente As Integer
		Dim Posizione As Integer
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
	Public SceltiPerCreazione As String = ""
	Public GiocataPartita As String = ""

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
								   Conn As Object, Connessione As String, MP As String, Modalita As String, PartitaJolly As Integer, idPartitaScelta As Integer,
								   ListaSorprese As List(Of StrutturaSorprese), SoloControllo As String) As String
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
		Dim PuntiSorpresa2 As Integer = 0

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
			Dim RisultatoCasa As Integer = -1
			Dim RisultatoFuori As Integer = -1
			If Risultato <> "" And Risultato.Contains("-") Then
				Dim r2() As String = Risultato.Split("-")
				RisultatoCasa = r2(0)
				RisultatoFuori = r2(1)
			End If
			Dim RisultatoSegno As String = Campi2(4)
			Dim PartitaSospesa As Boolean = IIf(Campi2(5) = "S", True, False)
			Dim PartitaTrovata As Boolean = False

			Dim SegnoPreso As Integer = 0
			Dim RisultatoEsattoPartita As Integer = 0
			Dim RisultatoCasaPartita As Integer = 0
			Dim RisultatoFuoriPartita As Integer = 0
			Dim SommaGoalPartita As Integer = 0
			Dim DifferenzaGoalPartita As Integer = 0

			Dim PartitaSorpresa As Boolean = False
			Dim SegniSorpresa As String = ""
			For Each Sorpresa As StrutturaSorprese In ListaSorprese
				Dim idSorpresa As Integer = Sorpresa.idPartita
				If idSorpresa = idPartita2 Then
					If PartitaSorpresa = False Then
						PartitaSorpresa = True
					End If
					SegniSorpresa &= Sorpresa.Segno & ";"
				End If
			Next
			Dim PuntiSorpresa As Integer = 0

			For Each Pronostico As String In Pronostici
				Dim Jolly As Integer = 0
				Dim PuntiPartitaScelta As Integer = 0
				'Pronostici.Add(Rec("idPartita").Value & ";" & Rec("Risultato").Value & ";" & Rec("Segno").Value)
				Dim Campi() As String = Pronostico.Split(";")
				Dim idPartita As Integer = Campi(0)

				If PartitaSospesa = False Then
					If idPartita = idPartita2 Then
						PartitaTrovata = True
						Dim Pronostico2 As String = Campi(1)
						Dim r() As String = Pronostico2.Split("-")
						Dim PronosticoCasa As Integer = r(0)
						Dim PronosticoFuori As Integer = r(1)
						Dim PronosticoSegno As String = Campi(2)

						If PartitaSorpresa Then
							If SegniSorpresa.Contains(PronosticoSegno & ";") Then
								If RisultatoSegno = PronosticoSegno Then
									Punti += 15
									PuntiSorpresa += 15
								End If
							End If
						End If

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
						Ritorno &= Punti & ";" & Jolly & ";" & PuntiPartitaScelta & ";" & PuntiSorpresa & "§"

						PuntiTotali += Punti

						Jolly2 += Jolly
						PuntiPartitaScelta2 += PuntiPartitaScelta
						PuntiSorpresa2 += PuntiSorpresa
					End If
				End If
			Next

			If Not PartitaTrovata Then
				Ritorno &= idPartita2 & ";" & Squadra1 & ";" & Squadra2 & ";" & Risultato & ";" & RisultatoSegno & ";;;;;;;;;0;§"
			End If
		Next
		Ritorno = idUtente & ";" & SistemaStringaPerRitorno(NickName) & ";" & PuntiTotali & "|" & Ritorno

		If Modalita <> "Controllato" Then
			Dim Altro As String = ""
			If SoloControllo = "SI" Then
				Altro = "Sim"
			End If

			Dim Sql As String = ""
			If SoloControllo = "SI" Then
				Sql = "Delete From Risultati" & Altro
			Else
				Sql = "Delete From Risultati Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
			End If
			Dim Ritorno2 As String = Conn.EsegueSql(MP, Sql, Connessione, False)

			Sql = "Insert Into Risultati" & Altro & " Values(" & idAnno & ", " & idConcorso & ", " & idUtente & ", " & PuntiTotali & "," &
					" " & SegniPresi & ", " & RisultatoEsatto & ", " & RisultatoCasaTot & ", " & RisultatoFuoriTot & "," &
					" " & SommaGoal & ", " & DifferenzaGoal & ", " & PuntiPartitaScelta2 & ", " & PuntiSorpresa2 & ")"
			Ritorno2 = Conn.EsegueSql(MP, Sql, Connessione, False)
			If Not Ritorno2.Contains("ERROR") Then
				Sql = "Select * From RisultatiAltro" & Altro & " Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
				Dim Rec As Object = CreaRecordset(MP, Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Sql = "Insert Into RisultatiAltro" & Altro & " Values (" & idAnno & ", " & idConcorso & ", " & idUtente & ", 0, 0, " & Jolly2 & ")"
						Ritorno2 = Conn.EsegueSql(MP, Sql, Connessione, False)
					Else
						Sql = "Update RisultatiAltro" & Altro & " Set Jolly = " & Jolly2 & " Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
						Ritorno2 = Conn.EsegueSql(MP, Sql, Connessione, False)
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	Public Function RitornaClassificaGenerale(Mp As String, idAnno As Integer, idGiornata As Integer, Conn As Object, Connessione As String, SoloUnaGiornata As Boolean,
											  MostraFinto As String, Simulazione As String) As String
		Dim Ritorno As String = ""

		Dim Confronto As String = "<="
		If SoloUnaGiornata Then
			Confronto = "="
		End If

		Dim Altro As String = ""
		If MostraFinto = "N" Then
			Altro = "Where idTipologia <> 2 "
		End If

		Dim Altro2 As String = ""
		If Simulazione = "SI" Then
			Altro2 = "Sim"
		End If

		Dim sql As String = "Select * From (" &
			"SELECT A.idUtente, NickName, Sum(Punti) As Punti, Sum(RisultatiEsatti) As RisultatiEsatti, " &
			"Sum(RisultatiCasaTot) As RisCasaTot, Sum(RisultatiFuoriTot) As RisFuoriTot, " &
			"Sum(SegniPresi) As Segni, Sum(SommeGoal) As SommaGoal, Sum(DifferenzeGoal) As DifferenzeGoal, " &
			"(SELECT Count(*) FROM Pronostici Where idAnno = A.idAnno And idUtente = A.idUtente And idPartita = 1 And idConcorso " & Confronto & " A.idConcorso) As Giocate, " &
			"Coalesce(Sum(C.Vittorie),0) As Vittorie, Coalesce(Sum(C.Ultimo),0) As Ultimo, Coalesce(Sum(C.Jolly), 0) As Jolly, " &
			"Coalesce(Sum(A.PuntiPartitaScelta), 0) As PuntiPartitaScelta, B.idTipologia, Coalesce(Sum(A.PuntiSorpresa), 0) As PuntiSorpresa, " &
			"(Select Posizione From PosizioniClassifica Where idAnno=1 And idUtente=A.idUtente And idConcorso=2) As PosAttuale, " &
			"(Select Posizione From PosizioniClassifica Where idAnno=1 And idUtente=A.idUtente And idConcorso=1) As PosPrecedente " &
			"FROM Risultati" & Altro2 & " A " &
			"Left Join Utenti B On A.idUtente = B.idUtente And A.idAnno = B.idAnno " &
			"Left Join RisultatiAltro" & Altro2 & " C On A.idAnno = C.idAnno And A.idConcorso = C.idConcorso And A.idUtente = C.idUtente " &
			"Where A.idAnno=" & idAnno & " And A.idConcorso " & Confronto & " " & idGiornata & " " &
			"Group By A.idUtente, NickName, B.idTipologia " &
			"Union ALL " &
			"Select idUtente, NickName, 0 As Punti, 0 As RisultatiEsatti, " &
			"0 As RisCasaTot, 0 As RisFuoriTot, " &
			"0 As Segni, 0 As SommaGoal, 0 As DifferenzeGoal, 0 As Giocate, " &
			"0 As Vittorie,0 As Ultimo, 0 As Jolly, 0 As PuntiPartitaScelta, idTipologia, 0 As PuntiSorpresa, " &
			"0 As PosAttuale, 0 As PosPrecedente " &
			"From Utenti Where idUtente Not In (Select idUtente From Risultati" & Altro2 & ") " &
			") As A " &
			" " & Altro & " " &
			"Order By 3 Desc, 4 Desc, 7 Desc, 5 Desc, 6 Desc, 8 Desc, 9 Desc, 10 Desc, 12 Desc, 13 Desc, 14 Desc, 16 Desc, 2, idTipologia"
		Dim Rec As Object = CreaRecordset(Mp, Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun utente rilevato"
			Else
				Dim Posizioni As New List(Of StrutturaPosizioni)
				Dim Pos As Integer = 1
				Dim VecchiPunti As Integer = -1

				Do Until Rec.Eof
					Ritorno &= Rec("idUtente").Value & ";" & SistemaStringaPerRitorno(Rec("NickName").Value) & ";" & Rec("Punti").Value & ";" &
						Rec("RisultatiEsatti").Value & ";" & Rec("RisCasaTot").Value & ";" & Rec("RisFuoriTot").Value & ";" &
						Rec("Segni").Value & ";" & Rec("SommaGoal").Value & ";" & Rec("DifferenzeGoal").Value & ";" & Rec("Giocate").Value & ";" &
						Rec("Vittorie").Value & ";" & Rec("Ultimo").Value & ";" & Rec("Jolly").Value & ";" & Rec("PuntiPartitaScelta").Value & ";" &
						Rec("PuntiSorpresa").Value & ";" & Rec("PosAttuale").Value & ";" & Rec("PosPrecedente").Value &
						"§"

					Dim s As New StrutturaPosizioni
					s.idUtente = Rec("idUtente").Value
					s.Posizione = Pos
					Posizioni.Add(s)

					If VecchiPunti <> Rec("Punti").Value Then
						VecchiPunti = Rec("Punti").Value
						Pos += 1
					End If

					Rec.MoveNext
				Loop
				Rec.Close

				If Not SoloUnaGiornata And MostraFinto = "S" And Simulazione <> "SI" Then
					sql = "Select * From PosizioniClassifica Where idAnno=" & idAnno & " And idConcorso=" & idGiornata
					Rec = CreaRecordset(Mp, Conn, sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							For Each s As StrutturaPosizioni In Posizioni
								sql = "Insert Into PosizioniClassifica Values(" & idAnno & ", " & idGiornata & ", " & s.idUtente & ", " & s.Posizione & ")"
								Dim Ritorno2 As String = Conn.EsegueSql(Mp, sql, Connessione, False)
							Next
						Else
						End If
					End If
				End If
			End If

		End If

		Return Ritorno
	End Function

	Public Function PrendeSorprese(Mp As String, Conn As Object, Connessione As String, idAnno As String, idGiornata As String) As List(Of StrutturaSorprese)
		Dim Lista As New List(Of StrutturaSorprese)
		Dim Sql As String = "Select * From (Select idPartita, Prima, Seconda, Segno, Quanti, TotalePartite, Round((Quanti / TotalePartite) * 100) As Media From ( " &
			"Select *, (Select Count(*) From Pronostici Where idAnno= " & idAnno & " And idConcorso = " & idGiornata & " And idPartita=1) As TotalePartite From ( " &
			"Select A.idPartita, B.Prima, B.Seconda, A.Segno, Count(*) As Quanti From Pronostici As A " &
			"Left Join Concorsi B On A.idAnno = B.idAnno And A.idConcorso = B.idConcorso And A.idPartita = B.idPartita " &
			"Where A.idAnno = " & idAnno & " And A.idConcorso=" & idGiornata & " " &
			"Group By A.idPartita, A.Segno " &
			") As A " &
			") As B " &
			") As C Where Media < 10"
		Dim Rec As Object = CreaRecordset(Mp, Conn, Sql, Connessione)
		Do Until Rec.Eof
			Dim s As New StrutturaSorprese
			s.idPartita = Rec("idPartita").Value
			s.Casa = Rec("Prima").Value
			s.Fuori = Rec("Seconda").Value
			s.Segno = Rec("Segno").Value
			s.Percentuale = Rec("Media").Value
			Lista.Add(s)

			Rec.MoveNext
		Loop
		Rec.Close

		Return Lista
	End Function

	Public Function CreaColonnaUtenteFinto(Mp As String, idAnno As String, idGiornata As String, Conn As Object, Connessione As String) As String
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From Utenti Where idTipologia=2 And idAnno=" & idAnno
		Dim Rec As Object = CreaRecordset(Mp, Conn, sql, Connessione)

		If TypeOf (Rec) Is String Then
		Else
			If Rec.Eof Then
				' Ritorno = "ERROR: Nessun fintone rilevato"
			Else
				Dim ids As New List(Of String)

				Do Until Rec.Eof
					ids.Add(Rec("idUtente").Value)

					Rec.MoveNext
				Loop
				Rec.Close

				sql = "Select Coalesce(Count(*), 0) As Quante From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idGiornata
				Rec = CreaRecordset(Mp, Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
				Else
					If Rec.Eof Then
						' Ritorno = "ERROR: Nessun concorso rilevato"
					Else
						Dim Quante As Integer = Rec("Quante").Value
						Rec.Close

						For Each id As String In ids
							Dim Pron As New List(Of String)

							For i As Integer = 1 To Quante
								sql = "Select Pronostico, Segno From Pronostici Where idAnno=" & idAnno & " And idConcorso=" & idGiornata & " And idPartita=" & i
								Rec = CreaRecordset(Mp, Conn, sql, Connessione)
								If TypeOf (Rec) Is String Then
								Else
									If Rec.Eof Then
										' Ritorno = "ERROR: Nessun pronostico rilevato"
									Else
										Dim q As Integer = 0
										Dim TotC As Integer = 0
										Dim TotF As Integer = 0
										Dim Segni1 As Integer = 0
										Dim SegniX As Integer = 0
										Dim Segni2 As Integer = 0

										Do Until Rec.Eof
											Dim Pronostico As String = Rec("Pronostico").Value

											If Pronostico <> "" And Pronostico.Contains("-") Then
												Dim P() As String = Pronostico.Split("-")
												Dim Casa As Integer = Val(P(0))
												Dim Fuori As Integer = Val(P(1))

												Select Case Rec("Segno").Value
													Case "1"
														Segni1 += 1
													Case "X"
														SegniX += 1
													Case "2"
														Segni2 += 1
												End Select

												TotC += Casa
												TotF += Fuori
												q += 1
											End If

											Rec.MoveNext
										Loop
										Rec.Close

										If q > 0 Then
											Dim MediaC As Integer = Math.Floor(TotC / q)
											Dim MediaF As Integer = Math.Floor(TotF / q)
											Dim Segno As String = ""

											If Segni1 > SegniX And Segni1 > Segni2 Then
												Segno = "1"
											Else
												If SegniX > Segni1 And SegniX > Segni2 Then
													Segno = "X"
												Else
													If Segni2 > Segni1 And Segni2 > SegniX Then
														Segno = "2"
													Else
														If MediaC > MediaF Then
															Segno = "1"
														Else
															If MediaC < MediaF Then
																Segno = "2"
															Else
																Segno = "X"
															End If
														End If
													End If
												End If
											End If

											Pron.Add(MediaC & "-" & MediaF & ";" & Segno)
										End If
									End If
								End If
							Next

							Dim idPartita As Integer = 1

							For Each p As String In Pron
								Dim PP() As String = p.Split(";")

								sql = "Insert Into Pronostici Values (" &
									" " & idAnno & ", " &
									" " & id & ", " &
									" " & idGiornata & ", " &
									" " & idPartita & ", " &
									"'" & PP(0) & "', " &
									"'" & PP(1) & "' " &
									")"
								Dim Rit As String = Conn.EsegueSql(Mp, sql, Connessione, False)
								If Rit.Contains("ERROR") Then
									Ritorno = Rit
									Exit For
								End If

								idPartita += 1
							Next
						Next
					End If
				End If
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
		Dim Sql As String = "Select Coalesce(Count(*), 0) As Quante From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
		Dim Rec As Object = CreaRecordset(Mp, Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun concorso rilevato"
			Else
				Dim Quante As Integer = Rec("Quante").Value
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
								"'" & SistemaStringaPerDB(Squadra) & "', " &
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
				"Where A.idAnno = " & idAnno & " And Eliminato = 'N' And " & Operazione & "='S' And idTipologia<>2 "
		Else
			Sql = "Select * From Utenti A " &
				"Left Join UtentiMails B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
				"Where A.idAnno = " & idAnno & " And Eliminato = 'N' And idTipologia<>2"
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

	Public Function SistemaNumeroDaDB(Numero As Object, Decimale As Boolean) As String
		Dim N As String = Numero

		If Decimale Then
			N = CInt(Val(N) * 100) / 100
		End If
		N = N.ToString.Replace(",", ".")
		N = N.Replace(".0000", "")

		Return N
	End Function

	Public Function GeneraSquadrePrese(Mp As String, idAnno As String, Conn As Object, Connessione As String) As String
		Dim Sql As String = ""
		Dim Rec As Object
		Dim Ritorno As String = "["

		Sql = "Select * From Utenti " &
			IIf(idAnno <> "", "Where idAnno=" & idAnno & " And Eliminato = 'N' And idTIpologia <> 2", "Where Eliminato = 'N' And idTIpologia <> 2")
		Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Dim Utenti As New List(Of String)
			Dim NickName As New List(Of String)

			Do Until Rec.Eof
				Utenti.Add(Rec("idUtente").Value)
				NickName.Add(Rec("NickName").Value)

				Rec.MoveNext
			Loop
			Rec.Close

			Dim Altro As String = ""
			If idAnno <> "" Then Altro = " A.idAnno = 1 And "

			Dim conta2 As Integer = 0
			Dim Riga2 As Boolean = True
			For Each id As String In Utenti
				Sql = "Select Squadra, Sum(Quante) As Prese, (Select Count(*) From Concorsi Where " & Altro.Replace("A.", "") & " (Prima=Squadra Or Seconda=Squadra) Group By Squadra) As Tot " &
					"From ( " &
					"SELECT Prima As Squadra, Count(*) As Quante FROM Pronostici As A " &
					"Left Join Concorsi B On A.idAnno = B.idAnno And A.idConcorso = B.idConcorso And A.idPartita = B.idPartita " &
					"Where " & Altro & " A.idUtente = 1 And A.Segno = B.Segno " &
					"Group By Prima " &
					"Union ALL " &
					"Select Seconda As Squadra, Count(*) As Quante FROM Pronostici As A " &
					"Left Join Concorsi B On A.idAnno = B.idAnno And A.idConcorso = B.idConcorso And A.idPartita = B.idPartita " &
					"Where " & Altro & " A.idUtente = 1 And A.Segno = B.Segno " &
					"Group By Prima " &
					"Union ALL " &
					"Select Prima As Squadra, 0 As Quante FROM Pronostici As A " &
					"Left Join Concorsi B On A.idAnno = B.idAnno And A.idConcorso = B.idConcorso And A.idPartita = B.idPartita " &
					"Where " & Altro & " A.idUtente = 1 And A.Segno <> B.Segno " &
					"Group By Prima " &
					"Union ALL " &
					"Select Seconda As Squadra, 0 As Quante FROM Pronostici As A " &
					"Left Join Concorsi B On A.idAnno = B.idAnno And A.idConcorso = B.idConcorso And A.idPartita = B.idPartita " &
					"Where " & Altro & " A.idUtente = 1 And A.Segno <> B.Segno " &
					"Group By Prima " &
					") As A " &
					"Group By Squadra " &
					"Order By 2 Desc, 3 "
				Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Dim Prese As New List(Of SquadrePrese)

					Do Until Rec.Eof
						Dim p As New SquadrePrese
						p.Squadra = Rec("Squadra").Value
						p.Quante = Rec("Prese").Value
						p.Totale = Rec("Tot").Value
						Prese.Add(p)

						Rec.MoveNext
					Loop
					Rec.Close

					Ritorno &= "{"

					Ritorno &= "" & Chr(34) & "idUtente" & Chr(34) & ": " & id & ", "
					Ritorno &= "" & Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & NickName.Item(conta2) & Chr(34) & ", "
					Ritorno &= "" & Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & Riga2 & Chr(34) & ", "
					Ritorno &= "" & Chr(34) & "Prese" & Chr(34) & ": ["
					Riga2 = Not Riga2

					Dim Ritorno2 As String = ""
					Dim Riga As Boolean = True

					For Each s As SquadrePrese In Prese
						Ritorno2 &= "{"
						Ritorno2 &= "" & Chr(34) & "Squadra" & Chr(34) & ": " & Chr(34) & s.Squadra & Chr(34) & ", "
						Ritorno2 &= "" & Chr(34) & "Prese" & Chr(34) & ": " & s.Quante & ", "
						Ritorno2 &= "" & Chr(34) & "Totale" & Chr(34) & ": " & s.Totale & ", "
						Ritorno2 &= "" & Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & Riga & Chr(34) & ", "
						Ritorno2 &= "" & Chr(34) & "Visibile" & Chr(34) & ": " & Chr(34) & "False" & Chr(34) & " "
						Ritorno2 &= "}, "
						Riga = Not Riga
					Next

					If Ritorno2.Length > 0 Then
						Ritorno2 = Mid(Ritorno2, 1, Ritorno2.Length - 2)
					End If
					Ritorno &= Ritorno2

					Ritorno &= "]"
					Ritorno &= "},"

					conta2 += 1
				End If
			Next
		End If
		If Ritorno.Length > 0 Then
			Ritorno = Mid(Ritorno, 1, Ritorno.Length - 1)
		End If
		Ritorno &= "]"

		Return Ritorno
	End Function
End Module
