﻿Imports Microsoft.SqlServer
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
								   Conn As Object, Connessione As String, MP As String, Modalita As String) As String
		Dim PuntiTotali As Integer = 0
		Dim Ritorno As String = ""
		Dim SegniPresi As Integer = 0
		Dim RisultatoEsatto As Integer = 0
		Dim RisultatoCasaTot As Integer = 0
		Dim RisultatoFuoriTot As Integer = 0
		Dim SommaGoal As Integer = 0
		Dim DifferenzaGoal As Integer = 0

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
						Punti += 5
						SegniPresi += 1
						SegnoPreso = 1
					End If

					If PronosticoCasa = RisultatoCasa And PronosticoFuori = RisultatoFuori Then
						Punti += 10
						RisultatoEsatto += 1
						RisultatoCasaTot += 1
						RisultatoFuoriTot += 1
						RisultatoEsattoPartita = 1
						RisultatoCasaPartita = 1
						RisultatoFuoriPartita = 1
					Else
						If PronosticoCasa = RisultatoCasa And PronosticoCasa <> RisultatoFuori Then
							Punti += 3
							RisultatoCasaTot += 1
							RisultatoCasaPartita = 1
						Else
							If PronosticoCasa <> RisultatoCasa And PronosticoFuori = RisultatoFuori Then
								Punti += 3
								RisultatoFuoriTot += 1
								RisultatoFuoriPartita = 1
							End If
						End If
					End If

					If PronosticoCasa + PronosticoFuori = RisultatoCasa + RisultatoFuori Then
						Punti += 2
						SommaGoal += 1
						SommaGoalPartita = 1
					End If

					If Math.Abs(PronosticoCasa - PronosticoFuori) = Math.Abs(RisultatoCasa - RisultatoFuori) Then
						Punti += 2
						DifferenzaGoal += 1
						DifferenzaGoalPartita = 1
					End If

					Ritorno &= idPartita & ";" & Squadra1 & ";" & Squadra2 & ";" & Risultato & ";" & RisultatoSegno & ";" & Pronostico2 & ";" & PronosticoSegno & ";" &
						SegnoPreso & ";" & RisultatoEsattoPartita & ";" & RisultatoCasaPartita & ";" & RisultatoFuoriPartita & ";" & SommaGoalPartita & ";" &
						DifferenzaGoalPartita & ";"
					Ritorno &= Punti & "§"

					PuntiTotali += Punti
				End If
			Next

			If Not PartitaTrovata Then
				Ritorno &= idPartita2 & ";" & Squadra1 & ";" & Squadra2 & ";" & Risultato & ";" & RisultatoSegno & ";;;;;;;;;0§"
			End If
		Next
		Ritorno = idUtente & ";" & SistemaStringaPerRitorno(NickName) & ";" & PuntiTotali & "|" & Ritorno

		If Modalita <> "Controllato" Then
			Dim Sql As String = "Delete From Risultati Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " And idUtente=" & idUtente
			Dim Ritorno2 As String = Conn.EsegueSql(MP, Sql, Connessione, False)

			Sql = "Insert Into Risultati Values(" & idAnno & ", " & idConcorso & ", " & idUtente & ", " & PuntiTotali & "," &
				" " & SegniPresi & ", " & RisultatoEsatto & ", " & RisultatoCasaTot & ", " & RisultatoFuoriTot & "," &
				" " & SommaGoal & ", " & DifferenzaGoal & ")"
			Ritorno2 = Conn.EsegueSql(MP, Sql, Connessione, False)
		End If

		Return Ritorno
	End Function

	Public Function RitornaClassificaGenerale(Mp As String, idAnno As Integer, idGiornata As Integer, Conn As Object, Connessione As String) As String
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From (" &
			"SELECT A.idUtente, NickName, Sum(Punti) As Punti, Sum(RisultatiEsatti) As RisultatiEsatti, " &
			"Sum(RisultatiCasaTot) As RisCasaTot, Sum(RisultatiFuoriTot) As RisFuoriTot, " &
			"Sum(SegniPresi) As Segni, Sum(SommeGoal) As SommaGoal, Sum(DifferenzeGoal) As DifferenzeGoal " &
			"FROM Risultati A Left Join Utenti B On A.idUtente = B.idUtente And A.idAnno = B.idAnno " &
			"Where A.idAnno=" & idAnno & " And idConcorso <= " & idGiornata & " " &
			"Group By A.idUtente, NickName " &
			"Union ALL " &
			"Select idUtente, NickName, 0 As Punti, 0 As RisultatiEsatti, " &
			"0 As RisCasaTot, 0 As RisFuoriTot, " &
			"0 As Segni, 0 As SommaGoal, 0 As DifferenzeGoal " &
			"From Utenti Where idUtente Not In (Select idUtente From Risultati) " &
			") As A " &
			"Order By 3 Desc, 4 Desc, 7 Desc, 5 Desc, 6 Desc, 8 Desc, 9 Desc, 2"
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
						Rec("Segni").Value & ";" & Rec("SommaGoal").Value & ";" & Rec("DifferenzeGoal").Value & "§"

					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function
End Module
