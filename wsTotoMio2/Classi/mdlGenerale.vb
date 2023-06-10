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
End Module
