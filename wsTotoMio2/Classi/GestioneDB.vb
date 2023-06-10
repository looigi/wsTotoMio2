Imports System.Reflection
Imports System.Timers
Imports ADOR
Imports MySqlConnector

Public Class clsGestioneDB
	Private Structure LogStruct
		Dim Cosa As String
		Dim Dove As String
	End Structure

	Private mdb As clsMariaDB
	Private TipoDB As String
	Private effettuaLog As Boolean = True
	Private nomeFileLogExec As String = ""
	Private nomeFileLogQuery As String = ""
	Private listaLog As New List(Of LogStruct)
	Private timerLog As Timers.Timer = Nothing

	Public Sub New(Tipo As String)
		Me.TipoDB = Tipo
	End Sub

	Public Function ApreDB(ByVal Connessione As String) As Object
		' Routine che apre il DB e vede se ci sono errori
		Dim Conn As Object
		' Dim TipoDB As String = LeggeTipoDB()

		If TipoDB = "SQLSERVER" Then
			Conn = CreateObject("ADODB.Connection")
			Try
				Conn.Open(Connessione)
				Conn.CommandTimeout = 0
			Catch ex As Exception
				Conn = StringaErrore & " " & ex.Message
			End Try
		Else
			mdb = New clsMariaDB

			Try
				Conn = mdb.apreConnessione(Connessione)
			Catch ex As Exception
				Conn = StringaErrore & " " & ex.Message
			End Try
		End If

		Return Conn
	End Function

	Public Function EsegueSql(MP As String, Sql As String, Connessione As String, Optional ModificaQuery As Boolean = True) As String
		Dim Ritorno As String = "*"
		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
		Dim gf As New GestioneFilesDirectory

		Connessione = Connessione.Replace(vbCrLf, "")
		Connessione = Connessione.Replace("*", ";")
		Connessione = Connessione.Replace("^", "=")

		Dim Errore As String = ""

		Dim Conn As Object = ApreDB(Connessione)
		If TypeOf (Conn) Is String Then
			Errore = Conn
		End If

		Dim Sql2 As String = ""

		If ModificaQuery Then
			If TipoDB = "SQLSERVER" Then
				Sql2 = Sql
			Else
				Sql2 = ConverteStringaPerLinux(Sql)
			End If
		Else
			Sql2 = Sql
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			nomeFileLogExec = MP & "\Logs\Exec_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"

			ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogExec)
			ThreadScriveLog(Datella & ": Esecuzione SQL", nomeFileLogExec)
			ThreadScriveLog(Datella & ": Tipo db: " & TipoDB, nomeFileLogExec)
			ThreadScriveLog(Datella & ": Connessione: " & Connessione, nomeFileLogExec)
			ThreadScriveLog(Datella & ": SQL = " & Sql2, nomeFileLogExec)
			If Errore <> "" Then
				ThreadScriveLog(Datella & ": " & Errore, nomeFileLogExec)
			End If
			' End If
		End If

		If Errore = "" Then
			' Routine che esegue una query sul db
			If TipoDB = "SQLSERVER" Then
				Try
					Conn.Execute(Sql2)
					If effettuaLog Then
						ThreadScriveLog(Datella & ": OK", nomeFileLogExec)
					End If
				Catch ex As Exception
					If effettuaLog Then
						ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message, nomeFileLogExec)
					End If
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			Else
				Try
					Ritorno = mdb.EsegueSql(Sql2, ModificaQuery)
					If Ritorno.ToUpper.Trim <> "OK" Then
						Ritorno = StringaErrore & " " & Ritorno
					End If
					If effettuaLog Then
						ThreadScriveLog(Datella & ": " & Ritorno, nomeFileLogQuery)
					End If
				Catch ex As Exception
					If effettuaLog Then
						ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message, nomeFileLogExec)
					End If
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If

			Try
				ChiudeDB(Conn)
			Catch ex As Exception
				Ritorno = StringaErrore & " " & ex.Message
			End Try
		Else
			Ritorno = Errore
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogExec)
			ThreadScriveLog("", nomeFileLogExec)
		End If

		Return Ritorno
	End Function

	Public Sub Close()

	End Sub

	Public Function LeggeQuery(MP As String, Sql As String, Connessione As String, Optional ModificaQuery As Boolean = True) As Object
		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
		Dim gf As New GestioneFilesDirectory
		' Dim TipoDB As String = LeggeTipoDB()

		Connessione = Connessione.Replace(vbCrLf, "")
		Connessione = Connessione.Replace("*", ";")
		Connessione = Connessione.Replace("^", "=")

		Dim Errore As String = ""

		Dim Conn As Object = ApreDB(Connessione)
		If TypeOf (Conn) Is String Then
			Errore = Conn
		End If

		Dim Sql2 As String = ""

		If ModificaQuery Then
			If TipoDB = "SQLSERVER" Then
				Sql2 = Sql
			Else
				Sql2 = Sql ' ConverteStringaPerLinux(Sql)
			End If
		Else
			Sql2 = Sql
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			nomeFileLogQuery = MP & "\Logs\Query_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"

			ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogQuery)
			ThreadScriveLog(Datella & ": Lettura Query", nomeFileLogQuery)
			ThreadScriveLog(Datella & ": Tipo db: " & TipoDB, nomeFileLogQuery)
			ThreadScriveLog(Datella & ": Connessione: " & Connessione, nomeFileLogQuery)
			ThreadScriveLog(Datella & ": SQL = " & Sql2, nomeFileLogQuery)
			If Errore <> "" Then
				ThreadScriveLog(Datella & ": ERROR: " & Errore, nomeFileLogQuery)
			End If
			'End If
		End If

		'Return "Lettura " & Indice & " -> " & mdb.Length

		Dim Rec As Object

		If Errore = "" Then
			If TipoDB = "SQLSERVER" Then
				Rec = New Recordset

				Try
					Rec.Open(Sql2, Conn)
				Catch ex As Exception
					Rec = StringaErrore & " " & ex.Message
					If effettuaLog Then
						ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message, nomeFileLogQuery)
					End If
				End Try
			Else
				Try
					Rec = mdb.Lettura(Sql2, ModificaQuery)
					If TypeOf (Rec) Is String Then
						If effettuaLog Then
							ThreadScriveLog(Datella & ": ERRORE SQL -> " & Rec, nomeFileLogQuery)
						End If
					End If
				Catch ex As Exception
					If effettuaLog Then
						ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message, nomeFileLogQuery)
					End If
					Rec = StringaErrore & " " & ex.Message
				End Try
			End If

			Try
				ChiudeDB(Conn)
			Catch ex As Exception
				Rec = StringaErrore & " " & ex.Message
			End Try
		Else
			Rec = Errore
		End If

		If effettuaLog And Not HttpContext.Current Is Nothing Then
			ThreadScriveLog(Datella & "--------------------------------------------------------------------------", nomeFileLogQuery)
			ThreadScriveLog("", nomeFileLogQuery)
		End If

		Return Rec
	End Function

	Private Function ConverteStringaPerLinux(Sql As String) As String
		Dim Sql2 As String = Sql

		If Sql2.ToUpper.Trim.StartsWith("INSERT INTO ") Then
			Dim a As Integer = Sql2.ToUpper.IndexOf(" VALUES")

			If a = 0 Then
				a = Sql2.ToUpper.IndexOf(" SELECT")
			End If
			If a > 0 Then
				Dim inizio As String = Mid(Sql2, 1, a)
				Dim modificato As String = inizio.ToLower
				Sql2 = Sql2.Replace(inizio, modificato)
			End If
		Else
			If Sql2.ToUpper.Trim.StartsWith("UPDATE ") Then
				Dim a As Integer = Sql2.ToUpper.IndexOf(" SET ")

				If a > 0 Then
					Dim inizio As String = Mid(Sql2, 1, a)
					Dim modificato As String = inizio.ToLower
					Sql2 = Sql2.Replace(inizio, modificato)
				End If
			Else
				Sql2 = Sql2.ToLower()
			End If
		End If

		'Sql2 = Sql2.Replace("[", "")
		'Sql2 = Sql2.Replace("]", "")
		Sql2 = Sql2.Replace("dbo.", "")

		'Sql2 = Sql2.Replace("generale", "Generale")

		Return Sql2
	End Function

	'Private Function ControllaAperturaConnessione(ByRef Conn As Object, ByVal Connessione As String, Indice As Integer) As Boolean
	'	Dim Ritorno As Boolean = False

	'	If Conn Is Nothing Then
	'		If TipoDB = "SQLSERVER" Then
	'			Ritorno = True
	'			Conn = ApreDB(Connessione, Indice)
	'		Else
	'			Ritorno = True
	'			Conn = ApreDB(Connessione, Indice)
	'		End If
	'	End If

	'	Return Ritorno
	'End Function

	Public Sub ChiudeDB(Conn As Object)
		If TipoDB = "SQLSERVER" Then
			Conn.Close()
		Else
			mdb.ChiudiConn(Conn)
		End If
	End Sub

	Private Sub ThreadScriveLog(s As String, dove As String)
		Dim e As New LogStruct
		e.Cosa = s
		e.Dove = dove
		listaLog.Add(e)

		avviaTimerLog()
	End Sub

	Private Sub avviaTimerLog()
		If timerLog Is Nothing Then
			timerLog = New Timer(100)
			AddHandler timerLog.Elapsed, New ElapsedEventHandler(AddressOf scodaLog)
			timerLog.Start()
		End If
	End Sub

	Private Sub scodaLog()
		timerLog.Enabled = False
		Dim ls As LogStruct = listaLog.Item(0)
		Dim Dove As String = ls.Dove
		Dim sLog As String = ls.Cosa

		Dim gf As New GestioneFilesDirectory
		gf.ApreFileDiTestoPerScrittura(Dove)
		gf.ScriveTestoSuFileAperto(sLog)
		gf.ChiudeFileDiTestoDopoScrittura()

		listaLog.RemoveAt(0)
		If listaLog.Count > 0 Then
			timerLog.Enabled = True
		Else
			timerLog = Nothing
			listaLog = New List(Of LogStruct)
		End If
	End Sub
End Class
