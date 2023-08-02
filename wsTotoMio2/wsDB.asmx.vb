Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Claims
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Windows.Forms

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://dbTotoMIO.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsDB
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaBackups() As String
		Dim Cartella As String = LeggePathBackup()
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim Barra As String = "/"

		For Each Dir As String In Directory.GetDirectories(Cartella & Barra)
			Dir = gf.TornaNomeFileDaPath(Dir)
			Ritorno &= Dir & "§"
		Next

		Return Ritorno
	End Function

	Private Function LeggePathBackup() As String
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = gf.LeggeFileIntero(Server.MapPath(".") & "/PathBackup.txt")
		Ritorno = Ritorno.Replace(vbCrLf, "")
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EffettuaBackup() As String
		Dim Cartella As String = LeggePathBackup()
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = "OK"
		Dim Barra As String = "/"
		Dim Sql As String = "Show Tables"
		Dim gf As New GestioneFilesDirectory
		Dim QualeBackup As String = "1;" & Now.Year & "-" & Format(Now.Month, "00") & "-" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & "-" & Format(Now.Minute, "00") & "-" & Format(Now.Second, "00")
		Dim Max As Integer = 0

		For Each Dir As String In Directory.GetDirectories(Cartella & Barra)
			Dir = gf.TornaNomeFileDaPath(Dir)
			Dim c() As String = Dir.Split(";")
			If Val(c(0)) > Max Then
				Max = Val(c(0))
			End If
		Next
		If Max > 0 Then
			QualeBackup = (Max + 1).ToString.Trim & ";" & Now.Year & "-" & Format(Now.Month, "00") & "-" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & "-" & Format(Now.Minute, "00") & "-" & Format(Now.Second, "00")
		End If

		gf.CreaDirectoryDaPercorso(Cartella & Barra & QualeBackup & Barra)
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessuna tabella rilevata"
			Else
				Dim NomeTabella As New List(Of String)

				Do Until Rec.Eof
					NomeTabella.Add(Rec(0).Value)

					Rec.MoveNext
				Loop
				Rec.CLose

				For Each nt As String In NomeTabella
					Sql = "SHOW COLUMNS FROM " & nt
					Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
						Exit For
					Else
						Dim Campi As String = ""
						Dim QuantiCampi As Integer = 0
						Dim Tipologia As New List(Of String)

						Do Until Rec.Eof
							Campi &= Rec(0).Value & ","
							Dim Tipo As String = Rec(1).Value & ";"
							If Rec(3).Value <> "" Then
								Tipo &= Rec(3).Value
							End If
							Tipologia.Add(Rec(0).Value & ";" & Tipo)
							QuantiCampi += 1

							Rec.MoveNext
						Loop
						Rec.Close

						If Campi.Length > 0 Then
							Dim NomeFileStrutt As String = Cartella & Barra & QualeBackup & Barra & nt & "_Strutt.txt"

							gf.EliminaFileFisico(NomeFileStrutt)
							gf.ApreFileDiTestoPerScrittura(NomeFileStrutt)
							For Each t As String In Tipologia
								gf.ScriveTestoSuFileAperto(t)
							Next
							gf.ChiudeFileDiTestoDopoScrittura()

							Campi = Mid(Campi, 1, Campi.Length - 1)
							Sql = "Select " & Campi & " From " & nt
							Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
								Exit For
							Else
								Dim NomeFile As String = Cartella & Barra & QualeBackup & Barra & nt & ".txt"

								gf.EliminaFileFisico(NomeFile)
								gf.ApreFileDiTestoPerScrittura(NomeFile)
								Do Until Rec.Eof
									Dim Riga As String = ""
									For i As Integer = 0 To QuantiCampi - 1
										Riga &= SistemaStringaPerRitorno(Rec(i).Value) & ";"
									Next
									gf.ScriveTestoSuFileAperto(Riga)

									Rec.MoveNext
								Loop
								Rec.Close
								gf.ChiudeFileDiTestoDopoScrittura()
							End If
						End If
					End If
				Next

				Ritorno = QualeBackup
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EffettuaRestore(QualeBackup As String, EsegueBackup As String) As String
		Dim Cartella As String = LeggePathBackup()
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = "OK"
		Dim Barra As String = "\"
		Dim NomeFileFinale As String = Server.MapPath(".") & Barra & "Backups" & Barra & "Esecuzione.txt"
		Dim Esecuzione As String = ""
		Dim gf As New GestioneFilesDirectory
		gf.CreaDirectoryDaPercorso(Cartella & Barra)
		gf.ScansionaDirectorySingola(Cartella & Barra)
		Dim Filetti() As String = gf.RitornaFilesRilevati
		Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati
		Dim Ok As Boolean = True

		Dim sql As String = "Start transaction"
		If EsegueBackup = "SI" Then
			Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
		End If

		For i As Integer = 1 To qFiletti
			If Filetti(i).ToUpper.Contains("_STRUTT") Then
				Dim NomeTabella As String = gf.TornaNomeFileDaPath(Filetti(i)).Replace("_Strutt.txt", "")
				Dim Contenuto As String = gf.LeggeFileIntero(Filetti(i))
				Dim Righe() As String = Contenuto.Split(vbCrLf)
				Dim Chiave As String = ""
				Dim Crea As String = ""
				For Each r As String In Righe
					If r <> "" And r <> vbLf Then
						Dim Campi() As String = r.Split(";")
						Dim NomeCampo As String = Campi(0).Replace(vbLf, "")
						Dim Dimensione As String = Campi(1).Replace(vbLf, "")
						Dim Nulla As String = ""
						If Campi(2) <> "" Then
							Chiave &= Campi(0).Replace(vbLf, "") & ", "
							Nulla = "NOT NULL"
						End If

						Crea &= NomeCampo & " " & Dimensione & " " & Nulla & ", "
					End If
				Next
				If Chiave.Length > 0 Then
					Chiave = Chiave.Substring(0, Chiave.Length - 2)
					Chiave = ", PRIMARY KEY (" & Chiave & ")"
				End If
				If Crea.Length > 0 Then
					Crea = Crea.Substring(0, Crea.Length - 2)

					If EsegueBackup = "SI" Then
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Ritorno <> "OK" Then
							Ok = False
							Exit For
						End If
					Else
						Esecuzione &= "DROP TABLE " & NomeTabella & vbCrLf
					End If

					Crea = "CREATE TABLE " & NomeTabella & " (" & Crea & " " & Chiave & ") ENGINE = InnoDB"
					If EsegueBackup = "SI" Then
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Ritorno <> "OK" Then
							Ok = False
							Exit For
						End If
					Else
						Esecuzione &= Crea & vbCrLf
					End If
				End If
			End If
		Next

		If Ok Then
			For i As Integer = 1 To qFiletti
				If Not Filetti(i).ToUpper.Contains("_STRUTT") Then
					Dim NomeTabella As String = gf.TornaNomeFileDaPath(Filetti(i)).Replace(".txt", "")
					Dim Contenuto As String = gf.LeggeFileIntero(Filetti(i))
					Dim Righe() As String = Contenuto.Split(vbCrLf)
					For Each r As String In Righe
						If r <> "" And r <> vbLf Then
							Dim Campi() As String = r.Split(";")
							Dim Riga As String = ""
							For Each c As String In Campi
								If c <> "" Then
									If ControllaNumerico(c) Then
										Riga &= c.Replace(vbLf, "").Replace(vbCrLf, "") & ","
									Else
										Riga &= "'" & SistemaStringaPerRitorno2(c.Replace(vbLf, "").Replace(vbCrLf, "")) & "',"
									End If
								End If
							Next
							If Riga <> "" Then
								Riga = Mid(Riga, 1, Riga.Length - 1)
								sql = "Insert Into " & NomeTabella & " Values (" & Riga & ")"
								If EsegueBackup = "SI" Then
									Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
									If Ritorno <> "OK" Then
										Ok = False
										Exit For
									End If
								Else
									Esecuzione &= sql & vbCrLf
								End If
							End If
						End If
					Next
				End If
			Next
			If EsegueBackup <> "SI" Then
				gf.CreaAggiornaFile(NomeFileFinale, Esecuzione)
			End If
		End If

		If EsegueBackup = "SI" Then
			If Ok Then
				sql = "commit"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
				Ritorno = "OK"
			Else
				sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			End If
		End If

		Return Ritorno
	End Function

	Private Function ControllaNumerico(Campo As String) As Boolean
		Dim c As Integer = Val(Campo)
		If c > 0 And Not Campo.Contains("/") And Not Campo.Contains("-") And Not Campo.Contains(":") Then
			Return True
		Else
			If Campo = "0" Then
				Return True
			Else
				Return False
			End If
		End If
	End Function
End Class