Imports System.ComponentModel
Imports System.IO
Imports System.Web.Services
Imports System.Web.Services.Protocols

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
		Return ritorno
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

						Do Until Rec.Eof
							Campi &= Rec(0).Value & ","
							QuantiCampi += 1

							Rec.MoveNext
						Loop
						Rec.Close

						If Campi.Length > 0 Then
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
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EffettuaRestore(QualeBackup As String) As String
		Dim Cartella As String = LeggePathBackup()
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = "OK"
		Dim Barra As String = "\"
		Dim Sql As String = ""
		Dim Rec As Object
		Dim gf As New GestioneFilesDirectory
		gf.CreaDirectoryDaPercorso(Cartella & Barra)
		gf.ScansionaDirectorySingola(Cartella & Barra)
		Dim Filetti() As String = gf.RitornaFilesRilevati
		Dim qFiletti As Integer = gf.RitornaQuantiFilesRilevati
		For i As Integer = 1 To qFiletti
			Dim NomeTabella As String = gf.TornaNomeFileDaPath(Filetti(i)).Replace(".txt", "")
			Dim Contenuto As String = gf.LeggeFileIntero(Filetti(i))
			Dim Righe() As String = Contenuto.Split(vbCrLf)
			For Each r As String In Righe
				If r <> "" Then
					Dim Campi() As String = r.Split(";")
					Dim Riga As String = ""
					For Each c As String In Campi
						If c <> "" Then
							If ControllaNumerico(c) Then
								Riga &= c & ","
							Else
								Riga &= "'" & SistemaStringaPerRitorno2(c) & "',"
							End If
						End If
					Next
					If Riga <> "" Then
						Riga = Mid(Riga, 1, Riga.Length - 1)
						Sql = "Insert Into " & NomeTabella & " Values (" & Riga & ")"
					End If
				End If
			Next
		Next
		Return Ritorno
	End Function

	Private Function ControllaNumerico(Campo As String) As Boolean
		Dim c As Integer = Val(Campo)
		If c > 0 Then
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