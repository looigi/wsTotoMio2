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
							Ritorno = idUtente

							Dim gi As New GestioneImmagini
							gi.CreaAvatar(Server.MapPath("."), idAnno, idUtente, NickName, Nome, Cognome)
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
				Do Until Rec.Eof
					Ritorno &= Rec("idPartita").Value & ";" & Rec("Pronostico").Value & ";" & Rec("Segno").Value & "§"

					Rec.MoveNext
				Loop
				Rec.Close
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
	Public Function SalvaPronosticoUtente(idAnno As String, idUtente As String, idConcorso As String, Dati As String) As String
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

End Class