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
								   Password As String, Mail As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		If ControllaValiditaMail(Mail) Then
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
							"'N'" &
							")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Not Ritorno.Contains(StringaErrore) Then
							Ritorno = idUtente
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaUtente(idAnno As String, idUtente As String, NickName As String, Cognome As String, Nome As String,
								   Password As String, Mail As String) As String
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
				"Mail='" & SistemaStringaPerDB(Mail) & "' " &
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
	Public Function RitornaUtenti(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Select * From Utenti Where idAnno=" & idAnno & " And Eliminato='N' Order By idUtente"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Do Until Rec.Eof
				Ritorno &= SistemaStringaPerRitorno(Rec("idUtente").Value) & ";"
				Ritorno &= SistemaStringaPerRitorno(Rec("NickName").Value) & ";"
				Ritorno &= SistemaStringaPerRitorno(Rec("Cognome").Value) & ";"
				Ritorno &= SistemaStringaPerRitorno(Rec("Nome").Value) & ";"
				Ritorno &= SistemaStringaPerRitorno(Rec("Password").Value) & ";"
				Ritorno &= SistemaStringaPerRitorno(Rec("Mail").Value) & "§"

				Rec.MoveNext
			Loop
			Rec.Close

			If Ritorno = "" Then
				Ritorno = "ERROR: Nessun utente rilevato"
			End If
		End If

		Return Ritorno
	End Function

End Class