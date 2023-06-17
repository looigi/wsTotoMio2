Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looConcorsiTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsConcorsi
	Inherits System.Web.Services.WebService

	'<WebMethod()>
	'Public Function AggiungeConcorso(idAnno As String, Dati As String) As String
	'	Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
	'	Dim Conn As Object = New clsGestioneDB(TipoServer)
	'	Dim Ritorno As String = ""
	'	Dim sql As String = ""

	'	sql = "Start transaction"
	'	Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
	'	If Not Ritorno.Contains("ERROR:") Then
	'		sql = "Select Coalesce(Max(idGiornata)+1,1) From Concorsi Where idAnno=" & idAnno
	'		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
	'		If TypeOf (Rec) Is String Then
	'			Ritorno = Rec
	'		Else
	'			Dim idGiornata As String = Rec(0).Value

	'			Dim Dati2() As String = Dati.Split("§")
	'			For Each D As String In Dati2
	'				Dim D2() As String = D.Split(";")

	'				' idAnno	idGiornata	idPartita	Prima	Seconda	Risultato	Segno
	'				sql = "Insert Into Concorsi Values (" &
	'					" " & idAnno & ", " &
	'					" " & idGiornata & ", " &
	'					" " & D2(0) & ", " &
	'					"'" & SistemaStringaPerDB(D2(1)) & "', " &
	'					"'" & SistemaStringaPerDB(D2(2)) & "', " &
	'					"'" & SistemaStringaPerDB(D2(3)) & "', " &
	'					"'' " &
	'					")"
	'				Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
	'				If Ritorno.Contains(StringaErrore) Then
	'					Exit For
	'				End If
	'			Next
	'		End If

	'		If Ritorno = "OK" Then
	'			sql = "commit"
	'			Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione)
	'		Else
	'			sql = "rollback"
	'			Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), sql, Connessione)
	'		End If
	'	End If

	'	Return Ritorno
	'End Function

	<WebMethod()>
	Public Function ApreConcorso(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From ModalitaConcorso Where Descrizione='Aperto'"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessuna tipologia rilevata"
			Else
				Dim idModalita As String = Rec("idModalitaConcorso").Value
				Dim Descrizione As String = Rec("Descrizione").Value
				Rec.Close

				sql = "Update Globale Set idModalitaConcorso=" & idModalita & " Where idAnno=" & idAnno
				Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
				If Not Ritorno.Contains("ERROR") Then
					Ritorno = idModalita & ";" & Descrizione
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaConcorso(idAnno As String, idConcorso As String, Dati As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""

		sql = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
		If Not Ritorno.Contains("ERROR:") Then
			sql = "Delete From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idConcorso
			Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			If Not Ritorno.Contains("ERROR:") Then
				Dim Dati2() As String = Dati.Split("§")
				For Each D As String In Dati2
					If D <> "" Then
						Dim D2() As String = D.Split(";")

						' idAnno	idGiornata	idPartita	Prima	Seconda	Risultato	Segno
						sql = "Insert Into Concorsi Values (" &
							" " & idAnno & ", " &
							" " & idConcorso & ", " &
							" " & D2(0) & ", " &
							"'" & SistemaStringaPerDB(D2(1)) & "', " &
							"'" & SistemaStringaPerDB(D2(2)) & "', " &
							"'" & SistemaStringaPerDB(D2(3)) & "', " &
							"'" & SistemaStringaPerDB(D2(4)) & "' " &
							")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
						If Ritorno.Contains(StringaErrore) Then
							Exit For
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

	<WebMethod()>
	Public Function RitornoConcorso(idAnno As String, idConcorso As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""

		sql = "Select * From Concorsi Where idAnno=" & idAnno & " And idConcorso=" & idConcorso & " Order By idPartita"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun dato presente"
			Else
				Do Until Rec.Eof
					Ritorno &= Rec("idPartita").Value & ";" & SistemaStringaPerRitorno(Rec("Prima").Value) & ";" &
						SistemaStringaPerRitorno(Rec("Seconda").Value) & ";" & Rec("Risultato").Value & ";" &
						Rec("Segno").Value & "§"
					Rec.MoveNext
				Loop
				Rec.Close
			End If
		End If

		Return Ritorno
	End Function

End Class