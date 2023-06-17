Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports ADODB

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsTotoMIO2
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function Test() As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = "Select * From Utenti"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)

		Do Until Rec.eof
			Ritorno &= Rec(2).Value & " " & Rec(3).Value & ";"

			Rec.movenext
		Loop
		Rec.Close

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaDatiGenerali(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Select *, C.Descrizione As ModalitaConcorso From Globale A " &
			"Left Join Anni B On A.idAnno = B.idAnno " &
			"Left Join ModalitaConcorso C On C.idModalitaConcorso = A.idModalitaConcorso " &
			"Where A.idAnno=" & idAnno
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun valore ritornato"
			Else
				Ritorno = Rec("idGiornata").Value & ";" &
						Rec("idModalitaConcorso").Value & ";" &
						Rec("ModalitaConcorso").Value & ";" &
						"|"
				Rec.Close

				Sql = "Select * From Anni Order By idAnno Desc"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Do Until Rec.Eof
						Ritorno &= Rec("idAnno").Value & ";" & Rec("Descrizione").Value & "§"

						Rec.MoveNext
					Loop
					Rec.Close
				End If
			End If
		End If
		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AggiornaDatiGenerali(idAnno As String, idGiornata As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Update Globale Set " &
			"idGiornata=" & idGiornata & " " &
			"Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
		If Not Ritorno.Contains(StringaErrore) Then
			Ritorno = "*"
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AvanzaGiornata(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Dim Sql As String = "Start transaction"
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
		If Not Ritorno.Contains("Error:") Then
			Sql = "Update Globale Set " &
				"idGiornata=idGiornata+1 " &
				"Where idAnno=" & idAnno
			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
			If Not Ritorno.Contains(StringaErrore) Then
				Sql = "Select Coalesce(Count(*),0) From Utenti Where idAnno=" & idAnno & " And Eliminato='N'"
				Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					Dim QuantiGiocatori As Integer = Rec(0).Value
					Rec.Close

					If QuantiGiocatori = 0 Then
						Ritorno = "ERROR: Nessun giocatore presente"
					Else
						Ritorno = AggiungeGiornataDiCampionatoAEventi(Conn, Connessione, idAnno)
						If Not Ritorno.Contains("ERROR") Then
							Ritorno = RitornaEventiGiornata(idAnno)
						End If
					End If
				End If
			End If

			If Not Ritorno.Contains("Error") Then
				Sql = "commit"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			Else
				Sql = "rollback"
				Dim Rit As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaEventiGiornata(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""

		Dim idGiornata As String = RitornaGiornata(Server.MapPath("."), Conn, Connessione, idAnno)
		If Not idGiornata.Contains("ERROR") Then
			Dim Sql As String = "Select A.idAnno, A.idGiornata, A.idEvento, A.idTipologia, B.Descrizione As Evento, C.Descrizione As Tipologia From " &
				"EventiCalendario As A " &
				"Left Join Eventi As B On A.idEvento = B.idEvento " &
				"Left Join EventiTipologie As C On A.idTipologia = C.idTipologia " &
				"Where A.idAnno=" & idAnno & " And A.idGiornata=" & idGiornata
			Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Not Rec.Eof Then
					Do Until Rec.Eof
						Ritorno &= Rec("idAnno").Value & ";"
						Ritorno &= Rec("idGiornata").Value & ";"
						Ritorno &= Rec("idEvento").Value & ";"
						Ritorno &= Rec("idTipologia").Value & ";"
						Ritorno &= SistemaStringaPerRitorno(Rec("Evento").Value) & ";"
						Ritorno &= SistemaStringaPerRitorno(Rec("Tipologia").Value) & "§"

						Rec.MoveNext
					Loop
				Else
					Ritorno = "Nessun evento in programma per la giornata " & idGiornata
				End If
				Rec.Close
			End If
		Else
			Ritorno = idGiornata
		End If

		Return Ritorno
	End Function

	Public Function AggiungeGiornataDiCampionatoAEventi(Conn As Object, Connessione As String, idAnno As String) As String
		Dim Ritorno As String = ""
		Dim idGiornata As String = RitornaGiornata(Server.MapPath("."), Conn, Connessione, idAnno)
		If Not idGiornata.Contains("ERROR") Then
			' Inserimento Partita di campionato
			Dim Sql As String = "Insert Into EventiCalendario Values (" &
				" " & idAnno & ", " &
				" " & idGiornata & ", " &
				"1, " &
				"1 " &
				")"
			Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
			If Not Ritorno.Contains(StringaErrore) Then
				' Gestire tutti gli altri eventi giornata per giornata
				Select Case idGiornata
					Case 10
				End Select

				Ritorno = "*"
			End If
		Else
			Ritorno = idGiornata
		End If

		Return Ritorno
	End Function
End Class