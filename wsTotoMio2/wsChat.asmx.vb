Imports System.ComponentModel
Imports System.Diagnostics.Eventing.Reader
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looChatTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsChat
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function InviaMessaggio(idAnno As String, Destinatari As String, idMittente As String, Messaggio As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""
		Dim Rec As Object

		sql = "Select Coalesce(Max(Progressivo) + 1, 1) From Chat Where idAnno=" & idAnno
		Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Dim Progressivo As Integer = Rec(0).Value
			Rec.Close

			sql = "Select * From Utenti Where idAnno=" & idAnno & " And idUtente=" & idMittente
			Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				Dim Mittente As String = Rec("NickName").Value
				Rec.Close

				Destinatari = Destinatari.Replace("*PV*", ";")
				Messaggio = Messaggio.Replace("*PV*", ";")

				Dim Dest() As String = Destinatari.Split(";")
				For Each d As String In Dest
					If d <> "" Then
						sql = "Select * From Utenti Where idAnno=" & idAnno & " And idUtente=" & d
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Dim NickName As String = Rec("NickName").Value
							Dim Mail As String = Rec("Mail").Value
							Rec.Close

							Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

							sql = "Insert Into Chat Values (" &
								" " & idAnno & ", " &
								" " & d & ", " &
								" " & idMittente & ", " &
								" " & Progressivo & ", " &
								"'N', " &
								"'" & SistemaStringaPerDB(Messaggio) & "', " &
								"'" & datella & "' " &
								")"
							Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
							If Ritorno.Contains(StringaErrore) Then
								Exit For
							Else
								Dim Testo As String = ""
								Testo = "Hai ricevuto un nuovo messaggio da parte di " & Mittente & "<br /><br />"
								Testo &= "Il testo iniziale è:<br />" & Messaggio.Substring(0, 15) & "...<br /><br />"
								Testo &= "Per entrare nel sito e vedere il resto: <a href=""" & IndirizzoSito & """>Click QUI</a>"

								Dim m As New mail(Server.MapPath("."))
								m.SendEmail(Server.MapPath("."), Mail, "TotoMIO: Nuovo messaggio da " & Mittente, Testo, Nothing)
							End If
							Progressivo += 1
						End If
					End If
				Next
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function LeggeNuoviMessaggi(idAnno As String, idUtente As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""
		Dim Rec As Object

		sql = "Select Coalesce(Count(*),0) As Quanti From Chat Where idAnno=" & idAnno & " And idUtente=" & idUtente & " And Letto='N'"
		Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Ritorno = Rec("Quanti").Value
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaTuttiIMessaggi(idAnno As String, idUtente As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""
		Dim Rec As Object

		sql = "Select * From Chat A " &
			"Left Join Utenti B On A.idAnno=B.idAnno And A.idMittente=B.idUtente " &
			"Where A.idAnno=" & idAnno & " And A.idUtente=" & idUtente & " Order By Progressivo Desc"
		Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Do Until Rec.Eof
				Ritorno &= Rec("Progressivo").Value & ";" & Rec("Letto").Value & ";" & Rec("idMittente").Value & ";" & SistemaStringaPerRitorno(Rec("NickName").Value) & ";" &
					SistemaStringaPerRitorno(Rec("Messaggio").Value) & ";" & Rec("Data").Value & "§"

				Rec.MoveNext
			Loop
			Rec.Close
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SegnaComeLetto(idAnno As String, Progressivo As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""

		sql = "Update Chat Set Letto='S' Where idAnno=" & idAnno & " And Progressivo=" & Progressivo
		Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)

		Return Ritorno
	End Function
End Class