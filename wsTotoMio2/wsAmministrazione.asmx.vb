Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports Microsoft.SqlServer.Server

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looAdminTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsAmministrazione
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function InviaPromemoria(idAnno As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""
		Dim Rec As Object

		sql = "Select * From Globale Where idAnno=" & idAnno
		Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Dim idConcorso As String = Rec("idGiornata").Value
			Dim Scadenza As String = Rec("Scadenza").Value
			Rec.Close

			Dim Frasi() As String = {
				"***NOME***, ti ricordo che c'è un concorso aperto e devi ancora compilare la schedina...",
				"'***NOME***, se non giochi la colonna è meglio per me...' cit. Looigi",
				"Allora? Questa schedina la vogliamo giocare? Daje ***NOME***",
				"***NOME*** ti ricordo di NON giocare la schedina almeno vinco io...",
				"***NOME***, sta scadendo il termine per giocare la schedina e ancora non l'hai giocata... Che vogliamo fare?",
				"***NOME***... Schedina!!!",
				"***NOME***... Daje co' la schedina!!!",
				"***NOME***... Te devo venì a pija a schiaffi pe' fatte fa la schedina???",
				"Aoh... A ***NOME***... Che volemo fa con questa colonna?",
				"***NOME***, ***NOME***, ***NOME***... Ti devo sempre ricordare della colonna da giuocare...",
				"***NOME***, e allora? 'Sta schedina la vogliamo compilare si o no?",
				"Mi hanno detto che c'è qualcuno che deve ancora compilare la colonna... Non è che per caso sei tu ***NOME***?",
				"Aiuto ***NOME***!!!, se non giochi la schedina qualcuno vincerà al posto tuo...",
				"Questo astensionismo di ***NOME*** dalla giocata della colonna mi manda al manicomo...",
				"Sei sempre tu, ***NOME***, che ti dimentichi di compilare la colonna...",
				"***NOME***, se hai tempo di leggere questa mail, hai anche tempo di compilare la colonna... Forza!!!",
				"Tutto bene... ***NOME***, non giocare la schedina e permetti ad un altro di vincere",
				"***NOME***, ha detto il mio db di fiducia che non riesce a trovare la tua colonna della settimana... Sei sicuro di averla giocata?",
				"Uhm... Sento odore di astensionismo... Qualcuno non vuole giocare la schedina della settimana... Tipo ***NOME***",
				"Daje secco... Ce la puoi fare a giocare la schedina... Su ***NOME*** non ti fare pregare",
				"***NOME*** fai vincere chi ti sta davanti e butta i soldi che hai puntato... Non giocare la schedina...",
				"***NOME***!!! La schedina ti sta chiamando!!!",
				"***NOME***, ***NOME*** perchè abbandonare tutto questo proprio ora che stai andando così bene? Dai, fai questa schedina",
				"***NOME*** ti ricordo che se non compili la colonna il montepremi se lo becca qualcun altro",
				"***NOME*** sei sempre il solito... La vuoi fare o no questa schedina?",
				"La radio ha appena detto che ***NOME*** non ha ancora fatto la schedina... Sarà vero?",
				"Un tale, al bar, diceva che ***NOME*** non fa mai la schedina... Io non ci credo però, magari, dai un'occhiata..."
			}
			Dim x As Integer = GetRandom(0, Frasi.Count - 1)
			Dim Frase As String = Frasi(x)

			sql = "SELECT Distinct A.idUtente, NickName, Mail FROM Utenti A " &
				"Left Join UtentiMails B On A.idAnno = B.idAnno And A.idUtente=B.idUtente " &
				"Where A.idAnno = " & idAnno & " And A.idUtente Not In (Select C.idUtente From Pronostici As C Where C.idAnno = " & idAnno & " And C.idConcorso = " & idConcorso & ") " &
				"And B.Reminder='S'"
			Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				'Ritorno = Rec
			Else
				Do Until Rec.Eof
					Dim Frase2 As String = Frase.Replace("***NOME***", Rec("NickName").Value)
					Dim Testo As String = ""
					Testo = Frase2
					Testo &= "<br />(Questo sopra è un pensiero del server dal quale il povero programmatore si dissocia...)<br /><br />La scadenza del concorso è " & Scadenza & "<br /><br />"
					Testo &= "Per entrare nel sito e vedere il resto: <a href=""" & IndirizzoSito & """>Click QUI</a>"

					Dim m As New mail(Server.MapPath("."))
					m.SendEmail(Server.MapPath("."), Rec("Mail").Value, "TotoMIO: Reminder concorso " & idConcorso, Testo, Nothing)

					Rec.MoveNext
				Loop
				Rec.Close

				Ritorno = "OK"
			End If
		End If

		Return Ritorno
	End Function

End Class