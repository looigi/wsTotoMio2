Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
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

End Class