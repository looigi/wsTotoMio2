Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looAnniTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsAnni
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function AggiungeAnno(Descrizione As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""
		Dim Rec As Object

		sql = "Select Coalesce(Max(idAnno) + 1, 1) From Anni"
		Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			Dim idAnno As String = Rec(0).Value

			sql = "Insert Into Anni Values (" &
						" " & idAnno & ", " &
						"'" & SistemaStringaPerDB(Descrizione) & "' " &
						")"
			Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
			If Not Ritorno.Contains(StringaErrore) Then
				Ritorno = idAnno
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaAnno(idAnno As String, Descrizione As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim sql As String = ""
		sql = "Update Anni Set " &
			"Descrizione='" & SistemaStringaPerDB(Descrizione) & "' " &
			"Where idAnno=" & idAnno
		Ritorno = Conn.EsegueSql(Server.MapPath("."), sql, Connessione, False)
		If Not Ritorno.Contains(StringaErrore) Then
			Ritorno = "*"
		End If

		Return Ritorno
	End Function
End Class