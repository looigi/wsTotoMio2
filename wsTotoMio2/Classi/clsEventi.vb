Public Class clsEventi

	Public Function GestioneEventi(Mp As String, idAnno As Integer, idGiornata As Integer, idEvento As Integer, Conn As Object, Connessione As String) As String
		Dim Ritorno As String = ""
		Dim Sql As String
		Dim Rec As Object

		Sql = "Select * From Eventi Where idEvento=" & idEvento
		Rec = CreaRecordset(Mp, Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Not Rec.Eof Then
				Dim Evento As String = Rec("Descrizione").Value
				Rec.Close

				Select Case Evento
					Case ""
				End Select
			End If
		End If

		Return Ritorno
	End Function

End Class
