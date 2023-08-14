Imports System.ComponentModel
Imports System.Net
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://looBilancioTotoMio.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsBilancio
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaBilancio(idAnno As String, idUtente As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Select * From Utenti Where idAnno=" & idAnno & " And idUtente=" & idUtente
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: Nessun utente rilevato"
			Else
				Dim idTipologia As Integer = Rec("idTipologia").Value
				Rec.Close

				Dim Altro As String = ""

				If idTipologia > 0 Then
					Altro = "And A.idUtente=" & idUtente
				End If

				Sql = "Select C.idMovimento, C.Descrizione, A.idUtente, B.NickName, A.Importo, A.Data, A.Note, A.Progressivo, " &
					"D.idPosizione, D.Descrizione As Posizione, Presentati " &
					"From Bilancio A " &
					"Left Join Movimenti C On A.idMovimento=C.idMovimento " &
					"Left Join Utenti B On A.idAnno=B.idAnno And A.idUtente=B.idUtente " &
					"Left Join Posizioni D On A.idPosizione = D.idPosizione " &
					"Left Join Presentati E On A.idAnno = E.idAnno And A.idUtente = E.idUtente " &
					"Where A.idAnno=" & idAnno & " " & Altro & " And A.Eliminato='N' Order By Progressivo"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = "ERROR: Nessun movimento rilevato"
					Else
						Do Until Rec.Eof
							Ritorno &= Rec("idMovimento").Value & ";" & Rec("Descrizione").Value & ";" & Rec("idUtente").Value & ";" &
								SistemaStringaPerRitorno(Rec("NickName").Value) & ";" & Rec("Importo").Value & ";" &
								Rec("Data").Value & ";" & SistemaStringaPerRitorno(Rec("Note").Value) & ";" & Rec("Progressivo").Value & ";" &
								Rec("idPosizione").Value & ";" & Rec("Posizione").Value & ";" & Rec("Presentati").Value & "§"
							Rec.MoveNext
						Loop
						Rec.Close

						Ritorno &= "|"

						Sql = "Select idUtente, NickName, Sum(Totale) As Totale From ( " &
							"SELECT A.idUtente, C.NickName, A.idMovimento, B.Descrizione, -Importo As Totale FROM Bilancio As A " &
							"Left Join Movimenti B On A.idMovimento = B.idMovimento " &
							"Left Join Utenti C On A.idAnno = C.idAnno And A.idUtente = C.idUtente " &
							"Where B.Descrizione = 'Entrata' And A.idAnno= " & idAnno & " " & Altro & " And C.Eliminato='N' " &
							"Union All " &
							"SELECT A.idUtente, C.NickName, A.idMovimento, B.Descrizione, Importo As Totale FROM Bilancio As A " &
							"Left Join Movimenti B On A.idMovimento = B.idMovimento " &
							"Left Join Utenti C On A.idAnno = C.idAnno And A.idUtente = C.idUtente " &
							"Where B.Descrizione = 'Uscita' And A.idAnno= " & idAnno & " " & Altro & " And C.Eliminato='N' " &
							"Union All " &
							"SELECT A.idUtente, C.NickName, A.idMovimento, B.Descrizione, Importo As Totale FROM Bilancio As A " &
							"Left Join Movimenti B On A.idMovimento = B.idMovimento " &
							"Left Join Utenti C On A.idAnno = C.idAnno And A.idUtente = C.idUtente " &
							"Where B.Descrizione = 'Vincita' And A.idAnno= " & idAnno & " " & Altro & " And C.Eliminato='N' " &
							"Union All " &
							"SELECT A.idUtente, C.NickName, A.idMovimento, B.Descrizione, Importo As Totale FROM Bilancio As A " &
							"Left Join Movimenti B On A.idMovimento = B.idMovimento " &
							"Left Join Utenti C On A.idAnno = C.idAnno And A.idUtente = C.idUtente " &
							"Where B.Descrizione = 'Presentazione' And A.idAnno= " & idAnno & " " & Altro & " And C.Eliminato='N' " &
							") As A " &
							"Group By idUtente, NickName " &
							"Order By Totale Desc"
						Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = "ERROR: Nessun movimento 2 rilevato"
							Else
								Do Until Rec.Eof
									Ritorno &= Rec("idUtente").Value & ";" & Rec("NickName").Value & ";" & Rec("Totale").Value & "§"

									Rec.MoveNext
								Loop
								Rec.Close
							End If
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaMovimento(idAnno As String, Progressivo As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = "Update Bilancio Set Eliminato='S' Where idAnno=" & idAnno & " And Progressivo=" & Progressivo
		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaMovimenti() As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = ""

		Sql = "Select * From Movimenti Order By idMovimento"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				Ritorno = "ERROR: nessun movimento rilevato"
			Else
				Do Until Rec.Eof
					Ritorno &= Rec("idMovimento").Value & ";" & SistemaStringaPerRitorno(Rec("Descrizione").Value) & "§"

					Rec.MoveNext
				Loop
				Rec.CLose

				Ritorno &= "|"

				Sql = "Select * From Posizioni Order By idPosizione"
				Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = "ERROR: nessuna posizione rilevata"
					Else
						Do Until Rec.Eof
							Ritorno &= Rec("idPosizione").Value & ";" & SistemaStringaPerRitorno(Rec("Descrizione").Value) & "§"

							Rec.MoveNext
						Loop
						Rec.CLose
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ScriveModificaMovimento(idAnno As String, idUtente As String, idMovimento As String, Importo As String,
											Data As String, Note As String, Progressivo As String, idPosizione As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Sql As String = ""

		If Progressivo = "" Then
			Sql = "Select Coalesce(Max(Progressivo) + 1, 1) As Massimo From Bilancio Where idAnno=" & idAnno
			Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Rec.Eof Then
					Ritorno = "ERROR: Problemi con la max del bilancio"
				Else
					Dim Massimo As Integer = Rec("Massimo").Value
					Rec.Close

					Sql = "Insert Into Bilancio Values (" &
						" " & idAnno & ", " &
						" " & idUtente & ", " &
						" " & Massimo & ", " &
						" " & idMovimento & ", " &
						" " & Importo.Replace(",", ".") & ", " &
						"'" & SistemaStringaPerDB(Data) & "', " &
						"'" & SistemaStringaPerDB(Note) & "', " &
						"'N', " &
						" " & idPosizione & " " &
						")"
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
					If Not Ritorno.Contains(StringaErrore) Then
						Sql = "Select * From Movimenti Where idMovimento=" & idMovimento
						Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = "ERROR: Nessun idMovimento rilevato"
							Else
								Dim Movimento As String = Rec("Descrizione").Value
								Rec.Close

								Sql = "Select * From Posizioni Where idPosizione=" & idPosizione
								Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Rec.Eof Then
										Ritorno = "ERROR: Nessuna posizione rilevata"
									Else
										Dim Posizione As String = Rec("Descrizione").Value
										Rec.Close

										Sql = "Select * From Utenti Where idAnno=" & idAnno & " And idUtente=" & idUtente
										Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
										If TypeOf (Rec) Is String Then
											Ritorno = Rec
										Else
											If Rec.Eof Then
												Ritorno = "ERROR: Nessun Utente rilevato"
											Else
												Dim NickName As String = Rec("NickName").Value
												Dim Mail As String = Rec("Mail").Value
												Rec.Close

												Dim Testo As String = "Movimento di bilancio:<br /><br /><style=""font-weight: bold;"">" & NickName & "</style><br />" &
													"" & Movimento & ": <style=""font-weight: bold;"">" & Importo & "</style><br />" &
													"<style=""font-weight: bold;"">Data: </style>" & Data & "<br />" &
													"<style=""font-weight: bold;"">Note: </style>" & Note & "<br />" &
													"<style=""font-weight: bold;"">Modalità di pagamento: </style>" & Posizione & "<br />"
												Testo &= "<br /><br />Per accedere: <a href=""" & IndirizzoSito & """>Click QUI</a>"

												Dim m As New mail(Server.MapPath("."))

												Sql = "Select * From Utenti Where idAnno=" & idAnno & " And Eliminato='N' And idTipologia=0"
												Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													Dim Mails As New List(Of String)
													Dim mmm As String = ""
													Mails.Add(Mail)
													mmm &= Mail & ";"
													Do Until Rec.Eof
														If Not mmm.Contains(Rec("Mail").Value & ";") Then
															Mails.Add(Rec("Mail").Value)
															mmm &= Rec("Mail").Value
														End If

														Rec.MoveNext
													Loop
													Rec.Close
													For Each mm As String In Mails
														m.SendEmail(Server.MapPath("."), mm, "TotoMIO: Movimento di bilancio", Testo, {})
													Next
												End If
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		Else
			Sql = "Select * From Bilancio Where idAnno=" & idAnno & " And idUtente=" & idUtente & " And Progressivo=" & Progressivo
			Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				If Rec.Eof Then
					Ritorno = "ERROR: Nessun movimento rilevato"
				Else
					Rec.Close

					Sql = "Update Bilancio Set " &
						"idMovimento=" & idMovimento & ", " &
						"Importo=" & Importo.Replace(",", ".") & ", " &
						"Data='" & Data & "'," &
						"Note='" & SistemaStringaPerDB(Note) & "'" &
						"Where idAnno=" & idAnno & " And idUtente=" & idUtente & " And Progressivo=" & Progressivo
					Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione, False)
				End If
			End If
		End If

		Return Ritorno
	End Function
End Class