Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols

' Per consentire la chiamata di questo servizio Web dallo script utilizzando ASP.NET AJAX, rimuovere il commento dalla riga seguente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://totomio.statistiche.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsStatistiche
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function Statistiche(idAnno As Integer) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim Ritorno1 As String = ElaboraStatistiche(idAnno, Conn, Connessione)
		Dim Ritorno2 As String = ElaboraStatistiche("", Conn, Connessione)
		If Not Ritorno1.Contains(StringaErrore) And Not Ritorno2.Contains(StringaErrore) Then
			Ritorno = "{"
			Ritorno &= "" & Chr(34) & "Annuale" & Chr(34) & ": " & Ritorno1 & ","
			Ritorno &= "" & Chr(34) & "Storico" & Chr(34) & ": " & Ritorno2
			Ritorno &= "}"
		End If
		Return Ritorno
	End Function

	Private Function ElaboraStatistiche(idAnno As String, Conn As Object, Connessione As String) As String
		Dim Ritorno As String = ""
		Dim Altro As String = ""
		Dim Altro2 As String = ""
		If idAnno <> "" Then
			Altro = "Where A.idAnno = " & idAnno & " "
			Altro2 = "A.idAnno = " & idAnno & " And"
		End If

		Dim sql As String = "SELECT A.idUtente, B.NickName, Coalesce(Avg(Punti), 0) As MediaPunti, Coalesce(Avg(SegniPresi), 0) As MediaSegni, " &
			"Coalesce(Avg(RisultatiEsatti), 0) As MediaRisEsatti, Coalesce(Avg(RisultatiCasaTot), 0) As MediaRisCasa, " &
			"Coalesce(Avg(RisultatiFuoriTot), 0) As MediaRisFuori, Coalesce(Avg(SommeGoal), 0) As MediaSomma, " &
			"Coalesce(Avg(DifferenzeGoal), 0) As MediaDiff, Coalesce(Avg(PuntiPartitaScelta), 0) As MediaPuntiPS,  " &
			"Coalesce(Sum(Punti), 0) As SommaPunti, Coalesce(Sum(SegniPresi), 0) As SommaSegni, " &
			"Coalesce(Sum(RisultatiEsatti), 0) As SommaRisEsatti, Coalesce(Sum(RisultatiCasaTot), 0) As SommaRisCasa, " &
			"Coalesce(Sum(RisultatiFuoriTot), 0) As SommaRisFuori, Coalesce(Sum(SommeGoal), 0) As SommaSomma, " &
			"Coalesce(Sum(DifferenzeGoal), 0) As SommaDiff, Coalesce(Sum(PuntiPartitaScelta), 0) As SommaPuntiPS, Coalesce(Sum(PuntiSorpresa), 0) As PuntiSorpresa " &
			"FROM Risultati As A " &
			"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
			Altro & " " &
			"Group By A.idUtente, B.NickName " &
			"Order By A.idUtente"
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			'If Rec.Eof Then
			'	Ritorno = "ERROR: Nessun utente rilevato"
			'Else
			Dim StatisticheRisultati As String = "["
			Dim Riga As Boolean = True
			Do Until Rec.Eof
				StatisticheRisultati &= "{"
				StatisticheRisultati &= "" & Chr(34) & "idUtente" & Chr(34) & ": " & Rec("idUtente").Value & ", "
				StatisticheRisultati &= "" & Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
				StatisticheRisultati &= "" & Chr(34) & "SommaPunti" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaPunti").Value, False) & ", "
				StatisticheRisultati &= "" & Chr(34) & "SommaSegni" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaSegni").Value, False) & ", "
				StatisticheRisultati &= "" & Chr(34) & "SommaRisEsatti" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaRisEsatti").Value, False) & ", "
				StatisticheRisultati &= "" & Chr(34) & "SommaRisCasa" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaRisCasa").Value, False) & ", "
				StatisticheRisultati &= "" & Chr(34) & "SommaRisFuori" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaRisFuori").Value, False) & ", "
				StatisticheRisultati &= "" & Chr(34) & "SommaSomma" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaSomma").Value, False) & ", "
				StatisticheRisultati &= "" & Chr(34) & "SommaDiff" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaDiff").Value, False) & ", "
				StatisticheRisultati &= "" & Chr(34) & "SommaPuntiPS" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaPuntiPS").Value, False) & ", "
				StatisticheRisultati &= "" & Chr(34) & "MediaPunti" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaPunti").Value, True) & ", "
				StatisticheRisultati &= "" & Chr(34) & "MediaSegni" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaSegni").Value, True) & ", "
				StatisticheRisultati &= "" & Chr(34) & "MediaRisEsatti" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaRisEsatti").Value, True) & ", "
				StatisticheRisultati &= "" & Chr(34) & "MediaRisCasa" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaRisCasa").Value, True) & ", "
				StatisticheRisultati &= "" & Chr(34) & "MediaRisFuori" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaRisFuori").Value, True) & ", "
				StatisticheRisultati &= "" & Chr(34) & "MediaSomma" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaSomma").Value, True) & ", "
				StatisticheRisultati &= "" & Chr(34) & "MediaDiff" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaDiff").Value, True) & ", "
				StatisticheRisultati &= "" & Chr(34) & "MediaPuntiPS" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaPuntiPS").Value, True) & ", "
				StatisticheRisultati &= "" & Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & Riga & Chr(34) & ", "
				StatisticheRisultati &= "" & Chr(34) & "PuntiSorpresa" & Chr(34) & ": " & Chr(34) & SistemaNumeroDaDB(Rec("PuntiSorpresa").Value, True) & Chr(34) & " "
				StatisticheRisultati &= "}, "
				Riga = Not Riga

				Rec.MoveNext
			Loop
			If StatisticheRisultati <> "[" Then
				StatisticheRisultati = Mid(StatisticheRisultati, 1, StatisticheRisultati.Length - 2)
			End If
			StatisticheRisultati &= "]"
			Rec.Close

			sql = "SELECT A.idUtente, B.NickName, Coalesce(Avg(Vittorie), 0) As MediaVittorie,  " &
				"Coalesce(Avg(Ultimo), 0) As MediaUltimo, Coalesce(Avg(Jolly), 0) As MediaJolly, " &
				"Coalesce(Sum(Vittorie), 0) As SommaVittorie, " &
				"Coalesce(Sum(Ultimo), 0) As SommaUltimo, Coalesce(Sum(Jolly), 0) As SommaJolly " &
				"From RisultatiAltro As A " &
				"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
				Altro & " " &
				"Group By A.idUtente, B.idUtente " &
				"Order By A.idUtente"
			Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
			If TypeOf (Rec) Is String Then
				Ritorno = Rec
			Else
				'If Rec.Eof Then
				'	Ritorno = "ERROR: Nessun utente rilevato"
				'Else
				Dim StatisticheRisultatiA As String = "["
				Riga = True
				Do Until Rec.Eof
					StatisticheRisultatiA &= "{"
					StatisticheRisultatiA &= "" & Chr(34) & "idUtente" & Chr(34) & ": " & Rec("idUtente").Value & ", "
					StatisticheRisultatiA &= "" & Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
					StatisticheRisultatiA &= "" & Chr(34) & "SommaVittorie" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaVittorie").Value, False) & ", "
					StatisticheRisultatiA &= "" & Chr(34) & "SommaUltimo" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaUltimo").Value, False) & ", "
					StatisticheRisultatiA &= "" & Chr(34) & "SommaJolly" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaJolly").Value, False) & ", "
					StatisticheRisultatiA &= "" & Chr(34) & "MediaVittorie" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaVittorie").Value, True) & ", "
					StatisticheRisultatiA &= "" & Chr(34) & "MediaUltimo" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaUltimo").Value, True) & ", "
					StatisticheRisultatiA &= "" & Chr(34) & "MediaJolly" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("MediaJolly").Value, True) & ", "
					StatisticheRisultatiA &= "" & Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & Riga & Chr(34) & " "
					StatisticheRisultatiA &= "}, "
					Riga = Not Riga

					Rec.MoveNext
				Loop
				If StatisticheRisultatiA <> "[" Then
					StatisticheRisultatiA = Mid(StatisticheRisultatiA, 1, StatisticheRisultatiA.Length - 2)
				End If
				StatisticheRisultatiA &= "]"
				Rec.Close

				sql = "Select idUtente, NickName, Sum(Vinte) As Vinte, Sum(Pareggiate) As Pareggiate, Sum(Perse) As Perse, (Sum(Vinte) + Sum(Pareggiate) + Sum(Perse)) As Giocate From (" &
					"SELECT B.idUtente, B.NickName, Coalesce(Count(*), 0) As Vinte, 0 As Pareggiate, 0 As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore1 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 1 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"SELECT B.idUtente, B.NickName, Coalesce(Count(*), 0) As Vinte, 0 As Pareggiate, 0 As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore2 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 2 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"Select B.idUtente, B.NickName, 0 As Vinte, Coalesce(Count(*), 0) As Pareggiate, 0 As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore1 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 0 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"SELECT B.idUtente, B.NickName, 0 As Vinte, Coalesce(Count(*), 0) As Pareggiate, 0 As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore2 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 0 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"SELECT B.idUtente, B.NickName, 0 As Vinte, 0 As Pareggiate, Coalesce(Count(*), 0) As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore1 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 2 " &
					"Group By B.idUtente, B.NickName " &
					"Union ALL " &
					"SELECT B.idUtente, B.NickName, 0 As Vinte, 0 As Pareggiate, Coalesce(Count(*), 0) As Perse FROM EventiPartite As A " &
					"Left Join Utenti As B On A.idAnno = B.idAnno And A.idGiocatore2 = B.idUtente " &
					"Where " & Altro2 & " A.idVincente = 1 " &
					"Group By B.idUtente, B.NickName " &
					") As A " &
					"Group By idUtente, NickName " &
					"Order By 3 Desc, 2 Desc"
				Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					'If Rec.Eof Then
					'	Ritorno = "ERROR: Nessun utente rilevato"
					'Else
					Dim StatisticheScontriDiretti As String = "["
					Riga = True
					Do Until Rec.Eof
						StatisticheScontriDiretti &= "{"
						StatisticheScontriDiretti &= "" & Chr(34) & "idUtente" & Chr(34) & ": " & Rec("idUtente").Value & ", "
						StatisticheScontriDiretti &= "" & Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
						StatisticheScontriDiretti &= "" & Chr(34) & "Vinte" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("Vinte").Value, False) & ", "
						StatisticheScontriDiretti &= "" & Chr(34) & "Pareggiate" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("Pareggiate").Value, False) & ", "
						StatisticheScontriDiretti &= "" & Chr(34) & "Perse" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("Perse").Value, False) & ", "
						StatisticheScontriDiretti &= "" & Chr(34) & "Giocate" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("Giocate").Value, False) & ", "
						StatisticheScontriDiretti &= "" & Chr(34) & "MediaVinte" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("Vinte").Value / Rec("Giocate").Value, False) & ", "
						StatisticheScontriDiretti &= "" & Chr(34) & "MediaPareggiate" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("Pareggiate").Value / Rec("Giocate").Value, False) & ", "
						StatisticheScontriDiretti &= "" & Chr(34) & "MediaPerse" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("Perse").Value / Rec("Giocate").Value, False) & ", "
						StatisticheScontriDiretti &= "" & Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & Riga & Chr(34) & " "
						StatisticheScontriDiretti &= "}, "
						Riga = Not Riga

						Rec.MoveNext
					Loop
					If StatisticheScontriDiretti <> "[" Then
						StatisticheScontriDiretti = Mid(StatisticheScontriDiretti, 1, StatisticheScontriDiretti.Length - 2)
					End If
					StatisticheScontriDiretti &= "]"
					Rec.Close

					Dim idGiornata As String = ""
					If idAnno <> "" Then
						sql = "Select * From Globale Where idAnno=" & idAnno
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							idGiornata = Rec("idGiornata").Value
							Rec.Close
						End If
					Else
						idGiornata = "Nessuna"
					End If

					Dim QuantiAnni As String = "1"
					If idAnno = "" Then
						sql = "Select Coalesce(Count(*), 0) As Quanti From Anni"
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							QuantiAnni = Rec("Quanti").Value
							Rec.Close
						End If
					End If

					sql = "Select idUtente, NickName, Sum(Entrate) As SommaEntrate, Sum(Uscite) As SommaUscite, Sum(Vincita) As SommaVincite, " &
						"(Sum(Entrate) + Sum(Vincita)) - Sum(Uscite) As SommaBilancio From ( " &
						"SELECT A.idUtente, B.NickName, Sum(Importo) As Entrate, 0 As Uscite, 0 As Vincita FROM Bilancio As A " &
						"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
						"Where " & Altro2 & " idMovimento = 1 And A.Eliminato = 'N' " &
						"Group By A.idUtente, B.NickName " &
						"Union ALL " &
						"SELECT A.idUtente, B.NickName, 0 Entrate, Sum(Importo) As Uscite, 0 As Vincita FROM Bilancio As A " &
						"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
						"Where " & Altro2 & " idMovimento = 2 And A.Eliminato = 'N' " &
						"Group By A.idUtente, B.NickName " &
						"Union ALL " &
						"SELECT A.idUtente, B.NickName, 0 Entrate, 0 As Uscite, Sum(Importo) As Vincita FROM Bilancio As A " &
						"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
						"Where " & Altro2 & " idMovimento = 3 And A.Eliminato = 'N' " &
						"Group By A.idUtente, B.NickName " &
						") AS A  " &
						"Group By idUtente, NickName " &
						"Order By 6 Desc"
					Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						'If Rec.Eof Then
						'	Ritorno = "ERROR: Nessun utente rilevato"
						'Else
						Dim StatisticheBilancio As String = "["
						Riga = True
						Do Until Rec.Eof
							StatisticheBilancio &= "{"
							StatisticheBilancio &= "" & Chr(34) & "idUtente" & Chr(34) & ": " & Rec("idUtente").Value & ", "
							StatisticheBilancio &= "" & Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
							StatisticheBilancio &= "" & Chr(34) & "Entrate" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaEntrate").Value, True) & ", "
							StatisticheBilancio &= "" & Chr(34) & "Uscite" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaUscite").Value, True) & ", "
							StatisticheBilancio &= "" & Chr(34) & "Vincite" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaVincite").Value, True) & ", "
							StatisticheBilancio &= "" & Chr(34) & "Bilancio" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("SommaBilancio").Value, True) & ", "
							StatisticheBilancio &= "" & Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & Riga & Chr(34) & " "
							StatisticheBilancio &= "}, "
							Riga = Not Riga

							Rec.MoveNext
						Loop
						If StatisticheBilancio <> "[" Then
							StatisticheBilancio = Mid(StatisticheBilancio, 1, StatisticheBilancio.Length - 2)
						End If
						StatisticheBilancio &= "]"
						Rec.Close

						Dim Anno As String = ""
						If idAnno <> "" Then
							Anno = "1"
						Else
							Anno = "Tutti"
						End If

						Dim Quanti As String = "0"
						sql = "SELECT Coalesce(Count(*), 0) As Quanti FROM Concorsi " &
							IIf(idAnno <> "", "Where idAnno = " & idAnno & " And idPartita=1 Group By idAnno", "Where idPartita=1")
						Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Not Rec.Eof Then
								Quanti = Rec("Quanti").Value
							End If
							Rec.Close

							sql = "SELECT A.idUtente, NickName, Coalesce(Count(*), 0) As Quanti FROM Pronostici As A " &
								"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
								"Where " & Altro2 & " A.idPartita = 1 " &
								"Group By A.idUtente, NickName"
							Rec = CreaRecordset(Server.MapPath("."), Conn, sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								'If Rec.Eof Then
								'	Ritorno = "ERROR: Nessun utente rilevato"
								'Else
								Dim StatistichePronostici As String = "["
								Riga = True
								Do Until Rec.Eof
									StatistichePronostici &= "{"
									StatistichePronostici &= "" & Chr(34) & "idUtente" & Chr(34) & ": " & Rec("idUtente").Value & ", "
									StatistichePronostici &= "" & Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & Rec("NickName").Value & Chr(34) & ", "
									StatistichePronostici &= "" & Chr(34) & "Giocate" & Chr(34) & ": " & SistemaNumeroDaDB(Rec("Quanti").Value, False) & ", "
									StatistichePronostici &= "" & Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & Riga & Chr(34) & " "
									StatistichePronostici &= "}, "
									Riga = Not Riga

									Rec.MoveNext
								Loop
								If StatistichePronostici <> "[" Then
									StatistichePronostici = Mid(StatistichePronostici, 1, StatistichePronostici.Length - 2)
								End If
								StatistichePronostici &= "]"
								Rec.Close

								Dim SquadrePrese As String = GeneraSquadrePrese(Server.MapPath("."), idAnno, Conn, Connessione)

								Dim PuntiPresi As String = RitornaRecord(Conn, Connessione, idAnno, "Punti")
								Dim SegniPresi As String = RitornaRecord(Conn, Connessione, idAnno, "SegniPresi")
								Dim RisultatiEsattiPresi As String = RitornaRecord(Conn, Connessione, idAnno, "RisultatiEsatti")
								Dim RisultatiCasaTotPresi As String = RitornaRecord(Conn, Connessione, idAnno, "RisultatiCasaTot")
								Dim RisultatiFuoriTotPresi As String = RitornaRecord(Conn, Connessione, idAnno, "RisultatiFuoriTot")
								Dim SommeGoalPresi As String = RitornaRecord(Conn, Connessione, idAnno, "SommeGoal")
								Dim DifferenzeGoalPresi As String = RitornaRecord(Conn, Connessione, idAnno, "DifferenzeGoal")
								Dim PuntiPartitaSceltaPresi As String = RitornaRecord(Conn, Connessione, idAnno, "PuntiPartitaScelta")
								Dim PuntiSorpresaPresi As String = RitornaRecord(Conn, Connessione, idAnno, "PuntiSorpresa")
								Dim VittorePresi As String = RitornaRecordAltro(Conn, Connessione, idAnno, "Vittorie")
								Dim UltimoPresi As String = RitornaRecordAltro(Conn, Connessione, idAnno, "Ultimo")
								Dim JollyPresi As String = RitornaRecordAltro(Conn, Connessione, idAnno, "Jolly")

								Dim StatistichePresi As String = "["
								StatistichePresi &= PuntiPresi & ","
								StatistichePresi &= SegniPresi & ","
								StatistichePresi &= RisultatiEsattiPresi & ","
								StatistichePresi &= RisultatiCasaTotPresi & ","
								StatistichePresi &= RisultatiFuoriTotPresi & ","
								StatistichePresi &= SommeGoalPresi & ","
								StatistichePresi &= DifferenzeGoalPresi & ","
								StatistichePresi &= PuntiPartitaSceltaPresi & ","
								StatistichePresi &= PuntiSorpresaPresi & ","
								StatistichePresi &= VittorePresi & ","
								StatistichePresi &= UltimoPresi & ","
								StatistichePresi &= JollyPresi
								StatistichePresi &= "]"

								Ritorno &= "{"
								Ritorno &= "" & Chr(34) & "Anno" & Chr(34) & ": " & Chr(34) & Anno & Chr(34) & ", "
								Ritorno &= "" & Chr(34) & "Anni" & Chr(34) & ": " & QuantiAnni & ", "
								Ritorno &= "" & Chr(34) & "Giornata" & Chr(34) & ": " & Chr(34) & idGiornata & Chr(34) & ", "
								Ritorno &= "" & Chr(34) & "ConcorsiAperti" & Chr(34) & ": " & Quanti & ", "
								Ritorno &= "" & Chr(34) & "StatistichePresi" & Chr(34) & ": " & StatistichePresi & ", "
								Ritorno &= "" & Chr(34) & "Risultati" & Chr(34) & ": " & StatisticheRisultati & ", "
								Ritorno &= "" & Chr(34) & "RisultatiAltro" & Chr(34) & ": " & StatisticheRisultatiA & ", "
								Ritorno &= "" & Chr(34) & "ScontriDiretti" & Chr(34) & ": " & StatisticheScontriDiretti & ", "
								Ritorno &= "" & Chr(34) & "Pronostici" & Chr(34) & ": " & StatistichePronostici & ", "
								Ritorno &= "" & Chr(34) & "SquadrePrese" & Chr(34) & ": " & SquadrePrese & ", "
								Ritorno &= "" & Chr(34) & "Bilancio" & Chr(34) & ": " & StatisticheBilancio ' & ", "
								'Ritorno &= "" & Chr(34) & "Grafici" & Chr(34) & ": " & PrendeGrafici(Conn, Connessione, idAnno, idGiornata)
								Ritorno &= "}"
							End If
						End If

						'End If
					End If
				End If
				'End If
			End If
			'End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function PrendeGrafici(idUtente As String, idAnno As String, idGiornata As String, Cosa As String, Altro As String) As String
		Dim Connessione As String = RitornaPercorso(Server.MapPath("."), 5)
		Dim Conn As Object = New clsGestioneDB(TipoServer)
		Dim Ritorno As String = ""
		Dim idUtenti As New List(Of String)

		If idAnno <> "" Then
			Dim Sql As String = "Select * From Utenti Where idAnno=" & idAnno & " And Eliminato='N'"
			Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				' Ritorno = Rec
			Else
				Do Until Rec.Eof
					idUtenti.Add(Rec("idUtente").Value & ";" & Rec("NickName").Value)

					Rec.MoveNext
				Loop
				Rec.Close

				Ritorno &= "{""DatiGrafico"": " & PrendeGraficiPunti(Conn, Connessione, idUtente, idAnno, idGiornata, idUtenti, Cosa, Altro) & "}"
			End If
		Else
			Ritorno = "[{}]"
		End If

		Return Ritorno
	End Function

	Private Function PrendeGraficiPunti(Conn As Object, Connessione As String, idUtente As String, idAnno As String, idGiornata As String, idUtenti As List(Of String),
										Cosa As String, Altro As String) As String
		Dim Ritorno As String = "{"
		Ritorno &= """Tipologia"": """ & Cosa & """, "
		Ritorno &= """data"": ["
		Dim Ritorno3 As String = ""
		Dim Sql As String = ""
		Dim Rec As Object

		For Each id As String In idUtenti
			Dim iid() As String = id.Split(";")
			If Cosa = "Posizioni" Then
				Sql = "Select A.idUtente, A.idConcorso, A.Posizione As Valore From PosizioniClassifica A " &
					"Where A.idAnno = " & idAnno & " And A.idUtente=" & iid(0) & " And A.idConcorso <=" & idGiornata & " " &
					"Order By A.idConcorso"
			Else
				If Cosa = "Classifica" Then
					Sql = "Select A.idUtente, A.idConcorso, A.Punti As Valore From Risultati A " &
						"Where A.idAnno = " & idAnno & " And A.idUtente=" & iid(0) & " And A.idConcorso <=" & idGiornata & " " &
						"Order By A.idConcorso"
				Else
					Sql = "Select A.idUtente, A.idConcorso, A." & Cosa & " As Valore From Risultati" & Altro & " A " &
						"Where A.idAnno = " & idAnno & " And A.idUtente=" & iid(0) & " And A.idConcorso <=" & idGiornata & " " &
						"Order By A.idConcorso"
				End If
			End If
			Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				' Ritorno = Rec
			Else
				Ritorno3 &= "{"
				If Val(idUtente) = Val(iid(0)) Then
					Ritorno3 &= """visible"": true,"
				Else
					Ritorno3 &= """visible"": false,"
				End If
				Ritorno3 &= """type"": ""line"","
				Ritorno3 &= """axisYType"": ""primary"","
				Ritorno3 &= """name"": """ & iid(1) & ""","
				Ritorno3 &= """showInLegend"": ""true"","
				Ritorno3 &= """dataPoints"": ["

				Dim Ritorno2 As String = ""
				Dim Quanto As Integer = 0

				Do Until Rec.Eof
					If Cosa = "Classifica" Then
						Quanto += Rec("Valore").Value
					Else
						Quanto = Rec("Valore").Value
					End If

					Ritorno2 &= "{"
					Ritorno2 &= Chr(34) & "x" & Chr(34) & ": " & Rec("idConcorso").Value & ","
					Ritorno2 &= Chr(34) & "y" & Chr(34) & ": " & Quanto
					Ritorno2 &= "},"

					Rec.MoveNext
				Loop
				If Ritorno2.Length > 0 Then
					Ritorno2 = Mid(Ritorno2, 1, Ritorno2.Length - 1)
				End If
				Ritorno3 &= Ritorno2
				Ritorno3 &= "]},"

				Rec.Close
			End If

		Next

		If Ritorno3.Length > 0 Then
			Ritorno3 = Mid(Ritorno3, 1, Ritorno3.Length - 1)
		End If
		Ritorno &= Ritorno3
		Ritorno &= "]}"

		Return Ritorno
	End Function

	Private Function RitornaRecord(Conn As Object, Connessione As String, idAnno As String, Campo As String) As String
		Dim RisultatoMigliore As String = ""
		Dim RisultatoPeggiore As String = ""
		Dim Ritorno As String = ""
		Dim P As Boolean = True

		Dim Sql As String = "SELECT B.idUtente, A.idConcorso, B.NickName, " & Campo & " As Valore FROM Risultati A " &
							"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
							"Where " & Campo & " > 0 *** Order By " & Campo & " Desc, A.idConcorso Limit 1"
		If idAnno <> "" Then
			Sql = Sql.Replace("***", "And A.idAnno=" & idAnno & " ")
		Else
			Sql = Sql.Replace("***", "")
		End If
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Not Rec.eof Then
				RisultatoMigliore = "{" & Chr(34) & "Titolo" & Chr(34) & ": " & Chr(34) & Campo & "Max" & Chr(34) & "," &
									Chr(34) & "idUtente" & Chr(34) & ": " & Chr(34) & Rec("idUtente").Value & Chr(34) & ", " &
									Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & Rec("NickName").Value & Chr(34) & ", " &
									Chr(34) & "Giornata" & Chr(34) & ": " & Rec("idConcorso").Value & ", " &
									Chr(34) & "Valore" & Chr(34) & ": " & Rec("Valore").Value & "," &
									Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & P & Chr(34) & " " &
									"}"
			End If
			Rec.Close
		End If
		P = Not P

		Sql = "SELECT B.idUtente, A.idConcorso, B.NickName, " & Campo & " As Valore FROM Risultati A " &
							"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
							"Where " & Campo & " > 0 *** Order By " & Campo & ", A.idConcorso Limit 1"
		If idAnno <> "" Then
			Sql = Sql.Replace("***", "And A.idAnno=" & idAnno & " ")
		Else
			Sql = Sql.Replace("***", "")
		End If
		Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Not Rec.eof Then
				RisultatoPeggiore = "{" & Chr(34) & "Titolo" & Chr(34) & ": " & Chr(34) & Campo & "Min" & Chr(34) & "," &
									Chr(34) & "idUtente" & Chr(34) & ": " & Chr(34) & Rec("idUtente").Value & Chr(34) & ", " &
									Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & Rec("NickName").Value & Chr(34) & ", " &
									Chr(34) & "Giornata" & Chr(34) & ": " & Rec("idConcorso").Value & ", " &
									Chr(34) & "Valore" & Chr(34) & ": " & Rec("Valore").Value & "," &
									Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & P & Chr(34) & " " &
									"}"
			End If
			Rec.Close
		End If

		Ritorno = RisultatoMigliore & "," & RisultatoPeggiore

		Return Ritorno
	End Function

	Private Function RitornaRecordAltro(Conn As Object, Connessione As String, idAnno As String, Campo As String) As String
		Dim RisultatoMigliore As String = ""
		Dim RisultatoPeggiore As String = ""
		Dim Ritorno As String = ""
		Dim P As Boolean = True

		Dim Sql As String = "SELECT B.idUtente, A.idConcorso, B.NickName, " & Campo & " As Valore FROM RisultatiAltro A " &
							"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
							"Where " & Campo & " > 0 *** Order By " & Campo & ", A.idConcorso Desc Limit 1"
		If idAnno <> "" Then
			Sql = Sql.Replace("***", "And A.idAnno=" & idAnno & " ")
		Else
			Sql = Sql.Replace("***", "")
		End If
		Dim Rec As Object = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Not Rec.eof Then
				RisultatoMigliore = "{" & Chr(34) & "Titolo" & Chr(34) & ": " & Chr(34) & Campo & "Max" & Chr(34) & "," &
									Chr(34) & "idUtente" & Chr(34) & ": " & Chr(34) & Rec("idUtente").Value & Chr(34) & ", " &
									Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & Rec("NickName").Value & Chr(34) & ", " &
									Chr(34) & "Giornata" & Chr(34) & ": " & Rec("idConcorso").Value & ", " &
									Chr(34) & "Valore" & Chr(34) & ": " & Rec("Valore").Value & "," &
									Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & P & Chr(34) & " " &
									"}"
			End If
			Rec.Close
		End If
		P = Not P

		Sql = "SELECT B.idUtente, A.idConcorso, B.NickName, " & Campo & " As Valore FROM RisultatiAltro A " &
			"Left Join Utenti B On A.idAnno = B.idAnno And A.idUtente = B.idUtente " &
			"Where " & Campo & " > 0 *** Order By " & Campo & ", A.idConcorso Limit 1"
		If idAnno <> "" Then
			Sql = Sql.Replace("***", "And A.idAnno=" & idAnno & " ")
		Else
			Sql = Sql.Replace("***", "")
		End If
		Rec = CreaRecordset(Server.MapPath("."), Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Not Rec.eof Then
				RisultatoPeggiore = "{" & Chr(34) & "Titolo" & Chr(34) & ": " & Chr(34) & Campo & "Min" & Chr(34) & "," &
									Chr(34) & "idUtente" & Chr(34) & ": " & Chr(34) & Rec("idUtente").Value & Chr(34) & ", " &
									Chr(34) & "NickName" & Chr(34) & ": " & Chr(34) & Rec("NickName").Value & Chr(34) & ", " &
									Chr(34) & "Giornata" & Chr(34) & ": " & Rec("idConcorso").Value & ", " &
									Chr(34) & "Valore" & Chr(34) & ": " & Rec("Valore").Value & "," &
									Chr(34) & "Pari" & Chr(34) & ": " & Chr(34) & P & Chr(34) & " " &
									"}"
			End If
			Rec.Close
		End If

		Ritorno = RisultatoMigliore & "," & RisultatoPeggiore

		Return Ritorno
	End Function
End Class