Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Timers

Public Class mail
	Public Function SendEmail(Mp As String, Destinatario As String, ByVal oggetto As String, ByVal newBody As String, ByVal Allegato() As String) As String
		Dim Ritorno As String = "*"
		Dim s As New strutturaMail
		s.Destinatario = Destinatario
		s.Oggetto = oggetto
		s.newBody = newBody
		s.Allegato = Allegato

		pathMail = Mp & "/Logs/"
		path1 = Mp & "/"

		listaMails.Add(s)

		If effettuaLogMail Then
			Dim gf As New GestioneFilesDirectory
			gf.CreaDirectoryDaPercorso(pathMail)
			nomeFileLogmail = pathMail & "logMail_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"
		End If

		avviaTimer()

		Return Ritorno
	End Function

	Private Sub avviaTimer()
		If timerMails Is Nothing Then
			timerMails = New Timer(5000)
			AddHandler timerMails.Elapsed, New ElapsedEventHandler(AddressOf scodaMessaggi)
			timerMails.Start()

			If effettuaLogMail And nomeFileLogmail <> "" Then
				Dim gf As New GestioneFilesDirectory
				Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

				gf.ApreFileDiTestoPerScrittura(nomeFileLogmail)
				gf.ScriveTestoSuFileAperto(Datella & " - Timer avviato. Mail da scodare: " & listaMails.Count)
				gf.ChiudeFileDiTestoDopoScrittura()
			End If
		End If
	End Sub

	Private Sub scodaMessaggi()
		timerMails.Enabled = False
		Dim mail As strutturaMail = listaMails.Item(0)

		Dim gf As New GestioneFilesDirectory
		If effettuaLogMail Then
			Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

			gf.ApreFileDiTestoPerScrittura(nomeFileLogmail)
			gf.ScriveTestoSuFileAperto(Datella & " - Scodo Mail: " & mail.Destinatario & "/" & mail.Oggetto)
			gf.ChiudeFileDiTestoDopoScrittura()
		End If

		Dim Ritorno As String = SendEmailAsincrona(mail.Destinatario, mail.Oggetto, mail.newBody, mail.Allegato, gf)
		listaMails.RemoveAt(0)
		If listaMails.Count > 0 Then
			timerMails.Enabled = True
		Else
			timerMails = Nothing
			listaMails = New List(Of strutturaMail)
		End If
	End Sub

	Private Function SendEmailAsincrona(Destinatario As String, ByVal oggetto As String, ByVal newBody As String,
										ByVal Allegato() As String,
										gf As GestioneFilesDirectory) As String
		'Dim myStream As StreamReader = New StreamReader(Server.MapPath(ConfigurationManager.AppSettings("VirtualDir") & "mailresponsive.html"))
		'Dim newBody As String = ""
		'newBody = myStream.ReadToEnd()
		'newBody = newBody.Replace("$messaggioemail", body)
		'myStream.Close()
		'myStream.Dispose()

		Dim Ritorno As String = ""
		Dim mail As MailMessage = New MailMessage()
		Dim Credenziali As String = gf.LeggeFileIntero(path1 & "CredenzialiPosta.txt")
		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

		If effettuaLogMail Then
			gf.ApreFileDiTestoPerScrittura(nomeFileLogmail)
			gf.ScriveTestoSuFileAperto(Datella & " - Inizio")
		End If

		Try
			Dim cr() As String = Credenziali.Split(";")
			Dim Utenza As String = cr(0)
			Dim Password As String = cr(1).Replace(vbCrLf, "")

			If effettuaLogMail Then
				gf.ScriveTestoSuFileAperto(Datella & " - Inizio 1")
			End If

			'Dim newNewBody As String = ""
			'Dim c() As String = newBody.Split(";")
			'For Each cc As String In c
			'	If cc <> "" Then
			'		newNewBody &= Chr(cc)
			'	End If
			'Next

			mail.From = New MailAddress("looigi@gmail.com")
			mail.[To].Add(New MailAddress(Destinatario))
			mail.Subject = oggetto
			mail.IsBodyHtml = True
			If newBody <> "" Then
				mail.Body = newBody ' CreaCorpoMail(Squadra, mail, newBody)
			Else
				mail.Body = ""
			End If

			If effettuaLogMail Then
				gf.ScriveTestoSuFileAperto(Datella & " - Inizio 2")
			End If

			mail.Body &= "<br><hr />"

			'Dim Data As Attachment = Nothing
			'If Allegato.Length > 0 Then
			'	For Each All As String In Allegato
			'		If All <> "" Then
			'			Dim Allegatone As String = All
			'			Dim paths As String = ""

			'			If effettuaLogMail Then
			'				gf.ScriveTestoSuFileAperto(Datella & " - Aggiunge Allegato: " & Allegatone)
			'			End If

			'			Data = New Attachment(Allegatone, MediaTypeNames.Application.Octet)
			'			Dim disposition As ContentDisposition = Data.ContentDisposition
			'			disposition.CreationDate = System.IO.File.GetCreationTime(Allegatone)
			'			disposition.ModificationDate = System.IO.File.GetLastWriteTime(Allegatone)
			'			disposition.ReadDate = System.IO.File.GetLastAccessTime(Allegatone)
			'			mail.Attachments.Add(Data)
			'		End If

			'		If effettuaLogMail Then
			'			gf.ScriveTestoSuFileAperto(Datella & " - Inizio 2-1")
			'		End If
			'	Next
			'End If

			If effettuaLogMail Then
				gf.ScriveTestoSuFileAperto(Datella & " - Inizio 3")
			End If
			'mail.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8")
			'Dim plainView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(System.Text.RegularExpressions.Regex.Replace(newBody, "< (.|\n) *?>", String.Empty), Nothing, "text/plain")
			'Dim htmlView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(newBody, Nothing, "text/html")
			'mail.AlternateViews.Add(plainView)
			'mail.AlternateViews.Add(htmlView)
			Dim smtpClient As SmtpClient = New SmtpClient("smtp.gmail.com")

			smtpClient.EnableSsl = True
			smtpClient.Port = 25
			smtpClient.UseDefaultCredentials = False
			smtpClient.Credentials = New System.Net.NetworkCredential(Utenza, Password)
			smtpClient.Send(mail)
			smtpClient = Nothing

			If effettuaLogMail Then
				gf.ScriveTestoSuFileAperto(Datella & " - Invio in corso")
			End If

			'If Allegato.Length > 0 And Not Data Is Nothing Then
			'	Try
			'		Data.Dispose()
			'	Catch ex As Exception

			'	End Try
			'End If

			Ritorno = "*"
			If effettuaLogMail Then
				gf.ScriveTestoSuFileAperto(Datella & " - Invio effettuato")
			End If
		Catch ex As Exception
			Ritorno = "ERROR: " & ex.Message

			If effettuaLogMail Then
				gf.ScriveTestoSuFileAperto(Datella & " - Errore nell'invio: " & ex.Message)
			End If
		End Try
		'smtpClient.Dispose()

		If effettuaLogMail Then
			gf.ScriveTestoSuFileAperto(Datella & "-----------------------------------------------------------------")
			gf.ScriveTestoSuFileAperto(Datella & "")
			gf.ChiudeFileDiTestoDopoScrittura()
		End If

		Return Ritorno
	End Function

End Class
