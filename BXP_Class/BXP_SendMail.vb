Public Class BXP_SendMail

    Public Property From As String
    Public Property Pass As String
    Public Property Server As String
    Public Property Port As Integer
    Public Property EnableSSl As Boolean

    'Contructor
    Public Sub New(ByVal sFrom As String, _
                   ByVal sPass As String, _
                   ByVal sServer As String, _
                   ByVal iPort As Integer, _
                   ByVal bSSL As Boolean)

        Me.From = sFrom
        Me.Pass = sPass
        Me.Server = sServer
        Me.Port = iPort
        Me.EnableSSl = bSSL

    End Sub

    Public Function SendMail(ByVal sTo As String, _
                             ByVal sSubject As String, _
                             ByVal sBody As String, _
                             Optional ByVal sFile As String = "" _
                                                              ) As Boolean
        Dim SmtpServer As New Net.Mail.SmtpClient()
        Dim mail As New Net.Mail.MailMessage()

        Try
            SmtpServer.Credentials = New Net.NetworkCredential(Me.From, Me.Pass)
            SmtpServer.Port = Me.Port
            SmtpServer.Host = Me.Server
            SmtpServer.EnableSsl = Me.EnableSSl
        Catch ex As Exception
            Dim Log As New BXP_Log
            Log.ErrorLog("Error al enviar correo 1 " + ex.Message)
            Return False
        End Try

        Try
            mail = New Net.Mail.MailMessage()
            mail.From = New Net.Mail.MailAddress(Me.From)
            mail.To.Add(sTo)
            mail.Subject = sSubject
            mail.Body = sBody
            mail.IsBodyHtml = False
        Catch ex As Exception
            Dim Log As New BXP_Log
            Log.ErrorLog("Error al enviar correo 2 " + ex.Message)
            Return False
        End Try

        Try
            If Not String.IsNullOrEmpty(sFile) Then
                Dim ms As New IO.MemoryStream()
                mail.Attachments.Add(New Net.Mail.Attachment(sFile))
                Try
                    SmtpServer.Send(mail)
                Catch ex As Exception
                    Return False
                End Try

                ms.Close()
                ms.Dispose()
                ms = Nothing
                mail.Attachments.Clear()
                GC.Collect()
                Return True
            Else
                SmtpServer.Send(mail)
                SmtpServer.Dispose()
                GC.Collect()
                Return True
            End If
        Catch ex As System.Net.Mail.SmtpException
            Dim Log As New BXP_Log
            Log.ErrorLog("Error al enviar correo 3 " + ex.Message)
            Return False
        End Try

    End Function
End Class
