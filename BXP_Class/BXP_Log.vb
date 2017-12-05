Public Class BXP_Log

    Public Function ErrorLog(ByVal Err As String) As Boolean
        Dim FileManagement As New BXP_Common

        Dim StrError As String = ""
        Try

            Dim Ws As New BXP_WebServiceSQL

            StrError = DateTime.Now.ToString("dd/MM/yyyy") + " - " + DateTime.Now.ToString("HH:mm:ss tt") + " - " + Err
            'FileManagement.Write(StrError, "Error_" + DateTime.Now.ToString("ddMMyyyy") + ".txt", "ErrorLog")
            Dim Query As String = "INSERT INTO BXP_Polimeros.dbo.BXP_Error (Error, Fecha, Hora, Usuario) VALUES ('" + FileManagement.Sanea(Err) + "', '" + DateTime.Now.ToString("dd/MM/yyyy") + "', '" + DateTime.Now.ToString("HH:mm:ss tt") + "', '" + "')"

            Ws.Query3(Query)

            Return True
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function

    Public Function EventLog(ByVal Evento As String) As Boolean
        Dim FileManagement As New BXP_Common

        Dim StrError As String = ""
        Try

            Dim Ws As New BXP_WebServiceSQL

            StrError = DateTime.Now.ToString("dd/MM/yyyy") + " - " + DateTime.Now.ToString("HH:mm:ss tt") + " - " + Evento
            'FileManagement.Write(StrError, "Event_" + DateTime.Now.ToString("ddMMyyyy") + ".txt", "EventLog")

            Dim Query As String = "INSERT INTO BXP_Polimeros.dbo.BXP_Event (Error, Fecha, Hora, Usuario) VALUES ('" + FileManagement.Sanea(Evento) + "', '" + DateTime.Now.ToString("dd/MM/yyyy") + "', '" + DateTime.Now.ToString("HH:mm:ss tt") + "', '" + "')"

            Ws.Query3(Query)

            Return True
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function

    Public Function EventLog(ByVal Evento As String, ByVal Usuario As String) As Boolean
        Dim FileManagement As New BXP_Common

        Dim StrError As String = ""
        Try

            Dim Ws As New BXP_WebServiceSQL

            StrError = DateTime.Now.ToString("dd/MM/yyyy") + " - " + DateTime.Now.ToString("HH:mm:ss tt") + " - " + Evento
            'FileManagement.Write(StrError, "Event_" + DateTime.Now.ToString("ddMMyyyy") + "_" + CStr(Usuario) + ".txt", "EventLog")

            Dim Query As String = "INSERT INTO BXP_Polimeros.dbo.BXP_Event (Error, Fecha, Hora, Usuario) VALUES ('" + FileManagement.Sanea(Evento) + "', '" + DateTime.Now.ToString("dd/MM/yyyy") + "', '" + DateTime.Now.ToString("HH:mm:ss tt") + "', '" + Usuario + "')"

            Ws.Query3(Query)

            Return True
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function

End Class
