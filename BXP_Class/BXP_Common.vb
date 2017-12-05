Public Class BXP_Common

    Public Function Existdirectory(ByVal Nombre As String) As Boolean
        If System.IO.Directory.Exists(Nombre) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CreatDirectory(ByVal Nombre As String) As Boolean
        Try
            System.IO.Directory.CreateDirectory(Nombre)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function Path() As String
        Return Replace(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
    End Function

    Public Function Sanea(ByVal Texto As String) As String
        Return Replace(Texto, "'", "''")
    End Function

    Public Function Write(ByVal Texto As String, ByVal NombreArchivo As String, ByVal Directorio As String) As Boolean
        Dim File As System.IO.StreamWriter
        Dim PathCompleto As String
        Try

            PathCompleto = CStr(Path()) & "\" & CStr(Directorio)
            If Existdirectory(PathCompleto) = False Then
                CreatDirectory(PathCompleto)
            End If

            File = New System.IO.StreamWriter(Path() & "\" & Directorio & "\" & NombreArchivo, True)
            File.WriteLine(Texto)
            File.Flush()
            File.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function

    Public Function GetConfigValue(ByVal sPattern As String) As String
        Dim xmlDoc As System.Xml.XmlDocument
        Dim sRes As String = Nothing
        Try
            xmlDoc = New System.Xml.XmlDocument
            xmlDoc.Load(Path() & "\Config\Config.xml")
            If Not String.IsNullOrEmpty(xmlDoc.Item("Config").Item(sPattern).InnerText) Then
                sRes = xmlDoc.Item("Config").Item(sPattern).InnerText
            End If
        Catch ex As Exception
            Throw New Exception("GetConfigValue: " & ex.Message)
        Finally
            xmlDoc = Nothing
        End Try
        Return sRes
    End Function

End Class
