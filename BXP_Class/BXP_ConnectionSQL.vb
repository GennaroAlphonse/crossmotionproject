Imports System.Data
Imports System.Data.SqlClient

Public Class BXP_ConnectionSQL

    Private Function ConvertirTableASet(ByVal Tabla As DataTable) As DataSet
        Dim myDataSet = New DataSet()
        myDataSet.Tables.Add(Tabla)
        Return myDataSet
    End Function

    Public Function Query1(ByVal Query As String) As DataSet
        Dim Con As SqlConnection

        Try
            Con = New SqlConnection(GetConnString)

            Con.Open()

            Dim cmd = New SqlCommand(Query)
            cmd.Connection = Con
            Dim dr As SqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            Dim dt As New DataTable
            dt.Load(dr)
            Return ConvertirTableASet(dt)

        Catch ex As Exception
            Dim Log As New BXP_Log
            Log.ErrorLog("Query1 - " + ex.Message + " - Query = " + Query)
            Return Nothing
        Finally
            If Con.State = ConnectionState.Open Then
                Con.Close()
                Con.Dispose()
                Con = Nothing
                GC.Collect()
            End If
        End Try
    End Function

    Public Function Query2(ByVal Query As String) As String
        Dim Con As SqlConnection

        Try
            Con = New SqlConnection(GetConnString)

            Con.Open()

            Dim cmd = New SqlCommand(Query)
            cmd.Connection = Con
            Return cmd.ExecuteScalar

        Catch ex As Exception
            'Dim Log As New BXP_Log
            'Log.ErrorLog("Query2 - " + ex.Message + " - Query = " + Query)
            Return Nothing
        Finally
            If Con.State = ConnectionState.Open Then
                Con.Close()
                Con.Dispose()
                Con = Nothing
                GC.Collect()
            End If
        End Try
    End Function

    Public Function Query3(ByVal Query As String) As Boolean
        Dim Con As SqlConnection

        Try
            Con = New SqlConnection(GetConnString)

            Con.Open()

            Dim cmd = New SqlCommand(Query)
            cmd.Connection = Con
            If cmd.ExecuteNonQuery Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            'Dim Log As New BXP_Log
            'Log.ErrorLog("Query3 - " + ex.Message + " - Query = " + Query)

            Return False
        Finally
            If Con.State = ConnectionState.Open Then
                Con.Close()
                Con.Dispose()
                Con = Nothing
                GC.Collect()
            End If
        End Try
    End Function

    Public Function GetConnString() As String
        Try
            Dim FileConfig As New BXP_ObjFileConfig
            If FileConfig.TrustedConection = False Then
                GetConnString = "Data Source = " & Trim(FileConfig.Server) & "; Initial Catalog = " & Trim(FileConfig.DB) & "; " _
                              & "User Id = " & Trim(FileConfig.User) & "; Password = " & Trim(FileConfig.Pass) & "; " _
                              & "Trusted_Connection = False; "
            Else
                GetConnString = ""
            End If
        Catch ex As Exception
            'Dim Log As New BXP_Log
            'Log.ErrorLog("GetConnString - " + ex.Message + " - Conn = " + GetConnString)
            GetConnString = ""
        End Try
    End Function

End Class
