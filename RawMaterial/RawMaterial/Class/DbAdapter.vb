Imports Npgsql

Public Class DbAdapter
    Public Property userid As String
    Public Property password As String
    Private Shared Property myInstance As DbAdapter

    Public Shared Function getInstance() As DbAdapter
        If myInstance Is Nothing Then
            myInstance = New DbAdapter
        End If
        Return myInstance
    End Function

    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date)
        Dim result As Object
        Dim builder As New NpgsqlConnectionStringBuilder()
        builder.ConnectionString = "host=hon14nt;port=5432;database=LogisticDb;commandTimeout=1000;Timeout=1000;"
        builder.Add("User Id", _userid)
        builder.Add("password", _password)
        Dim Connectionstring = builder.ConnectionString
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = _userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function
End Class
