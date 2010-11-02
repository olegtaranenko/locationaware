Public Class ConnectionClass

    Sub New()
        
    End Sub

    'connect string 
	Friend Const ConnectionString As String = "packet size=4096;" & _
			   "integrated security=SSPI;data source=localhost;persist security info=False;initial catalog=pim;" & _
			   "Connection Lifetime=300"

    Friend SQLConnection As System.Data.SqlClient.SqlConnection

    'get connect to the sql server 
    Public Function GetSQLConnection() As Boolean
        Try
            If SQLConnection Is Nothing Then
                SQLConnection = New System.Data.SqlClient.SqlConnection(ConnectionString)
            End If

            If SQLConnection.State = ConnectionState.Closed Then
                SQLConnection.Open()
                Return True
            Else
                Return True
            End If

        Catch excp As Exception
            Return False
            Exit Function
        End Try

    End Function

    'close connect to the sql server 
    Function CloseSQLConnection() As Boolean
        If Not SQLConnection Is Nothing Then
            If SQLConnection.State = ConnectionState.Open Then
                SQLConnection.Close()
                SQLConnection = Nothing
                Return True
            End If
        End If
    End Function

End Class
