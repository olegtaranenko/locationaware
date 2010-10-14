
Public Class DBConnectionStringBuilder

    Private configService As ConfigService

    Sub New(ByRef configService As configService)
        Me.configService = configService
    End Sub


    Public Function buildConnStr() As String

        Dim uid As String = configService.getDBUser
        Dim pwd As String = configService.getDBPassword
        Dim server As String = configService.getDBServer
        Dim db As String = configService.getDBName
        Dim connLifetime As Long = configService.getConnectionLifetime
        Dim persistSecInfo As Boolean = configService.isPersistSecurityInfo
        Dim integratedSec As Boolean = configService.isIntegratedSecurity

        Return build(uid, pwd, server, db, connLifetime, integratedSec, persistSecInfo)
    End Function



    Public Function build(ByVal uid, ByVal pwd, ByVal server, ByVal db, ByVal connLifetime, ByVal integratedSec, ByVal persistSecInfo) As String
        Dim connString As String
		connString = "initial catalog=" & db & ";Connection Lifetime=" & connLifetime & _
					 ";persist security info=" & persistSecInfo & ";data source=" & server _
					 & ";Network Library=DBMSSOCN"


        If (integratedSec) Then
            connString = connString & ";integrated security=SSPI"
        Else
            connString = connString & ";integrated security=false;uid=" & uid & ";pwd=" & pwd
        End If

        Return connString
    End Function


    Public Function setConfigService(ByRef configService As configService)
        Me.configService = configService
    End Function

End Class
