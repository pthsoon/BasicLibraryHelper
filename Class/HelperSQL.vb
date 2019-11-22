Imports System.Configuration

Public Class HelperSQL

    Public Shared Function GetConnectionString(ByVal CONNECTIONSTRINGKEY As String) As String
        Return ConfigurationManager.ConnectionStrings(CONNECTIONSTRINGKEY).ConnectionString()
    End Function

    Public Function GetConnectionString(ByVal DataSource As String, ByVal InitialCatalog As String) As String
        Dim sqlbuilder As New System.Data.SqlClient.SqlConnectionStringBuilder
        sqlbuilder.DataSource = DataSource
        sqlbuilder.InitialCatalog = InitialCatalog
        sqlbuilder.IntegratedSecurity = True

        Return sqlbuilder.ConnectionString
    End Function

    Public Function GetConnectionString(ByVal DataSource As String, ByVal InitialCatalog As String, ByVal UserId As String, ByVal Password As String) As String
        Dim sqlbuilder As New System.Data.SqlClient.SqlConnectionStringBuilder
        sqlbuilder.DataSource = DataSource
        sqlbuilder.InitialCatalog = InitialCatalog
        sqlbuilder.IntegratedSecurity = False
        sqlbuilder.UserID = UserId
        sqlbuilder.Password = Password


        Return sqlbuilder.ConnectionString
    End Function


End Class
