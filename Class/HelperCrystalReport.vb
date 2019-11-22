Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class HelperCrystalReport

 

    Public Shared Sub SetDBLogonForReport(ByVal myReportDocument As ReportDocument, ByVal ConnectionString As String)


        Dim sqlbuilder As New System.Data.SqlClient.SqlConnectionStringBuilder
        Dim myConnectionInfo As ConnectionInfo = New ConnectionInfo()


        'Get SqlConnectionStringBuilder
        sqlbuilder.ConnectionString = ConnectionString


        'Get ConnectionInfo Info
        If sqlbuilder.IntegratedSecurity = True Then
            myConnectionInfo.DatabaseName = sqlbuilder.InitialCatalog
            myConnectionInfo.IntegratedSecurity = True

        Else
            myConnectionInfo.DatabaseName = sqlbuilder.InitialCatalog
            myConnectionInfo.UserID = sqlbuilder.UserID
            myConnectionInfo.Password = sqlbuilder.Password
        End If

        'Set Report Document ConnectionInfo
        Dim myTables As Tables = myReportDocument.Database.Tables
        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            Dim myTableLogonInfo As TableLogOnInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next


    End Sub

    Public Shared Function crformulajoin(ByVal separator As String, ByVal formula1 As String, ByVal formula2 As String) As String
        If formula1.Trim = "" Then
            Return formula2
        ElseIf formula2.Trim = "" Then
            Return formula1
        Else
            Return String.Join(separator, formula1, formula2)
        End If

    End Function

    Public Shared Function SetMultiplePamaValues(Parameter As ParameterFieldDefinition, ValuesList As Object())
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        Parameter.EnableAllowMultipleValue = True
        crParameterValues = Parameter.CurrentValues
        crParameterValues.Clear()

        For Each value As Object In ValuesList
            crParameterDiscreteValue = New ParameterDiscreteValue
            crParameterDiscreteValue.Value = value
            crParameterValues.Add(crParameterDiscreteValue)
        Next

        Parameter.ApplyCurrentValues(crParameterValues)

        Return Parameter
    End Function

    Public Shared Function SetMultiplePamaValues(ByVal Parameter As ParameterFieldDefinition, ByVal datatable As DataTable, ByVal ColumnName As String)
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        Parameter.EnableAllowMultipleValue = True
        crParameterValues = Parameter.CurrentValues
        crParameterValues.Clear()

        For Each dr As DataRow In datatable.Rows

            crParameterDiscreteValue = New ParameterDiscreteValue
            crParameterDiscreteValue.Value = dr(ColumnName)
            crParameterValues.Add(crParameterDiscreteValue)
        Next

        Parameter.ApplyCurrentValues(crParameterValues)

        Return Parameter
    End Function

    Public Shared Function SetMultiplePamaValues(ByVal Parameter As ParameterFieldDefinition, Dataset As DataSet, ByVal TableName As String, ByVal ColumnName As String)
        Return SetMultiplePamaValues(Parameter, Dataset.Tables(TableName), ColumnName)
    End Function



End Class
