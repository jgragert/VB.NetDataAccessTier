Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks
Imports System.Text
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient

Public Class SharedProcedures

    Public Shared Function GetDbConnectionString() As String

        Dim strbldrConnectionString As StringBuilder = New StringBuilder()

        strbldrConnectionString.Append(" Data Source=" & System.Configuration.ConfigurationManager.AppSettings("DatabaseServerIpOrName") & ";")
        strbldrConnectionString.Append(" Initial Catalog=" & System.Configuration.ConfigurationManager.AppSettings("DatabaseInitialCatalog") & ";")
        strbldrConnectionString.Append(" User Id=" & System.Configuration.ConfigurationManager.AppSettings("DatabaseUserId") & ";" _
            & " Password=" & System.Configuration.ConfigurationManager.AppSettings("DatabasePassword") & ";")

        Return strbldrConnectionString.ToString()

    End Function

    Public Shared Function GetGlobalValueFromDb(instrValueName As String) As String

        Dim strValue As String = ""
        Dim dsGlobalValues As DataSet

        dsGlobalValues = GetDataSetForDynamicSql("SELECT Value FROM GlobalValues WHERE ValueName LIKE '" + instrValueName + "'")
        strValue = dsGlobalValues.Tables(0).Rows(0)("Value").ToString()

        Return strValue

    End Function

    'NOTE: This procedure is intended for simple, fast-performing queries and that do not include user-entry (so that sql injection is not possible) only.
    Public Shared Function GetDataSetForDynamicSql(instrSql As String) As DataSet
        Dim dsResultSet As DataSet = New DataSet()

        Using conDbConnection As SqlConnection = New SqlConnection()
            Using cmdSqlCommand As SqlCommand = New SqlCommand()
                Using Adapter As SqlDataAdapter = New SqlDataAdapter()

                    cmdSqlCommand.CommandType = CommandType.Text
                    cmdSqlCommand.CommandText = instrSql
                    conDbConnection.ConnectionString = GetDbConnectionString()
                    cmdSqlCommand.Connection = conDbConnection
                    Adapter.SelectCommand = cmdSqlCommand
                    Adapter.Fill(dsResultSet, "dummy")

                End Using
            End Using
        End Using

        Return dsResultSet

    End Function

    'NOTE: This is intended for simple, fast-performing queries only, that have no user-entered values (ie, no sql injection risk). 
    'More complex (and non-select queries including user entry) sql should be executed via stored procedure.
    Public Shared Sub ExecuteNonQueryForDynamicSql(instrSql As String)

        Dim dsResultSet As DataSet = New DataSet()
        Dim conDbConnection As SqlConnection = New SqlConnection()
        Dim cmdSqlCommand As SqlCommand = New SqlCommand()

        cmdSqlCommand.CommandType = CommandType.Text
        cmdSqlCommand.CommandText = instrSql

        conDbConnection.ConnectionString = GetDbConnectionString()
        conDbConnection.Open()

        cmdSqlCommand.Connection = conDbConnection
        cmdSqlCommand.ExecuteNonQuery()
        conDbConnection.Close()
        conDbConnection.Dispose()
        cmdSqlCommand.Dispose()

    End Sub

    'NOTE: This is intended for simple, fast-performing, scalar queries only. 
    'More complex (and non-select queries, to prevent sql injection) sql is executed via stored procedure.
    Public Shared Function GetScalarAsStringForDynamicSql(instrSql As String) As String
        Dim strReturnValue As String = ""
        Dim dsResultSet As DataSet = New DataSet()
        Dim conDbConnection As SqlConnection = New SqlConnection()
        Dim cmdSqlCommand As SqlCommand = New SqlCommand()
        Dim Adapter As SqlDataAdapter = New SqlDataAdapter()

        cmdSqlCommand.CommandType = CommandType.Text
        cmdSqlCommand.CommandText = instrSql

        conDbConnection.ConnectionString = GetDbConnectionString()

        cmdSqlCommand.Connection = conDbConnection
        Adapter.SelectCommand = cmdSqlCommand
        Adapter.Fill(dsResultSet, "dummy")

        Try
            strReturnValue = Convert.ToString(dsResultSet.Tables(0).Rows(0)(0))
        Catch ex As Exception
            'Take no action, strReturnValue retains empty string if value not found.
        End Try

        Return strReturnValue

    End Function

    Public Shared Function GetDataSetForSp(instrStoredProcedureName As String, lstParameterCollection As List(Of SqlParameter)) As DataSet

        Dim dsResultSet As DataSet = New DataSet()

        Using conDbConnection As SqlConnection = New SqlConnection()
            Using cmdSqlCommand As SqlCommand = New SqlCommand()
                Using Adapter As SqlDataAdapter = New SqlDataAdapter()

                    cmdSqlCommand.CommandType = CommandType.StoredProcedure
                    cmdSqlCommand.CommandText = instrStoredProcedureName
                    For Each param As SqlParameter In lstParameterCollection
                        cmdSqlCommand.Parameters.Add(param)
                    Next
                    conDbConnection.ConnectionString = GetDbConnectionString()
                    cmdSqlCommand.Connection = conDbConnection
                    Adapter.SelectCommand = cmdSqlCommand
                    Adapter.Fill(dsResultSet, "dummy")

                End Using
            End Using
        End Using

        Return dsResultSet

    End Function

    Public Shared Function ExecuteInsertReturnIdentityForSp(instrStoredProcedureName As String, lstParameterCollection As List(Of SqlParameter)) As Int64

        Dim int64InsertIdentity As Int64 = 0
        Dim int32NumberOfRowsAffected As Int32 = 0
        Dim conDbConnection As SqlConnection = New SqlConnection()
        Dim cmdSqlCommand As SqlCommand = New SqlCommand()

        cmdSqlCommand.CommandType = CommandType.StoredProcedure
        cmdSqlCommand.CommandText = instrStoredProcedureName
        For Each param As SqlParameter In lstParameterCollection
            cmdSqlCommand.Parameters.Add(param)
        Next

        conDbConnection.ConnectionString = GetDbConnectionString()
        conDbConnection.Open()
        cmdSqlCommand.Connection = conDbConnection
        int32NumberOfRowsAffected = cmdSqlCommand.ExecuteNonQuery()

        'If insert succeeded, retrieve identity value.
        If int32NumberOfRowsAffected > 0 Then
            int64InsertIdentity = Convert.ToInt64(cmdSqlCommand.Parameters("@outint64IdentityValue").Value)
        End If

        conDbConnection.Close()
        conDbConnection.Dispose()
        cmdSqlCommand.Dispose()

        Return int64InsertIdentity

    End Function

    Public Shared Sub ExecuteNonQueryForSp(instrStoredProcedureName As String, lstParameterCollection As List(Of SqlParameter))

        Dim conDbConnection As SqlConnection = New SqlConnection()
        Dim cmdSqlCommand As SqlCommand = New SqlCommand()

        cmdSqlCommand.CommandType = CommandType.StoredProcedure
        cmdSqlCommand.CommandText = instrStoredProcedureName
        For Each param As SqlParameter In lstParameterCollection
            cmdSqlCommand.Parameters.Add(param)
        Next

        conDbConnection.ConnectionString = GetDbConnectionString()
        conDbConnection.Open()
        cmdSqlCommand.Connection = conDbConnection
        cmdSqlCommand.ExecuteNonQuery()

        conDbConnection.Close()
        conDbConnection.Dispose()
        cmdSqlCommand.Dispose()

    End Sub

    Public Shared Function GetScalarAsStringForSp(instrStoredProcedureName As String, lstParameterCollection As List(Of SqlParameter)) As String

        Dim dsResultSet As DataSet = New DataSet()
        Dim decReturnValue As Decimal = 0
        Dim strReturnValue As String = ""

        Using conDbConnection As SqlConnection = New SqlConnection()
            Using cmdSqlCommand As SqlCommand = New SqlCommand()

                cmdSqlCommand.CommandType = CommandType.StoredProcedure
                cmdSqlCommand.CommandText = instrStoredProcedureName
                conDbConnection.ConnectionString = GetDbConnectionString()
                cmdSqlCommand.Connection = conDbConnection
                For Each param As SqlParameter In lstParameterCollection
                    cmdSqlCommand.Parameters.Add(param)
                Next
                cmdSqlCommand.Connection.Open()
                decReturnValue = Convert.ToDecimal(cmdSqlCommand.ExecuteScalar())
                strReturnValue = Convert.ToString(decReturnValue)
                Return strReturnValue

            End Using
        End Using

    End Function

End Class
