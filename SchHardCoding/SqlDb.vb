Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.IO
Imports System.Diagnostics

Public Class SqlDb

    '/ <summary>
    '/ update.xml的路径 
    '/ </summary>
    Public Shared strUpdateXmlPath As String = Application.StartupPath + "\update\conf\update.xml"

    Function GetconnectionString() As String
        Dim connectionString As String = ""
        'connectionString = Mod_Function.Decryption(System.Configuration.ConfigurationSettings.AppSettings("acmeConn"))   'SQL Server
        'connectionString = Mod_Function.Decrypt(ERP.Autoupdate.getConfigValue(strUpdateXmlPath, "ConnectionString"))
        '提案改善，資料登錄后資料庫的鏈接不能被修改
        connectionString = connection.con

        Return connectionString
    End Function

#Region "执行查询语句，返回DataSet"
    Public Function FillDataSet(ByVal SQLString As String) As DataSet
        Using connection As New SqlConnection(GetconnectionString())
            Dim ds As New DataSet()
            Try
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                End If
                connection.Open()
                Dim command As New SqlDataAdapter(SQLString, connection)
                command.SelectCommand.CommandTimeout = 0
                command.Fill(ds, "ds")
            Catch E As System.Data.SqlClient.SqlException
                Throw New Exception(E.Message)
            Finally
                connection.Close()
            End Try
            Return ds
        End Using
    End Function
#End Region

#Region "执行單條SQL语句"
    Public Sub ExecuteSql(ByVal sql As String)
        Using conn As New SqlConnection(GetconnectionString())
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            conn.Open()
            Dim cmmd As SqlCommand
            cmmd = New SqlCommand(sql, conn)
            cmmd.ExecuteNonQuery()
            conn.Close()
        End Using
    End Sub
#End Region

#Region "执行SQL语句返回值"
    Function GetExecSQL(ByVal Sql As String) As String
        GetExecSQL = 0
        Using conn As New SqlConnection(GetconnectionString())
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            conn.Open()
            Dim cmmd As SqlCommand
            cmmd = New SqlCommand(Sql, conn)
            GetExecSQL = cmmd.ExecuteScalar
            conn.Close()
        End Using
    End Function
#End Region

#Region "执行多条SQL语句，实现数据库事务"
    Public Sub ExecuteSqlTran(ByVal SQLStringList As ArrayList)
        Using conn As New SqlConnection(GetconnectionString())
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            conn.Open()
            Dim cmd As New SqlCommand()
            cmd.CommandTimeout = 0
            cmd.Connection = conn
            Dim tx As SqlTransaction = conn.BeginTransaction()
            cmd.Transaction = tx
            Try
                For n As Integer = 0 To SQLStringList.Count - 1
                    Dim strsql As String = SQLStringList(n).ToString()
                    If strsql.Trim().Length > 1 Then
                        cmd.CommandText = strsql
                        cmd.ExecuteNonQuery()
                    End If
                Next
                tx.Commit()
            Catch E As System.Data.SqlClient.SqlException
                tx.Rollback()
                Throw New Exception(E.Message)
            Finally
                conn.Close()
            End Try
        End Using
    End Sub
#End Region

#Region "執行帶參數的存儲過程"
    Public Sub ExecuteProcedure(ByVal storedProcName As String, ByVal cmdParms As SqlParameter())
        Using connection As New SqlConnection(GetconnectionString())
            If connection.State = ConnectionState.Open Then
                connection.Close()
            End If
            connection.Open()
            Dim command As SqlCommand = BuildQueryCommand(connection, storedProcName, cmdParms)
            command.CommandTimeout = 0
            command.ExecuteNonQuery()
            connection.Close()
        End Using
    End Sub
#End Region

#Region "执行存储过程返回結果集"
    Public Function RunProcedure(ByVal storedProcName As String, ByVal cmdParms As SqlParameter()) As DataSet
        Using connection As New SqlConnection(GetconnectionString())
            Dim ds As New DataSet()
            If connection.State = ConnectionState.Open Then
                connection.Close()
            End If
            connection.Open()
            Dim trans As SqlTransaction = connection.BeginTransaction()
            Dim da As New SqlDataAdapter()
            da.SelectCommand = BuildQueryCommand(connection, storedProcName, cmdParms)
            da.SelectCommand.Transaction = trans
            da.SelectCommand.CommandTimeout = 0
            Try
                da.Fill(ds)
            Catch ex As Exception
            Finally
                connection.Close()
            End Try
            Return ds
        End Using
    End Function
#End Region

#Region "执行存储过程返回結果集,不带属性 Cathy add at 20140317"
    Public Function RunProcedure(ByVal storedProcName As String) As DataSet
        Using connection As New SqlConnection(GetconnectionString())
            Dim ds As New DataSet()
            If connection.State = ConnectionState.Open Then
                connection.Close()
            End If
            connection.Open()
            Dim trans As SqlTransaction = connection.BeginTransaction()
            Dim da As New SqlDataAdapter()
            da.SelectCommand = BuildQueryCommand(connection, storedProcName)
            da.SelectCommand.Transaction = trans
            Try
                da.Fill(ds)
            Catch ex As Exception
            Finally
                connection.Close()
            End Try
            Return ds
        End Using
    End Function
#End Region

#Region "构建 SqlCommand 对象"
    Private Function BuildQueryCommand(ByVal connection As SqlConnection, ByVal storedProcName As String, ByVal ParamArray cmdParms As SqlParameter()) As SqlCommand
        Dim command As New SqlCommand(storedProcName, connection)
        command.CommandType = CommandType.StoredProcedure
        If Not cmdParms Is Nothing Then
            For Each parm As SqlParameter In cmdParms
                command.Parameters.Add(parm)
            Next
        End If
        Return command
    End Function
#End Region

#Region "報表調用"
    Public Function GetRptSet(ByVal SQLString As String, ByVal dt As String) As DataSet
        Using connection As New SqlConnection(GetconnectionString())
            Dim ds As New DataSet()
            Try
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                End If
                connection.Open()
                Dim command As New SqlDataAdapter(SQLString, connection)
                command.Fill(ds, dt)
            Catch E As System.Data.SqlClient.SqlException
                Throw New Exception(E.Message)
            Finally
                connection.Close()
            End Try
            Return ds
        End Using
    End Function
    Public Sub GetDataSet_byref(ByVal Sqlstr As String, ByRef ds As DataSet, ByVal tbname As String)
        Using connection As New SqlConnection(GetconnectionString())
            Try
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                End If
                connection.Open()
                Dim adp As New SqlDataAdapter(Sqlstr, connection)
                adp.SelectCommand.CommandTimeout = 0
                adp.Fill(ds, tbname)

            Catch E As System.Data.SqlClient.SqlException
                Throw New Exception(E.Message)
            Finally
                connection.Close()
            End Try

        End Using

    End Sub
#End Region

#Region "大批量數據快速導入"
    ''' <summary>
    ''' 注：sourceTable欄位（包括名稱\順序）須和數據庫中表完全一致
    ''' </summary>
    ''' <param name="sourceTable">來源Table</param>
    ''' <param name="destinationTable">目標Table</param>
    ''' <param name="connectionString">連接字符串</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ExecuteBulkCopy(ByRef sourceTable As DataTable, ByVal destinationTable As String, ByVal connectionString As String) As Boolean
        Using conn As New SqlConnection(connectionString)
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            conn.Open()
            Dim sqlTran As SqlTransaction = conn.BeginTransaction()
            Using sbc As SqlBulkCopy = New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, sqlTran)
                sbc.BatchSize = 1000
                sbc.BulkCopyTimeout = 180
                '将DataTable表名作为待导入库中的目标表名         
                sbc.DestinationTableName = destinationTable
                Try
                    sbc.WriteToServer(sourceTable)
                    sqlTran.Commit()
                    Return True
                Catch ex As Exception
                    sqlTran.Rollback()
                    Throw New Exception(ex.Message)
                End Try

            End Using

        End Using
    End Function
#End Region

#Region "执行帶output類型參數的存储过程，返回output字典"
    Public Function RunProcedureOutPut(ByVal storedProcName As String, ByVal cmdParms As SqlParameter()) As Dictionary(Of String, Object)
        Dim dict As Dictionary(Of String, Object) = New Dictionary(Of String, Object)
        Using connection As New SqlConnection(GetconnectionString())
            If connection.State = ConnectionState.Open Then
                connection.Close()
            End If
            connection.Open()
            Dim command As SqlCommand = BuildQueryCommand(connection, storedProcName, cmdParms)
            command.ExecuteNonQuery()
            For index = 0 To cmdParms.Count - 1
                If cmdParms(index).Direction = ParameterDirection.Output Then
                    dict.Add(cmdParms(index).ParameterName, command.Parameters(cmdParms(index).ParameterName).Value)
                End If
            Next
            Return dict
        End Using
    End Function
#End Region
#Region "查詢表單簽核对象"
    Public Shared Function Query(ByVal SQLString As String) As DataSet
        Dim connectionString As String = connection.con
        Using connection As New SqlConnection(connectionString)
            Dim ds As New DataSet()
            connection.Open()
            Dim command As New SqlDataAdapter(SQLString, connection)
            command.Fill(ds, "ds")
            Return ds
        End Using
    End Function
#End Region
End Class
