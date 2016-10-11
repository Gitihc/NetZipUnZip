Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Data.OleDb
Imports System.Collections
Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions
Imports System.Data.Common
Public NotInheritable Class DBUtil
    'SqlServer
    Shared dbType As String = "" 'ConfigurationManager.AppSettings("dbType") ' "SqlServer" '"Oracle","MySql"
    Private Shared conn As DbConnection
    Private Shared ReadOnly Property connectionString() As String
        Get
            Select Case dbType
                Case "SqlServer"
                    Return "server=192.168.2.10;user id=sa;pwd=123456;database=MyWebData;MultipleActiveResultSets=true" 'ConfigurationManager.AppSettings("connString")
                Case "Oracle"
                    Return "Provider=OraOLEDB.Oracle.1;Data Source=XE;User Id=plus;Password=sa"
                Case "MySql"
                    Return "server=localhost; user id=sa; password=sa; database=plusoft_test;"
                Case Else
                    Return "server=localhost; database=plusoft_test; Integrated Security=True; "
            End Select
        End Get
    End Property
    Public Shared Function getConn() As DbConnection
        If IsNothing(conn) OrElse conn.State <> ConnectionState.Open Then
            If (dbType = "MySql") Then
                conn = New MySqlConnection(connectionString)
            ElseIf (dbType = "Oracle") Then
                conn = New OleDbConnection(connectionString)
            ElseIf (dbType = "SqlServer") Then
                conn = New SqlConnection(connectionString)
            End If
            conn.Open()
        End If
        Return conn
    End Function

    ''' <summary>
    ''' 执行数据数据查询，并返回数据集
    ''' </summary>
    ''' <remarks>传入要执行的SQL语句</remarks>
    Public Shared Function ExecuteHasQuery(ByVal strSql As String) As DataTable
        Dim con As DbConnection = getConn()
        If IsNothing(con) Then Return Nothing
        Dim tbl As DataTable = New DataTable()
        If dbType = "MySql" Then
            Using cmd As MySqlCommand = New MySqlCommand(strSql, con)
                Using adapter As MySqlDataAdapter = New MySqlDataAdapter(strSql, connectionString)
                    adapter.Fill(tbl)
                End Using
            End Using
        ElseIf (dbType = "SqlServer") Then
            Using cmd As SqlCommand = New SqlCommand(strSql, con)
                'Using SqlAdpt As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter(strSql, con)
                '    SqlAdpt.Fill(tbl)
                '    SqlAdpt.Dispose()
                'End Using
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    tbl.Load(reader)
                    reader.Close()
                End Using
            End Using
        ElseIf (dbType = "Oracle") Then
                Using cmd As OleDbCommand = New OleDbCommand(strSql, con)
                    Using adapter As OleDbDataAdapter = New OleDbDataAdapter(strSql, connectionString)
                        adapter.Fill(tbl)
                    End Using
                End Using
        End If
        Return tbl
    End Function

    ''' <summary>
    ''' 执行数据库更新动作，并返回受影响的行数
    ''' </summary>
    ''' <remarks>传入要执行的SQL语句</remarks>
    Public Shared Function ExecuteNonQuery(ByVal strSql As String) As Integer
        Dim Effect As Integer = 0
        Try
            Dim con As DbConnection = getConn()
            If IsNothing(con) Then Return Effect
            If dbType = "MySql" Then
                Using MyTran As MySqlTransaction = con.BeginTransaction
                    Try
                        Using cmd As MySqlCommand = New MySqlCommand(strSql, con)
                            cmd.Transaction = MyTran
                            Effect = cmd.ExecuteNonQuery()
                            MyTran.Commit()
                            cmd.Dispose()
                        End Using
                    Catch ex As Exception
                        MyTran.Rollback()
                    End Try
                End Using
            ElseIf (dbType = "SqlServer") Then
                Using MyTran As SqlClient.SqlTransaction = con.BeginTransaction
                    Try
                        Using cmd As New SqlClient.SqlCommand(strSql, con)
                            cmd.Transaction = MyTran
                            Effect = cmd.ExecuteNonQuery()
                            MyTran.Commit()
                            cmd.Dispose()
                        End Using
                    Catch ex As Exception
                        MyTran.Rollback()
                    End Try
                End Using
            ElseIf (dbType = "Oracle") Then
                Using MyTran As OleDbTransaction = con.BeginTransaction
                    Try
                        Using cmd As OleDbCommand = New OleDbCommand(strSql, con)
                            cmd.Transaction = MyTran
                            Effect = cmd.ExecuteNonQuery()
                            MyTran.Commit()
                            cmd.Dispose()
                        End Using
                    Catch ex As Exception
                        MyTran.Rollback()
                    End Try
                End Using
            End If
        Catch ex As Exception
        End Try
        Return Effect
    End Function
    Public Shared Function ExecuteReader(ByVal strSql As String) As Object
        Dim Effect As Object = Nothing
        Try
            Dim con As DbConnection = getConn()
            If IsNothing(con) Then Return Nothing
            If dbType = "MySql" Then
                Using cmd As MySqlCommand = New MySqlCommand(strSql, con)
                    Effect = cmd.ExecuteReader()
                    cmd.Dispose()
                End Using
            ElseIf (dbType = "SqlServer") Then
                Using cmd As New SqlClient.SqlCommand(strSql, con)
                    Effect = cmd.ExecuteReader()
                    cmd.Dispose()
                End Using
            ElseIf (dbType = "Oracle") Then
                Using cmd As OleDbCommand = New OleDbCommand(strSql, con)
                    Effect = cmd.ExecuteReader()
                    cmd.Dispose()
                End Using
            End If
        Catch ex As Exception
        End Try
        Return Effect
    End Function
    Public Shared Function ExecuteAdapter(ByVal strSql As String) As Object
        Dim Effect As Object = Nothing
        Try
            Dim con As DbConnection = getConn()
            If IsNothing(con) Then Return Nothing
            If dbType = "MySql" Then
                Effect = New MySqlDataAdapter(strSql, con)
            ElseIf (dbType = "SqlServer") Then
                Effect = New SqlDataAdapter(strSql, con)
            ElseIf (dbType = "Oracle") Then
                Effect = New OleDbDataAdapter(strSql, con)
            End If
        Catch ex As Exception
        End Try
        Return Effect
    End Function
    ''' <summary>
    ''' 执行数据库查询动作，并返回第一行第一列的值
    ''' </summary>
    ''' <remarks>传入要查询的SQL语句</remarks>
    Public Shared Function ExecuteScalar(ByVal strSql As String) As Object
        Dim Effect As Object = Nothing
        Try
            Dim con As DbConnection = getConn()
            If IsNothing(con) Then Return Effect
            If dbType = "MySql" Then
                Using MyTran As MySqlTransaction = con.BeginTransaction
                    Try
                        Using cmd As MySqlCommand = New MySqlCommand(strSql, con)
                            cmd.Transaction = MyTran
                            Effect = cmd.ExecuteScalar()
                            MyTran.Commit()
                            cmd.Dispose()
                        End Using
                    Catch ex As Exception
                        MyTran.Rollback()
                    End Try
                End Using
            ElseIf (dbType = "SqlServer") Then
                If strSql.Trim.StartsWith("Select", StringComparison.CurrentCultureIgnoreCase) Then '一般查询不用起事务
                    Try
                        Using cmd As New SqlClient.SqlCommand(strSql, con)
                            Effect = cmd.ExecuteScalar()
                            cmd.Dispose()
                        End Using
                    Catch ex As Exception

                    End Try
                Else
                    Using MyTran As SqlClient.SqlTransaction = con.BeginTransaction '更新才起事务
                        Try
                            Using cmd As New SqlClient.SqlCommand(strSql, con)
                                cmd.Transaction = MyTran
                                Effect = cmd.ExecuteScalar()
                                MyTran.Commit()
                                cmd.Dispose()
                            End Using
                        Catch ex As Exception
                            MyTran.Rollback()
                        End Try
                    End Using
                End If
            ElseIf (dbType = "Oracle") Then
                Using MyTran As OleDbTransaction = con.BeginTransaction
                    Try
                        Using cmd As OleDbCommand = New OleDbCommand(strSql, con)
                            cmd.Transaction = MyTran
                            Effect = cmd.ExecuteScalar()
                            MyTran.Commit()
                            cmd.Dispose()
                        End Using
                    Catch ex As Exception
                        MyTran.Rollback()
                    End Try
                End Using
            End If
        Catch ex As Exception
        End Try
        Return Effect
    End Function

    Private Sub formatSql(ByVal sql As String, ByVal args As Hashtable, ByVal cmd As IDbCommand)
        If dbType = "MySql" Then
            Dim CS As MatchCollection = Regex.Matches(sql, "@\w+")
            For Each m As Match In CS
                Dim key As String = m.Value
                Dim newKey As String = "?" + key.Substring(1)
                sql = sql.Replace(key, newKey)
                Dim value As Object = args(key)
                If IsNothing(value) Then
                    value = args(key.Substring(1))
                End If
                cmd.Parameters.Add(New MySqlParameter(newKey, value))
            Next
            cmd.CommandText = sql
        ElseIf dbType = "Oracle" Then
            Dim CS As MatchCollection = Regex.Matches(sql, "@\w+")
            Dim i As Integer = 1
            For Each m As Match In CS
                Dim key As String = m.Value
                Dim newKey As String = "@P" + (i + 1)
                sql = sql.Replace(key, "?")

                Dim value As Object = args(key)
                If IsNothing(value) Then
                    value = args(key.Substring(1))
                End If
                cmd.Parameters.Add(New OleDbParameter(newKey, value))
            Next
            cmd.CommandText = sql

        ElseIf dbType = "SqlServer" Then
            Dim CS As MatchCollection = Regex.Matches(sql, "@\w+")
            For Each m As Match In CS
                Dim key As String = m.Value
                Dim value As Object = args(key)
                If IsNothing(value) Then value = args(key.Substring(1))
                If IsNothing(value) Then value = DBNull.Value
                cmd.Parameters.Add(New SqlParameter(key, value))
            Next
            cmd.CommandText = sql
        Else
            cmd.CommandText = sql
        End If
    End Sub
End Class
