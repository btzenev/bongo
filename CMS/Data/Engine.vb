Imports System.Configuration

Namespace Data

  Public Class Engine
    Implements IDisposable

#Region " Declarations "

    Private mobj_connstr As New ConnectionStringSettings
    Private mobj_factory As DbProviderFactory
    Private mstr_conn As String
    Private mb_connected As Boolean
    Private mobj_conn As DbConnection
    Private mb_trans As Boolean = False
    Private mint_retry As Integer
    Private mint_cmdtimeout As Integer
    Private mb_closing As Boolean = False
    Private mb_errors As Boolean
    Private mobj_trans As DbTransaction
    Private mobj_cmd As DbCommand
    Private mobj_error As System.Exception

#End Region

#Region " Properties "

    Public ReadOnly Property Factory() As DbProviderFactory
      Get
        On Error Resume Next
        If mobj_factory Is Nothing Then
          mobj_factory = DbProviderFactories.GetFactory(MyClass.Provider)
        End If
        Return mobj_factory
      End Get
    End Property

    Public Property Provider() As String
      Get
        Return mobj_connstr.ProviderName
      End Get
      Set(ByVal Value As String)
        mobj_connstr.ProviderName = Value
      End Set
    End Property

    Public Property Connection() As DbConnection
      Get
        On Error Resume Next
        If mobj_conn Is Nothing Then
          mobj_conn = Factory.CreateConnection()
          MyClass.ConnectionString = MyClass.ConnectionString
          mobj_conn.ConnectionString = MyClass.ConnectionString
        End If
        Return mobj_conn
      End Get
      Set(ByVal vobj_conn As DbConnection)
        mobj_conn = vobj_conn
      End Set
    End Property

    Public Property ConnectionString() As String
      Get
        Return mobj_connstr.ConnectionString
      End Get
      Set(ByVal vstr_conn As String)
        mobj_connstr = New ConnectionStringSettings(mobj_connstr.Name, vstr_conn.Replace("~\", MapPath()), mobj_connstr.ProviderName)
      End Set
    End Property

    Public Property CommandTimeout() As Integer
      Get
        Return mint_cmdtimeout
      End Get
      Set(ByVal vint_time As Integer)
        If mint_cmdtimeout > 0 Then
          mint_cmdtimeout = vint_time
        End If
      End Set
    End Property

    Public Property ConnectionRetry() As Integer
      Get
        Return mint_retry
      End Get
      Set(ByVal vint_retry As Integer)
        If mint_retry >= 0 Then
          mint_retry = vint_retry
        End If
      End Set
    End Property

    Public ReadOnly Property Errors() As Boolean
      Get
        Return mb_errors
      End Get
    End Property

    Public ReadOnly Property LastError() As System.Exception
      Get
        Return mobj_error
      End Get
    End Property

#Region " Get "

    Public ReadOnly Property [Get](ByVal vstr_table As String, ByVal vstr_field As String, ByVal vstr_where As String, ByVal vobj_value As Object) As Object
      Get
        Dim obj_query As New QueryBuilder(vstr_table)
        obj_query.Add(vstr_field)
        obj_query.Where(vstr_where, vobj_value)
        Return MyClass.Get(obj_query)
      End Get
    End Property

    Public ReadOnly Property [Get](ByVal vstr_sql As String) As Object
      Get
        Return ExecuteScalar(vstr_sql)
      End Get
    End Property

    Public ReadOnly Property [Get](ByVal vobj_query As QueryBuilder) As Object
      Get
        Return ExecuteScalar(vobj_query)
      End Get
    End Property

    Public ReadOnly Property [Get](ByVal vobj_cmd As DbCommand) As Object
      Get
        Return ExecuteScalar(vobj_cmd)
      End Get
    End Property

#End Region

#Region " Row "
        Public ReadOnly Property Row(ByVal vstr_table As String, ByVal vstr_where As String, ByVal vobj_value As Object) As DataRow
            Get
                Dim obj_query As New QueryBuilder(vstr_table)
                obj_query.Where(vstr_where, vobj_value)
                Return Row(obj_query)
            End Get
        End Property

    Public ReadOnly Property Row(ByVal vstr_sql As String) As DataRow
      Get
        On Error Resume Next
        Dim obj_cmd As DbCommand = GetCommand(vstr_sql, CommandType.Text)
        Return Row(obj_cmd)
      End Get
    End Property

    Public ReadOnly Property Row(ByVal vobj_query As QueryBuilder) As DataRow
      Get
        Return Row(vobj_query.ToCommand(Me))
      End Get
    End Property

    Public ReadOnly Property Row(ByVal vobj_cmd As DbCommand) As DataRow
      Get
        Try
          Using vobj_table As DataTable = Table(vobj_cmd)
            If vobj_table IsNot Nothing Then
              If vobj_table.Rows.Count > 0 Then Return vobj_table.Rows(0)
            End If
          End Using
        Catch ex As Exception
          Exception("Data", "Unable to retrieve database row.", ex)
        End Try
      End Get
    End Property

#End Region

#Region " Table "

    Public ReadOnly Property Table(ByVal vstr_sql As String) As DataTable
      Get
        Dim obj_cmd As DbCommand = GetCommand(vstr_sql, CommandType.Text)
        Return Table(obj_cmd)
      End Get
    End Property

    Public ReadOnly Property Table(ByVal vobj_query As QueryBuilder) As DataTable
      Get
        Return Table(vobj_query.ToCommand(Me))
      End Get
    End Property

    Public ReadOnly Property Table(ByVal vobj_cmd As DbCommand) As DataTable
      Get
        Try
          Using obj_reader As DbDataReader = ExecuteReader(vobj_cmd)
            Dim obj_table As New DataTable
            obj_table.Load(obj_reader)
            Return obj_table
          End Using
        Catch ex As Exception
          Exception("Data", "Unable to retrieve database table.", ex)
        End Try
      End Get
    End Property

#End Region

#Region " Rows "

    Public ReadOnly Property Rows(ByVal vstr_sql As String) As System.Data.DataRowCollection
      Get
        Dim obj_table As DataTable = Table(vstr_sql)
        If obj_table Is Nothing Then obj_table = New DataTable
        Return obj_table.Rows
      End Get
    End Property

    Public ReadOnly Property Rows(ByVal vobj_query As QueryBuilder) As System.Data.DataRowCollection
      Get
        Dim obj_table As DataTable = Table(vobj_query)
        If obj_table Is Nothing Then obj_table = New DataTable
        Return obj_table.Rows
      End Get
    End Property

    Public ReadOnly Property Rows(ByVal vobj_cmd As DbCommand) As System.Data.DataRowCollection
      Get
        Dim obj_table As DataTable = Table(vobj_cmd)
        If obj_table Is Nothing Then obj_table = New DataTable
        Return obj_table.Rows
      End Get
    End Property

#End Region

#End Region

#Region " New "

    Public Sub New(ByVal vstr_name As String)
      MyBase.New()
      On Error Resume Next
      mobj_connstr = ConfigurationManager.ConnectionStrings(vstr_name)
      MyClass.Provider = MyClass.Provider
      MyClass.ConnectionString = MyClass.ConnectionString
      ConnectionRetry = ToInt(Config("data.retry"), 3)
      CommandTimeout = ToInt(Config("data.cmdtimeout"), 0)
    End Sub

    Public Sub New()
      MyBase.New()
    End Sub

#End Region

#Region " Open/Close "

    Public Function Open() As Boolean
      If Not mobj_connstr Is Nothing Then
        If NotEmpty(MyClass.Provider) And NotEmpty(MyClass.ConnectionString) Then
          If Not Connection.State = ConnectionState.Open Then
            Try
              Connection.Open()
            Catch ex As Exception
              mobj_error = ex
              Exception("DATA", "Unable to connect to the database.", ex)
            End Try
          End If
          Return Connection.State = ConnectionState.Open
        End If
      End If
    End Function

    Public Sub Close()
      On Error Resume Next
      Connection.Close()
    End Sub

#End Region

#Region " Factory Methods "

#Region " GetCommand "

    Public Function GetCommand(ByVal vstr_sql As String) As DbCommand
      Return GetCommand(vstr_sql, CommandType.Text)
    End Function

    Public Function GetCommand(ByVal vstr_sql As String, ByVal vobj_type As CommandType) As DbCommand
      Try
        Dim obj_cmd As DbCommand
        If mb_trans Then
          obj_cmd = mobj_cmd
        Else
          obj_cmd = Factory.CreateCommand
        End If
        obj_cmd.Connection = Connection
        obj_cmd.CommandTimeout = CommandTimeout
        obj_cmd.CommandType = vobj_type
        obj_cmd.CommandText = vstr_sql
        Return obj_cmd
      Catch ex As Exception
        Exception("Data", "Unable to create database command object.", ex)
      End Try
    End Function

#End Region

#Region " GetDataAdapter "

    Public Function GetDataAdapter() As DbDataAdapter
      Return GetDataAdapter(Nothing)
    End Function

    Public Function GetDataAdapter(ByVal vobj_cmd As DbCommand) As DbDataAdapter
      Try
        Dim obj_adp As DbDataAdapter = mobj_factory.CreateDataAdapter
        obj_adp.MissingSchemaAction = MissingSchemaAction.AddWithKey
        'obj_adp.MissingMappingAction = MissingMappingAction.Passthrough
        obj_adp.SelectCommand = vobj_cmd
        Return obj_adp
      Catch ex As Exception
        Exception("Data", "Unable to create database adapter object.", ex)
      End Try
    End Function

#End Region

#Region " CreateParameter "

    Public Function CreateParameter() As DbParameter
      Try
        Return mobj_factory.CreateParameter
      Catch ex As Exception
        Exception("Data", "Unable to create database parameter object.", ex)
      End Try
    End Function

    Public Function CreateParameter(ByVal vstr_name As String, ByVal vobj_value As Object) As DbParameter
      Dim obj_param As DbParameter = CreateParameter()
      If obj_param IsNot Nothing Then
        obj_param.ParameterName = vstr_name.TrimStart("@"c).ToLower
        obj_param.Value = vobj_value
      End If
      Return obj_param
    End Function

    Public Function CreateParameter(ByVal vstr_name As String, ByVal vobj_value As Object, ByVal vobj_type As DbType) As DbParameter
      Dim obj_param As DbParameter = CreateParameter(vstr_name, vobj_value)
      If obj_param IsNot Nothing Then
        obj_param.DbType = vobj_type
      End If
      Return obj_param
    End Function

#End Region

#End Region

#Region " Execute Methods "

#Region " ExecuteNonQuery "

    Public Function ExecuteNonQuery(ByVal vobj_query As QueryBuilder) As Integer
      Return ExecuteNonQuery(vobj_query.ToCommand(Me))
    End Function

    Public Function ExecuteNonQuery(ByVal vstr_sql As String) As Integer
      Return ExecuteNonQuery(vstr_sql, CommandType.Text)
    End Function

    Public Function ExecuteNonQuery(ByVal vstr_sql As String, ByVal vobj_type As CommandType) As Integer
      Return ExecuteNonQuery(vstr_sql, vobj_type, Nothing)
    End Function

    Public Function ExecuteNonQuery(ByVal vstr_sql As String, ByVal vobj_params As List(Of DbParameter)) As Integer
      Return ExecuteNonQuery(vstr_sql, CommandType.Text, vobj_params)
    End Function

    Public Function ExecuteNonQuery(ByVal vstr_sql As String, ByVal vobj_type As CommandType, ByVal vobj_params As List(Of DbParameter)) As Integer
      Dim obj_cmd As DbCommand = Me.GetCommand(vstr_sql, vobj_type)
      obj_cmd.Parameters.Clear()
      If vobj_params IsNot Nothing Then
        obj_cmd.Parameters.AddRange(vobj_params.ToArray)
      End If
      Return ExecuteNonQuery(obj_cmd)
    End Function

    Public Function ExecuteNonQuery(ByVal vobj_cmd As DbCommand) As Integer
      Try
        Using vobj_cmd
          mb_errors = False
          If Not vobj_cmd.Connection.State = ConnectionState.Open Then vobj_cmd.Connection.Open()
          Return vobj_cmd.ExecuteNonQuery
        End Using
      Catch ex As Exception
        mb_errors = True : mobj_error = ex
        Exception("Data", "Unable to execute nonquery command.", ex)
        Return -1
      End Try
    End Function

#End Region

#Region " ExecuteScalar "

    Public Function ExecuteScalar(ByVal vobj_query As QueryBuilder) As Object
      Return ExecuteScalar(vobj_query.ToCommand(Me))
    End Function

    Public Function ExecuteScalar(ByVal vstr_sql As String) As Object
      Return ExecuteScalar(vstr_sql, CommandType.Text)
    End Function

    Public Function ExecuteScalar(ByVal vstr_sql As String, ByVal vobj_type As CommandType) As Object
      Return ExecuteScalar(vstr_sql, vobj_type, Nothing)
    End Function

    Public Function ExecuteScalar(ByVal vstr_sql As String, ByVal vobj_params As List(Of DbParameter)) As Object
      Return ExecuteScalar(vstr_sql, CommandType.Text, vobj_params)
    End Function

    Public Function ExecuteScalar(ByVal vstr_sql As String, ByVal vobj_type As CommandType, ByVal vobj_params As List(Of DbParameter)) As Object
      Dim obj_cmd As DbCommand = Me.GetCommand(vstr_sql, vobj_type)
      obj_cmd.Parameters.Clear()
      If vobj_params IsNot Nothing Then
        obj_cmd.Parameters.AddRange(vobj_params.ToArray)
      End If
      Return ExecuteScalar(obj_cmd)
    End Function

    Public Function ExecuteScalar(ByVal vobj_cmd As DbCommand) As Object
      Try
        Using vobj_cmd
          mb_errors = False
          If Not vobj_cmd.Connection.State = ConnectionState.Open Then vobj_cmd.Connection.Open()
          Return vobj_cmd.ExecuteScalar
        End Using
      Catch ex As Exception
        mb_errors = True : mobj_error = ex
        Exception("Data", "Unable to execute scalar command.", ex)
      End Try
    End Function

#End Region

#Region " ExecuteTable "

    Public Function ExecuteTable(ByVal vobj_query As QueryBuilder) As DataTable
      Return ExecuteTable(vobj_query.ToCommand(Me))
    End Function

    Public Function ExecuteTable(ByVal vstr_sql As String) As DataTable
      Return ExecuteTable(vstr_sql, CommandType.Text)
    End Function

    Public Function ExecuteTable(ByVal vstr_sql As String, ByVal vobj_type As CommandType) As DataTable
      Return ExecuteTable(vstr_sql, vobj_type, Nothing)
    End Function

    Public Function ExecuteTable(ByVal vstr_sql As String, ByVal vobj_params As List(Of DbParameter)) As DataTable
      Return ExecuteTable(vstr_sql, CommandType.Text, vobj_params)
    End Function

    Public Function ExecuteTable(ByVal vstr_sql As String, ByVal vobj_type As CommandType, ByVal vobj_params As List(Of DbParameter)) As DataTable
      Dim obj_cmd As DbCommand = Me.GetCommand(vstr_sql, vobj_type)
      obj_cmd.Parameters.Clear()
      If vobj_params IsNot Nothing Then
        obj_cmd.Parameters.AddRange(vobj_params.ToArray)
      End If
      Return ExecuteTable(obj_cmd)
    End Function

    Public Function ExecuteTable(ByVal vobj_cmd As DbCommand) As DataTable
      Try
        Using vobj_cmd
          mb_errors = False
          If Not vobj_cmd.Connection.State = ConnectionState.Open Then vobj_cmd.Connection.Open()
          Dim obj_adp As DbDataAdapter = GetDataAdapter(vobj_cmd)
          Dim obj_table As New DataTable
          obj_adp.Fill(obj_table)
          Return obj_table
        End Using
      Catch ex As Exception
        mb_errors = True : mobj_error = ex
        Exception("Data", "Unable to execute data table command.", ex)
      End Try
    End Function

#End Region

#Region " ExecuteReader "

    Public Function ExecuteReader(ByVal vobj_query As QueryBuilder) As DbDataReader
      Return ExecuteReader(vobj_query.ToCommand(Me))
    End Function

    Public Function ExecuteReader(ByVal vstr_sql As String) As DbDataReader
      Return ExecuteReader(vstr_sql, CommandType.Text)
    End Function

    Public Function ExecuteReader(ByVal vstr_sql As String, ByVal vobj_type As CommandType) As DbDataReader
      Return ExecuteReader(vstr_sql, vobj_type, Nothing)
    End Function

    Public Function ExecuteReader(ByVal vstr_sql As String, ByVal vobj_params As List(Of DbParameter)) As DbDataReader
      Return ExecuteReader(vstr_sql, CommandType.Text, vobj_params)
    End Function

    Public Function ExecuteReader(ByVal vstr_sql As String, ByVal vobj_type As CommandType, ByVal vobj_params As List(Of DbParameter)) As DbDataReader
      Dim obj_cmd As DbCommand = Me.GetCommand(vstr_sql, vobj_type)
      obj_cmd.Parameters.Clear()
      If vobj_params IsNot Nothing Then
        obj_cmd.Parameters.AddRange(vobj_params.ToArray)
      End If
      Return ExecuteReader(obj_cmd)
    End Function

    Public Function ExecuteReader(ByVal vobj_cmd As DbCommand) As DbDataReader
      Try
        Return ExecuteTable(vobj_cmd).CreateDataReader
      Catch ex As Exception
        mb_errors = True : mobj_error = ex
        Exception("Data", "Unable to execute data reader command.", ex)
      End Try
    End Function

#End Region

#Region " ExecuteProc "

    Public Function ExecuteProc(ByVal vobj_query As QueryBuilder) As Object
      Return ExecuteProc(vobj_query.ToCommand(Me))
    End Function

    Public Function ExecuteProc(ByVal vstr_sql As String) As Object
      Return ExecuteProc(vstr_sql, Nothing)
    End Function

    Public Function ExecuteProc(ByVal vstr_sql As String, ByVal vobj_params As List(Of DbParameter)) As Object
      Dim obj_cmd As DbCommand = Me.GetCommand(vstr_sql)
      obj_cmd.Parameters.Clear()
      If vobj_params IsNot Nothing Then
        obj_cmd.Parameters.AddRange(vobj_params.ToArray)
      End If
      Return ExecuteProc(obj_cmd)
    End Function

    Public Function ExecuteProc(ByVal vobj_cmd As DbCommand) As Object
      Try
        vobj_cmd.CommandType = CommandType.StoredProcedure
        Return ExecuteNonQuery(vobj_cmd)
      Catch ex As Exception
        mb_errors = True : mobj_error = ex
        Exception("Data", "Unable to execute stored procedure.", ex)
      End Try
    End Function

#End Region

#Region " Execute "

    Public Function Execute(ByVal vobj_query As QueryBuilder) As Boolean
      Return Execute(vobj_query.ToCommand(Me))
    End Function

    Public Function Execute(ByVal vstr_sql As String) As Boolean
      Return Execute(vstr_sql, CommandType.Text)
    End Function

    Public Function Execute(ByVal vstr_sql As String, ByVal vobj_type As CommandType) As Boolean
      Return Execute(vstr_sql, vobj_type, Nothing)
    End Function

    Public Function Execute(ByVal vstr_sql As String, ByVal vobj_params As List(Of DbParameter)) As Boolean
      Return Execute(vstr_sql, CommandType.Text, vobj_params)
    End Function

    Public Function Execute(ByVal vstr_sql As String, ByVal vobj_type As CommandType, ByVal vobj_params As List(Of DbParameter)) As Boolean
      Dim obj_cmd As DbCommand = Me.GetCommand(vstr_sql, vobj_type)
      obj_cmd.Parameters.Clear()
      If vobj_params IsNot Nothing Then
        obj_cmd.Parameters.AddRange(vobj_params.ToArray)
      End If
      Return Execute(obj_cmd)
    End Function

    Public Function Execute(ByVal vobj_cmd As DbCommand) As Boolean
            Try
                mb_errors = False
                Dim int_count As Integer = ExecuteNonQuery(vobj_cmd)
                Return Not mb_errors
                'If int_count > 0 And Not mb_errors Then
                'Return True
                'Else
                'If Not mb_errors Then Return True
                'End If
            Catch ex As Exception
                mb_errors = True : mobj_error = ex
                Exception("Data", "Unable to execute command.", ex)
            End Try
    End Function

#End Region

#Region " Schema "

    Public Function Schema() As DataTable
      Try
        If Not Connection.State = ConnectionState.Open Then Connection.Open()
        Return Connection.GetSchema()
      Catch ex As Exception
        Exception("Data", "Unable retrieve database schema.", ex)
      End Try
    End Function

    Public Function Schema(ByVal vstr_name As String) As DataColumnCollection
            Try
    
                    Using obj_reader As DbDataReader = ExecuteReader("select top 1 * from " & vstr_name.Replace("'", "''"))
                        Dim obj_table As New DataTable
                        obj_table.Load(obj_reader)
                        Return obj_table.Columns
                    End Using
            Catch ex As Exception
                Exception("Data", "Unable retrieve database table schema. [" & vstr_name & "]", ex)
            End Try
    End Function

#End Region

#End Region

#Region " Transaction "

    Public ReadOnly Property Started() As Boolean
      Get
        Return mb_trans
      End Get
    End Property

    Public Sub Cancel()
      On Error Resume Next
      mobj_trans.Dispose()
      mobj_cmd.Dispose()
      mb_trans = False
    End Sub

    Public Sub Begin()
      mobj_trans = Connection.BeginTransaction()
      mobj_cmd = Factory.CreateCommand
      mobj_cmd.Transaction = mobj_trans
      mb_trans = True
    End Sub

    Public Sub Commit()
      If mb_trans Then
        mobj_trans.Commit()
        Cancel()
      Else
        Exception("Data", "Transaction.Begin must be called before commit or rollback.")
      End If
    End Sub

    Public Sub Rollback()
      Try
        mobj_trans.Rollback()
      Catch ex As System.Exception
        Exception("Data", "Unable to rollback transaction", ex)
      End Try
      Cancel()
    End Sub

#End Region

#Region " CRUD Methods "

    Public Function Identity() As Integer
      Return ToInt(MyClass.Get("SELECT @@IDENTITY AS x"), 0)
    End Function

#Region " Insert "

    Public Function Insert(ByVal vobj_query As QueryBuilder, ByRef vint_id As Integer) As Boolean
      Return Insert(vobj_query.ToCommand(Me), vint_id)
    End Function

    Public Function Insert(ByVal vobj_query As QueryBuilder) As Integer
      Return Insert(vobj_query.ToCommand(Me))
    End Function

    Public Function Insert(ByVal vstr_sql As String) As Integer
      Return Insert(vstr_sql, CommandType.Text)
    End Function

    Public Function Insert(ByVal vstr_sql As String, ByVal vobj_type As CommandType) As Integer
      Return Insert(vstr_sql, vobj_type, Nothing)
    End Function

    Public Function Insert(ByVal vstr_sql As String, ByVal vobj_params As List(Of DbParameter)) As Integer
      Return Insert(vstr_sql, CommandType.Text, vobj_params)
    End Function

    Public Function Insert(ByVal vstr_sql As String, ByVal vobj_type As CommandType, ByVal vobj_params As List(Of DbParameter)) As Integer
      On Error Resume Next
      If Execute(vstr_sql, vobj_type, vobj_params) Then Return Identity()
    End Function

    Public Function Insert(ByVal vobj_cmd As DbCommand) As Integer
      On Error Resume Next
      Dim int_id As Integer = 0
      If Insert(vobj_cmd, int_id) Then Return int_id
    End Function

    Public Function Insert(ByVal vobj_cmd As DbCommand, ByRef vint_id As Integer) As Boolean
            'On Error Resume Next
            Try
                If Execute(vobj_cmd) Then
                    vint_id = Identity()
                    Return (vint_id > 0)
                End If
            Catch ex As System.Exception
                mb_errors = True : mobj_error = ex
                Exception("Data", "Unable to insert records.", ex)
            End Try
        End Function

#End Region

#Region " Update "

    Public Function Update(ByVal vstr_table As String, ByVal vstr_old As String, ByVal vobj_old As Object, ByVal vstr_new As String, ByVal vobj_new As Object) As Boolean
      Dim obj_query As New QueryBuilder(vstr_table, QueryType.Update)
      obj_query.Add(vstr_new, vobj_new)
      obj_query.Where(vstr_old, vobj_old)
      Return Update(obj_query)
    End Function

    Public Function Update(ByVal vstr_table As String, ByVal vstr_field As String, ByVal vobj_old As Object, ByVal vobj_new As Object) As Boolean
      Dim obj_query As New QueryBuilder(vstr_table, QueryType.Update)
      obj_query.Add(vstr_field, vobj_new)
      obj_query.Where(vstr_field, vobj_old)
      Return Update(obj_query)
    End Function

    Public Function Update(ByVal vobj_query As QueryBuilder) As Boolean
      If vobj_query.QueryType = QueryType.Update Then
        Return Update(vobj_query.ToCommand(Me))
      Else
        Exception("Data", "Update query is required.")
      End If
    End Function

    Public Function Update(ByVal vstr_sql As String) As Boolean
      Return Update(vstr_sql, CommandType.Text)
    End Function

    Public Function Update(ByVal vstr_sql As String, ByVal vobj_type As CommandType) As Boolean
      Return Update(vstr_sql, vobj_type, Nothing)
    End Function

    Public Function Update(ByVal vstr_sql As String, ByVal vobj_params As List(Of DbParameter)) As Boolean
      Return Update(vstr_sql, CommandType.Text, vobj_params)
    End Function

    Public Function Update(ByVal vstr_sql As String, ByVal vobj_type As CommandType, ByVal vobj_params As List(Of DbParameter)) As Boolean
      Return Execute(vstr_sql, vobj_type, vobj_params)
    End Function

        Public Function Update(ByVal vobj_cmd As DbCommand) As Boolean
            Try
                Return Execute(vobj_cmd)
            Catch ex As System.Exception
                mb_errors = True : mobj_error = ex
                Exception("Data", "Unable to insert records.", ex)
            End Try
        End Function

#End Region

#Region " Delete "

    Public Function Delete(ByVal vstr_table As String, ByVal vstr_field As String, ByVal vobj_value As Object) As Boolean
      Dim obj_query As New QueryBuilder(vstr_table, QueryType.Delete)
      obj_query.Where(vstr_field, vobj_value)
      Return Delete(obj_query)
    End Function

    Public Function Delete(ByVal vobj_query As QueryBuilder) As Boolean
      If vobj_query.QueryType = QueryType.Delete Then
        Return Delete(vobj_query.ToCommand(Me))
      Else
        Exception("Data", "Delete query is required.")
      End If
    End Function

    Public Function Delete(ByVal vstr_sql As String) As Boolean
      Return Delete(vstr_sql, CommandType.Text)
    End Function

    Public Function Delete(ByVal vstr_sql As String, ByVal vobj_type As CommandType) As Boolean
      Return Delete(vstr_sql, vobj_type, Nothing)
    End Function

    Public Function Delete(ByVal vstr_sql As String, ByVal vobj_params As List(Of DbParameter)) As Boolean
      Return Delete(vstr_sql, CommandType.Text, vobj_params)
    End Function

    Public Function Delete(ByVal vstr_sql As String, ByVal vobj_type As CommandType, ByVal vobj_params As List(Of DbParameter)) As Boolean
      Return Execute(vstr_sql, vobj_type, vobj_params)
    End Function

        Public Function Delete(ByVal vobj_cmd As DbCommand) As Boolean
            mb_errors = False
            Return Execute(vobj_cmd) And Not mb_errors
        End Function

#End Region

#End Region

#Region " Dispose "
    Public Sub Dispose() Implements IDisposable.Dispose
      Close()
    End Sub
#End Region

  End Class

End Namespace
