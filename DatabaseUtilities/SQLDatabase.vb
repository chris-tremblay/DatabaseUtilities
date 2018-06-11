Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.ComponentModel.Design
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks

Namespace DatabaseUtilities

    <Serializable>
    Public Class SQLDatabase

#Region "Fields"

        Protected _connectionString As String = ""
        Private Shared _Cache As Cache = Cache.GetInstance()
        Private Shared rand = New Random()

#End Region

#Region "Constructors"

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

#End Region

#Region "Properties"

        Protected ReadOnly Property ConnectionString() As String
            Get
                Return _connectionString
            End Get
        End Property

        Protected Shared ReadOnly Property Cache As Cache
            Get
                Return _Cache
            End Get
        End Property


#End Region

#Region "Protected Methods"

        Protected Shared Function GetConnectionStringFromConfig(name As String) As String
            Try
                Dim value As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(name)

                If value Is Nothing Then
                    Throw New Exception("SQLDatabase requires a Connection String Added to the configuration file (app.config or web.config) with the name of 'ConnectionString'. ")
                Else
                    Return value.ConnectionString
                End If
            Catch ex As Exception
                Throw New Exception("")
            End Try
        End Function

#End Region

#Region "CreateDataReader"

        Public Function CreateDataReader(cnn As SqlConnection, query As String, ByVal ParamArray parameters As Object()) As SqlDataReader
            Return CreateDataReader(cnn, Nothing, query, parameters)
        End Function

        Public Function CreateDataReader(transaction As SqlTransaction, query As String, ByVal ParamArray parameters As Object()) As SqlDataReader
            Return CreateDataReader(transaction.Connection, transaction, query, parameters)
        End Function

        Public Function CreateDataReader(cnn As SqlConnection, transaction As SqlTransaction, query As String, ByVal ParamArray parameters As Object()) As SqlDataReader
            Using command As SqlCommand = CreateCommand(cnn, transaction, query, parameters)
                Dim value As SqlDataReader

                If command.Connection.State = ConnectionState.Closed Then command.Connection.Open()

                value = command.ExecuteReader()

                Return value
            End Using
        End Function

#End Region

        Public Shared Function CreateBinaryParam(paramName As String, data As Byte()) As SqlParameter
            Dim fileParam As New SqlParameter(String.Concat("@", paramName), Nothing)

            fileParam.SqlDbType = SqlDbType.VarBinary

            If data Is Nothing Then
                fileParam.Value = DBNull.Value
                fileParam.Size = -1
            Else
                fileParam.Value = data.ToArray()
            End If

            Return fileParam
        End Function

        ''' <summary>
        ''' Creates a new Parameterized SqlCommand object. Query should be structured as it would if String.Format was used. Parameters should be referenced by their index in the <paramref  name="Parameters"/> argument. (ex: {0}, {1}, etc...)
        ''' </summary>
        ''' <param name="Query"></param>
        ''' <param name="Parameters"></param>
        ''' <returns>Using this function is safe from SQL injection attaks because it creates a parameterized command.</returns>
        ''' <remarks></remarks>
        Public Function CreateSqlCommand(query As String, ByVal ParamArray parameters As Object()) As SqlCommand
            Return CreateCommand(Nothing, Nothing, query, parameters)
        End Function

        Public Function CreateCommand(cnn As SqlConnection, query As String, ByVal ParamArray parameters As Object()) As SqlCommand
            Return CreateCommand(cnn, Nothing, query, parameters)
        End Function

        Public Function CreateCommand(transaction As SqlTransaction, query As String, ByVal ParamArray parameters As Object()) As SqlCommand
            Return CreateCommand(transaction.Connection, transaction, query, parameters)
        End Function

        Public Function CreateCommand(cnn As SqlConnection, transaction As SqlTransaction, query As String, ByVal ParamArray parameters As Object()) As SqlCommand
            Dim command As New SqlCommand(query)
            Dim ex = New Regex("[\\\!\@\#|$\^\&\*\(\)\{\}\[\]\:\;\""\'\<\>\,\.\?\/\~\-\+\=\`]", RegexOptions.Compiled + RegexOptions.Singleline)

            If cnn Is Nothing Then
                command.Connection = New SqlConnection(_connectionString)
            Else
                command.Connection = cnn
            End If

            command.Transaction = transaction

            If parameters IsNot Nothing Then
                Dim parameterNames(parameters.Length - 1) As String

                For counter = 0 To parameters.Length - 1
                    Select Case True
                        Case TypeOf parameters(counter) Is Dictionary(Of String, Object)
                            For Each entry In parameters(counter)
                                command.Parameters.AddWithValue(If(entry.Key.StartsWith("@"), entry.Key, $"@{entry.Key}"), GetParameter(entry.Value))
                            Next

                        Case TypeOf parameters(counter) Is KeyValuePair(Of String, Object)
                            Dim entry = CType(parameters(counter), KeyValuePair(Of String, Object))

                            command.Parameters.AddWithValue(If(entry.Key.StartsWith("@"), entry.Key, $"@{entry.Key}"), GetParameter(entry.Value))
                        Case TypeOf parameters(counter) Is IEnumerable AndAlso Not TypeOf parameters(counter) Is String
                            Dim items = CType(parameters(counter), IEnumerable)
                            'Dim counter2 = 0
                            Dim sb As New StringBuilder()

                            For Each item As Object In items
                                If IsNumber(item) Then
                                    sb.Append($"{item},")
                                ElseIf TypeOf item Is String AndAlso Not ex.IsMatch(item) Then
                                    sb.Append($"'{item}',")
                                Else
                                    'dim parameterName = CreateParameterName()
                                    'sb.Append($"@P{counter}_{counter2},")
                                    'command.Parameters.AddWithValue(String.Concat("@P", counter, "_", counter2), GetParameter(item))
                                    Dim parameterName = CreateParameterName()
                                    sb.Append($"{parameterName},")
                                    command.Parameters.AddWithValue(parameterName, GetParameter(item))
                                    'counter2 += 1
                                End If

                            Next

                            If sb.Length > 0 Then sb.Length -= 1

                            parameterNames(counter) = sb.ToString()
                        Case TypeOf parameters(counter) Is SqlParameter
                            Dim p = CType(parameters(counter), SqlParameter)

                            If String.IsNullOrEmpty(p.ParameterName) Then p.ParameterName = String.Concat("@P", counter)

                            parameterNames(counter) = p.ParameterName
                            command.Parameters.Add(p)

                        Case Else
                            Dim item = parameters(counter)

                            If IsNumber(item) Then
                                parameterNames(counter) = item
                            ElseIf TypeOf item Is String AndAlso Not ex.IsMatch(item) Then
                                parameterNames(counter) = $"'{item}'"
                            Else
                                Dim parameterName = CreateParameterName()
                                command.Parameters.AddWithValue(parameterName, GetParameter(item))
                                parameterNames(counter) = parameterName
                            End If
                            'Dim parameterName = CreateParameterName()
                            'command.Parameters.AddWithValue(parameterName, GetParameter(parameters(counter)))
                            'parameterNames(counter) = parameterName

                            'command.Parameters.AddWithValue(String.Concat("@P", counter), GetParameter(parameters(counter)))
                            'parameterNames(counter) = String.Concat("@P", counter)
                    End Select
                Next

                command.CommandText = String.Format(command.CommandText, parameterNames)
            End If

            Return command
        End Function

        Private Function CreateParameterName() As String
            Dim value As New StringBuilder("@")
            Dim characters = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}

            SyncLock rand
                value.Append(characters(rand.Next(0, 27)))

                For i = 0 To 9
                    value.Append(characters(rand.Next(0, 36)))
                Next
            End SyncLock

            Return value.ToString()
        End Function

        Private Function IsNumber(obj As Object) As Boolean
            Dim t = obj?.GetType()

            Return t = GetType(Integer) _
                OrElse t = GetType(Long) _
                OrElse t = GetType(Short) _
                OrElse t = GetType(Double) _
                OrElse t = GetType(Decimal)
        End Function

        ''' <summary>
        ''' Creates a new (closed) SQLConnection to the database.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Function CreateConnection() As SqlConnection
            Return New SqlConnection(_connectionString)
        End Function

#Region "ExecuteNonQuery"

        Public Function ExecuteNonQuery(query As String, ByVal ParamArray parameters() As Object) As Integer
            Using command As SqlCommand = CreateSqlCommand(query, parameters)
                Using command.Connection
                    Return ExecuteNonQuery(command)
                End Using
            End Using
        End Function

        Public Function ExecuteNonQuery(cnn As SqlConnection, sql As String, ByVal ParamArray parameters() As Object) As Integer
            Using command As SqlCommand = CreateCommand(cnn, sql, parameters)
                Return ExecuteNonQuery(command)
            End Using
        End Function

        Public Function ExecuteNonQuery(transaction As SqlTransaction, sql As String, ByVal ParamArray parameters() As Object) As Integer
            Using command As SqlCommand = CreateCommand(transaction.Connection, transaction, sql, parameters)
                Return ExecuteNonQuery(command)
            End Using
        End Function

        Public Function ExecuteNonQuery(cnn As SqlConnection, transaction As SqlTransaction, sql As String, ByVal ParamArray parameters() As Object) As Integer
            Using command As SqlCommand = CreateCommand(cnn, transaction, sql, parameters)
                Return ExecuteNonQuery(command)
            End Using
        End Function

        Public Function ExecuteNonQuery(command As SqlCommand) As Integer
            If command.Connection Is Nothing Then
                Using cnn As SqlConnection = New SqlConnection(_connectionString)
                    command.Connection = cnn

                    cnn.Open()
                    Return command.ExecuteNonQuery()
                End Using
            Else
                Dim closeConnection As Boolean = (command.Connection.State = ConnectionState.Closed)

                Try
                    If closeConnection Then command.Connection.Open()
                    Return command.ExecuteNonQuery()
                Finally
                    If closeConnection Then command.Connection.Close()
                End Try
            End If
        End Function

#End Region

#Region "ExecuteScalar"

        Public Function ExecuteScalar(query As String, ByVal ParamArray parameters() As Object) As Object
            Using command As SqlCommand = CreateSqlCommand(query, parameters)
                Using command.Connection
                    Return ExecuteScalar(command)
                End Using
            End Using
        End Function

        Public Function ExecuteScalar(cnn As SqlConnection, sql As String, ByVal ParamArray parameters() As Object) As Object
            Using command As SqlCommand = CreateCommand(cnn, sql, parameters)
                Return ExecuteScalar(command)
            End Using
        End Function

        Public Function ExecuteScalar(cnn As SqlConnection, transaction As SqlTransaction, sql As String, ByVal ParamArray parameters() As Object) As Object
            Using command As SqlCommand = CreateCommand(cnn, transaction, sql, parameters)
                Return ExecuteScalar(command)
            End Using
        End Function

        Public Function ExecuteScalar(command As SqlCommand) As Object
            If command.Connection Is Nothing Then
                Using cnn As SqlConnection = New SqlConnection(_connectionString)
                    command.Connection = cnn

                    cnn.Open()
                    Return command.ExecuteScalar()
                End Using
            Else
                Dim closeConnection As Boolean = command.Connection.State = ConnectionState.Closed

                Try
                    If command.Connection.State <> ConnectionState.Open Then command.Connection.Open()
                    Return command.ExecuteScalar()
                Finally
                    If command.Connection.State = ConnectionState.Open AndAlso closeConnection Then command.Connection.Close()
                End Try
            End If
        End Function

#End Region

        Public Function CreateDataTable(name As String, command As SqlCommand) As DataTable
            Dim value As New DataTable(name)

            If command.Connection Is Nothing Then
                Using cnn = New SqlConnection(_connectionString)
                    command.Connection = cnn

                    Using adpt = New SqlDataAdapter(command)
                        adpt.Fill(value)
                    End Using
                End Using
            Else
                Dim closeConnection As Boolean = command.Connection.State = ConnectionState.Closed

                Try
                    If command.Connection.State <> ConnectionState.Open Then command.Connection.Open()
                    Using adpt = New SqlDataAdapter(command)
                        adpt.Fill(value)
                    End Using
                Finally
                    If command.Connection.State = ConnectionState.Open AndAlso closeConnection Then command.Connection.Close()
                End Try
            End If

            Return value
        End Function

        Public Function GetParameter(value As Object) As Object
            If value Is Nothing Or value Is DBNull.Value Then
                Return DBNull.Value
            Else
                'If Nullable.GetUnderlyingType(Value) Is Nothing Then
                Return value
                'Else

                'End If
            End If
        End Function

#Region "GetObject"

        Public Function GetObject(Of TValueType As {IDBObject(Of TValueType), New})(commandText As String, ByVal ParamArray parameters As Object()) As TValueType
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return GetObject(Of TValueType)(command)
                End Using
            End Using
        End Function

        'Public Function GetObject(Of ValueType As {IDBObject(Of ValueType), New})(ByVal cnn As SqlConnection, ByVal CommandText As String, ByVal ParamArray Parameters As Object()) As ValueType
        '    Using Command As SqlCommand = CreateCommand(cnn, CommandText, Parameters)
        '        Using Command.Connection
        '            Return GetObject(Of ValueType)(Command)
        '        End Using
        '    End Using
        'End Function

        ''' <summary>
        ''' Creates an object of the specified type from the values contained within the results of the specified Command.
        ''' </summary>
        ''' <typeparam name="TValueType"></typeparam>
        ''' <param name="Command"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetObject(Of TValueType As {IDBObject(Of TValueType), New})(command As SqlCommand) As TValueType
            Dim value As New TValueType
            Dim closeConnection As Boolean

            Try
                If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)
                If command.Connection.State = ConnectionState.Closed Then
                    closeConnection = True
                    command.Connection.Open()
                Else
                    closeConnection = False
                End If

                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        value = value.FromDataReaderInstance(reader)
                    Else
                        Return Nothing
                    End If
                End Using
            Finally
                If closeConnection Then command.Connection.Close()
            End Try

            Return value
        End Function

#End Region

#Region "GetObjectArray"

        Public Function GetObjectArray(Of TValueType)(commandText As String, ByVal ParamArray parameters() As Object) As TValueType()
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return GetObjectArray(Of TValueType)(command)
                End Using
            End Using
        End Function

        Public Function GetObjectArray(Of TValueType)(command As SqlCommand) As TValueType()
            Dim value As New List(Of TValueType)
            Dim closeConnection As Boolean = False

            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)

            Try
                If command.Connection.State = ConnectionState.Closed Then
                    command.Connection.Open()
                    closeConnection = True
                End If

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        value.Add(CType(reader(0), TValueType))
                    End While
                End Using

            Finally
                If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
            End Try

            Return value.ToArray()
        End Function

#End Region

#Region "GetObjectList"

        'Public Function GetObjectList(Of TValueType As {IDBObject(Of TValueType), New})(commandText As String, ByVal ParamArray parameters() As Object) As List(Of TValueType)
        '    Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
        '        Using command.Connection
        '            Return GetObjectList(Of List(Of TValueType), TValueType)(command)
        '        End Using
        '    End Using
        'End Function

        'Public Function GetObjectList(Of TValueType As {IDBObject(Of TValueType), New})(command As SqlCommand) As List(Of TValueType)
        '    Return GetObjectList(Of List(Of TValueType), TValueType)(command)
        'End Function

        'Public Function GetObjectList(Of TListType As {New, IList(Of TValueType)}, TValueType As {IDBObject(Of TValueType), New})(commandText As String, ByVal ParamArray parameters() As Object) As TListType
        '    Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
        '        Using command.Connection
        '            Return GetObjectList(Of TListType, TValueType)(command)
        '        End Using
        '    End Using
        'End Function

        'Public Function GetObjectList(Of TListType As {New, IList(Of TValueType)}, TValueType As {IDBObject(Of TValueType), New})(command As SqlCommand) As TListType
        '    Dim value As New TListType
        '    Dim item As New TValueType
        '    Dim closeConnection As Boolean = False

        '    If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)

        '    Try
        '        If command.Connection.State = ConnectionState.Closed Then
        '            command.Connection.Open()
        '            closeConnection = True
        '        End If

        '        Using reader As SqlDataReader = command.ExecuteReader()
        '            While reader.Read()
        '                value.Add(item.FromDataReaderInstance(reader))
        '            End While
        '        End Using

        '    Finally
        '        If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
        '    End Try

        '    Return value
        'End Function

        'Public Function GetObjectList(Of TValueType As New)(cnn as SqlConnection, commandText As String, ByVal ParamArray parameters() As Object) As List(Of TValueType)
        '    Using command As SqlCommand = CreateSqlCommand(cnn, commandText, parameters)
        '        Using command.Connection
        '            Return GetObjectList(Of List(Of TValueType), TValueType)(command)
        '        End Using
        '    End Using
        'End Function

        Public Function GetObjectList(Of TValueType As New)(commandText As String, ByVal ParamArray parameters() As Object) As List(Of TValueType)
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return GetObjectList(Of List(Of TValueType), TValueType)(command)
                End Using
            End Using
        End Function

        Public Async Function GetObjectListAsync(Of TValueType As New)(commandText As String, ByVal ParamArray parameters() As Object) As Task(Of List(Of TValueType))
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return Await GetObjectListAsync(Of List(Of TValueType), TValueType)(command)
                End Using
            End Using
        End Function

        Public Function GetObjectList(Of TValueType As New)(command As SqlCommand) As List(Of TValueType)
            Return GetObjectList(Of List(Of TValueType), TValueType)(command)
        End Function

        Public Async Function GetObjectListAsync(Of TValueType As New)(command As SqlCommand) As Task(Of List(Of TValueType))
            Return Await GetObjectListAsync(Of List(Of TValueType), TValueType)(command)
        End Function

        Public Function GetObjectList(Of TListType As {New, IList(Of TValueType)}, TValueType As New)(commandText As String, ByVal ParamArray parameters() As Object) As TListType
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return GetObjectList(Of TListType, TValueType)(command)
                End Using
            End Using
        End Function

        Public Async Function GetObjectListAsync(Of TListType As {New, IList(Of TValueType)}, TValueType As New)(commandText As String, ByVal ParamArray parameters() As Object) As Task(Of TListType)
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return Await GetObjectListAsync(Of TListType, TValueType)(command)
                End Using
            End Using
        End Function

        Public Function GetObjectList(Of TListType As {New, IList(Of TValueType)}, TValueType As New)(command As SqlCommand) As TListType
            Dim value As New TListType
            Dim closeConnection = False

            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)

            Try
                If command.Connection.State = ConnectionState.Closed Then
                    command.Connection.Open()
                    closeConnection = True
                End If

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        value.Add(CreateObject(Of TValueType)(reader))
                    End While
                End Using

            Finally
                If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
            End Try

            Return value
        End Function

        Public Async Function GetObjectListAsync(Of TListType As {New, IList(Of TValueType)}, TValueType As New)(command As SqlCommand) As Task(Of TListType)
            Dim value As New TListType
            Dim closeConnection = False

            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)

            Try
                If command.Connection.State = ConnectionState.Closed Then
                    Await command.Connection.OpenAsync()
                    closeConnection = True
                End If

                Using reader As SqlDataReader = Await command.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        value.Add(CreateObject(Of TValueType)(reader))
                    End While
                End Using

            Finally
                If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
            End Try

            Return value
        End Function

        Public Function GetObjectListFromFactory(Of TValueType As {New}, TFactoryType As {IObjectFactory(Of TValueType), New})(commandText As String, ByVal ParamArray parameters() As Object) As IEnumerable(Of TValueType)
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return GetObjectListFromFactory(Of TValueType, TFactoryType)(command)
                End Using
            End Using
        End Function

        Public Async Function GetObjectListFromFactoryAsync(Of TValueType As {New}, TFactoryType As {IObjectFactory(Of TValueType), New})(commandText As String, ByVal ParamArray parameters() As Object) As Task(Of IEnumerable(Of TValueType))
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return Await GetObjectListFromFactoryAsync(Of TValueType, TFactoryType)(command)
                End Using
            End Using
        End Function

        Public Function GetObjectListFromFactory(Of TValueType As {New}, TFactoryType As {IObjectFactory(Of TValueType), New})(command As SqlCommand) As IEnumerable(Of TValueType)
            Dim value As New List(Of TValueType)
            Dim closeConnection = False
            Dim factory As New TFactoryType

            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)

            Try
                If command.Connection.State = ConnectionState.Closed Then
                    command.Connection.Open()
                    closeConnection = True
                End If

                'command.CommandType = CommandType.StoredProcedure

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        value.Add(factory.FromDataReader(reader))
                    End While
                End Using
            Catch ex As Exception
                Dim a = ex
            Finally
                If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
            End Try

            Return value
        End Function

        Public Async Function GetObjectListFromFactoryAsync(Of TValueType As {New}, TFactoryType As {IObjectFactory(Of TValueType), New})(command As SqlCommand) As Task(Of IEnumerable(Of TValueType))
            Dim value As New List(Of TValueType)
            Dim closeConnection = False
            Dim factory As New TFactoryType

            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)

            Try
                If command.Connection.State = ConnectionState.Closed Then
                    Await command.Connection.OpenAsync()
                    closeConnection = True
                End If

                'command.CommandType = CommandType.StoredProcedure

                Using reader As SqlDataReader = Await command.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        value.Add(factory.FromDataReader(reader))
                    End While
                End Using
                'Catch ex As Exception
                '    Dim a = ex
            Finally
                If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
            End Try

            Return value
        End Function

#End Region

#Region "GetObjectDictionary"

        Public Shared Function GetObjectDictionary(Of TKeyType, TValueType As IDBIdentity)(items As IEnumerable(Of TValueType)) As Dictionary(Of TKeyType, TValueType)
            Dim value As New Dictionary(Of TKeyType, TValueType)

            For Each item As TValueType In items
                If Not value.ContainsKey(item.GetKey()) Then value.Add(item.GetKey(), item)
            Next

            Return value
        End Function

        Public Function GetObjectDictionary(Of TKeyType, TValueType As {IDBObject(Of TValueType), New})(commandText As String, ByVal ParamArray parameters() As Object) As Dictionary(Of TKeyType, TValueType)
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return GetObjectDictionary(Of TKeyType, TValueType)(command)
                End Using
            End Using
        End Function

        Public Function GetObjectDictionary(Of TKeyType, TValueType As {IDBObject(Of TValueType), New})(keyColumn As String, sql As String, params As Object()) As Dictionary(Of TKeyType, TValueType)
            Using command As SqlCommand = CreateSqlCommand(sql, params)
                Using command.Connection
                    Return GetObjectDictionary(Of TKeyType, TValueType)(keyColumn, command)
                End Using
            End Using

        End Function

        Public Function GetObjectDictionary(Of TKeyType, TValueType As {IDBObject(Of TValueType), New})(command As SqlCommand) As Dictionary(Of TKeyType, TValueType)
            Dim value As New Dictionary(Of TKeyType, TValueType)
            Dim item As New TValueType
            Dim closeConnection = False

            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)


            Try
                If command.Connection.State = ConnectionState.Closed Then
                    command.Connection.Open()
                    closeConnection = True
                End If

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        item = item.FromDataReaderInstance(reader)
                        value.Add(CType(item.GetKey, TKeyType), item)
                    End While
                End Using

            Finally
                If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
            End Try

            Return value
        End Function

        Public Function GetObjectDictionary(Of TKeyType, TValueType As {IDBObject(Of TValueType), New})(keyColumn As String, command As SqlCommand) As Dictionary(Of TKeyType, TValueType)
            Dim value As New Dictionary(Of TKeyType, TValueType)
            Dim item As New TValueType
            Dim closeConnection = False

            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)

            Try
                If command.Connection.State = ConnectionState.Closed Then
                    command.Connection.Open()
                    closeConnection = True
                End If

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        item = item.FromDataReaderInstance(reader)
                        value.Add(CType(reader(keyColumn), TKeyType), item)
                    End While
                End Using
            Finally
                If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
            End Try

            Return value
        End Function

#End Region

#Region "GetValueList"

        Public Function GetValueList(Of TValueType)(commandText As String, ByVal ParamArray parameters() As Object) As List(Of TValueType)
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return GetValueList(Of TValueType)(command)
                End Using
            End Using
        End Function

        Public Function GetValueList(Of TValueType)(command As SqlCommand) As List(Of TValueType)
            Dim value As New List(Of TValueType)
            Dim closeConnection = False

            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)

            Try
                If command.Connection.State = ConnectionState.Closed Then
                    command.Connection.Open()
                    closeConnection = True
                End If

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        value.Add(reader(0))
                    End While
                End Using

            Finally
                If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
            End Try

            Return value
        End Function

#End Region

#Region "GetDictionary"

        Public Function GetDictionary(Of TKeyType, TValueType)(commandText As String, keyColumn As String, valueColumn As String, ByVal ParamArray parameters() As Object) As Dictionary(Of TKeyType, TValueType)
            Using command As SqlCommand = CreateSqlCommand(commandText, parameters)
                Using command.Connection
                    Return GetDictionary(Of TKeyType, TValueType)(command, keyColumn, valueColumn)
                End Using
            End Using
        End Function

        Public Function GetDictionary(Of TKeyType, TValueType)(command As SqlCommand, keyColumn As String, valueColumn As String) As Dictionary(Of TKeyType, TValueType)
            Dim value As New Dictionary(Of TKeyType, TValueType)
            Dim closeConnection = False

            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)

            Try
                If command.Connection.State = ConnectionState.Closed Then
                    command.Connection.Open()
                    closeConnection = True
                End If

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        value.Add(reader(keyColumn), reader(valueColumn))
                    End While
                End Using

            Finally
                If closeConnection And command.Connection.State = ConnectionState.Open Then command.Connection.Close()
            End Try

            Return value
        End Function

#End Region

#Region "Exists"

        ''' <summary>
        ''' Determines if any records are returned from the specified query.
        ''' </summary>
        ''' <param name="CommandText"></param>
        ''' <param name="Parameters"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Exists(commandText As String, ByVal ParamArray parameters() As Object) As Boolean
            Using command As SqlCommand = CreateCommand(Nothing, Nothing, commandText, parameters)
                Using command.Connection
                    Return Exists(command)
                End Using
            End Using
        End Function

        Public Function Exists(cnn As SqlConnection, commandText As String, ByVal ParamArray parameters() As Object) As Boolean
            Using command As SqlCommand = CreateCommand(cnn, commandText, parameters)
                Return Exists(command)
            End Using
        End Function

        Public Function Exists(cnn As SqlConnection, transaction As SqlTransaction, commandText As String, ByVal ParamArray parameters() As Object) As Boolean
            Using command As SqlCommand = CreateCommand(cnn, transaction, commandText, parameters)
                Return Exists(command)
            End Using
        End Function

        Public Function Exists(command As SqlCommand) As Boolean
            If command.Connection Is Nothing Then command.Connection = New SqlConnection(_connectionString)
            Dim closeConnection As Boolean = (command.Connection.State = ConnectionState.Closed)


            Try
                If closeConnection Then command.Connection.Open()

                command.CommandText = String.Concat("SELECT CAST( CASE WHEN EXISTS(", command.CommandText, ") THEN 1 ELSE 0 End AS BIT)")

                Return command.ExecuteScalar()
            Finally
                If closeConnection Then command.Connection.Close()
            End Try
        End Function

#End Region

#Region "CreateParameters"

        Public Shared Sub CreateParameters(parameterName As String, items As ICollection, ByRef commandText As String, ByRef parameters As SqlParameterCollection)
            commandText = CreateParameters(parameterName, items, parameters)
        End Sub

        Public Shared Function CreateParameters(items As ICollection, ByRef parameters As SqlParameterCollection) As String
            Return CreateParameters("IndexedParam", items, parameters)
        End Function

        Public Shared Function CreateParameters(parameterName As String, items As ICollection, ByRef parameters As SqlParameterCollection) As String
            Dim value As New StringBuilder()

            For index = 0 To items.Count - 1
                value.AppendFormat("@{0}{1},", parameterName, index)
                parameters.AddWithValue(String.Concat("@", parameterName, index), items(index))
            Next

            If value.Chars(value.Length - 1) = "," Then value.Remove(value.Length - 1, 1)

            Return value.ToString()
        End Function

#End Region

        Public Shared Function CreateObject(Of TValueType As New)(reader As IDataReader) As TValueType
            Dim obj As New TValueType

            If TypeOf obj Is IDBObject(Of TValueType) Then
                Return DirectCast(obj, IDBObject(Of TValueType)).FromDataReaderInstance(reader)
            Else
                Dim props = obj.GetType().GetProperties()

                For Each p In props
                    If (p.CanWrite) Then
                        Dim attr = CType(p.GetCustomAttribute(GetType(ColumnAttribute)), ColumnAttribute)

                        If attr IsNot Nothing Then
                            If (Not String.IsNullOrEmpty(attr.Name)) AndAlso reader(attr.Name) IsNot DBNull.Value Then
                                p.SetValue(obj, CTypeDynamic(reader(attr.Name), p.PropertyType))
                            End If
                        End If
                    End If
                Next

                Return obj
            End If
        End Function


        'Public Shared Function BuildUpdateString(tableName As String, map As Dictionary(Of String, Object), idColumn As String) As String
        '    Dim sql = New StringBuilder($"UPDATE {tableName} SET ")

        '    For Each entry In map
        '        Dim columnName = entry.Key
        '        Dim parameterName = entry.Key

        '        If columnName.StartsWith("@") Then columnName = columnName.TrimStart(New Char() {"@"})
        '        If Not parameterName.StartsWith("@") Then parameterName = $"@{parameterName}"

        '        sql.Append($"{columnName} = {parameterName}, ")
        '    Next

        '    sql.Length -= 2

        '    sql.Append($" WHERE {idColumn} = {If(map.ContainsKey(idColumn), map(idColumn), map("@" + idColumn))}")


        '    Return sql.ToString()
        'End Function

        'Public Shared Function BuildInsertString(tableName As String, map As Dictionary(Of String, Object), idColumn As String) As String
        '    Dim sql = New StringBuilder($"INSERT INTO {tableName} ")

        '    For Each entry In map
        '        Dim columnName = entry.Key

        '        If(columnName != idColumn)

        '        If columnName.StartsWith("@") Then columnName = columnName.TrimStart(New Char() {"@"})

        '        sql.Append($"{columnName}, ")
        '    Next

        '    sql.Length -= 2

        '    For Each entry In map
        '        Dim parameterName = entry.Key

        '        If Not parameterName.StartsWith("@") Then parameterName = $"@{parameterName}"

        '        sql.Append($"{parameterName}, ")
        '    Next

        '    Return sql.ToString()
        'End Function


    End Class

End Namespace
