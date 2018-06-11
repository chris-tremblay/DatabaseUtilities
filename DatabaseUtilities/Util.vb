Imports System.Collections.Generic
Imports System.Data.SqlClient

Namespace DatabaseUtilities
    Public Module Util

        Public Function CreateIntegerListTVP(parameterName As String, values As IEnumerable(Of Integer), Optional typeName As String = "Integer_List") As SqlParameter
            Dim value As New SqlParameter()
            Dim table As New DataTable()

            Table.Columns.Add("Value", GetType(Integer))

            For Each Val As Integer In Values
                Table.Rows.Add(Val)
            Next

            Value.Direction = ParameterDirection.Input
            Value.ParameterName = ParameterName
            Value.SqlDbType = SqlDbType.Structured
            Value.TypeName = typeName
            Value.Value = Table


            Return Value
        End Function

        Public Function CreateStringListTVP(parameterName As String, values As IEnumerable(Of String), Optional typeName As String = "String_List") As SqlParameter
            Dim value As New SqlParameter()
            Dim table As New DataTable()

            table.Columns.Add("Value", GetType(String))

            For Each Val As String In Values
                table.Rows.Add(New String() {Val})
            Next

            value.Direction = ParameterDirection.Input
            value.ParameterName = ParameterName
            value.SqlDbType = SqlDbType.Structured

            'If expectedLength.HasValue AndAlso expectedLength <= 20 Then
            'Value.TypeName = "String_List_20"
            'Else
            value.TypeName = typeName
            'End If

            value.Value = table

            Return value
        End Function

        Public Function CreateStringListTVP(parameterName As String, expectedLength As Integer?, values As IEnumerable(Of String), Optional typeName As String = "String_List") As SqlParameter
            return CreateStringListTVP(parameterName, values, typeName)
        End Function

        ''' <summary>
        ''' Creates a parameter 
        ''' </summary>
        ''' <param name="ParameterName">The name of the parameter that will be added to the <paramref name="CommandText"/>.</param>
        ''' <param name="Items">The list of items that will have parameters created for them.</param>
        ''' <param name="CommandText">The String that will hold new list of paramters.</param>
        ''' <param name="Parameters">The SQLParameterCollection that the parameters will be added to.</param>
        ''' <remarks></remarks>
        Public Function CreateParameters(ByVal parameterName As String, ByVal items As ICollection, ByRef commandText As String, ByRef parameters As SqlParameterCollection) As String
            Dim counter As Integer = 0

            For Each item As Object In Items
                CommandText += "@" & ParameterName & counter & ","
                Parameters.AddWithValue("@" & ParameterName & counter, item)
                counter += 1
            Next

            If CommandText.EndsWith(",") Then CommandText = CommandText.Remove(CommandText.Length - 1)

            Return CommandText
        End Function

        ''' <summary>
        ''' Creates a parameter 
        ''' </summary>
        ''' <param name="ParameterName"></param>
        ''' <param name="Items">The list of items that will have parameters created for them.</param>
        ''' <param name="Parameters">The SQLParameterCollection that the parameters will be added to.</param>
        ''' <remarks></remarks>
        Public Function CreateParameters(ByVal parameterName As String, ByVal items As ICollection, ByRef parameters As SqlParameterCollection) As String
            Dim counter As Integer = 0
            Dim value As String = ""

            For Each item As Object In items
                value += "@" & parameterName & counter & ","
                parameters.AddWithValue("@" & parameterName & counter, item)
                counter += 1
            Next

            If Value.EndsWith(",") Then Value = Value.Remove(Value.Length - 1)

            Return Value
        End Function

        Public Function CreatePagedSelectCommand(ByVal selectCommand As SqlCommand, ByVal pageNumber As Integer, ByVal recordsPerPage As Integer, ByRef numberOfRecords As Integer, ByVal defaultOrderColumn As String) As SqlCommand
            Dim cnn As SqlConnection
            Dim pagedCommandText As String = SelectCommand.CommandText
            Dim countCommandText As String = SelectCommand.CommandText
            Dim orderBy As String
            Dim sqlOption As String

            '************************************************************************************
            '* Determine if a new connection must be established, or if the existing connection *
            '* can be used.                                                                     *
            '************************************************************************************
            If SelectCommand.Connection Is Nothing Then
                Throw New ArgumentException("The 'Connection' property of the SelectCommand has not been initialized. You must set the Connection property to a valid SQLConnection instance.")
            Else
                cnn = SelectCommand.Connection
            End If

            If cnn.State = ConnectionState.Closed Then cnn.Open()

            '************************************************************************************
            '* Determine the total number of records returned by the query before being paged.  *
            '************************************************************************************
            If countCommandText.LastIndexOf(" OPTION ") <> -1 Then countCommandText = countCommandText.Remove(countCommandText.LastIndexOf(" OPTION "))
            If countCommandText.LastIndexOf(" ORDER BY ") > -1 Then countCommandText = countCommandText.Remove(countCommandText.LastIndexOf(" ORDER BY "))

            countCommandText = String.Concat("SELECT Count(*) FROM (", countCommandText, ") Results")

            SelectCommand.CommandText = countCommandText
            NumberOfRecords = CInt(SelectCommand.ExecuteScalar())

            '************************************************************************************
            '* INSERT the paging SQL into the SelectCommand to page the results.                *
            '************************************************************************************
            If pagedCommandText.LastIndexOf(" OPTION ") <> -1 Then
                sqlOption = pagedCommandText.Substring(pagedCommandText.LastIndexOf(" OPTION "))
                pagedCommandText = pagedCommandText.Remove(pagedCommandText.LastIndexOf(" OPTION "))
            End If

            If pagedCommandText.LastIndexOf("ORDER BY") <> -1 Then
                '*** The ORDER BY statement must be moved to the outer query ***
                orderBy = pagedCommandText.Substring(pagedCommandText.LastIndexOf("ORDER BY") + 9)
                pagedCommandText = pagedCommandText.Remove(pagedCommandText.LastIndexOf("ORDER BY"))
            Else
                '*** Default Order ***
                orderBy = DefaultOrderColumn
            End If

            pagedCommandText = pagedCommandText.Insert(pagedCommandText.IndexOf(" FROM ") + 1, ", ROW_NUMBER() OVER (ORDER BY " & orderBy & ") As Rownum ")
            pagedCommandText = String.Concat("SELECT * FROM (", pagedCommandText, ") AS Results WHERE Rownum >= @Paging_RowStart AND Rownum <= @Paging_RowEnd")

            SelectCommand.Parameters.AddWithValue("@Paging_RowStart", ((PageNumber - 1) * RecordsPerPage) + 1)
            SelectCommand.Parameters.AddWithValue("@Paging_RowEnd", PageNumber * RecordsPerPage)
            SelectCommand.CommandText = pagedCommandText

            Return SelectCommand
        End Function

        Public Function CreatePagedSelectCommand(selectCommand As SqlCommand, pageNumber As Integer, recordsPerPage As Integer, defaultOrderColumn As String) As SqlCommand
            Dim pagedCommandText As String = SelectCommand.CommandText
            Dim orderBy As String
            
            '************************************************************************************
            '* INSERT the paging SQL into the SelectCommand to page the results.                *
            '************************************************************************************
            'If pagedCommandText.LastIndexOf(" OPTION ") <> -1 Then
            '    //sqlOption = pagedCommandText.Substring(pagedCommandText.LastIndexOf(" OPTION "))
            '    pagedCommandText = pagedCommandText.Remove(pagedCommandText.LastIndexOf(" OPTION "))
            'End If

            If pagedCommandText.LastIndexOf("ORDER BY") <> -1 Then
                '*** The ORDER BY statement must be moved to the outer query ***
                orderBy = pagedCommandText.Substring(pagedCommandText.LastIndexOf("ORDER BY") + 9)
                pagedCommandText = pagedCommandText.Remove(pagedCommandText.LastIndexOf("ORDER BY"))
            Else
                '*** Default Order ***
                orderBy = DefaultOrderColumn
            End If

            pagedCommandText = pagedCommandText.Insert(pagedCommandText.IndexOf(" FROM ") + 1, ", ROW_NUMBER() OVER (ORDER BY " & orderBy & ") As Rownum ")
            pagedCommandText = String.Concat("SELECT * FROM (", pagedCommandText, ") AS Results WHERE Rownum >= @Paging_RowStart AND Rownum <= @Paging_RowEnd")

            SelectCommand.Parameters.AddWithValue("@Paging_RowStart", ((PageNumber - 1) * RecordsPerPage) + 1)
            SelectCommand.Parameters.AddWithValue("@Paging_RowEnd", PageNumber * RecordsPerPage)
            SelectCommand.CommandText = pagedCommandText

            Return SelectCommand
        End Function

        ''' <summary>
        ''' Executes a command and adds 2 additional columns: row_num and row_count.
        ''' </summary>
        ''' <param name="SelectCommand"></param>
        ''' <param name="PageNumber"></param>
        ''' <param name="RecordsPerPage"></param>
        ''' <param name="DefaultOrderColumn"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExecutePagedReader(selectCommand As SqlCommand, pageNumber As Integer, recordsPerPage As Integer, defaultOrderColumn As String) As SqlDataReader
            Dim pagedCommandText As String = selectCommand.CommandText
            Dim orderBy As String
            Dim sqlOption As String = nothing

            '************************************************************************************
            '* INSERT the paging SQL into the SelectCommand to page the results.                *
            '************************************************************************************
            If PagedCommandText.LastIndexOf(" OPTION ") <> -1 Then
                sqlOption = PagedCommandText.Substring(PagedCommandText.LastIndexOf(" OPTION "))
                PagedCommandText = PagedCommandText.Remove(PagedCommandText.LastIndexOf(" OPTION "))
            End If

            If PagedCommandText.LastIndexOf("ORDER BY") <> -1 Then
                '*** The ORDER BY statement must be moved to the outer query ***
                OrderBy = PagedCommandText.Substring(PagedCommandText.LastIndexOf("ORDER BY") + 9)
                PagedCommandText = PagedCommandText.Remove(PagedCommandText.LastIndexOf("ORDER BY"))
            Else
                '*** Default Order ***
                OrderBy = DefaultOrderColumn
            End If

            PagedCommandText = PagedCommandText.Insert(PagedCommandText.IndexOf(" FROM ") + 1, ", ROW_NUMBER() OVER (ORDER BY " & OrderBy & ") As Row_num, row_Count = Count(*) Over() ")
            PagedCommandText = String.Concat("SELECT * FROM (", PagedCommandText, ") AS Results WHERE Row_num >= @Paging_RowStart AND Row_num <= @Paging_RowEnd")

            SelectCommand.Parameters.AddWithValue("@Paging_RowStart", ((PageNumber - 1) * RecordsPerPage) + 1)
            SelectCommand.Parameters.AddWithValue("@Paging_RowEnd", PageNumber * RecordsPerPage)
            SelectCommand.CommandText = String.Concat(PagedCommandText, sqlOption)

            Return SelectCommand.ExecuteReader()
        End Function

         

        Public Function CreateTVP(typeName As String, parameterName As String, table As DataTable) As SqlParameter
            Dim value As New SqlParameter()

            Value.Direction = ParameterDirection.Input
            Value.ParameterName = ParameterName
            Value.SqlDbType = SqlDbType.Structured
            Value.TypeName = TypeName
            Value.Value = Table

            Return Value
        End Function

    End Module

End Namespace

