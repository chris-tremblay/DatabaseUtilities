
Namespace DatabaseUtilities

    ''' <summary>
    ''' Creates object of a given type.
    ''' </summary>
    ''' <typeparam name="TObjectType"></typeparam>
    ''' <remarks></remarks>
    Public Interface IObjectFactory(Of TObjectType)

#Region "Public Methods"

        ''' <summary>
        ''' Creates an object of the specified type from a data reader.
        ''' </summary>
        ''' <param name="reader"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function FromDataReader(reader As IDataReader) As TObjectType
        
#End Region

    End Interface

End Namespace

