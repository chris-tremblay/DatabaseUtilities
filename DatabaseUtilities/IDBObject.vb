Namespace DatabaseUtilities

    Public Interface IDBObject
        Inherits IDbIdentity

        Function FromDataReader(ByVal reader As IDataReader) As Object

    End Interface

    Public Interface IDBObject(Of T)
        Inherits IDbIdentity, IDbObject

        Function FromDataReaderInstance(ByVal reader As IDataReader) As T

    End Interface

End Namespace



