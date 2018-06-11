Imports System.Runtime.CompilerServices

<Extension>
Public Module IDataRecordExtensions

    <Extension>
    Public Function HasColumn(dr As IDataRecord, columnName As String) As Boolean

        For i = 0 To dr.FieldCount - 1
            If (dr.GetName(i).Equals(columnName, StringComparison.InvariantCultureIgnoreCase)) Then
                Return True
            End If
        Next

        Return False
    End Function

End Module