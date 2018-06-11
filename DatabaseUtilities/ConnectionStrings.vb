Namespace DatabaseUtilities
    Public Class ConnectionStrings
        Public Const SQLServer_TrustedConnection As String = "Data Source=LEWSQL01;Initial Catalog=ELMETEPA;Integrated Security=SSPI;"
        Public Const ExcelXLS As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=Yes;"""
        Public Const ExcelXLSX As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;"""
        Public Const ExcelCSV As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR={1};FMT=Delimited"""
        Public Const Access As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};User Id={1};Password={2};"

        Public Shared Function GetTextConnectionString(ByVal Folder As String) As String
            Return GetExcelCSVConnectionString(Folder, True)
        End Function

        Public Shared Function GetSQLConnectionString(ByVal databaseName As String, ByVal server As String, ByVal userName As String, ByVal password As String) As String
            Return String.Format("Password={3};User ID={2};Initial Catalog={0};Data Source={1}", DatabaseName, Server, UserName, Password)
        End Function

        Public Shared Function GetTextConnectionString(ByVal folder As String, ByVal includesHeaderRow As Boolean) As String
            Return GetExcelCSVConnectionString(Folder, IncludesHeaderRow)
        End Function

        Public Shared Function GetExcelXLSConnectionString(ByVal fileName As String) As String
            Return String.Format(ExcelXLS, FileName)
        End Function

        Public Shared Function GetExcelXLSXConnectionString(ByVal fileName As String) As String
            Return String.Format(ExcelXLSX, FileName)
        End Function

        ''' <summary>
        ''' Creates the connection string for opening a Comma-Separated File (.csv)
        ''' </summary>
        ''' <param name="Folder">The path of the folder to establish a connection with.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetExcelCSVConnectionString(ByVal folder As String) As String
            Return String.Format(ExcelCSV, Folder, "YES")
        End Function

        ''' <summary>
        ''' Creates the connection string for opening a Comma-Separated File (.csv)
        ''' </summary>
        ''' <param name="FileName">The filename to establish a connection with.</param>
        ''' <param name="IncludesHeaderRow">Specifies whether or not the file has a header row containing the names of the columns.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetExcelCSVConnectionString(ByVal fileName As String, ByVal includesHeaderRow As Boolean) As String
            If IncludesHeaderRow Then
                Return String.Format(ExcelCSV, FileName, "YES")
            Else
                Return String.Format(ExcelCSV, FileName, "NO")
            End If
        End Function

        Public Shared Function GetAccessConnectionString(ByVal fileName As String) As String
            Return String.Format(Access, FileName, "", "")
        End Function

        Public Shared Function GetAccessConnectionString(ByVal fileName As String, ByVal userName As String, ByVal password As String) As String
            Return String.Format(Access, FileName, UserName, Password)
        End Function
    End Class

End Namespace
