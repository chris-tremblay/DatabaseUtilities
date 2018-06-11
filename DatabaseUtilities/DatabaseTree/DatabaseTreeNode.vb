Imports System.Collections.ObjectModel 

Namespace DatabaseUtilities
    Namespace DatabaseTree

        <Serializable()> _
        Public Class DBTreeNode

            Public Sub New()

            End Sub

            Public Property Node() As Integer = -1
            Public Property Nodes As ObservableCollection(Of DBTreeNode) = New ObservableCollection(Of DBTreeNode)()
            Public Property ParentNode() As Integer = -1
            Public Property DateCreated() As Date
            Public Property DateModified() As Date
            Public Property Depth() As Integer = 0
            Public Property Title() As String = 0


            Public ReadOnly Property FullPath() As String
                Get
                    Return _Lineage & _Node & "/"
                End Get
            End Property

            Public Property Lineage() As String = ""


            Public ReadOnly Property IsRootCategory() As Boolean
                Get
                    Return (_ParentNode = -1)
                End Get
            End Property

            Public Shared Function FromDataReader(ByVal dtr As System.Data.IDataReader) As DBTreeNode
                Dim value As New DBTreeNode()

                DBTreeNode.FillFromDataReader(value, dtr)

                Return value
            End Function

            Public Shared Sub FillFromDataReader(ByVal Node As DBTreeNode, ByVal dtr As System.Data.IDataReader)
                If dtr.GetSchemaTable().Select("ColumnName = 'Node'").Length > 0 Then
                    If Not dtr.IsDBNull(dtr.GetOrdinal("Node")) Then Node.Node = dtr!Node
                    If Not dtr.IsDBNull(dtr.GetOrdinal("ParentNode")) Then Node.ParentNode = dtr!ParentNode
                    If Not dtr.IsDBNull(dtr.GetOrdinal("Depth")) Then Node.Depth = dtr!Depth
                    If Not dtr.IsDBNull(dtr.GetOrdinal("Lineage")) Then Node.Lineage = dtr!Lineage
                    If Not dtr.IsDBNull(dtr.GetOrdinal("CategoryTitle")) Then Node.Title = dtr!CategoryTitle
                    If Not dtr.IsDBNull(dtr.GetOrdinal("DateCreated")) Then Node.DateCreated = dtr!DateCreated
                    If Not dtr.IsDBNull(dtr.GetOrdinal("DateModified")) Then Node.DateModified = dtr!DateModified
                End If
            End Sub

            Public Shared Function FromDataRow(ByVal dRow As DataRow) As DBTreeNode
                Dim Value As New DBTreeNode()

                If Not dRow.IsNull("Node") Then Value.Node = dRow!Node
                If Not dRow.IsNull("ParentNode") Then Value.ParentNode = dRow!ParentNode
                If Not dRow.IsNull("Depth") Then Value.Depth = dRow!Depth
                If Not dRow.IsNull("Lineage") Then Value.Lineage = dRow!Lineage
                If Not dRow.IsNull("CategoryTitle") Then Value.Title = dRow!CategoryTitle
                If Not dRow.IsNull("DateCreated") Then Value.DateCreated = dRow!DateCreated
                If Not dRow.IsNull("DateModified") Then Value.DateModified = dRow!DateModified

                Return Value
            End Function
        End Class
    End Namespace

End Namespace

