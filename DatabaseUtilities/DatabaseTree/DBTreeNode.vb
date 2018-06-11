Imports System.Collections.Generic

Namespace DatabaseUtilities
    Namespace DatabaseTree
        Public Class DBTreeNode

#Region "Private Variables"

            Private _ID As Integer
            Private _Children As New List(Of DBTreeNode)
            Private _ParentNode As Integer
            Private _Parent As DBTreeNode
            Private _Lineage As String
            Private _LineageText As String
            Private _Name As String
            Private _Depth As Integer
            Private _IsProtected As Boolean = False

#End Region

#Region "Constructors"

            Public Sub New()

            End Sub

#End Region

#Region "Properties"

            Public ReadOnly Property Children() As List(Of DBTreeNode)
                Get
                    Return _Children
                End Get
            End Property

            Public Property ID() As Integer
                Get
                    Return _ID
                End Get
                Set(ByVal value As Integer)
                    _ID = value
                End Set
            End Property

            Public Property Parent() As DBTreeNode
                Get
                    Return _Parent
                End Get
                Set(ByVal value As DBTreeNode)
                    _Parent = value
                End Set
            End Property

            Public Property ParentNode() As Integer
                Get
                    Return _ParentNode
                End Get
                Set(ByVal value As Integer)
                    _ParentNode = value
                End Set
            End Property

            Public Property Lineage() As String
                Get
                    Return _Lineage
                End Get
                Set(ByVal value As String)
                    _Lineage = value
                End Set
            End Property

            Public Property LineageText() As String
                Get
                    Return _LineageText
                End Get
                Set(ByVal value As String)
                    _LineageText = value
                End Set
            End Property

            Public Property Name() As String
                Get
                    Return _Name
                End Get
                Set(ByVal value As String)
                    _Name = value
                End Set
            End Property

            Public Property Depth() As Integer
                Get
                    Return _Depth
                End Get
                Set(ByVal value As Integer)
                    _Depth = value
                End Set
            End Property

            Public Property IsProtected() As Boolean
                Get
                    Return _IsProtected
                End Get
                Set(ByVal value As Boolean)
                    _IsProtected = value
                End Set
            End Property

            Public ReadOnly Property FullPath() As String
                Get
                    Return _Lineage & _ID & "\"
                End Get
            End Property

            Public ReadOnly Property FullPathText() As String
                Get
                    Return _LineageText & _Name & "\"
                End Get
            End Property

#End Region

#Region "Methods"

            Public Function CountChildren(ByVal Recursive As Boolean) As Integer
                If Not Recursive Then
                    Return Me.Children.Count
                Else
                    Dim Value As Integer = Me.Children.Count

                    For Each Child As DBTreeNode In Me.Children
                        Value += Child.CountChildren(True)
                    Next

                    Return Value
                End If
            End Function

            Public Shared Function FromDataReader(ByVal Reader As IDataReader) As DBTreeNode
                Dim Value As New DBTreeNode()

                Try
                    Value.Depth = Reader.GetInt32(Reader.GetOrdinal("Depth"))
                    Value.Lineage = Reader.GetString(Reader.GetOrdinal("Lineage"))
                    Value.LineageText = Reader.GetString(Reader.GetOrdinal("LineageText"))
                    Value.ID = Reader.GetInt32(Reader.GetOrdinal("Node"))
                    If Not Reader.IsDBNull(Reader.GetOrdinal("Parent")) Then Value.ParentNode = Reader.GetInt32(Reader.GetOrdinal("Parent"))
                    Value.Name = Reader.GetString(Reader.GetOrdinal("Name"))
                    Value.IsProtected = Reader.GetBoolean(Reader.GetOrdinal("IsProtected"))
                Catch ex As Exception

                End Try

                Return Value
            End Function

#End Region

        End Class
    End Namespace
End Namespace


