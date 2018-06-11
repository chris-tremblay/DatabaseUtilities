Imports System.Collections.Generic 
Imports System.Data.SqlClient

Namespace DatabaseUtilities
    Namespace DatabaseTree
        Public Class DBTree
            Inherits SQLDatabase

#Region "Private Variables"

            Private _TreeTableName As String

#End Region

#Region "Constructors"

            Public Sub New(ByVal TreeTableName As String, ConnectionString As String)
                MyBase.New(ConnectionString)
                _TreeTableName = "SettingCategories"
            End Sub

            Public Sub New(ByVal TreeTableName As String, ByVal cnn As IDbConnection)
                MyBase.New(cnn.ConnectionString)
                _TreeTableName = TreeTableName
            End Sub

#End Region

#Region "Properties"

            Public Property TreeTableName() As String
                Get
                    Return _TreeTableName
                End Get
                Set(ByVal value As String)
                    _TreeTableName = value
                End Set
            End Property

#End Region

            Shared Function CreateParameter(ByVal Command As IDbCommand, ByVal ParameterName As String, ByVal ParameterValue As Object) As IDbDataParameter
                Dim Value As IDbDataParameter = Command.CreateParameter()

                Value.ParameterName = ParameterName

                If ParameterValue Is Nothing Then
                    Value.Value = System.DBNull.Value
                Else
                    Value.Value = ParameterValue
                End If


                Return Value
            End Function

            ''' <summary>
            ''' Gets the sub-tree for the specified path.
            ''' </summary>
            ''' <param name="Path"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function GetTree(ByVal Path As String) As DBTreeNode
                Dim Index As New Dictionary(Of Integer, DBTreeNode)
                Dim Value As New DBTreeNode()
                Dim Node As DBTreeNode

                Using cnn As IDbConnection = CreateConnection()

                    cnn.Open()
                    Using Command As IDbCommand = cnn.CreateCommand()
                        Command.CommandText = String.Format("SELECT * FROM {0} WHERE Node = @Node", _TreeTableName)

                        '*****************************
                        '* Get the root node details *
                        '*****************************
                        Command.Parameters.Add(CreateParameter(Command, "@Node", Me.GetNode(Path)))

                        Using Reader As IDataReader = Command.ExecuteReader()
                            If Reader.Read() Then
                                Value = DBTreeNode.FromDataReader(Reader)
                                Index.Add(Value.ID, Value)
                            Else
                                Return Nothing
                            End If
                        End Using
                    End Using

                    '************************************************************************
                    '* Get the sub-tree for the selcted node. We build the tree by indexing *
                    '* each node in a dictionary. For each Sub-item of the selected tree,   *
                    '* Add the node to the index, and then see if the parent node belongs   *
                    '* to the table. Since the data is sorted by depth, the parent node     *
                    '* should always be available unless records were deleted to corrupt    *
                    '* the data or remove a branch. If the parent node does exist, then     *
                    '* set the parent node of the current node and add the current node to  *
                    '* the children of the parent node.                                     *
                    '************************************************************************
                    Using Command As IDbCommand = cnn.CreateCommand()
                        Command.CommandText = String.Format("SELECT * FROM {0} WHERE LineageText LIKE @LineageText ORDER BY Depth, Name", _TreeTableName)
                        Command.Parameters.Add(CreateParameter(Command, "@LineageText", Path & "%"))

                        Using Reader As IDataReader = Command.ExecuteReader()
                            While Reader.Read()
                                Node = DBTreeNode.FromDataReader(Reader)

                                Index.Add(Node.ID, Node)

                                If Index.ContainsKey(Node.ParentNode) Then
                                    Node.Parent = Index(Node.ParentNode)
                                    Node.Parent.Children.Add(Node)
                                End If
                            End While
                        End Using
                    End Using

                    Return Value
                End Using


                Return Value
            End Function

            '''<summary>
            '''Gets the index of the node with the specified path. Returns -1 if not found.
            '''</summary>
            Public Function GetNode(ByVal Path As String) As Integer
                '*****************************************************************
                '* Created By     : Chris Tremblay                               *
                '* Date Created   : 8/1/2006                                     *
                '* Date Modified  : 8/1/2006                                     *
                '* Description    :                                              *
                '* Dependencies   :                                              *
                '*****************************************************************
                Dim Lineage As String
                Dim Category As String
                Dim Index As Integer

                Using cnn As IDbConnection = CreateConnection()
                    Try


                        If Path.EndsWith("\") And Path.Length > 1 Then
                            Index = InStrRev(Path, "\", Path.Length() - 1)
                            Lineage = Left(Path, Index)
                            Category = Right(Path, Path.Length() - Index)
                        Else
                            Index = InStrRev(Path, "\")
                            Lineage = Left(Path, Index)
                            Category = Right(Path, Path.Length() - Index)
                        End If

                        If Not Lineage.StartsWith("\") Then Lineage = "\" & Lineage
                        If Category.EndsWith("\") Then Category = Left(Category, Len(Category) - 1)

                        cnn.Open()

                        Using Command As IDbCommand = cnn.CreateCommand()
                            Command.CommandText = String.Format("SELECT Node FROM {0} WHERE LineageText = @LineageText AND Name = @Category", TreeTableName)

                            Command.Parameters.Add(CreateParameter(Command, "@LineageText", Lineage))
                            Command.Parameters.Add(CreateParameter(Command, "@Category", Category))

                            Using reader As IDataReader = Command.ExecuteReader()
                                If reader.Read() Then
                                    Return reader.GetInt32(0)
                                Else
                                    Return -1
                                End If
                            End Using
                        End Using
                    Catch ex As Exception
                        Return -1
                    End Try
                End Using


                Return -1

            End Function

            Protected Function CreateNode(ByVal Path As String, ByRef NodeName As String) As Integer
                Return Me.CreateNode(Path, NodeName, False)
            End Function

            '''<summary>
            '''Creates a Category at the specified path and returns its index. Returns -1 if the specified path does not exist.
            '''</summary>
            Protected Function CreateNode(ByVal Path As String, ByRef NodeName As String, ByVal IsProtected As Boolean) As Integer
                '*****************************************************************
                '* Created By     : Chris Tremblay                               *
                '* Date Created   : 8/1/2006                                     *
                '* Date Modified  : 8/1/2006                                     *
                '* Description    :                                              *
                '* Dependencies   :                                              *
                '*****************************************************************
                Dim ParentLineage As String
                Dim ParentCategory As String
                Dim Success As Boolean = False

                Using cnn As IDbConnection = CreateConnection()
                    Try
                        cnn.Open()
                        '*** Make sure that the path starts and ends with a back-slash ***
                        If Not Path.StartsWith("\") Then Path = String.Concat("\", Path)
                        If Not Path.EndsWith("\") Then Path = String.Concat(Path, "\")

                        ParentLineage = Left(Path, InStrRev(Left(Path, Path.Length() - 1), "\"))
                        ParentCategory = Right(Left(Path, Path.Length() - 1), Path.Length() - InStrRev(Left(Path, Len(Path) - 1), "\") - 1)


                        While Not Success
                            '*** Get the index of the destination path ***
                            Using Command As IDbCommand = cnn.CreateCommand()
                                Command.CommandText = String.Format("SELECT * FROM {0} WHERE LineageText = @ParentLineage AND Name = @ParentCategory", _TreeTableName)
                                Command.Parameters.Add(CreateParameter(Command, "@ParentLineage", ParentLineage))
                                Command.Parameters.Add(CreateParameter(Command, "@ParentCategory", ParentCategory))

                                Using Reader As IDataReader = Command.ExecuteReader()
                                    If Reader.Read() Then
                                        Success = True

                                        '*** Insert the category ***
                                        Command.CommandText = String.Format("INSERT INTO {0} (Depth, Parent, Lineage, LineageText, Name) VALUES(" & CInt(Reader!Depth) + 1 & ", " & CStr(Reader!Node) & ", '" & CStr(Reader!Lineage) & CStr(Reader!Node) & "\', @Path, @Name);SELECT Scope_Identity();", _TreeTableName)
                                        Reader.Close()

                                        Command.Parameters.Add(CreateParameter(Command, "@Name", NodeName))
                                        Command.Parameters.Add(CreateParameter(Command, "@Path", Path))

                                        Return CInt(Command.ExecuteScalar())
                                    Else
                                        If Path = "\" Then
                                            Reader.Close()

                                            Command.CommandText = String.Format("INSERT INTO {0} (Depth, Parent, Lineage, LineageText, Name) VALUES(0, -1, @Path, @Path, @Name);SELECT Scope_Identity();", _TreeTableName)
                                            Command.Parameters.Add(CreateParameter(Command, "@Name", NodeName))
                                            Command.Parameters.Add(CreateParameter(Command, "@Path", Path))

                                            Return CInt(Command.ExecuteScalar())
                                        Else
                                            CreateNode(ParentLineage, ParentCategory)
                                        End If


                                    End If
                                End Using
                            End Using
                        End While

                    Catch ex As Exception
                        
                    End Try
                End Using

                return -1
            End Function

            '''<summary>
            '''Creates a Category at the specified path and returns its index. Returns -1 if the specified path does not exist.
            '''</summary>
            Public Function CreateNode(node As DBTreeNode) As Integer
                Return CreateNode(Node.LineageText, Node.Name, Node.IsProtected)
            End Function

            ''' <summary>
            ''' Updates the Name and IsProtected properties of the specified Node
            ''' </summary>
            ''' <param name="Node"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function UpdateNode(ByVal Node As DBTreeNode) As Boolean
                Using cnn As IDbConnection = CreateConnection()
                    Try
                        RenameNode(Node)

                        cnn.Open()
                        Using Command As IDbCommand = cnn.CreateCommand()
                            Command.CommandText = String.Format("UPDATE {0} SET IsProtected = @IsProtected WHERE Node = @Node", _TreeTableName)

                            Command.Parameters.Add(CreateParameter(Command, "@IsProtected", Node.IsProtected))
                            Command.Parameters.Add(CreateParameter(Command, "@Node", Node.ID))

                            Return Command.ExecuteNonQuery() > 0
                        End Using
                    Catch ex As Exception
                        Return False
                    Finally
                        If cnn.State = ConnectionState.Open Then cnn.Close()
                    End Try
                End Using

            End Function

            ''' <summary>
            ''' Renames a node in a tree.
            ''' </summary>
            ''' <param name="Node">The node containing the ID of the node to be updated and the new name.</param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function RenameNode(ByVal Node As DBTreeNode) As Boolean
                Return Me.RenameNode(Node.ID, Node.Name)
            End Function

            ''' <summary>
            ''' Renames a node in a tree.
            ''' </summary>
            ''' <param name="NodeID">The ID of the node to be updated.</param>
            ''' <param name="NewName">The new name of the node.</param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function RenameNode(ByVal NodeID As Integer, ByVal NewName As String) As Boolean
                If String.IsNullOrEmpty(NewName) Then Return False

                Dim LineageText As String
                Dim CurrentName As String

                Using cnn As IDbConnection = CreateConnection()

                    cnn.Open()

                    Using Transaction As IDbTransaction = cnn.BeginTransaction(IsolationLevel.ReadUncommitted)
                        Try
                            Using Command As IDbCommand = cnn.CreateCommand()
                                '********************************************************************************
                                '* Get the lineage and the current name of the node to be updated for correctly *
                                '* updating the lineage (path) of the sub-nodes.                                *
                                '********************************************************************************
                                Command.CommandText = String.Format("SELECT LineageText, Name FROM {0} WHERE Node = @Node", _TreeTableName)
                                Command.Parameters.Add(CreateParameter(Command, "@Node", NodeID))
                                Using Reader As IDataReader = Command.ExecuteReader()
                                    If Not Reader.Read() Then Return False

                                    CurrentName = Reader.GetString(1)
                                    LineageText = Reader.GetString(0)
                                End Using
                            End Using

                            Using Command As IDbCommand = cnn.CreateCommand()
                                '********************************************************************************
                                '* Determine if there is already a node with the NewName at the specified path. *
                                '* If so, do not rename the node.                                               *
                                '********************************************************************************
                                Command.CommandText = String.Format("SELECT * FROM {0} WHERE LineageText = @LineageText AND Name = @NewName AND Node <> @ID", _TreeTableName)
                                Command.Parameters.Add(CreateParameter(Command, "@LineageText", LineageText))
                                Command.Parameters.Add(CreateParameter(Command, "@NewName", NewName))
                                Command.Parameters.Add(CreateParameter(Command, "@ID", NodeID))

                                Using Reader As IDataReader = Command.ExecuteReader()
                                    If Reader.Read() Then Return False
                                End Using
                            End Using

                            Using Command As IDbCommand = cnn.CreateCommand()
                                '********************************************************************************
                                '* Update the sub-nodes of the node to be updated so that the path is correct.  *
                                '********************************************************************************
                                Command.CommandText = String.Format("UPDATE {0} SET LineageText = @NewLineage WHERE LineageText LIKE @OldLineage", _TreeTableName)
                                Command.Parameters.Add(CreateParameter(Command, "@OldLineage", LineageText & CurrentName & "\"))
                                Command.Parameters.Add(CreateParameter(Command, "@NewLineage", LineageText & NewName & "\"))

                                Command.ExecuteNonQuery()
                            End Using

                            Using Command As IDbCommand = cnn.CreateCommand()
                                '********************************************************************************
                                '* Update the name of the specified node.                                       *
                                '********************************************************************************
                                Command.CommandText = String.Format("UPDATE {0} SET Name = @Name WHERE Node = @Node", _TreeTableName)
                                Command.Parameters.Add(CreateParameter(Command, "@Name", NewName))
                                Command.Parameters.Add(CreateParameter(Command, "@Node", NodeID))

                                If Command.ExecuteNonQuery() > 0 Then
                                    Command.Transaction.Commit()
                                    Return True
                                End If
                            End Using
                        Catch ex As Exception

                        End Try
                    End Using
                End Using

                return false
            End Function

            Public Function RemoveNode(node As DBTreeNode) As Boolean
                Using cnn As IDbConnection = CreateConnection()
                    cnn.Open()

                    Using command As IDbCommand = cnn.CreateCommand()
                        Command.CommandText = String.Format("DELETE FROM {0} WHERE Lineage = @FullPath OR Node = @Node", _TreeTableName)

                        Command.Parameters.Add(CreateParameter(Command, "@FullPath", Node.FullPath))
                        Command.Parameters.Add(CreateParameter(Command, "@Node", Node.ID))

                        Return Command.ExecuteNonQuery > 0
                    End Using
                End Using
            End Function
        End Class

    End Namespace
End Namespace
