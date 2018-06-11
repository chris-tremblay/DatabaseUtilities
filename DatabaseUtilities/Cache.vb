Imports System.Collections.Generic

Namespace DatabaseUtilities
    Public Class Cache
        Private _Cache As New Dictionary(Of Object, CacheEntry)
        Private Shared WithEvents _Instance As Cache = Nothing

        Private Sub New()

        End Sub

        Public Shared Function GetInstance() As Cache
            If _Instance Is Nothing Then
                _Instance = New Cache()
                RaiseEvent Initialized(_Instance, New System.EventArgs)
            End If

            Return _Instance
        End Function

        <Serializable()> _
        Private Class CacheEntry
            Private _Value As Object
            Private _ExpirtationDate As Date

            Public Sub New(ByVal Value As Object)
                _Value = Value
                _ExpirtationDate = Now.AddYears(10)
            End Sub

            Public Sub New(ByVal Value As Object, ByVal ExpirationDate As Date)
                _Value = Value
                _ExpirtationDate = ExpirationDate
            End Sub

            Public Property Value() As Object
                Get
                    Return _Value
                End Get
                Set(ByVal value As Object)
                    _Value = value
                End Set
            End Property

            Public Property ExpirationDate() As Date
                Get
                    Return _ExpirtationDate
                End Get
                Set(ByVal value As Date)
                    _ExpirtationDate = value
                End Set
            End Property
        End Class

        Public Function GetObject(ByVal Key As Object) As Object
            RaiseEvent BeforeGetObject(Me, New System.EventArgs())

            If _Cache.ContainsKey(Key) Then
                Dim Entry As CacheEntry = _Cache(Key)

                If Now < Entry.ExpirationDate Then
                    Return Entry.Value
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If

            RaiseEvent AfterGetObject(Me, New System.EventArgs())
        End Function

        Private Sub AddObject(ByVal Key As Object, ByVal Entry As CacheEntry)
            RaiseEvent BeforeAddObject(Me, New System.EventArgs())

            SyncLock _Cache
                If _Cache.ContainsKey(Key) Then
                    _Cache(Key) = Entry
                Else
                    _Cache.Add(Key, Entry)
                End If
            End SyncLock

            RaiseEvent AfterAddObject(Me, New System.EventArgs())
        End Sub

        Public Sub RemoveObject(ByVal Key As Object)
            SyncLock _Cache
                If _Cache.ContainsKey(Key) Then
                    _Cache(Key) = Nothing
                End If
            End SyncLock

            RaiseEvent AfterAddObject(Me, New System.EventArgs())
        End Sub

        Public Sub AddObject(ByVal Key As Object, ByVal Data As Object)
            AddObject(Key, New CacheEntry(Data))
        End Sub

        Public Sub AddObject(ByVal Key As Object, ByVal Data As Object, ByVal ExpirationDate As Date)
            AddObject(Key, New CacheEntry(Data, ExpirationDate))
        End Sub

        ''' <summary>
        ''' Loads the cache from a stream containing the serialized cache.
        ''' </summary>
        ''' <param name="Stream"></param>
        ''' <remarks></remarks>
        Public Sub LoadCache(ByVal Stream As IO.Stream)
            Try
                Dim bin As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()

                _Cache = CType(bin.Deserialize(Stream), Dictionary(Of Object, CacheEntry))
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' Serializes the cache to the specified stream.
        ''' </summary>
        ''' <param name="Stream"></param>
        ''' <remarks></remarks>
        Public Sub SaveCache(ByVal Stream As IO.Stream)
            Dim bin As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()

            bin.Serialize(Stream, _Cache)
        End Sub

#Region "Events"

        Public Event BeforeGetObject As EventHandler

        Public Event AfterGetObject As EventHandler

        Public Event AfterAddObject As EventHandler

        Public Event BeforeAddObject As EventHandler

        Public Shared Event Initialized As EventHandler

#End Region



    End Class

End Namespace