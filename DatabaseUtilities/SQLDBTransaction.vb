Imports System.Data.SqlClient

Namespace DatabaseUtilities

    Public Class SQLDBTransaction
        Implements IDbTransaction

#Region "Fields"

        Private _cnn As SqlConnection
        'Private _Commited As New List(Of Action)
        Private _isolationLevel As IsolationLevel = Data.IsolationLevel.Unspecified
        Private _Transaction As SqlTransaction

#End Region

#Region "Constructors"

        Public Sub New(cnn As SqlConnection, Transaction As SqlTransaction)
            _cnn = cnn
            _Transaction = Transaction
        End Sub

        Public Sub New(cnn As SqlConnection, Transaction As SqlTransaction, isolationLevel As IsolationLevel)
            Me.New(cnn, Transaction)
            _isolationLevel = isolationLevel
        End Sub

#End Region

#Region "Properties"

        Public ReadOnly Property cnn As SqlConnection
            Get
                Return _cnn
            End Get
        End Property

        Private ReadOnly Property Connection As IDbConnection Implements IDbTransaction.Connection
            Get
                Return cnn
            End Get
        End Property

        Public ReadOnly Property Transaction As SqlTransaction
            Get
                Return _Transaction
            End Get
        End Property

#End Region

#Region "Public Methods"

        'Public Sub AddCommited(method As Action)
        '    _Commited.Add(method)
        'End Sub

        Public Sub Commit() Implements IDbTransaction.Commit
            If Transaction IsNot Nothing Then
                Transaction.Commit()

                'For Each method As Action In _Commited
                '    Try
                '        method.Invoke()
                '    Catch ex As Exception

                '    End Try
                'Next
            End If

        End Sub

        Public Sub Rollback() Implements IDbTransaction.Rollback
            If Transaction IsNot Nothing Then Transaction.Rollback()
        End Sub

#End Region

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If cnn IsNot Nothing Then cnn.Dispose()
                    If Transaction IsNot Nothing Then Transaction.Dispose()
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

        Public ReadOnly Property IsolationLevel As IsolationLevel Implements IDbTransaction.IsolationLevel
            Get
                return IsolationLevel.ReadCommitted
            End Get
        End Property
    End Class

End Namespace