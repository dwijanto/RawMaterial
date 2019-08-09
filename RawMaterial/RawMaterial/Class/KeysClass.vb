Public Class KeysClass
    Implements IDisposable

    Private mKeyid As String
    Private mUpdateby As String
    Private mSource As String
    Private mLinks As String
    Private mArea As String
    Private mCurrency As String
    Private mUnit As String

    Public Sub New(ByVal Keyid As String, ByVal Updateby As String, ByVal Source As String, _
                   ByVal Links As String, ByVal Area As String, ByVal Currency As String, ByVal Unit As String)
        mKeyid = Keyid
        mUpdateby = Updateby
        mSource = Source
        mLinks = Links
        mArea = Area
        mCurrency = Currency
        mUnit = Unit
    End Sub

    Public ReadOnly Property Currency() As String
        Get
            Return mCurrency
        End Get
    End Property
    Public ReadOnly Property Unit() As String
        Get
            Return mUnit
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return mKeyid
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
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
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
