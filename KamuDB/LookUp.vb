
Friend Class LookupObject
    Private ReadOnly _ID As Long
    Private ReadOnly _Name As String

    Public Sub New(ByVal lngID As Long, ByVal strName As String)
        Me._ID = lngID
        Me._Name = strName
    End Sub

    Public ReadOnly Property ID() As String
        Get
            Return _ID
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property

End Class
