Public Class User

    Private _DisplayName As String
    Public Property DisplayName() As String
        Get
            Return _DisplayName
        End Get
        Set(ByVal value As String)
            _DisplayName = value
        End Set
    End Property

    Private _Name As String
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property

    Private _Password As String
    Public Property Password() As String
        Get
            Return _Password
        End Get
        Set(ByVal value As String)
            _Password = value
        End Set
    End Property

    Private _Group As UserGroup
    Public Property Group() As UserGroup
        Get
            Return _Group
        End Get
        Set(ByVal value As UserGroup)
            _Group = value
        End Set
    End Property

    Sub New()

    End Sub

    Sub New(ByVal _Name As String)
        Name = _Name
    End Sub

    Sub New(ByVal _Group As UserGroup)
        Group = _Group
    End Sub

    Sub New(ByVal _Name As String, ByVal _Group As UserGroup)
        Name = _Name
        Group = _Group
    End Sub

End Class
