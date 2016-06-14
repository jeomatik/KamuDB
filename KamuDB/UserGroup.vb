Public Class UserGroup

    Private _Name As String
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property

    Private _Authorization As Authorization
    Public Property Authorization() As Authorization
        Get
            Return _Authorization
        End Get
        Set(ByVal value As Authorization)
            _Authorization = value
        End Set
    End Property

    Sub New()

    End Sub

    Sub New(ByVal _Name As String)
        Name = _Name
    End Sub

    Sub New(ByVal _Authorization As Authorization)
        Authorization = _Authorization
    End Sub

    Sub New(ByVal _Name As String, ByVal _Authorization As Authorization)
        Name = _Name
        Authorization = _Authorization
    End Sub

End Class
