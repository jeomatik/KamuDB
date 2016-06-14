Public Class Authorization

    Private _ParselRead As Boolean
    Public Property ParselRead() As Boolean
        Get
            Return _ParselRead
        End Get
        Set(ByVal value As Boolean)
            _ParselRead = value
        End Set
    End Property

    Private _ParselWrite As Boolean
    Public Property ParselWrite() As Boolean
        Get
            Return _ParselWrite
        End Get
        Set(ByVal value As Boolean)
            _ParselWrite = value
        End Set
    End Property

    Private _KisiRead As Boolean
    Public Property KisiRead() As Boolean
        Get
            Return _KisiRead
        End Get
        Set(ByVal value As Boolean)
            _KisiRead = value
        End Set
    End Property

    Private _KisiWrite As Boolean
    Public Property KisiWrite() As Boolean
        Get
            Return _KisiWrite
        End Get
        Set(ByVal value As Boolean)
            _KisiWrite = value
        End Set
    End Property

    Private _DavaRead As Boolean
    Public Property DavaRead() As Boolean
        Get
            Return _DavaRead
        End Get
        Set(ByVal value As Boolean)
            _DavaRead = value
        End Set
    End Property

    Private _DavaWrite As Boolean
    Public Property DavaWrite() As Boolean
        Get
            Return _DavaWrite
        End Get
        Set(ByVal value As Boolean)
            _DavaWrite = value
        End Set
    End Property

    Private _MustemilatRead As Boolean
    Public Property MustemilatRead() As Boolean
        Get
            Return _MustemilatRead
        End Get
        Set(ByVal value As Boolean)
            _MustemilatRead = value
        End Set
    End Property

    Private _MustemilatWrite As Boolean
    Public Property MustemilatWrite() As Boolean
        Get
            Return _MustemilatWrite
        End Get
        Set(ByVal value As Boolean)
            _MustemilatWrite = value
        End Set
    End Property

    Private _MevsimlikRead As Boolean
    Public Property MevsimlikRead() As Boolean
        Get
            Return _MevsimlikRead
        End Get
        Set(ByVal value As Boolean)
            _MevsimlikRead = value
        End Set
    End Property

    Private _MevsimlikWrite As Boolean
    Public Property MevsimlikWrite() As Boolean
        Get
            Return _MevsimlikWrite
        End Get
        Set(ByVal value As Boolean)
            _MevsimlikWrite = value
        End Set
    End Property

    Private _ProjeRead As Boolean
    Public Property ProjeRead() As Boolean
        Get
            Return _ProjeRead
        End Get
        Set(ByVal value As Boolean)
            _ProjeRead = value
        End Set
    End Property

    Private _ProjeWrite As Boolean
    Public Property ProjeWrite() As Boolean
        Get
            Return _ProjeWrite
        End Get
        Set(ByVal value As Boolean)
            _ProjeWrite = value
        End Set
    End Property

    Private _KamuRead As Boolean
    Public Property KamuRead() As Boolean
        Get
            Return _KamuRead
        End Get
        Set(ByVal value As Boolean)
            _KamuRead = value
        End Set
    End Property

    Private _KamuWrite As Boolean
    Public Property KamuWrite() As Boolean
        Get
            Return _KamuWrite
        End Get
        Set(ByVal value As Boolean)
            _KamuWrite = value
        End Set
    End Property

    Private _OdemeRead As Boolean
    Public Property OdemeRead() As Boolean
        Get
            Return _OdemeRead
        End Get
        Set(ByVal value As Boolean)
            _OdemeRead = value
        End Set
    End Property

    Private _OdemeWrite As Boolean
    Public Property OdemeWrite() As Boolean
        Get
            Return _OdemeWrite
        End Get
        Set(ByVal value As Boolean)
            _OdemeWrite = value
        End Set
    End Property

    Private _MalikSurecRead As Boolean
    Public Property MalikSurecRead() As Boolean
        Get
            Return _MalikSurecRead
        End Get
        Set(ByVal value As Boolean)
            _MalikSurecRead = value
        End Set
    End Property

    Private _MalikSurecWrite As Boolean
    Public Property MalikSurecWrite() As Boolean
        Get
            Return _MalikSurecWrite
        End Get
        Set(ByVal value As Boolean)
            _MalikSurecWrite = value
        End Set
    End Property

    Private _ParselSurecRead As Boolean
    Public Property ParselSurecRead() As Boolean
        Get
            Return _ParselSurecRead
        End Get
        Set(ByVal value As Boolean)
            _ParselSurecRead = value
        End Set
    End Property

    Private _ParselSurecWrite As Boolean
    Public Property ParselSurecWrite() As Boolean
        Get
            Return _ParselSurecWrite
        End Get
        Set(ByVal value As Boolean)
            _ParselSurecWrite = value
        End Set
    End Property

    Private _CanImport As Boolean
    Public Property CanImport() As Boolean
        Get
            Return _CanImport
        End Get
        Set(ByVal value As Boolean)
            _CanImport = value
        End Set
    End Property

    Private _CanExport As Boolean
    Public Property CanExport() As Boolean
        Get
            Return _CanExport
        End Get
        Set(ByVal value As Boolean)
            _CanExport = value
        End Set
    End Property

    Private _BasitAnaliz As Boolean
    Public Property BasitAnaliz() As Boolean
        Get
            Return _BasitAnaliz
        End Get
        Set(ByVal value As Boolean)
            _BasitAnaliz = value
        End Set
    End Property

    Private _GelismisAnaliz As Boolean
    Public Property GelismisAnaliz() As Boolean
        Get
            Return _GelismisAnaliz
        End Get
        Set(ByVal value As Boolean)
            _GelismisAnaliz = value
        End Set
    End Property

    Private _OdemeEmri As Boolean
    Public Property OdemeEmri() As Boolean
        Get
            Return _OdemeEmri
        End Get
        Set(ByVal value As Boolean)
            _OdemeEmri = value
        End Set
    End Property

    Private _BolgeID As Long
    Public Property BolgeID() As Long
        Get
            Return _BolgeID
        End Get
        Set(ByVal value As Long)
            _BolgeID = value
        End Set
    End Property

    Private _TakpasSorgu As Boolean
    Public Property TakpasSorgu() As Boolean
        Get
            Return _TakpasSorgu
        End Get
        Set(ByVal value As Boolean)
            _TakpasSorgu = value
        End Set
    End Property

    Private _LogView As Boolean
    Public Property LogView() As Boolean
        Get
            Return _LogView
        End Get
        Set(ByVal value As Boolean)
            _LogView = value
        End Set
    End Property

    Private _ManageUsers As Boolean
    Public Property ManageUsers() As Boolean
        Get
            Return _ManageUsers
        End Get
        Set(ByVal value As Boolean)
            _ManageUsers = value
        End Set
    End Property
End Class
