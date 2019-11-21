Public Enum Connections
    OleDbConnection
    SqlConnection
    PgSqlConnection
End Enum

Public Class ConnectionInfo
    Private _ConnectionType As Connections
    Private _ConnectionString As String
    Private _AktifDosya As String
    Private _Database As String
    Private _Server As String
    Private _Port As String
    Private _User As String
    Private _Password As String

    Public Property User() As String
        Get
            Return _User
        End Get
        Set(ByVal value As String)
            _User = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return _Password
        End Get
        Set(ByVal value As String)
            _Password = value
        End Set
    End Property

    Public Property Port() As String
        Get
            Return _Port
        End Get
        Set(ByVal value As String)
            _Port = value
        End Set
    End Property

    Public Property DataBase() As String
        Get
            Return _Database
        End Get
        Set(ByVal value As String)
            _Database = value
        End Set
    End Property

    Public Property Server() As String
        Get
            Return _Server
        End Get
        Set(ByVal value As String)
            _Server = value
        End Set
    End Property

    Public Property ConnectionType() As Connections
        Get
            Return _ConnectionType
        End Get
        Set(ByVal value As Connections)
            _ConnectionType = value
        End Set
    End Property

    Public Property ConnectionString() As String
        Get
            Return _ConnectionString
        End Get
        Set(ByVal value As String)
            _ConnectionString = value
        End Set
    End Property

    Public Property AktifDosya() As String
        Get
            Return _AktifDosya
        End Get
        Set(ByVal value As String)
            _AktifDosya = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal _Host As String, ByVal _Database As String, ByVal _Port As String, ByVal _UserName As String, ByVal _Password As String)
        Me.DataBase = _Database
        Me.Server = _Host
        Me.Port = _Port
        Me.User = _UserName
        Me.Password = _Password
        Me.ConnectionString = "Server=" + _Host + ";Port=" + _Port + ";Database=" + _Database + ";Userid=" + _UserName + ";password=" + _Password ';Timeout=15;Pooling=true;MinPoolSize=1;MaxPoolSize=20;Encoding=UNICODE;SslMode=Disable"
        Me.ConnectionType = Connections.PgSqlConnection
        Me.AktifDosya = String.Empty
    End Sub

    Public Sub New(ByVal _DataSource As String, ByVal _InitialCatalog As String, ByVal _UserName As String, ByVal _Password As String)
        Me.DataBase = _InitialCatalog
        Me.Server = _DataSource
        Me.User = _UserName
        Me.Password = _Password
        Me.ConnectionString = "Server=" + _DataSource + ";Database=" + _InitialCatalog + ";User Id=" + _UserName + ";Password=" + _Password
        Me.ConnectionType = Connections.SqlConnection
        Me.AktifDosya = String.Empty
    End Sub

    Public Sub New(ByVal _DataSource As String, ByVal _InitialCatalog As String)
        Me.DataBase = _InitialCatalog
        Me.Server = _DataSource
        Me.ConnectionString = "Data Source=" + _DataSource + ";Initial Catalog=" + _InitialCatalog + ";Integrated Security=True"
        Me.ConnectionType = Connections.SqlConnection
        Me.AktifDosya = String.Empty
    End Sub

    Public Sub New(ByVal _FileName As String)
        Me.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _FileName & ";User Id=admin;Password=;"
        Me.ConnectionType = Connections.OleDbConnection
        Me.AktifDosya = _FileName
    End Sub

    Public Sub New(ByVal _FileName As String, ByVal IsACCDB As Boolean)
        Me.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _FileName & ";Persist Security Info=False;"
        Me.ConnectionType = Connections.OleDbConnection
        Me.AktifDosya = _FileName
    End Sub

End Class