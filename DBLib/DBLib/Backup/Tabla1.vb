Public Class Tabla1
    Private _campo1 As String
    Public Property Campo1() As String
        Get
            Return _campo1
        End Get

        Set(ByVal Value As String)
            _campo1 = Value
        End Set
    End Property

    Private _campo2 As String
    Public Property Campo2() As String
        Get
            Return _campo2
        End Get

        Set(ByVal Value As String)
            _campo2 = Value
        End Set
    End Property

    Private _campo3 As Integer
    Public Property Campo3() As Integer
        Get
            Return _campo3
        End Get

        Set(ByVal Value As Integer)
            _campo3 = Value
        End Set
    End Property
End Class

Public Class Sucursales
    Private _IDsucursal As Short
    Public Property IDsucursal() As Short
        Get
            Return _IDsucursal
        End Get

        Set(ByVal Value As Short)
            _IDsucursal = Value
        End Set
    End Property

    Private _nombre As String
    Public Property nombre() As String
        Get
            Return _nombre
        End Get

        Set(ByVal Value As String)
            _nombre = Value
        End Set
    End Property

    Private _direccion As String
    Public Property direccion() As String
        Get
            Return _direccion
        End Get

        Set(ByVal Value As String)
            _direccion = Value
        End Set
    End Property

    Private _provincia As String
    Public Property provincia() As String
        Get
            Return _provincia
        End Get

        Set(ByVal Value As String)
            _provincia = Value
        End Set
    End Property

    Private _localidad As String
    Public Property localidad() As String
        Get
            Return _localidad
        End Get

        Set(ByVal Value As String)
            _localidad = Value
        End Set
    End Property

    Private _cp As String
    Public Property cp() As String
        Get
            Return _cp
        End Get

        Set(ByVal Value As String)
            _cp = Value
        End Set
    End Property

    Private _cuit As String
    Public Property cuit() As String
        Get
            Return _cuit
        End Get

        Set(ByVal Value As String)
            _cuit = Value
        End Set
    End Property

    Private _telefono As String
    Public Property telefono() As String
        Get
            Return _telefono
        End Get

        Set(ByVal Value As String)
            _telefono = Value
        End Set
    End Property

    Private _fax As String
    Public Property fax() As String
        Get
            Return _fax
        End Get

        Set(ByVal Value As String)
            _fax = Value
        End Set
    End Property

    Private _posnet As String
    Public Property posnet() As String
        Get
            Return _posnet
        End Get

        Set(ByVal Value As String)
            _posnet = Value
        End Set
    End Property

    Private _correo As String
    Public Property correo() As String
        Get
            Return _correo
        End Get

        Set(ByVal Value As String)
            _posnet = Value
        End Set
    End Property
End Class