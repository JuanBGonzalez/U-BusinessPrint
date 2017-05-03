Public Class DatosRegistro
    Private _Senorg As String
    Private _Fecha As String
    Private _Cedula As String
    Private _Nombre As String
    Private _FechaInicioDesc As String
    Private _DeudaDesc As String
    Private _MontoDesc As String
    Private _Periodo As String
    Private _Porcentaje As String
    Private _LeEntregueDesc As String
    Private _DeudaTotal As String
    Private _MontoTotal As String

    Public Property Senorg() As String
        Get
            Return _Senorg
        End Get
        Set(ByVal value As String)
            _Senorg = value
        End Set
    End Property

    Public Property Fecha() As String
        Get
            Return _Fecha
        End Get
        Set(ByVal value As String)
            _Fecha = value
        End Set
    End Property

    Public Property Cedula() As String
        Get
            Return _Cedula
        End Get
        Set(ByVal value As String)
            _Cedula = value
        End Set
    End Property

    Public Property Nombre() As String
        Get
            Return _Nombre
        End Get
        Set(ByVal value As String)
            _Nombre = value
        End Set
    End Property

    Public Property FechaInicioDesc() As String
        Get
            Return _FechaInicioDesc
        End Get
        Set(ByVal value As String)
            _FechaInicioDesc = value
        End Set
    End Property

    Public Property DeudaDesc() As String
        Get
            Return _DeudaDesc
        End Get
        Set(ByVal value As String)
            _DeudaDesc = value
        End Set
    End Property

    Public Property MontoDesc() As String
        Get
            Return _MontoDesc
        End Get
        Set(ByVal value As String)
            _MontoDesc = value
        End Set
    End Property

    Public Property Periodo() As String
        Get
            Return _Periodo
        End Get
        Set(ByVal value As String)
            _Periodo = value
        End Set
    End Property

    Public Property Porcentaje() As String
        Get
            Return _Porcentaje
        End Get
        Set(ByVal value As String)
            _Porcentaje = value
        End Set
    End Property

    Public Property LeEntregueDesc() As String
        Get
            Return _LeEntregueDesc
        End Get
        Set(ByVal value As String)
            _LeEntregueDesc = value
        End Set
    End Property

    Public Property DeudaTotal() As String
        Get
            Return _DeudaTotal
        End Get
        Set(ByVal value As String)
            _DeudaTotal = value
        End Set
    End Property

    Public Property MontoTotal() As String
        Get
            Return _MontoTotal
        End Get
        Set(ByVal value As String)
            _MontoTotal = value
        End Set
    End Property

End Class
