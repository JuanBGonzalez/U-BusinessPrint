Imports MySql.Data
Imports MySql.Data.Types
Imports MySql.Data.MySqlClient
Module Coneccion_Global
    Public cadena As String
    Public _coneccion As New MySqlConnection

    Public Function Conexion_Global() As Boolean
        Dim ESTADO As Boolean = True
        Try
            cadena = ("server=localhost;User Id=root;password= ;Database=clientescs")
            _coneccion = New MySqlConnection(cadena)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            ESTADO = False
        End Try
        Return ESTADO
    End Function

    Public Sub Cerrar()
        _coneccion.Close()
    End Sub
End Module
