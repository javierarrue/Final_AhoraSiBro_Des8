Imports System.Data.SqlClient

Public Class ConexionBD
    Dim conexion As New SqlConnection("Data Source=RYL;Initial Catalog=finalDes5;Integrated Security=True")

    Sub enlace()
        Try
            conexion.Open()

        Catch ex As Exception
        End Try
    End Sub

    Sub cerrar()
        conexion.Close()
    End Sub
End Class
