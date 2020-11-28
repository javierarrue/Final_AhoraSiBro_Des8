Imports System.Data.SqlClient
Module mdNegativos
    Dim conexion = New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")
    Dim comando As SqlCommand
    Dim dt As New DataTable
    Dim dt2 As New DataTable
    Dim adapter As SqlDataAdapter
    Dim cadena As String

    Public Sub negativosEjecutar()
        conexion.open()
        calcularTotalNegativos()
        calcularTotalHombes()
        calcularTotalMujeres()
        conexion.close()
    End Sub


    Private Sub calcularTotalNegativos()
        Try
            Dim cadena2 = "pa_TotNegativo"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.negativo_lbTotal.Text = comando.Parameters("@total").Value

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    Private Sub calcularTotalHombes()
        Try
            Dim cadena2 = "pa_HombresNegativos"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.negativo_lbHombres.Text = comando.Parameters("@total").Value
            frmMenu.negativo_lbHombres.Text += " (" + CStr(Math.Round((CInt(frmMenu.negativo_lbHombres.Text) / CInt(frmMenu.negativo_lbTotal.Text)) * 100)) + ")%"

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    Private Sub calcularTotalMujeres()
        Try
            Dim cadena2 = "pa_MujeresNegativas"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.negativo_lbMujeres.Text = comando.Parameters("@total").Value
            frmMenu.negativo_lbMujeres.Text += " (" + CStr(Math.Round((CInt(frmMenu.negativo_lbMujeres.Text) / CInt(frmMenu.negativo_lbTotal.Text)) * 100)) + ")%"

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

End Module
