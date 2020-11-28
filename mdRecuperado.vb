Imports System.Data.SqlClient
Module mdRecuperado
    Dim conexion = New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")
    Dim comando As SqlCommand
    Dim dt As New DataTable
    Dim dt2 As New DataTable
    Dim adapter As SqlDataAdapter
    Dim cadena As String

    Public Sub recuperadosEjecutar()
        conexion.open
        resumen_totalRecuperados()
        calcularTotalMujeres()
        calcularTotalHombes()
        conexion.close
    End Sub

    Public Sub resumen_totalRecuperados()
        Dim comando As SqlCommand
        Try
            Dim cadena2 = "pa_TotRecuperados"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.recuperados_totalRecuperados.Text = comando.Parameters("@total").Value

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    Private Sub calcularTotalHombes()
        Try
            Dim cadena2 = "pa_HombresRecuperados"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.recuperados_pHombres.Text = comando.Parameters("@total").Value
            frmMenu.recuperados_pHombres.Text += " (" + CStr(Math.Round((CInt(frmMenu.recuperados_pHombres.Text) / CInt(frmMenu.recuperados_totalRecuperados.Text)) * 100)) + "%)"

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    Private Sub calcularTotalMujeres()
        Try
            Dim cadena2 = "pa_MujeresRecuperadas"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.recuperados_pMujeres.Text = comando.Parameters("@total").Value
            frmMenu.recuperados_pMujeres.Text += " (" + CStr(Math.Round((CInt(frmMenu.recuperados_pMujeres.Text) / CInt(frmMenu.recuperados_totalRecuperados.Text)) * 100)) + "%)"

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

End Module
