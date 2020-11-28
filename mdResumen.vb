Imports System.Data.SqlClient
Module mdResumen
    Dim conexion = New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")

    Public Sub resumenEjecutar()
        conexion.open
        resumen_totalTests()
        resumen_totalPendientes()
        resumen_totalNegativos()
        resumen_totalPositivos()
        resumen_totalRecuperados()
        resumen_totalHombresPositivos()
        resumen_totaMujeresPositivas()
        conexion.close
    End Sub

    ' -- TOTAL DE TESTS REALIZADOS'
    Public Sub resumen_totalTests()
        Dim comando As SqlCommand
        Try
            Dim cadena2 = "pa_TotalTests"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.resumen_lbTotalTest.Text = comando.Parameters("@total").Value

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    ' -- TOTAL DE PACIENTES EN LISTA DE ESPERA'
    Public Sub resumen_totalPendientes()
        Dim comando As SqlCommand
        Try
            Dim cadena2 = "pa_PacientesPendientes"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.resumen_lbEspera.Text = comando.Parameters("@total").Value

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    Public Sub resumen_totalNegativos()
        Dim comando As SqlCommand
        Try
            Dim cadena2 = "pa_TotNegativo"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.resumen_lbTotalNegativos.Text = comando.Parameters("@total").Value

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    Public Sub resumen_totalPositivos()
        Dim comando As SqlCommand
        Try
            Dim cadena2 = "pa_TotPositivo"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.resumen_lbTotalPositivos.Text = comando.Parameters("@total").Value

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
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
            frmMenu.resumen_lbTotalRecuperados.Text = comando.Parameters("@total").Value

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    Public Sub resumen_totalHombresPositivos()
        Dim comando As SqlCommand
        Try
            Dim cadena2 = "pa_HombresPositivos"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando

            frmMenu.resumen_HombresPositivos.Text = comando.Parameters("@total").Value

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    Public Sub resumen_totaMujeresPositivas()
        Dim comando As SqlCommand
        Try
            Dim cadena2 = "pa_MujeresPositivas"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando

            frmMenu.resumen_MujeresPositivas.Text = comando.Parameters("@total").Value

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub


End Module
