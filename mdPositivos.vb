Imports System.Data.SqlClient
Module mdPositivos

    Dim conexion = New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")
    Dim comando As SqlCommand

    Public Sub positivosEjecutar()

        conexion.open()
        frmMenu.positivos_lbTotal.Text = frmMenu.dgvPacientesPositivos.Rows.Count
        calcularTotalHombresPositivos()
        calcularTotalMujeresPositivos()
        conexion.close()

    End Sub

    Public Sub positivosActualizar()
        conexion.open()
        Try
            actualizarPaciente(CInt(frmMenu.positivo_txtIdCliente.Text))
            frmMenu.PacientesTableAdapter.Positivos(frmMenu.FinalDes5DataSet.pacientes)
            frmMenu.positivos_lbTotal.Text = frmMenu.dgvPacientesPositivos.Rows.Count

            calcularTotalHombresPositivos()
            calcularTotalMujeresPositivos()

        Catch ex As Exception
            MsgBox("Debe seleccionar un paciente de la lista", MessageBoxIcon.Error)
        End Try
        conexion.close()
    End Sub

    Private Sub actualizarPaciente(id As Integer)

        'Nueva conexion para ejecutar el procedimiento

        'Objeto comando.
        Dim command As SqlCommand = conexion.CreateCommand
        'TIpo de comando
        command.CommandType = CommandType.StoredProcedure
        'Nombre del procedimiento
        command.CommandText = "actualizarPacienteRecuperado"
        Dim params(0) As SqlParameter
        params(0) = New SqlParameter("@id_paciente", SqlDbType.VarChar)
        params(0).Value = id

        'Añadiendo parametros al comando
        command.Parameters.AddRange(params)

        'Abriendo conexion y ejecutando el comando(procedimiento almacenado)

        Dim ejecucion As Integer = command.ExecuteNonQuery()


    End Sub


    Private Sub calcularTotalHombresPositivos()
        Try
            Dim cadena2 = "totalHombresPositivos"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.positivo_lbHombres.Text = comando.Parameters("@total").Value
            frmMenu.positivo_lbHombres.Text += " (" + CStr(Math.Round((CInt(frmMenu.positivo_lbHombres.Text) / CInt(frmMenu.positivos_lbTotal.Text)) * 100)) + "%)"

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try

    End Sub

    Private Sub calcularTotalMujeresPositivos()
        Try
            Dim cadena2 = "totalMujeresPositivos"
            comando = New SqlCommand(cadena2, conexion) 'LLama al PA
            comando.CommandType = CommandType.StoredProcedure 'Ejecuta un PA
            'Parametro de salida para obtener los datos
            comando.Parameters.Add("@total", SqlDbType.Int)
            comando.Parameters(0).Direction = ParameterDirection.Output
            comando.ExecuteNonQuery() 'Ejecuta el comando
            frmMenu.positivo_lbMujeres.Text = comando.Parameters("@total").Value
            frmMenu.positivo_lbMujeres.Text += " (" + CStr(Math.Round((CInt(frmMenu.positivo_lbMujeres.Text) / CInt(frmMenu.positivos_lbTotal.Text)) * 100)) + "%)"

        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub


End Module
