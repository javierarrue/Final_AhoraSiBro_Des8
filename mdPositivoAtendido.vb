Imports System.Data.SqlClient
Module mdPositivoAtendido
    Dim conexion = New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")
    Dim comando As SqlCommand

    Public Sub positivoAtendidoEjecutar()
        conexion.open()
        actualizarAtendido()
        conexion.close()
    End Sub

    Private Sub actualizarAtendido()
        Try

            actualizarPaciente(frmMenu.atendido_idPaciente.Text)

            frmMenu.PacientesTableAdapter.NoAtendidos(frmMenu.FinalDes5DataSet.pacientes)
            frmMenu.PacientesTableAdapter1.Atendidos(frmMenu.FinalDes5DataSet1.pacientes)


        Catch ex As Exception
            MsgBox("Debes seleccionar un paciente de la lista", MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub actualizarPaciente(id As Integer)

        'Nueva conexion para ejecutar el procedimiento

        Dim conn = New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")
        'Objeto comando.
        Dim command As SqlCommand = conn.CreateCommand
        'TIpo de comando
        command.CommandType = CommandType.StoredProcedure
        'Nombre del procedimiento
        command.CommandText = "actualizarPacienteAtendido"
        Dim params(0) As SqlParameter
        params(0) = New SqlParameter("@id_paciente", SqlDbType.VarChar)
        params(0).Value = id


        'Añadiendo parametros al comando
        command.Parameters.AddRange(params)

        'Abriendo conexion y ejecutando el comando(procedimiento almacenado)
        conn.Open()
        Dim ejecucion As Integer = command.ExecuteNonQuery()
        conn.Close()

    End Sub




End Module
