Imports System.Net.Mail
Imports System.Data.SqlClient
Module mdRastreo
    Dim conexion = New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")
    Dim comando As SqlCommand
    Dim dt As New DataTable
    Public Sub rastreoEjecutar()
        conexion.open()

        frmMenu.rastreo_lblNombreCompleto.Text = frmMenu.positivo_txtNombre.Text + " " + frmMenu.positivo_txtApellido.Text
        frmMenu.rastreo_lblNombreCompleto2.Text = frmMenu.positivo_txtNombre.Text + " " + frmMenu.positivo_txtApellido.Text
        frmMenu.rastreo_cbCorregimiento.DisplayMember = "corregimiento"
        frmMenu.rastreo_cbCorregimiento.ValueMember = "id_equipo"

        cargarDataGridViewRastreo()
        conexion.close()

    End Sub

    Public Sub rastreoGuardar()
        Dim nombre As String = frmMenu.rastreo_txtNombre.Text
        Dim apellido As String = frmMenu.rastreo_txtApellido.Text
        Dim cedula As String = frmMenu.rastreo_txtCedula.Text
        Dim edad As Integer = frmMenu.rastreo_nudEdad.Value
        Dim genero As String = If(frmMenu.rastreo_radMasculino.Checked, "Masculino", "Femenino")
        Dim ubicacion As String = frmMenu.rastreo_txtUbicacion.Text
        Dim celular As String = frmMenu.rastreo_txtCelular.Text
        Dim correo As String = frmMenu.rastreo_txtCorreo.Text
        Dim estado As String = "pendiente"
        Dim id_equipo As String = frmMenu.rastreo_cbCorregimiento.SelectedValue.ToString
        Dim id_rastreo As Integer = frmMenu.positivo_txtIdCliente.Text

        If (nombre = "" Or apellido = "" Or cedula = "" Or edad = 0 Or ubicacion = "" Or celular = "" Or correo = "") Then
            MsgBox("Todos los campos son obligatorios", MessageBoxIcon.Warning)
        Else

            'Nueva conexion para ejecutar el procedimiento
            Dim conn As New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")

            'Objeto comando.
            Dim command As SqlCommand = conn.CreateCommand
            'TIpo de comando
            command.CommandType = CommandType.StoredProcedure
            'Nombre del procedimiento
            command.CommandText = "insertarPacienteRastreo"
            'Creando y cargando parametros
            Dim params(10) As SqlParameter
            params(0) = New SqlParameter("@nombre", SqlDbType.VarChar)
            params(0).Value = nombre

            params(1) = New SqlParameter("@apellido", SqlDbType.VarChar)
            params(1).Value = apellido

            params(2) = New SqlParameter("@cedula", SqlDbType.VarChar)
            params(2).Value = cedula

            params(3) = New SqlParameter("@edad", SqlDbType.Int)
            params(3).Value = edad

            params(4) = New SqlParameter("@genero", SqlDbType.VarChar)
            params(4).Value = genero

            params(5) = New SqlParameter("@ubicacion", SqlDbType.VarChar)
            params(5).Value = ubicacion

            params(6) = New SqlParameter("@celular", SqlDbType.VarChar)
            params(6).Value = celular

            params(7) = New SqlParameter("@correo", SqlDbType.VarChar)
            params(7).Value = correo

            params(8) = New SqlParameter("@estado", SqlDbType.VarChar)
            params(8).Value = estado

            params(9) = New SqlParameter("@id_equipo", SqlDbType.Int)
            params(9).Value = id_equipo

            params(10) = New SqlParameter("@id_rastreo", SqlDbType.Int)
            params(10).Value = id_rastreo

            'Añadiendo parametros al comando
            command.Parameters.AddRange(params)

            'Abriendo conexion y ejecutando el comando(procedimiento almacenado)
            conn.Open()
            Dim ejecucion As Integer = command.ExecuteNonQuery()
            'Vaciando campos rellenados
            enviarEmail("Usted ha sido agregado en la lista de rastreo de covid-19 debido a que tuvo contacto con: " + frmMenu.rastreo_lblNombreCompleto.Text, frmMenu.rastreo_txtCorreo.Text)
            frmMenu.rastreo_txtNombre.Text = ""
            frmMenu.rastreo_txtApellido.Text = ""
            frmMenu.rastreo_txtCedula.Text = ""
            frmMenu.rastreo_nudEdad.Value = 0
            frmMenu.rastreo_txtUbicacion.Text = ""
            frmMenu.rastreo_txtCelular.Text = ""
            frmMenu.rastreo_txtCorreo.Text = ""
            cargarDataGridViewRastreo()

            conn.Close()


        End If
    End Sub

    Private Sub cargarDataGridViewRastreo()
        Try
            'Cargar DataGridView
            Dim cadena As String = "select id_paciente,nombre, apellido,cedula,edad, genero, ubicacion,celular,correo,estado from pacientes where estado= 'pendiente' and id_rastreo =" + frmMenu.positivo_txtIdCliente.Text
            dt.Clear()
            comando = New SqlCommand(cadena, conexion)
            Using adapter As New SqlDataAdapter(comando)
                adapter.Fill(dt)
            End Using
            frmMenu.rastreo_dgvRastreo.DataSource = dt
        Catch ex As Exception
            MsgBox("Error: " + ex.Message)
        End Try
    End Sub

    Public Sub enviarEmail(mensaje As String, correo As String)

        'Enviar notificacion al email del paciente.
        Dim emailMessage As New MailMessage()
        Try
            emailMessage.From = New MailAddress("sistemaProyectoDes8@gmail.com")
            emailMessage.To.Add(correo)
            emailMessage.Subject = "Hisopado COVID - 19"
            emailMessage.Body = mensaje
            Dim SMTP As New SmtpClient("smtp.gmail.com")
            SMTP.Port = 587
            SMTP.EnableSsl = True
            SMTP.Credentials = New System.Net.NetworkCredential("sistemaProyectoDes8@gmail.com", "estoesunaprueba1")
            SMTP.Send(emailMessage)
        Catch ex As Exception

        End Try

    End Sub

    Public Sub actualizarPaciente(estado As String, id As Integer)

        'Nueva conexion para ejecutar el procedimiento
        Dim conn As New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")
        'Objeto comando.
        Dim command As SqlCommand = conn.CreateCommand
        'TIpo de comando
        command.CommandType = CommandType.StoredProcedure
        'Nombre del procedimiento
        command.CommandText = "actualizarPacienteRastreo"
        Dim params(1) As SqlParameter
        params(0) = New SqlParameter("@id_paciente", SqlDbType.Int)
        params(0).Value = id

        params(1) = New SqlParameter("@estado", SqlDbType.VarChar)
        params(1).Value = estado

        'Añadiendo parametros al comando
        command.Parameters.AddRange(params)

        'Abriendo conexion y ejecutando el comando(procedimiento almacenado)
        conn.Open()
        Dim ejecucion As Integer = command.ExecuteNonQuery()
        cargarDataGridViewRastreo()
        conn.Close()

    End Sub




End Module
