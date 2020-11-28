Imports System.Data.SqlClient
Imports System.Net.Mail
Module mdRegistrar
    Public Sub registrarEjecutar()
        frmMenu.registra_cbCorregimiento.DisplayMember = "corregimiento"
        frmMenu.registra_cbCorregimiento.ValueMember = "id_equipo"
    End Sub

    Public Sub guardarPaciente()

        Dim nombre As String = frmMenu.registra_txtNombre.Text
        Dim apellido As String = frmMenu.registra_txtApellido.Text
        Dim cedula As String = frmMenu.registra_txtCedula.Text
        Dim edad As Integer = frmMenu.registra_nudEdad.Value
        Dim genero As String = If(frmMenu.registra_radMasculino.Checked, "Masculino", "Femenino")
        Dim ubicacion As String = frmMenu.registra_txtUbicacion.Text
        Dim celular As String = frmMenu.registra_txtCelular.Text
        Dim correo As String = frmMenu.registra_txtCorreo.Text
        Dim estado As String = "pendiente"
        Dim id_equipo As String = frmMenu.registra_cbCorregimiento.SelectedValue.ToString

        If (nombre = "" Or apellido = "" Or cedula = "" Or edad = 0 Or ubicacion = "" Or celular = "" Or correo = "") Then
            MsgBox("Todos los campos son obligatorios", MessageBoxIcon.Warning)
        Else
            Dim conn = New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")

            'Objeto comando.
            Dim command As SqlCommand = conn.CreateCommand
            'TIpo de comando
            command.CommandType = CommandType.StoredProcedure
            'Nombre del procedimiento
            command.CommandText = "insertarPaciente"
            'Creando y cargando parametros
            Dim params(9) As SqlParameter
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

            'Añadiendo parametros al comando
            command.Parameters.AddRange(params)

            'Abriendo conexion y ejecutando el comando(procedimiento almacenado)
            conn.Open()
            Dim ejecucion As Integer = command.ExecuteNonQuery()
            conn.Close()

            If (ejecucion > 0) Then
                enviarCorreoRegistro(correo)
                MsgBox("Cliente registrado. Email de notificación enviado.", MessageBoxIcon.Information)

                'Vaciando campos rellenados
                frmMenu.registra_txtNombre.Text = ""
                frmMenu.registra_txtApellido.Text = ""
                frmMenu.registra_txtCedula.Text = ""
                frmMenu.registra_nudEdad.Value = 0
                frmMenu.registra_txtUbicacion.Text = ""
                frmMenu.registra_txtCelular.Text = ""
                frmMenu.registra_txtCorreo.Text = ""
            End If

        End If

    End Sub

    Private Sub enviarCorreoRegistro(correo As String)
        ' Enviar notificacion al email del paciente.
        Dim emailMessage As New MailMessage()
        Try
            emailMessage.From = New MailAddress("sistemaProyectoDes8@gmail.com")
            emailMessage.To.Add(correo)
            emailMessage.Subject = "Hisopado COVID - 19"
            emailMessage.Body = "Usted ha sido testeado para la prueba del virus covid-19. 
                                            Se le enviara otro correo con los resultados una vez haya sido procesado.
                                            No responder a este mensaje."
            Dim SMTP As New SmtpClient("smtp.gmail.com")
            SMTP.Port = 587
            SMTP.EnableSsl = True
            SMTP.Credentials = New System.Net.NetworkCredential("sistemaProyectoDes8@gmail.com", "estoesunaprueba1")
            SMTP.Send(emailMessage)
        Catch ex As Exception

        End Try
    End Sub

End Module
