Imports System.Data.SqlClient
Imports System.Net.Mail

Module mdEspera
    Dim conexion = New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")
    Dim comando As SqlCommand

    Public Sub esperaEjecutar()
        Try

            conexion.open()

            frmMenu.espera_lbTotal.Text = frmMenu.espera_dgvEspera.Rows.Count

            frmMenu.espera_lbHombres.Text = cantidadHombresTesteados()
            frmMenu.espera_lbHombres.Text += " (" + CStr(calcularPorcentaje(CInt(frmMenu.espera_lbHombres.Text), CInt(frmMenu.espera_lbTotal.Text))) + "%)"

            frmMenu.espera_lbMujeres.Text = cantidadMujeresTesteadas()
            frmMenu.espera_lbMujeres.Text += "(" + CStr(calcularPorcentaje(CInt(frmMenu.espera_lbMujeres.Text), CInt(frmMenu.espera_lbTotal.Text))) + "%)"

            conexion.close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Function cantidadHombresTesteados() As String
        Dim cantidad As String = ""

        'Nueva conexion para ejecutar el procedimiento
        Dim cmdObj As New SqlCommand("select*from cant_hombres_testeados", conexion)

        Try


            Dim readerObj As SqlDataReader = cmdObj.ExecuteReader

            While readerObj.Read
                cantidad = readerObj("cantidad").ToString
            End While
            readerObj.Close()


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return cantidad
    End Function

    Private Function cantidadMujeresTesteadas() As String
        Dim cantidad As String = ""

        'Nueva conexion para ejecutar el procedimiento
        Dim cmdObj As New SqlCommand("select*from cant_mujeres_testeadas", conexion)

        Try


            Dim readerObj As SqlDataReader = cmdObj.ExecuteReader

            While readerObj.Read
                cantidad = readerObj("cantidad").ToString
            End While
            readerObj.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        Return cantidad

    End Function

    Private Function calcularPorcentaje(numero As Integer, total As Integer) As Double

        Return Math.Round((numero / total) * 100)

    End Function


    Public Sub actualizarPositivo()

        If (frmMenu.espera_txtId.Text = "") Then
            MsgBox("Seleccione un paciente de la lista")
        Else
            enviarEmail("Usted es POSITIVO para Covid-19. En unos pocos días un equipo asignado estara visitando su casa para proveerle de insumos. ", frmMenu.espera_txtCorreo.Text)

            actualizarPacienteEspera("positivo", CInt(frmMenu.espera_txtId.Text))
            frmMenu.PacientesTableAdapter.Espera(frmMenu.FinalDes5DataSet.pacientes)
            esperaEjecutar()
            frmMenu.espera_txtId.Text = ""
        End If


    End Sub

    Public Sub actualizarNegativo()

        If (frmMenu.espera_txtId.Text = "") Then
            MsgBox("Seleccione un paciente de la lista")
        Else
            enviarEmail("Usted es NEGATIVO para Covid-19. Recuerde seguir cumpliendo con los consejos para combatir el virus Covid-19.", frmMenu.espera_txtCorreo.Text)

            actualizarPacienteEspera("negativo", CInt(frmMenu.espera_txtId.Text))
            frmMenu.PacientesTableAdapter.Espera(frmMenu.FinalDes5DataSet.pacientes)
            esperaEjecutar()
            frmMenu.espera_txtId.Text = ""
        End If

    End Sub

    Private Sub enviarEmail(mensaje As String, correo As String)

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

    Private Sub actualizarPacienteEspera(estado As String, id As Integer)

        'Nueva conexion para ejecutar el procedimiento
        Dim conn As New SqlConnection("server=(local) ; database=finalDes5 ; integrated security = true")
        'Objeto comando.
        Dim command As SqlCommand = conn.CreateCommand
        'TIpo de comando
        command.CommandType = CommandType.StoredProcedure
        'Nombre del procedimiento
        command.CommandText = "actualizarPaciente"
        Dim params(1) As SqlParameter
        params(0) = New SqlParameter("@id_paciente", SqlDbType.VarChar)
        params(0).Value = id

        params(1) = New SqlParameter("@estado", SqlDbType.VarChar)
        params(1).Value = estado

        'Añadiendo parametros al comando
        command.Parameters.AddRange(params)

        'Abriendo conexion y ejecutando el comando(procedimiento almacenado)
        conn.Open()
        Dim ejecucion As Integer = command.ExecuteNonQuery()
        conn.Close()

    End Sub

End Module
