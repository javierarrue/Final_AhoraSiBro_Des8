Imports System.Data.SqlClient

Public Class frmMenu

    Private Sub frmMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'FinalDes5DataSet.equipos' Puede moverla o quitarla según sea necesario.
        Me.EquiposTableAdapter.Fill(Me.FinalDes5DataSet.equipos)

    End Sub

    '-----Panel Menú Principal-----'
    Private Sub btnRegistrar_Click(sender As Object, e As EventArgs) Handles btnRegistrar.Click
        pnlRegistrar.Visible = True
        pnlEspera.Visible = False
        registrarEjecutar()
    End Sub

    Private Sub btnListaEspera_Click(sender As Object, e As EventArgs) Handles btnListaEspera.Click
        pnlRegistrar.Visible = True
        pnlEspera.Visible = True
        pnlPositivo.Visible = False
        Me.PacientesTableAdapter.Espera(Me.FinalDes5DataSet.pacientes)
        esperaEjecutar()
    End Sub

    Private Sub btnListaPositivos_Click(sender As Object, e As EventArgs) Handles btnListaPositivos.Click
        pnlRegistrar.Visible = True
        pnlEspera.Visible = True
        pnlPositivo.Visible = True
        pnlPositivoRastreo.Visible = False
        Me.PacientesTableAdapter.Positivos(Me.FinalDes5DataSet.pacientes)
        positivosEjecutar()
    End Sub

    Private Sub btnListaNegativos_Click(sender As Object, e As EventArgs) Handles btnListaNegativos.Click
        pnlRegistrar.Visible = True
        pnlEspera.Visible = True
        pnlPositivo.Visible = True
        pnlPositivoRastreo.Visible = True
        pnlPositivoAtendido.Visible = True
        pnlNegativos.Visible = True
        pnlRecuperados.Visible = False
        Me.PacientesTableAdapter.Negativos(Me.FinalDes5DataSet.pacientes)
        negativosEjecutar()
    End Sub

    Private Sub btnListaRecuperados_Click(sender As Object, e As EventArgs) Handles btnListaRecuperados.Click
        pnlRegistrar.Visible = True
        pnlEspera.Visible = True
        pnlPositivo.Visible = True
        pnlPositivoRastreo.Visible = True
        pnlPositivoAtendido.Visible = True
        pnlNegativos.Visible = True
        pnlRecuperados.Visible = True
        pnlResumen.Visible = False
        Me.PacientesTableAdapter.Recuperados(Me.FinalDes5DataSet.pacientes)
        recuperadosEjecutar()

    End Sub

    Private Sub btnResumen_Click(sender As Object, e As EventArgs) Handles btnResumen.Click
        pnlRegistrar.Visible = True
        pnlEspera.Visible = True
        pnlPositivo.Visible = True
        pnlPositivoRastreo.Visible = True
        pnlPositivoAtendido.Visible = True
        pnlNegativos.Visible = True
        pnlRecuperados.Visible = True
        pnlResumen.Visible = True
        pnlSobreNosotros.Visible = False
        mdResumen.resumenEjecutar()

    End Sub

    '-----Botones Superiores-----'
    Private Sub btnInicio_Click(sender As Object, e As EventArgs) Handles btnInicio.Click
        pnlRegistrar.Visible = True
        pnlEspera.Visible = True
        pnlPositivo.Visible = True
        pnlPositivoRastreo.Visible = True
        pnlPositivoAtendido.Visible = True
        pnlNegativos.Visible = True
        pnlRecuperados.Visible = True
        pnlResumen.Visible = True
        pnlSobreNosotros.Visible = True
        pnlInicio.Visible = True
    End Sub

    Private Sub btnVisualizar_Click(sender As Object, e As EventArgs) Handles btnVisualizar.Click
        pnlPrincipal.Visible = True
        pnlRegistrar.Visible = False
    End Sub

    Private Sub btnSobreNosotros_Click(sender As Object, e As EventArgs) Handles btnSobreNosotros.Click
        pnlRegistrar.Visible = True
        pnlEspera.Visible = True
        pnlPositivo.Visible = True
        pnlPositivoRastreo.Visible = True
        pnlPositivoAtendido.Visible = True
        pnlNegativos.Visible = True
        pnlRecuperados.Visible = True
        pnlResumen.Visible = True
        pnlSobreNosotros.Visible = True
        pnlInicio.Visible = False
    End Sub

    '-----Panel Registrar Paciente-----'
    Private Sub Registrar_btnRetroceder_Click(sender As Object, e As EventArgs) Handles Registrar_btnRetroceder.Click
        pnlPrincipal.Visible = True
        pnlRegistrar.Visible = False

    End Sub

    Private Sub Registrar_btnGuardar_Click(sender As Object, e As EventArgs) Handles Registrar_btnGuardar.Click
        mdRegistrar.guardarPaciente()
    End Sub


    '-----Panel En Espera-----'
    Private Sub Espera_btnRetroceder_Click(sender As Object, e As EventArgs) Handles Espera_btnRetroceder.Click
        pnlPrincipal.Visible = True
        pnlRegistrar.Visible = False
    End Sub

    Private Sub Espera_btnPositivo_Click(sender As Object, e As EventArgs) Handles Espera_btnPositivo.Click

        Dim opcion = MsgBox("Desea actualizar el estado del paciente a positivo", vbYesNo + vbQuestion, "Confirmación")
        If (opcion = DialogResult.Yes) Then
            mdEspera.actualizarPositivo()
        End If
    End Sub

    Private Sub Espera_btnNegativo_Click(sender As Object, e As EventArgs) Handles Espera_btnNegativo.Click

        Dim opcion = MsgBox("Desea actualizar el estado del paciente a negativo", vbYesNo + vbQuestion, "Confirmación")
        If (opcion = DialogResult.Yes) Then
            mdEspera.actualizarNegativo()
        End If

    End Sub

    Private Sub dgvPendienteClick(sender As Object, e As DataGridViewCellEventArgs) Handles espera_dgvEspera.CellClick

        espera_txtCorreo.Text = espera_dgvEspera.CurrentRow.Cells(8).Value.ToString()
        espera_txtId.Text = espera_dgvEspera.CurrentRow.Cells(0).Value.ToString()
    End Sub

    '-----Panel Pacientes Positivos----'
    Private Sub Positivos_btnRetroceder_Click(sender As Object, e As EventArgs) Handles Positivos_btnRetroceder.Click
        pnlPrincipal.Visible = True
        pnlRegistrar.Visible = False
    End Sub

    Private Sub btnIniciarRastreo_Click(sender As Object, e As EventArgs) Handles btnIniciarRastreo.Click
        pnlPositivoRastreo.Visible = True
        pnlPositivoAtendido.Visible = False
        mdRastreo.rastreoEjecutar()
    End Sub

    Private Sub rastreo_btnPositivo_Click(sender As Object, e As EventArgs) Handles rastreo_btnPositivo.Click

        If rastreo_txtIdRastreo.Text = "" Or rastreo_txtIdRastreo.Text = "" Then
            MsgBox("Escoja al paciente que desea actualizar", MessageBoxIcon.Warning)
        Else
            Dim opcion = MsgBox("Desea actualizar el estado del paciente a Pasivo", vbYesNo + vbQuestion, "Confirmación")
            If (opcion = DialogResult.Yes) Then

                mdRastreo.actualizarPaciente("positivo", CInt(rastreo_txtIdRastreo.Text))
                mdRastreo.enviarEmail("Usted es POSITIVO para Covid-19. En unos pocos días un equipo asignado estara visitando su casa para proveerle de insumos.", rastreo_txtCorreoRastreo.Text)
            End If
        End If

    End Sub

    Private Sub rastreo_btnNegativo_Click(sender As Object, e As EventArgs) Handles rastreo_btnNegativo.Click

        If rastreo_txtIdRastreo.Text = "" Or rastreo_txtIdRastreo.Text = "" Then
            MsgBox("Escoja al paciente que desea actualizar", MessageBoxIcon.Warning)
        Else
            Dim opcion = MsgBox("Desea actualizar el estado del paciente a Negativo", vbYesNo + vbQuestion, "Confirmación")
            If (opcion = DialogResult.Yes) Then

                mdRastreo.actualizarPaciente("negativo", CInt(rastreo_txtIdRastreo.Text))
                mdRastreo.enviarEmail("Usted es NEGATIVO para Covid-19. Recuerde seguir cumpliendo con los consejos para combatir el virus Covid-19.", rastreo_txtCorreoRastreo.Text)
            End If
        End If

    End Sub

    Private Sub rastreo_btnGuardar_Click(sender As Object, e As EventArgs) Handles rastreo_btnGuardar.Click
        mdRastreo.rastreoGuardar()
    End Sub

    Private Sub PositivoRastreo_btnRetroceder_Click(sender As Object, e As EventArgs) Handles PositivoRastreo_btnRetroceder.Click
        Me.PacientesTableAdapter.Positivos(Me.FinalDes5DataSet.pacientes)
        pnlPositivoRastreo.Visible = False
    End Sub

    Private Sub Positivos_btnReporte_Click(sender As Object, e As EventArgs) Handles Positivos_btnReporte.Click
        pnlPositivoRastreo.Visible = True
        pnlPositivoAtendido.Visible = True
        pnlNegativos.Visible = False
        Me.PacientesTableAdapter.NoAtendidos(Me.FinalDes5DataSet.pacientes)
        Me.PacientesTableAdapter1.Atendidos(Me.FinalDes5DataSet1.pacientes)
    End Sub

    Private Sub PositivosAtendidos_btnRetroceder_Click(sender As Object, e As EventArgs) Handles PositivosAtendidos_btnRetroceder.Click
        pnlPositivoAtendido.Visible = False
        pnlPositivoRastreo.Visible = False
    End Sub

    Private Sub positivos_btnActualizar_Click(sender As Object, e As EventArgs) Handles positivos_btnActualizar.Click
        Dim opcion = MsgBox("¿Deseas actualizar este cliente a Atendido?", vbYesNo + vbQuestion, "Actualizar paciente")
        If (opcion = DialogResult.Yes) Then
            mdPositivos.positivosActualizar()
        End If
    End Sub

    Private Sub positivos_ActualizarAtendido_Click(sender As Object, e As EventArgs) Handles positivos_ActualizarAtendido.Click

        Dim opcion = MsgBox("¿Deseas actualizar este cliente a Atendido?", vbYesNo + vbQuestion, "Actualizar paciente")
        If (opcion = DialogResult.Yes) Then
            mdPositivoAtendido.positivoAtendidoEjecutar()
            Me.PacientesTableAdapter.NoAtendidos(Me.FinalDes5DataSet.pacientes)
            Me.PacientesTableAdapter1.Atendidos(Me.FinalDes5DataSet1.pacientes)
        End If

    End Sub

    '-----Panel Pacientes Negativos----'
    Private Sub Negativos_btnRetroceder_Click(sender As Object, e As EventArgs) Handles Negativos_btnRetroceder.Click
        pnlPrincipal.Visible = True
        pnlRegistrar.Visible = False
    End Sub

    '-----Panel Pacientes Recuperados----'
    Private Sub Recuperados_btnRetroceder_Click(sender As Object, e As EventArgs) Handles Recuperados_btnRetroceder.Click
        pnlPrincipal.Visible = True
        pnlRegistrar.Visible = False
    End Sub

    '-----Panel Resumen Estadísticos----'
    Private Sub pnlResumen_btnRetroceder_Click(sender As Object, e As EventArgs) Handles pnlResumen_btnRetroceder.Click
        pnlPrincipal.Visible = True
        pnlRegistrar.Visible = False

    End Sub



    '-----Panel Sobre Nosotros----'
    Private Sub SobreNosotros_btnRetroceder_Click(sender As Object, e As EventArgs) Handles SobreNosotros_btnRetroceder.Click
        pnlPrincipal.Visible = True
        pnlRegistrar.Visible = False
    End Sub

    '-----Panel Inicio----'
    Private Sub Inicio_btnComienza_Click(sender As Object, e As EventArgs) Handles Inicio_btnComienza.Click
        pnlPrincipal.Visible = True
        pnlRegistrar.Visible = False
    End Sub

    '-----Botón para Salir----'
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MsgBox("¿Desea salir de la aplicación?", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
            End
        End If
    End Sub

    '-----Botón para Minimizar----'
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub recuperados_Buscador_TextChanged(sender As Object, e As EventArgs) Handles recuperados_txtBuscador.TextChanged
        PacientesBindingSource.Filter = "cedula LIKE '" + recuperados_txtBuscador.Text + "%'"
    End Sub

    Private Sub negativo_txtBusqueda_TextChanged(sender As Object, e As EventArgs) Handles negativo_txtBusqueda.TextChanged
        PacientesBindingSource.Filter = "cedula LIKE '" + negativo_txtBusqueda.Text + "%'"
    End Sub

    Private Sub positivos_txtBusqueda_TextChanged(sender As Object, e As EventArgs) Handles positivos_txtBusqueda.TextChanged
        PacientesBindingSource.Filter = "cedula LIKE '" + positivos_txtBusqueda.Text + "%'"
    End Sub

    Private Sub dgvPositivoClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPacientesPositivos.CellClick
        positivo_txtIdCliente.Text = dgvPacientesPositivos.CurrentRow.Cells(0).Value.ToString()
        positivo_txtNombre.Text = dgvPacientesPositivos.CurrentRow.Cells(1).Value.ToString()
        positivo_txtApellido.Text = dgvPacientesPositivos.CurrentRow.Cells(2).Value.ToString()
    End Sub

    Private Sub dgRastreoClick(sender As Object, e As DataGridViewCellEventArgs) Handles rastreo_dgvRastreo.CellClick
        rastreo_txtIdRastreo.Text = rastreo_dgvRastreo.CurrentRow.Cells(0).Value.ToString()
        rastreo_txtCorreoRastreo.Text = rastreo_dgvRastreo.CurrentRow.Cells(8).Value.ToString()
    End Sub

    Private Sub espera_txtBuscador_TextChanged(sender As Object, e As EventArgs) Handles espera_txtBuscador.TextChanged
        PacientesBindingSource.Filter = "cedula LIKE '" + espera_txtBuscador.Text + "%'"
    End Sub



End Class
