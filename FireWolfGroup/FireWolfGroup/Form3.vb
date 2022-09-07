Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'FireClientDataSet.CLIENTE' Puede moverla o quitarla según sea necesario.
        Me.CLIENTETableAdapter.Fill(Me.FireClientDataSet.CLIENTE)

        Button2.Enabled = False

        Num_clienteTextBox.Text = ""
        NombreTextBox.Text = ""
        ApellidoTextBox.Text = ""
        DniTextBox.Text = ""
        DireccionTextBox.Text = ""
        CelularTextBox.Text = ""
        PlanTextBox.Text = ""
        EstadoTextBox.Text = ""
        EquipoTextBox.Text = ""
        Modelo_exTextBox.Text = ""
        RouterTextBox.Text = ""

    End Sub

    Private Sub CLIENTEBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs) Handles CLIENTEBindingNavigatorSaveItem.Click

        Me.Validate()
        Me.CLIENTEBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.FireClientDataSet)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Num_clienteTextBox.Text = ""
        NombreTextBox.Text = ""
        ApellidoTextBox.Text = ""
        DniTextBox.Text = ""
        DireccionTextBox.Text = ""
        CelularTextBox.Text = ""
        PlanTextBox.Text = ""
        EstadoTextBox.Text = ""
        EquipoTextBox.Text = ""
        Modelo_exTextBox.Text = ""
        RouterTextBox.Text = ""

        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim B As String

        Button1.Enabled = True
        Button4.Enabled = True
        Button3.Enabled = True
        Button5.Enabled = True

        If Num_clienteTextBox.Text = "" And NombreTextBox.Text = "" And ApellidoTextBox.Text = "" And DniTextBox.Text = "" And DireccionTextBox.Text = "" And CelularTextBox.Text = "" And PlanTextBox.Text = "" And EstadoTextBox.Text = "" And EquipoTextBox.Text = "" And Modelo_exTextBox.Text = "" And RouterTextBox.Text = "" Then

            B = MsgBox("DEBE LLENAR TODOS LOS CAMPOS CRACK", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Else
            Me.CLIENTETableAdapter.AgregarClientes(Num_clienteTextBox.Text, NombreTextBox.Text, ApellidoTextBox.Text, DniTextBox.Text, DireccionTextBox.Text, CelularTextBox.Text, PlanTextBox.Text, EstadoTextBox.Text, Fecha_altaDateTimePicker.Value, Fecha_bajaDateTimePicker.Value, EquipoTextBox.Text, Modelo_exTextBox.Text, RouterTextBox.Text)
            Me.CLIENTETableAdapter.Fill(FireClientDataSet.CLIENTE)

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.CLIENTETableAdapter.ActualizarCliente(Num_clienteTextBox.Text, NombreTextBox.Text, ApellidoTextBox.Text, DniTextBox.Text, DireccionTextBox.Text, CelularTextBox.Text, PlanTextBox.Text, EstadoTextBox.Text, Fecha_altaDateTimePicker.Value, Fecha_bajaDateTimePicker.Value, EquipoTextBox.Text, Modelo_exTextBox.Text, RouterTextBox.Text, Num_clienteTextBox.Text)
        Me.CLIENTETableAdapter.Fill(FireClientDataSet.CLIENTE)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim A As String

        A = MsgBox("¿Desea eliminar este registro?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If A = MsgBoxResult.Yes Then

            Me.CLIENTETableAdapter.EliminarCliente(Num_clienteTextBox.Text)
            Me.CLIENTETableAdapter.Fill(FireClientDataSet.CLIENTE)

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.CLIENTETableAdapter.Fill(FireClientDataSet.CLIENTE)
    End Sub
End Class