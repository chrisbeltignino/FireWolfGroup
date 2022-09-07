Imports System.Windows.Forms

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Código generado por el Diseñador de Windows Forms "

    Public Sub New()
        MyBase.New()

        'El Diseñador de Windows Forms requiere esta llamada.
        InitializeComponent()

        'Agregar cualquier inicialización después de la llamada a InitializeComponent()

    End Sub

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms requiere el siguiente procedimiento
    'Puede modificarse utilizando el Diseñador de Windows Forms. 
    'No lo modifique con el editor de código.
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(16, 16)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(456, 20)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "TextBox1"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(16, 48)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(456, 20)
        Me.TextBox2.TabIndex = 1
        Me.TextBox2.Text = "TextBox2"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(16, 80)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(456, 20)
        Me.TextBox3.TabIndex = 2
        Me.TextBox3.Text = "0"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(400, 112)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 147)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim ob As New Tabla1
        Dim ob As New Sucursales
        Dim cn As String = "workstation id=PABLO;packet size=4096;integrated security=SSPI;data source=server;persist security info=False;initial catalog=Almarza"
        Dim tx As String

        'ob.Campo1 = TextBox1.Text
        'ob.Campo2 = TextBox2.Text
        'ob.Campo3 = CType(TextBox3.Text, Integer)

        'ob.IDsucursal = CType(TextBox1.Text, Short)
        'ob.nombre = TextBox2.Text
        'ob.direccion = TextBox3.Text

        'tx = Sql.Alta(cn, "Sucursales", ob, "")

        'If tx <> "" Then
        '    MessageBox.Show(tx)
        'Else
        '    MessageBox.Show("Registro grabado exitosamente")
        'End If

        tx = Sql.Cargar(cn, "sucursales", "idsucursal = " & TextBox1.Text, ob, "")

        If tx = "" Then
            TextBox1.Text = ob.IDsucursal
            TextBox2.Text = ob.nombre
            TextBox3.Text = ob.direccion
        Else
            MessageBox.Show(tx)
        End If
    End Sub
End Class
