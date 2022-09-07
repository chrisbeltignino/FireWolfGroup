Public Class Form1
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ProgressBar1.Value = ProgressBar1.Value + 1
        Label1.Text = ProgressBar1.Value & "%"

        If (ProgressBar1.Value = ProgressBar1.Maximum) Then
            Timer1.Enabled = False
            Me.Hide()
            Form2.Show()
        End If
    End Sub
End Class
