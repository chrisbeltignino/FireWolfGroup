Public Class Form1
    Dim op As String
    Dim result As Nullable(Of Double) = Nothing
    Dim valor As Nullable(Of Double) = Nothing
    Dim press As Boolean

    Private Sub Button0_Click(sender As Object, e As EventArgs) Handles Button0.Click
        Concat()
        txtResult.Text &= "0"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Concat()
        txtResult.Text &= "1"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Concat()
        txtResult.Text &= "2"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Concat()
        txtResult.Text &= "3"
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Concat()
        txtResult.Text &= "4"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Concat()
        txtResult.Text &= "5"
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Concat()
        txtResult.Text &= "6"
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Concat()
        txtResult.Text &= "7"
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Concat()
        txtResult.Text &= "8"
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Concat()
        txtResult.Text &= "9"
    End Sub

    Private Sub ButtonCls_Click(sender As Object, e As EventArgs) Handles ButtonCls.Click
        txtResult.Text = "0"
        valor = Nothing
        result = Nothing
    End Sub

    Private Sub ButtonDot_Click(sender As Object, e As EventArgs) Handles ButtonDot.Click
        If InStr(txtResult.Text, ".", CompareMethod.Text) = 0 Then
            txtResult.Text &= "."
        End If
    End Sub

    Private Sub ButtonSuma_Click(sender As Object, e As EventArgs) Handles ButtonSuma.Click
        Operacion()
        op = "+"
    End Sub

    Private Sub ButtonRest_Click(sender As Object, e As EventArgs) Handles ButtonRest.Click
        Operacion()
        op = "-"
    End Sub

    Private Sub ButtonMul_Click(sender As Object, e As EventArgs) Handles ButtonMul.Click
        Operacion()
        op = "*"
    End Sub

    Private Sub ButtonDiv_Click(sender As Object, e As EventArgs) Handles ButtonDiv.Click
        Operacion()
        op = "/"
    End Sub

    Private Sub ButtonIgual_Click(sender As Object, e As EventArgs) Handles ButtonIgual.Click
        Operacion()
        op = ""
    End Sub

    Private Sub txtResult_TextChanged(sender As Object, e As EventArgs) Handles txtResult.TextChanged

    End Sub

    Public Sub Operacion()
        press = True
        valor = Val(txtResult.Text)
        If result IsNot Nothing Then
            Select Case op
                Case "+"
                    result = result + valor
                Case "-"
                    result -= valor
                Case "*"
                    result *= valor
                Case "/"
                    result /= valor
            End Select
            txtResult.Text = result
        Else
            result = valor
        End If
    End Sub

    Public Sub Concat()
        If press = True Then
            txtResult.Text = ""
            press = False
        ElseIf txtResult.Text = "0" Then
            txtResult.Text = ""
        End If
    End Sub

End Class
