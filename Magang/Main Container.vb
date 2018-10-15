Public Class Form1

    Public session As Boolean

    Private Sub Form1_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        SessionChecker()
    End Sub

    Private Sub SessionChecker()
        If session = False Then
            Form2.Show()
        End If
    End Sub

End Class
