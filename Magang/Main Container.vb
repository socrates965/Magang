Public Class Form1
    Dim form2 As New Form2
    Dim form3 As New Form3
    Public session As Boolean


    Private Sub AToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load

        If session = False Then
            form2.MdiParent = Me
            form2.Show()
            form2.WindowState = FormWindowState.Maximized
        Else
            form3.MdiParent = Me
            form3.Show()
            form3.WindowState = FormWindowState.Maximized
        End If


    End Sub



End Class
