Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Connect()
        Dim Name As String
        Dim query As String = "SELECT * FROM [User] WHERE [Name] ='" & TextBox1.Text & "'AND [Password] = '" & TextBox2.Text & "'"

        cmd = New OleDb.OleDbCommand(query, cn)
        dr = cmd.ExecuteReader

        While dr.Read
            Form1.session = True
            'Access = dr("Access").ToString
        End While
        cn.Close()

        If Form1.session = True Then
            Me.Close()
            Form3.MdiParent = Form1
            Form3.WindowState = FormWindowState.Maximized
            Form3.Show()
        Else
            MsgBox("Username atau Pasword salah")
        End If
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MaximizeBox = False
    End Sub

    Private Sub Form2_Closed(sender As Object, e As EventArgs) Handles MyBase.Closed

    End Sub
End Class