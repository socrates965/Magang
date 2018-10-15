Imports System.Data.OleDb

Module Connection
    Public cn As OleDbConnection
    Public dr As OleDbDataReader
    Public cmd As OleDbCommand

    Public da As OleDbDataAdapter
    Public ds As DataSet

    Public Sub Connect()
        Static Dim acdb As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb;"
        cn = New OleDbConnection(acdb)
        If cn.State = ConnectionState.Closed Then
            cn.Open()
        End If

    End Sub

    Public Sub Logout()
        Form1.session = False
        Form.ActiveForm.ActiveMdiChild.Close()
        Form2.Show()
    End Sub

End Module
