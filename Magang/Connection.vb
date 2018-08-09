Imports System.Data.OleDb

Module Connection
    Public cn As OleDbConnection
    Public dr As OleDbDataReader
    Public cmd As OleDbCommand

    Public Sub Connect()
        Dim acdb As String
        acdb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb;"
        cn = New OleDbConnection(acdb)
        If cn.State = ConnectionState.Closed Then
            cn.Open()
        End If

    End Sub



End Module
