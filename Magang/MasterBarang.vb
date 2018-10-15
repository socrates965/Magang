Public Class MasterBarang

    Sub ComboBox()
        Call Connect()
        cmd = New OleDb.OleDbCommand("SELECT Kode_Barang FROM Inventori", cn)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Clear()
        Do While dr.Read
            ComboBox1.Items.Add(dr.Item(0))
        Loop
        cmd.Dispose()
        dr.Close()
        cn.Close()
    End Sub

    Sub DataGrid()
        Call Connect()
        da = New OleDb.OleDbDataAdapter("SELECT * FROM Inventori", cn)
        ds = New DataSet
        da.Fill(ds, "Inventori")
        DataGridView1.DataSource = ds.Tables("Inventori")
        DataGridView1.ReadOnly = True
        cn.Close()
    End Sub

    Sub Hapus()
        ComboBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Call Connect()
            cmd = New OleDb.OleDbCommand("SELECT Kode_Barang FROM Inventori WHERE Kode_Barang='" & ComboBox1.Text & "'", cn)
            dr = cmd.ExecuteReader
            dr.Read()

            If dr.HasRows Then
                MsgBox("Maaf, Data dengan kode tersebut sudah ada", MsgBoxStyle.Exclamation, "Warning!")
                dr.Close()
            Else
                Call Connect()
                Dim Simpan As String
                Simpan = "INSERT INTO Inventori(Kode_Barang,Nama_Barang,Harga_Barang) VALUES ('" & ComboBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "')"
                cmd = New OleDb.OleDbCommand(Simpan, cn)
                cmd.ExecuteNonQuery()
                cn.Close()
                Call DataGrid()
                Call Hapus()
            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try
    End Sub

    Private Sub MasterBarang_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormBorderStyle = FormBorderStyle.None
        MdiParent = Form1
        Dock = DockStyle.Fill
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Call DataGrid()
        Call ComboBox()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Hapus()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Call Connect()
            Dim ubah As String
            ubah = "UPDATE Inventori SET Nama_Barang = '" & TextBox2.Text & "', Harga_Barang = '" & TextBox3.Text & "', WHERE Kode_Barang ='" & ComboBox1.Text & "'"
            cmd = New OleDb.OleDbCommand(ubah, cn)
            cmd.ExecuteNonQuery()
            cn.Close()
            Call Hapus()
            Call DataGrid()
            Call ComboBox()
            ComboBox1.Enabled = True

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Connect()
        cmd = New OleDb.OleDbCommand("SELECT Kode_Barang,Nama_Barang, Harga_Barang FROM Inventori WHERE Kode_Barang = '" & ComboBox1.Text & "'", cn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            ComboBox1.Text = dr.Item(0)
            TextBox2.Text = dr.Item(1)
            TextBox3.Text = dr.Item(2)
            TextBox3.Focus()
        End If
        dr.Close()
        cn.Close()
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim value As Object = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        ComboBox1.Text = CType(value, String)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Call Connect()
            Dim Hapusaja As String
            Hapusaja = "DELETE FROM Inventori Where Kode_Barang='" & ComboBox1.Text & "'"
            cmd = New OleDb.OleDbCommand(Hapusaja, cn)
            cmd.ExecuteNonQuery()
            cn.Close()
            Call Hapus()
            Call DataGrid()
            Call ComboBox()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        MainMenu.Show()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
End Class