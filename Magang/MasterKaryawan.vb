Public Class MasterKaryawan
    Sub Hapus()
        ComboBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Sub ComboBox()
        Call Connect()
        cmd = New OleDb.OleDbCommand("SELECT Kode_Pegawai FROM Pegawai", cn)
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
        da = New OleDb.OleDbDataAdapter("SELECT * FROM Pegawai", cn)
        ds = New DataSet
        da.Fill(ds, "Pegawai")
        DataGridView1.DataSource = ds.Tables("Pegawai")
        DataGridView1.ReadOnly = False
        cn.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Connect()
        cmd = New OleDb.OleDbCommand("SELECT Nama_Pegawai FROM Pegawai WHERE Nomor_Pegawai = " & Integer.Parse(ComboBox1.Text), cn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox2.Text = dr.Item(0)
            TextBox2.Focus()
        End If
        cn.Close()
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        ComboBox1_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim value As Object = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        If Integer.TryParse(value, 0) = True Then
            ComboBox1.Text = CType(value, String)
        End If

    End Sub

    Private Sub MasterKaryawan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormBorderStyle = FormBorderStyle.None
        MdiParent = Form1
        Dock = DockStyle.Fill

        Call DataGrid()
        Call ComboBox()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
        MainMenu.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Call Connect()
            Dim Simpan As String
            Simpan = "INSERT INTO Pegawai(Nama_Pegawai,Jenis_Kelamin,Tempat_Lahir,Tanggal_Lahir,Nomor_Telepon,Alamat,Jabatan,Tanggal_Masuk) VALUES ( '" & TextBox2.Text & "')"
            cmd = New OleDb.OleDbCommand(Simpan, cn)
            cmd.ExecuteNonQuery()
            cn.Close()
            Call Hapus()
            InitializeComponent() 'load all the controls again
            MasterKaryawan_Load(e, e)

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Call Connect()
            Dim ubah As String
            ubah = "UPDATE Pegawai SET Nama_Pegawai = '" & TextBox2.Text & "' WHERE Nomor_Pegawai = " & Integer.Parse(ComboBox1.Text)
            cmd = New OleDb.OleDbCommand(ubah, cn)
            cmd.ExecuteNonQuery()
            cn.Close()
            Call Hapus()
            Call DataGrid()
            Call ComboBox()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Call Connect()
            Dim Hapusaja As String
            Hapusaja = "DELETE FROM Pegawai Where Nomor_Pegawai='" & ComboBox1.Text & "'"
            cmd = New OleDb.OleDbCommand(Hapusaja, cn)
            cmd.ExecuteNonQuery()
            Call Hapus()
            Call DataGrid()
            Call ComboBox()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Terjadi Kesalahan")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Hapus()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class