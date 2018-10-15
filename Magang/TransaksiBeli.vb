Public Class TransaksiBeli

    Sub ComboBox()
        Call Connect()
        cmd = New OleDb.OleDbCommand("SELECT Kode_Barang FROM Inventori", cn)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Clear()
        Do While dr.Read
            ComboBox3.Items.Add(dr.Item(0))
        Loop
        cmd.Dispose()
        dr.Close()
        cn.Close()

        Call Connect()
        cmd = New OleDb.OleDbCommand("SELECT Nama_Supplier FROM Supplier", cn)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Clear()
        Do While dr.Read
            ComboBox1.Items.Add(dr.Item(0))
        Loop
        cmd.Dispose()
        dr.Close()
        cn.Close()
    End Sub

    Public Sub ClearBoxes(parent As Control)
        For Each child As Control In parent.Controls
            ClearBoxes(child)
        Next
        If TryCast(parent, TextBox) IsNot Nothing Then
            TryCast(parent, TextBox).Text = ""
        End If
        If TryCast(parent, ComboBox) IsNot Nothing Then
            TryCast(parent, ComboBox).Text = ""
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = "" Then

        Else
            If ComboBox3.Text = "" Then
                MsgBox("Masukkan Kode Barang Dulu!")
            Else
                If Not Integer.TryParse(TextBox5.Text, Nothing) Then
                    MsgBox("Input Jumlah Harus Angka")
                    TextBox5.Text = ""
                Else
                    TextBox7.Text = Integer.Parse(TextBox5.Text) * Integer.Parse(TextBox3.Text)
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox5.Text = "" Then
            MsgBox("Input Jumlah Tidak Boleh Kosong")
        Else
            DataGridView1.Rows.Add(New Object() {TextBox1.Text, ComboBox3.Text, ComboBox2.Text, TextBox3.Text, TextBox5.Text, TextBox7.Text, TextBox4.Text, ComboBox4.Text, Format(DateTimePicker1.Value, "dd-MM-yyyy"), Format(DateTimePicker2.Value, "dd-MM-yyyy"), ComboBox1.Text})
        End If
    End Sub

    Private Sub TransaksiBeli_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormBorderStyle = FormBorderStyle.None
        MdiParent = Form1
        Dock = DockStyle.Fill
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Call ComboBox()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If cn.State = ConnectionState.Closed Then
            Call Connect()
            cmd = New OleDb.OleDbCommand("SELECT Nama_Barang, Harga_Barang, Stok FROM Inventori WHERE Kode_Barang = '" & ComboBox3.Text & "'", cn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                ComboBox2.Text = dr.Item(0)
                TextBox3.Text = dr.Item(1)
                TextBox6.Text = dr.Item(2)
            End If
            dr.Close()
            cn.Close()
        End If
        TextBox5.Text = ""
    End Sub

    Private Sub ComboBox2_keydown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call Connect()
            cmd = New OleDb.OleDbCommand("SELECT Nama_Barang FROM Inventori WHERE Nama_Barang LIKE '%" & ComboBox2.Text & "%' ", cn)
            dr = cmd.ExecuteReader
            ComboBox2.Items.Clear()
            Do While dr.Read
                ComboBox2.Items.Add(dr.Item(0))
            Loop
            cmd.Dispose()
            dr.Close()
            cn.Close()
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If cn.State = ConnectionState.Closed Then
            Call Connect()
            cmd = New OleDb.OleDbCommand("SELECT Kode_Barang, Harga_Barang, Stok FROM Inventori WHERE Nama_Barang = '" & ComboBox2.Text & "'", cn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                ComboBox3.Text = dr.Item(0)
                TextBox3.Text = dr.Item(1)
                TextBox6.Text = dr.Item(2)
            End If
            dr.Close()
            cn.Close()
        End If
        TextBox5.Text = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            DataGridView1.Rows.Remove(DataGridView1.SelectedRows(0))
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim x As Integer
        For x = 0 To DataGridView1.Rows.Count - 1
            Call Connect()

            Dim Simpan As String
            Simpan = "INSERT INTO TransaksiBeli(No_Faktur, Kode_Barang, Nama_Barang, Harga_Barang, Jumlah, Total_Harga, Kode_Batch, Kemasan, Tanggal_Beli, Tanggal_Expire, Supplier) VALUES ('" & DataGridView1.Rows(x).Cells(0).Value & "','" & DataGridView1.Rows(x).Cells(1).Value & "','" & DataGridView1.Rows(x).Cells(2).Value & "','" & Integer.Parse(DataGridView1.Rows(x).Cells(3).Value) & "','" & Integer.Parse(DataGridView1.Rows(x).Cells(4).Value) & "','" & Integer.Parse(DataGridView1.Rows(x).Cells(5).Value) & "','" & DataGridView1.Rows(x).Cells(6).Value & "','" & DataGridView1.Rows(x).Cells(7).Value & "','" & DataGridView1.Rows(x).Cells(8).Value & "','" & DataGridView1.Rows(x).Cells(9).Value & "','" & DataGridView1.Rows(x).Cells(10).Value & "')"
            cmd = New OleDb.OleDbCommand(Simpan, cn)
            cmd.ExecuteNonQuery()

            Dim update As String
            update = "UPDATE Inventori SET Stok = Stok +'" & Integer.Parse(DataGridView1.Rows(x).Cells(4).Value) & "' WHERE Kode_Barang = '" & DataGridView1.Rows(x).Cells(1).Value & "'"
            cmd = New OleDb.OleDbCommand(update, cn)
            cmd.ExecuteNonQuery()

            cn.Close()
        Next
        MsgBox("Data Berhasil di Masukkan ke Database")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.Rows.Clear()
        ClearBoxes(Parent)
    End Sub
End Class