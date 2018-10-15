Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office
Imports System.Text.RegularExpressions

Public Class Transaksijual
    Dim y, TempInt(100), z As Integer
    Dim TempStr(100) As String

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

    Private Sub TransaksiPenjualan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormBorderStyle = FormBorderStyle.None
        MdiParent = Form1
        Dock = DockStyle.Fill

        Form1.Size = New Size(Me.Width, Me.Height)
        Call ComboBox()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox5.Text = "" Then
            MsgBox("Input Jumlah Tidak Boleh Kosong")
        Else
            DataGridView1.Rows.Add(New Object() {TextBox1.Text,
                    ComboBox3.Text, ComboBox2.Text, TextBox3.Text,
                    TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox4.Text,
                    ComboBox4.Text, Format(DateTimePicker1.Value, "dd-MM-yyyy"),
                    Format(DateTimePicker2.Value, "dd-MM-yyyy"), ComboBox1.Text, TextBox2.Text})
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text Is Nothing Then

        ElseIf Not Integer.TryParse(TextBox5.Text, Nothing) Then
            MsgBox("Input Jumlah Harus Angka")
            TextBox5.Text = Nothing

        ElseIf ComboBox3.Text Is Nothing Then
            MsgBox("Masukkan Kode Barang")

        ElseIf TextBox8.Text = "" Then
            TextBox7.Text = Integer.Parse(TextBox5.Text) * Integer.Parse(TextBox3.Text)

        ElseIf TextBox8.Text > 100 Then
            MsgBox("Potongan Harga Tidak Boleh Lebih Besar Dari 100%!")

        Else
            TextBox7.Text = Integer.Parse(TextBox5.Text) * (Integer.Parse(TextBox3.Text) / 100 * (100 - Integer.Parse(TextBox8.Text)))

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

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If cn.State = ConnectionState.Closed Then
            Call Connect()
            cmd = New OleDb.OleDbCommand("SELECT Nama_Barang,  Stok FROM Inventori WHERE Kode_Barang = '" & ComboBox3.Text & "'", cn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                ComboBox2.Text = dr.Item(0)
                TextBox6.Text = dr.Item(1)
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
            Simpan = "INSERT INTO TransaksiJual(No_Faktur, Kode_Barang, Nama_Barang, Harga_Barang, Jumlah, 
                        Total_Harga, Potongan, Kode_Batch, Kemasan, Tanggal_Jual, Tanggal_Expire, Reseller, Sales) 
                        VALUES ('" & DataGridView1.Rows(x).Cells(0).Value & "','" & DataGridView1.Rows(x).Cells(1).Value & "',
                        '" & DataGridView1.Rows(x).Cells(2).Value & "','" & Integer.Parse(DataGridView1.Rows(x).Cells(3).Value) & "',
                        '" & Integer.Parse(DataGridView1.Rows(x).Cells(4).Value) & "','" & Integer.Parse(DataGridView1.Rows(x).Cells(5).Value) & "','" & DataGridView1.Rows(x).Cells(6).Value & "',
                        '" & DataGridView1.Rows(x).Cells(7).Value & "','" & DataGridView1.Rows(x).Cells(8).Value & "','" & DataGridView1.Rows(x).Cells(9).Value & "',
                        '" & DataGridView1.Rows(x).Cells(10).Value & "','" & DataGridView1.Rows(x).Cells(11).Value & "','" & DataGridView1.Rows(x).Cells(12).Value & "')"

            cmd = New OleDb.OleDbCommand(Simpan, cn)
            cmd.ExecuteNonQuery()
            Dim update As String
            update = "UPDATE Inventori SET Stok = Stok -'" & Integer.Parse(DataGridView1.Rows(x).Cells(4).Value) & "' WHERE Kode_Barang = '" & DataGridView1.Rows(x).Cells(1).Value & "'"
            cmd = New OleDb.OleDbCommand(update, cn)
            cmd.ExecuteNonQuery()
            cn.Close()
        Next
        MsgBox("Data Berhasil di Masukkan ke Database")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.Rows.Clear()
        ClearBoxes(Parent)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Show()
    End Sub

    Public Sub OpenExcelDemo(ByVal FileName As String, ByVal SheetName As String)
        If IO.File.Exists(FileName) Then
            Dim Proceed As Boolean = False
            Dim xlApp As Excel.Application = Nothing
            Dim xlWorkBooks As Excel.Workbooks = Nothing
            Dim xlWorkBook As Excel.Workbook = Nothing
            Dim xlWorkSheet As Excel.Worksheet = Nothing
            Dim xlWorkSheets As Excel.Sheets = Nothing
            Dim xlCells As Excel.Range = Nothing
            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            xlWorkBook = xlWorkBooks.Open(FileName)
            xlApp.Visible = True
            xlWorkSheets = xlWorkBook.Sheets
            For x As Integer = 1 To xlWorkSheets.Count
                xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)
                If xlWorkSheet.Name = SheetName Then
                    Console.WriteLine(SheetName)
                    Proceed = True
                    Exit For
                End If
                Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkSheet)
                xlWorkSheet = Nothing
            Next
            If Proceed Then
                xlWorkSheet.Activate()
                'MessageBox.Show("File is open, if you close Excel just opened outside of this program we will crash-n-burn.")
            Else
                MessageBox.Show(SheetName & " not found.")
            End If

            'No_Faktur
            xlWorkSheet.Cells(7, 2).Value = TextBox1.Text

            'Tanggal_Jual
            xlWorkSheet.Cells(7, 4).Value = Format(DateTimePicker1.Value, "dd-MM-yyyy")

            'Reseller
            xlWorkSheet.Cells(1, 11).Value = ComboBox1.Text

            'Kode_Barang
            For y = 0 To DataGridView1.Rows.Count - 1
                TempStr(y) = DataGridView1.Rows(y).Cells(1).Value
                xlWorkSheet.Cells(10 + y, 2).Value = TempStr(y)
            Next

            'Nama_Barang
            For y = 0 To DataGridView1.Rows.Count - 1
                TempStr(y) = DataGridView1.Rows(y).Cells(2).Value
                xlWorkSheet.Cells(10 + y, 3).Value = TempStr(y)
            Next

            'No_Batch
            For y = 0 To DataGridView1.Rows.Count - 1
                TempStr(y) = DataGridView1.Rows(y).Cells(7).Value
                xlWorkSheet.Cells(10 + y, 6).Value = TempStr(y)
            Next

            'Tanggal_Expire
            For y = 0 To DataGridView1.Rows.Count - 1
                TempStr(y) = DataGridView1.Rows(y).Cells(9).Value
                xlWorkSheet.Cells(10 + y, 7).Value = TempStr(y)
            Next

            'Jumlah_Barang
            For y = 0 To DataGridView1.Rows.Count - 1
                TempInt(y) = Integer.Parse(DataGridView1.Rows(y).Cells(4).Value)
                xlWorkSheet.Cells(10 + y, 8).Value = TempInt(y)
            Next

            'Harga_Satuan
            For y = 0 To DataGridView1.Rows.Count - 1
                TempInt(y) = Integer.Parse(DataGridView1.Rows(y).Cells(3).Value)
                xlWorkSheet.Cells(10 + y, 10).Value = TempInt(y)
            Next

            'Kemasan
            For y = 0 To DataGridView1.Rows.Count - 1
                TempStr(y) = DataGridView1.Rows(y).Cells(8).Value
                xlWorkSheet.Cells(10 + y, 9).Value = TempStr(y)
            Next

            'Potongan
            For y = 0 To DataGridView1.Rows.Count - 1
                TempStr(y) = DataGridView1.Rows(y).Cells(6).Value
                xlWorkSheet.Cells(10 + y, 11).Value = TempStr(y)
            Next

            'xlWorkBook.Close()
            'xlApp.UserControl = True
            'xlApp.Quit()
            ReleaseComObject(xlCells)
            ReleaseComObject(xlWorkSheets)
            ReleaseComObject(xlWorkSheet)
            ReleaseComObject(xlWorkBook)
            ReleaseComObject(xlWorkBooks)
            ReleaseComObject(xlApp)
        Else
            MessageBox.Show("'" & FileName & "' not located. Try one of the write examples first.")
        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If TextBox8.Text Is Nothing Then

        Else
            If Not Integer.TryParse(TextBox8.Text, Nothing) Then
                MsgBox("Input Jumlah Harus Angka")
                TextBox8.Text = ""
            Else
                If TextBox8.Text > 100 Then
                    MsgBox("Persentase Potongan Tidak Boleh Lebih Besar Dari 100%!")
                    TextBox8.Text = Nothing
                End If
            End If
        End If
    End Sub

    Public Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim FakturIrma As String = My.Application.Info.DirectoryPath + "\FakturIrmaModified.xlsm"
        OpenExcelDemo(FakturIrma, 713)
    End Sub
End Class