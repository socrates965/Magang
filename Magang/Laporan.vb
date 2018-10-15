Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office
Public Class Laporan

    Private Sub Laporan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormBorderStyle = FormBorderStyle.None
        MdiParent = Form1
        Dock = DockStyle.Fill
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FakturIrma As String = My.Application.Info.DirectoryPath + "\laporan.xlsx"
        OpenExcelDemo(FakturIrma, "Sheet1")
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

            Dim y As Integer
            Dim batch() As String

            'Get Data Jual From Database
            Call Connect()
            cmd = New OleDb.OleDbCommand("SELECT No_Faktur, Tanggal_Jual, Nama_Barang, Harga_Barang, Jumlah, Total_Harga, Reseller, Kode_Batch FROM TransaksiJual WHERE Tanggal_Jual >= #" & Format(DateTimePicker1.Value, "MM/dd/yyyy") & "# AND Tanggal_Jual <= #" & Format(DateTimePicker2.Value, "MM/dd/yyyy") & "#", cn)
            dr = cmd.ExecuteReader
            Do While dr.Read()
                y = y + 1

                'Nomor
                'xlWorkSheet.Cells(y + 1, 0).Value = y.ToString

                'No_Faktur_Jual
                xlWorkSheet.Cells(y + 1, 9).Value = dr.Item(0)

                'Tanggal_Jual
                xlWorkSheet.Cells(y + 1, 2).Value = dr.Item(1)

                'Nama_Barang
                xlWorkSheet.Cells(y + 1, 5).Value = dr.Item(2)

                'Harga_Jual
                xlWorkSheet.Cells(y + 1, 10).Value = dr.Item(3)

                'Qty
                xlWorkSheet.Cells(y + 1, 11).Value = dr.Item(4)

                'Total_Jual
                xlWorkSheet.Cells(y + 1, 12).Value = dr.Item(5)

                'Reseller
                xlWorkSheet.Cells(y + 1, 8).Value = dr.Item(6)

                'Get Kode_Batch For Search Beli
                ReDim Preserve batch(y - 1)
                batch(y - 1) = dr.Item(7)
            Loop
            dr.Close()
            cn.Close()

            'Get Data Beli From Database
            Dim i As Integer = 0
            Dim j As Integer
            While i < y
                Call Connect()
                cmd = New OleDb.OleDbCommand("SELECT No_Faktur, Harga_Barang, Supplier FROM TransaksiBeli WHERE Kode_Batch = '" & batch(i) & "'", cn)
                dr = cmd.ExecuteReader
                Do While dr.Read()
                    j = j + 1
                    'No_Faktur_Beli
                    xlWorkSheet.Cells(j + 1, 4).Value = dr.Item(0)

                    'Harga_Beli
                    xlWorkSheet.Cells(j + 1, 6).Value = dr.Item(1)

                    'Supplier
                    xlWorkSheet.Cells(j + 1, 3).Value = dr.Item(2)

                    'Jumlah_Modal
                    Dim modal As Integer = dr.Item(1) * xlWorkSheet.Cells(j + 1, 11).Value
                    xlWorkSheet.Cells(j + 1, 7).Value = modal
                Loop
                i = i + 1
                dr.Close()
                cn.Close()
            End While

            Dim z As Integer
            Do While y > z
                z = z + 1

                'Nomor
                xlWorkSheet.Cells(z + 1, 1).Value = z

                'Selisih
                Dim selisih As Integer = xlWorkSheet.Cells(z + 1, 12).Value - xlWorkSheet.Cells(z + 1, 7).Value
                xlWorkSheet.Cells(z + 1, 13).Value = selisih
            Loop

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

    Public Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub


End Class