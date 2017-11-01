Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel
Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ExcelApp = New Microsoft.Office.Interop.Excel.Application
        Dim Libro = ExcelApp.Workbooks.Add
        Dim Fila As Integer = 2
        Dim Columna As Integer = 1
        Dim RowCount = DataGridView1.Rows.Count - 2
        Dim ColumnCount = DataGridView1.Columns.Count - 1

        Try
            For nColumna As Integer = 0 To ColumnCount

                Libro.Worksheets("Hoja1").Cells(1, Columna) = DataGridView1.Columns(nColumna).HeaderText
                Libro.Worksheets("Hoja1").Cells(1, Columna).Font.Bold = True

                For nFila As Integer = 0 To RowCount
                    Libro.Worksheets("Hoja1").Cells(Fila, Columna) = DataGridView1.Rows(nFila).Cells(nColumna).Value
                    Fila = Fila + 1
                Next
                Columna = Columna + 1
                Fila = 2
            Next

            SaveFileDialog1.DefaultExt = "*.xlsx"
            SaveFileDialog1.FileName = "Libro1"
            SaveFileDialog1.Filter = "Libro de Excel(*.xlsx) | *.xlsx"

            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                Libro.SaveAs(SaveFileDialog1.FileName)
                MsgBox("Los registros se exportaron satisfactoriamente")
            End If



        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Libro.Saved() = True
            ExcelApp.Quit()
            Libro = Nothing
            ExcelApp = Nothing
        End Try



    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For i As Integer = 1 To 100
            DataGridView1.Rows.Add(New String() {"Cliente " & i, "correo" & i & "@mail.com"})
        Next



    End Sub
End Class