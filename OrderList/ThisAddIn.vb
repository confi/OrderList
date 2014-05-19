Imports Microsoft.Office.Interop.Excel

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        makeOrderList()
    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub


    Private Sub makeOrderList()
        Dim listWorkbook As Workbook
        Dim listWorksheet As Worksheet
        Dim listRange As Range
        listWorkbook = Me.Application.ActiveWorkbook

        Dim lastRow As Integer


        For Each listWorksheet In listWorkbook.Worksheets
            lastRow = listWorksheet.Range("print_area").Rows.Count
            Select Case listWorksheet.Name
                Case "设备清单"

                    listRange = listWorksheet.Range("A9:J" & CStr(lastRow))
                    listWorksheet.Range("A8").Value = "数量"
                    listWorksheet.Range("F8").Value = "编号"
                    combine(listRange, "C", "F")

                Case "仪表清单"
                    listRange = listWorksheet.Range("A9:L" & CStr(lastRow))
                    listWorksheet.Range("A8").Value = "数量"
                    listWorksheet.Range("L8").Value = "编号"
                    combine(listRange, "DE", "L")
                    listWorksheet.Columns(2).delete()
                Case "阀门清单"
                    listRange = listWorksheet.Range("A9:H" & CStr(lastRow))
                    listWorksheet.Range("A8").Value = "数量"
                    combine(listRange, "BC", "H")
            End Select
            
        Next
    End Sub

    Sub combine(ByVal combineRange As Range, ByVal criteriaCol As String, ByVal codeCol As String)
        Dim data(,) As Object = combineRange.Value
        Dim criteria() As Integer = {Asc(Left(criteriaCol, 1)) - 64, Asc(Right(criteriaCol, 1)) - 64}
        Dim code As Integer = Asc(codeCol) - 64
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim counter As Integer = 1


        For i = 1 To UBound(data, 1)
            data(i, code) = data(i, 1)
            For j = UBound(data, 1) To i + 1 Step -1
                If data(i, criteria(0)) = data(j, criteria(0)) And data(i, criteria(1)) = data(j, criteria(1)) Then
                    counter = counter + 1
                    data(i, code) = data(i, code) & data(j, 1)
                    For k = 1 To UBound(data, 2)
                        data(j, k) = ""
                    Next
                End If
            Next
            data(i, 1) = counter
            counter = 1
        Next


    End Sub
End Class
