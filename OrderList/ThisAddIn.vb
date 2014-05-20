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
        Dim i As Integer


        For Each listWorksheet In listWorkbook.Worksheets
            lastRow = listWorksheet.Range("print_area").Rows.Count
            Select Case listWorksheet.Name
                Case "设备清单"

                    listRange = listWorksheet.Range("A9:J" & CStr(lastRow))
                    listWorksheet.Range("A8").Value = "数量"
                    listWorksheet.Range("F8").Value = "编号"
                    combine(listRange, "C", "F")
                    listWorksheet.Range("F:F").WrapText = True

                Case "仪表清单"
                    listRange = listWorksheet.Range("A9:L" & CStr(lastRow))
                    listWorksheet.Range("A8").Value = "数量"
                    listWorksheet.Range("L8").Value = "编号"
                    For i = 9 To lastRow
                        If listWorksheet.Range("A9").Value <> "" Then
                            listWorksheet.Range("A" & i.ToString).Value = listWorksheet.Range("A" & i.ToString).Value.ToString _
                                                                        & listWorksheet.Range("B" & i.ToString).Value.ToString
                        End If
                    Next
                    combine(listRange, "DE", "L")
                    listWorksheet.Columns(2).delete()
                    listWorksheet.Range("K:K").WrapText = True
                Case "阀门清单"
                    listRange = listWorksheet.Range("A9:H" & CStr(lastRow))
                    listWorksheet.Range("A8").Value = "数量"
                    combine(listRange, "BC", "H")
                    listWorksheet.Range("H:H").WrapText = True

            End Select


        Next
    End Sub

    Sub combine(ByRef combineRange As Range, ByVal criteriaCol As String, ByVal codeCol As String)
        Dim data(,) As Object = combineRange.Value
        Dim criteria() As Integer = {Asc(Left(criteriaCol, 1)) - 64, Asc(Right(criteriaCol, 1)) - 64}
        Dim code As Integer = Asc(codeCol) - 64
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim counter As Integer = 1
        Dim blankCounter As Integer = 0
        Dim lastRow As Range

        For i = 1 To UBound(data, 1)
            If data(i, 1) <> "" Then                            '跳过空值
                data(i, code) = data(i, 1)
                For j = UBound(data, 1) To i + 1 Step -1
                    If data(i, criteria(0)).ToString = data(j, criteria(0)).ToString _
                    And data(i, criteria(1)).ToString = data(j, criteria(1)).ToString Then
                        counter = counter + 1
                        data(i, code) = data(i, code) & "," & data(j, 1)
                        For k = 1 To UBound(data, 2)
                            data(j, k) = ""
                        Next
                        blankCounter = blankCounter + 1
                    End If
                Next
                If data(i, criteria(0)) <> "" Then data(i, 1) = counter
                counter = 1
            End If
        Next

        '按关键列降序排列，空行排列于最后
        Dim swap As String
        For i = 1 To UBound(data, 1)
            For j = UBound(data, 1) To i + 1 Step -1
                If data(i, criteria(0)).ToString < data(j, criteria(0)).ToString Then
                    For k = 1 To UBound(data, 2)
                        swap = data(i, k)
                        data(i, k) = data(j, k)
                        data(j, k) = swap
                    Next
                End If
            Next
        Next
        For i = blankCounter To 2 Step -1
            lastRow = combineRange.Rows(combineRange.Rows.Count)
            lastRow.Delete()
        Next

        Dim firstCol As Range
        With combineRange
            .Value = data
            '调整格式
            .VerticalAlignment = XlVAlign.xlVAlignCenter
            .Rows.AutoFit()
            firstCol = .Columns(1)
        End With

        With firstCol
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
            .Font.Name = "Arial"
            .ColumnWidth = 4
        End With

    End Sub
End Class
