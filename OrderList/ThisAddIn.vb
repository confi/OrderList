Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports System.String

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup


    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    '合并仪表阀门清单
    Sub makeOrderList()
        Dim listWorkbook As Workbook
        Dim listWorksheet As Worksheet
        Dim listRange As Range
        listWorkbook = Me.Application.ActiveWorkbook

        Dim lastRow As Integer
        Dim i As Integer
        For Each listWorksheet In listWorkbook.Worksheets
            listRange = detectRange(listWorksheet)
            lastRow = listRange.Rows.Count

            '如未经整理则进行合并和格式调整。如已经整理则提示用户。
            If listWorksheet.Range("A8").Value = Nothing Then Exit For

            If listWorksheet.Range("A8").Value.ToString <> "数量" Then

                '格式化第一列
                Dim firstCol As Range = listWorksheet.Columns(1)
                With firstCol
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .Font.Name = "Arial"
                    .ColumnWidth = 4
                End With

                Select Case listWorksheet.Name

                    Case "设备清单"
                        listRange = listWorksheet.Range("A9:J" & lastRow.ToString)
                        listWorksheet.Range("A8").Value = "数量"
                        listWorksheet.Range("F8").Value = "编号"
                        listRange.Sort(listRange.Range("C1"), XlSortOrder.xlDescending, , , , , , _
                                       XlYesNoGuess.xlNo, , , XlSortOrientation.xlSortColumns)
                        combine1(listRange, "C", "F")
                        With listWorksheet.Range("F:F")
                            .WrapText = True
                            .ColumnWidth = 15
                            .Font.Name = "Arial"
                        End With
                        listWorksheet.Range("B:B").ColumnWidth = 23
                        listWorksheet.PageSetup.PrintArea = listRange.Address & "," & "A1:J9"


                    Case "仪表清单"
                        listRange = listWorksheet.Range("A9:L" & CStr(lastRow))
                        listWorksheet.Range("A8").Value = "数量"
                        listWorksheet.Range("L8").Value = "编号"
                        For i = 9 To lastRow
                            If listWorksheet.Range("A9").Value.ToString <> "" Then
                                listWorksheet.Range("A" & i.ToString).Value = listWorksheet.Range("A" & i.ToString).Value.ToString _
                                                                            & listWorksheet.Range("B" & i.ToString).Value.ToString
                            End If
                        Next
                        listRange.Sort(listRange.Range("D1"), XlSortOrder.xlDescending, _
                                       listRange.Range("E1"), , XlSortOrder.xlDescending, , , _
                                       XlYesNoGuess.xlNo, , , XlSortOrientation.xlSortColumns)
                        combine1(listRange, "DE", "L")
                        listWorksheet.Columns(2).delete()
                        With listWorksheet.Range("K:K")
                            .WrapText = True
                            .ColumnWidth = 15
                        End With
                        listWorksheet.Range("B:B").ColumnWidth = 25


                    Case "阀门清单"
                        listRange = listWorksheet.Range("A9:H" & CStr(lastRow))
                        listWorksheet.Range("A8").Value = "数量"
                        listRange.Sort(listRange.Range("B1"), XlSortOrder.xlAscending, _
                                       listRange.Range("C1"), , XlSortOrder.xlAscending, _
                                       listRange.Range("E1"), XlSortOrder.xlAscending, _
                                       XlYesNoGuess.xlNo, , , XlSortOrientation.xlSortColumns)
                        combine1(listRange, "BCE", "H")
                        With listWorksheet.Range("H:H")
                            .WrapText = True
                            .ColumnWidth = 13
                        End With
                        listWorksheet.Range("B:B").ColumnWidth = 21

                        listWorksheet.Range("B6:B7").Value = listWorksheet.Range("A6:A7").Value
                        listWorksheet.Range("A6:A7").Value = ""
                        listWorksheet.Range("E6:E7").Value = listWorksheet.Range("C6:C7").Value
                        listWorksheet.Range("C6:C7").Value = ""
                End Select
                listRange.Rows.AutoFit()
                listRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            Else
                MsgBox("本表已经过合并整理，无需再合并！", MsgBoxStyle.Information)
                Exit Sub
            End If

        Next
    End Sub

    '合并管道管件清单
    Sub makePipeLise()

        Dim pipeWorksheet As Worksheet = Me.Application.ActiveSheet
        Dim strTitle As String() = pipeWorksheet.Range("B6:D6").Value


    End Sub

    'combineRange为需要合并的单元区域，criteriaCol为合并条件列，codeCol为PID编号存储列
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
            If data(i, 1) IsNot Nothing And data(i, 1) <> "" Then                            '跳过空值
                data(i, code) = data(i, 1)

                For j = UBound(data, 1) To i + 1 Step -1
                    '避免空值错误
                    If data(i, criteria(0)) Is Nothing Then data(i, criteria(0)) = ""
                    If data(i, criteria(1)) Is Nothing Then data(i, criteria(1)) = ""
                    If data(j, criteria(0)) Is Nothing Then data(j, criteria(0)) = ""
                    If data(j, criteria(1)) Is Nothing Then data(j, criteria(1)) = ""
                    '比较关键值

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
                If data(i, criteria(0)).ToString <> "" Then data(i, 1) = counter
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
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            .Rows.AutoFit()
            firstCol = .Columns(1)
        End With

        With firstCol
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Font.Name = "Arial"
            .ColumnWidth = 4
        End With

    End Sub

    Sub combine1(ByVal combineRange As Range, ByVal criterialCol As String, ByVal codeCol As String, Optional ByVal countTotalCol As Char = "")

        Dim startRow As Integer = 1
        Dim endRow As Integer = 2

        Do While startRow <= combineRange.Rows.Count
            Dim indicator As Boolean = True
            Dim PIDcode As String = ""
            '将备注和PID号的内容赋给PIDcode
            If Not (combineRange.Range(codeCol & startRow.ToString).Value Is Nothing) Then
                PIDcode = combineRange.Range(codeCol & startRow.ToString).Value.ToString
            End If

            If Not (combineRange.Range("A" & startRow.ToString).Value Is Nothing) Then

                If PIDcode <> "" Then
                    PIDcode = PIDcode & vbLf & combineRange.Range("A" & startRow.ToString).Value.ToString
                Else
                    PIDcode = combineRange.Range("A" & startRow.ToString).Value.ToString
                End If

            End If

            '逐行判断关键数据是否相同，如相同则判断下一行，否则跳出循环并将行指针退回一行。
            Do While indicator

                For i = 1 To criterialCol.Length
                    Dim colName As String = Mid(criterialCol, i, 1)

                    Dim startValue As String = ""
                    Dim nextValue As String = ""

                    If Not (combineRange.Range(colName & startRow.ToString).Value2 Is Nothing) Then
                        startValue = combineRange.Range(colName & startRow.ToString).Value.ToString
                    End If

                    If Not (combineRange.Range(colName & endRow.ToString).Value2 Is Nothing) Then
                        nextValue = combineRange.Range(colName & endRow.ToString).Value.ToString
                    End If
                    '如果判断数据为空则向下进行比较，对该条记录不予合并。
                    If startValue = "" Or nextValue = "" Then
                        If i < 2 Then
                            indicator = False
                        End If
                    Else

                        If startValue <> nextValue Then
                            indicator = False
                            Exit For

                        End If
                    End If

                Next

                If indicator = True Then

                    If Not (combineRange.Range("A" & endRow.ToString).Value Is Nothing) Then
                        PIDcode = PIDcode & vbLf & combineRange.Range("A" & endRow.ToString).Value.ToString
                    End If
                    endRow = endRow + 1
                Else
                    endRow = endRow - 1
                End If
            Loop

            '将PID编号移到codeCol，计数
            combineRange.Range(codeCol & startRow.ToString).Value = PIDcode
            '如果累计项为空值，则无需累计任何值，只需计数。如累计项不为空，则需累计累计项后填入总计单元。
            If countTotalCol = "" Then
                combineRange.Range("A" & startRow.ToString).Value = endRow - startRow + 1
                startRow = startRow + 1
            Else
                Dim total As Double = 0
                For i As Integer = startRow To endRow
                    total += combineRange.Range(countTotalCol & i.ToString).Value
                Next
                combineRange.Range(countTotalCol & startRow.ToString).Value = total
            End If
            '删除重复行
            If endRow >= startRow Then
                combineRange.Rows(startRow.ToString & ":" & endRow.ToString).Delete(XlDeleteShiftDirection.xlShiftUp)
            End If
            endRow = startRow + 1
            If endRow > (combineRange.Rows.Count + 2) Then Exit Do

        Loop

    End Sub

    Function detectRange(ByVal listWorksheet As Worksheet) As Range

        Dim rowNum As Integer = 8
        Dim colNum As Integer = 0
        Dim cell1 As Range = listWorksheet.Cells(1, 1)
        Dim cell2 As Range
        Dim value As String = "xxx"

        Do Until value = ""
            If listWorksheet.Cells(8, colNum + 1).value Is Nothing Then
                value = ""
            Else
                value = listWorksheet.Cells(8, colNum + 1).value.ToString
            End If
            colNum = colNum + 1
        Loop

        value = "xxx"

        Do Until value = ""
            If listWorksheet.Cells(rowNum + 1, 1).value Is Nothing Then
                value = ""
            Else
                listWorksheet.Cells(rowNum + 1, 1).value.ToString()
            End If
            rowNum = rowNum + 1
        Loop

        If rowNum > 9 And colNum > 1 Then
            cell2 = listWorksheet.Cells(rowNum - 1, colNum - 1)
        Else
            cell2 = listWorksheet.Cells(8, 1)
        End If
        Return listWorksheet.Range(cell1, cell2)
    End Function
    
End Class
