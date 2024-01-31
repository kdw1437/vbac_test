Attribute VB_Name = "Subroutines"
'서브루틴을 모아놓은 모듈입니다.
'sorting하는 subroutine (startRow: sorting될 첫번째 row, startColumn: 이 칼럼을 기준으로 sorting이 일어나게 된다, numRows: Sorting될 row의 수)
Sub SortTenorAndRate(ws As Worksheet, startRow As Integer, startColumn As Integer, numRows As Integer)
    Dim i As Integer, j As Integer
    Dim minIndex As Integer '각 sort시 가장 작은 Tenor값을 row의 인덱스에 저장한다.
    Dim tempTenor As Variant, tempRate As Variant

    ' Bubble Sort by Tenor
    For i = startRow To startRow + numRows - 1
        minIndex = i
        For j = i + 1 To startRow + numRows - 1
            If ws.Cells(j, startColumn).Value < ws.Cells(minIndex, startColumn).Value Then
                minIndex = j
            End If
        Next j
        ' Swap Tenor
        tempTenor = ws.Cells(minIndex, startColumn).Value
        ws.Cells(minIndex, startColumn).Value = ws.Cells(i, startColumn).Value
        ws.Cells(i, startColumn).Value = tempTenor
        ' Swap Rate
        tempRate = ws.Cells(minIndex, startColumn + 1).Value
        ws.Cells(minIndex, startColumn + 1).Value = ws.Cells(i, startColumn + 1).Value
        ws.Cells(i, startColumn + 1).Value = tempRate
    Next i
End Sub
'POST request하는 코드
Sub SendPostRequest(DataString As String, url As String)
    Dim xmlhttp As Object
    
    ' Create a new XML HTTP request
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Open the HTTP request as a POST method
    xmlhttp.Open "POST", url, False
    
    ' Set the request content-type header to application/x-www-form-urlencoded
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Send the request with the DataString
    xmlhttp.Send "a=" & DataString
    
    ' Check the status of the request
    If xmlhttp.Status = 200 Then
        ' If the request was successful, output the response
        MsgBox xmlhttp.responseText
    Else
        ' If the request failed, output the status
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
    End If
    
    ' Clean up
    Set xmlhttp = Nothing
End Sub

