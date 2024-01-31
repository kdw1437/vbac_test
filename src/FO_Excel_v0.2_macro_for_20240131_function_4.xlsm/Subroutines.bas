Attribute VB_Name = "Subroutines"
'서브루틴을 모아놓은 모듈입니다.
' selection sort(bubble sort와 유사) 알고리즘을 사용해서 특정 칼럼의 값에 근거, worksheet내의 특정 영역에 대해 sorting을 수행한다.
'
' @subroutine SortTenorAndRate
' @param {Worksheet} ws - 정렬될 데이터를 포함하는 worksheet
' @param {Integer} startRow - 정렬될 range의 시작 row index
' @param {Integer} startColumn - 이 column index에 근거해서 정렬이 수행되게 된다.
' @param {Integer} numRows - 정렬에 포함될 row의 수
' @usage - startColumn의 값에 근거해서 ascending order로 특정 영역을 정렬한다.
Sub SortTenorAndRate(ws As Worksheet, startRow As Integer, startColumn As Integer, numRows As Integer)
    Dim i As Integer, j As Integer 'Sorting 알고리즘에서 Counter로 사용되어 진다.
    
    Dim minIndex As Integer '각 sort시 가장 작은 Tenor값을 row의 인덱스에 저장한다.
    Dim tempTenor As Variant, tempRate As Variant '일시적으로 정렬 과정 중, cell을 바꿀 때, 값을 저장하기 위해 사용된다.

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
' 주어진 data string과 함께 POST request를 특정 URL에 보낸다. response는 message box에 보여진다.
'
' @subroutine SendPostRequest
' @param {String} DataString - POST request에서 보내질 데이터
' @param {String} url - POST request가 보내질 URL
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

