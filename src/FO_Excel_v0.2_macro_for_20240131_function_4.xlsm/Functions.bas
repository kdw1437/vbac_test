Attribute VB_Name = "Functions"
'GET request를 특정 URL에 보내고 parsing된 JSON response를 return한다.
'
' @method GetJsonResponse
' @param {String} url - GET request가 이 URL로 보내지게 된다.
' @return {Object} - 파싱된 JSON response. Dictionary나 Collection 객체가 보통 return된다.
' @usage - JSON response를 받고 파싱하기 위해서 유효한 URL과 함께 이 함수를 호출한다.
Public Function GetJsonResponse(url As String) As Object
    ' Variables to hold the HTTP request and response data
    Dim httpRequest As Object
    Dim JsonString As String
    Dim jsonResponse As Object

    ' Create the HTTP request
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", url, False
        .Send
        JsonString = .responseText
    End With

    ' Parse the JSON response
    Set jsonResponse = JsonConverter.ParseJson(JsonString)

    ' Return the parsed JSON response
    Set GetJsonResponse = jsonResponse
End Function

'셀의 범위를 변수로 받아서 (header포함), header의 값(vertical, horizontal)과 지표가 일치하는 경우 해당 corr값을 넣어주는 함수이다.
'
' @method UpdateCellsWithCorrelation
' @param {Worksheet} ws - 업데이트 될 cell들이 위치하고 있는 worksheet
' @param {Collection} selCorrelation - correlation data를 가지고 있는 collection
' @param {Integer} ColumnNameRow - column name이 위치하고 있는 row의 index
' @param {Integer} FXmarker - 셀을 업데이트하기 위한 시작 row index
' @param {Integer} LastContiguousRow - 업데이트를 위해 고려해야 할 마지막 row index
' @param {Integer} LastContiguousColumn - 업데이트를 위해 고려해야 할 마지막 column index
' @param {Integer} ColumnIndex - 셀을 업데이트하기 위한 시작 column index
' @return {Boolean} - 함수가 성공적으로 완료되면 True를 반환
Function UpdateCellsWithCorrelation(ws As Worksheet, selCorrelation As Collection, _
                                    ColumnNameRow As Integer, FXmarker As Integer, _
                                    LastContiguousRow As Integer, LastContiguousColumn As Integer, ColumnIndex As Integer) As Boolean
    Dim ColumnIndex2 As Integer
    Dim RowIndex2 As Integer

    For ColumnIndex2 = ColumnIndex To LastContiguousColumn
        Dim hheader2 As String
        hheader2 = ws.Cells(ColumnNameRow, ColumnIndex2).Value
        For RowIndex2 = FXmarker To LastContiguousRow
            Dim vheader2 As String
            vheader2 = ws.Cells(RowIndex2, 1).Value
            For i = 1 To selCorrelation.Count
                Dim data2 As Variant
                data2 = selCorrelation(i)("data")
                 
                Dim dataParts2 As Variant
                dataParts2 = Split(data2, "|")
                                 
                If (vheader2 = dataParts2(4) And hheader2 = dataParts2(5)) Or _
                   (vheader2 = dataParts2(5) And hheader2 = dataParts2(4)) Then
                    ws.Cells(RowIndex2, ColumnIndex2).Value = dataParts2(3)
                End If
            Next i
        Next RowIndex2
    Next ColumnIndex2
    
    UpdateCellsWithCorrelation = True
End Function

'String의 특수 문자를 Encoding해주는 함수이다.
'
' @method URLEncode
' @param {String} StringVal - URL인코딩될 string
' @param {Boolean} [SpaceAsPlus=False] - 스페이스가 True면 +로 False면 %20으로 인코딩 될지 여부를 결정한다.
' @return {String} - input string의 URL-encoded 버전
Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
    Dim StringLen As Long: StringLen = Len(StringVal)

    If StringLen > 0 Then
        ReDim result(StringLen) As String
        Dim i As Long, CharCode As Integer
        Dim Char As String, Space As String

        If SpaceAsPlus Then Space = "+" Else Space = "%20"

        For i = 1 To StringLen
            Char = Mid$(StringVal, i, 1)
            CharCode = Asc(Char)

            Select Case CharCode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                    result(i) = Char
                Case 32
                    result(i) = Space
                Case 0 To 15
                    result(i) = "%0" & Hex(CharCode)
                Case Else
                    result(i) = "%" & Hex(CharCode)
            End Select
        Next i

        URLEncode = Join(result, "")
    End If
End Function



