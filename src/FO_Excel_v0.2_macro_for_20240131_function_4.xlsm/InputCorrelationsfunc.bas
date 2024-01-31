Attribute VB_Name = "InputCorrelationsfunc"
'Correlation값을 칼럼에 맞춰 다이나믹하게 넣어주는 코드
Sub InputCorrelations()
    ' Variables to hold the HTTP request and response data
 
    ' Assuming you have a worksheet variable set to the target sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data") ' Change to your actual sheet name

    ' Retrieve the date value from cell A2 and format it as 'yyyymmdd'
    Dim targetDate As Date
    targetDate = ws.Range("A2").Value
    Dim dateParameter As String
    dateParameter = Format(targetDate, "yyyymmdd")

    ' Construct the full URL with the formatted date parameter
    Dim baseURL As String
    Dim url As String
    baseURL = "http://localhost:8080/val/v1/Correlations/official?basedt="
    url = baseURL & dateParameter

    Dim jsonResponse As Object
    Set jsonResponse = GetJsonResponse(url)
    
    Dim rowIndex As Integer
    Dim ColumnIndex As Integer
    
    Dim LastContiguousColumn As Integer
    LastContiguousColumn = 3 ' Start from column 3
    
    Dim equityRow As Integer
    equityRow = ws.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart).Row
    
    ' Starting row for writing data is 4 rows below 'Equity'
    Dim startRow As Integer
    startRow = equityRow + 4
    Dim ColumnNameRow As Integer
    ColumnNameRow = equityRow + 3
    
    While Not IsEmpty(ws.Cells(ColumnNameRow, LastContiguousColumn + 1))
        LastContiguousColumn = LastContiguousColumn + 1
    Wend
    
    Dim LastContiguousRow As Integer
    LastContiguousRow = startRow
    
    While Not IsEmpty(ws.Cells(LastContiguousRow + 1, 1))
        LastContiguousRow = LastContiguousRow + 1
    Wend
    'When I dont' know beforehand how many columns contain data.
    For ColumnIndex = 3 To LastContiguousColumn
        Dim headerValue As String
        headerValue = ws.Cells(ColumnNameRow, ColumnIndex).Value
        
        For rowIndex = startRow To LastContiguousRow
            If ws.Cells(rowIndex, 1).Value = headerValue Then
                ws.Cells(rowIndex, ColumnIndex).Value = 1
            End If
        Next rowIndex
    Next ColumnIndex

    ' Extract the correlation data from the JSON response
    Dim selCorrelation As Collection
    Set selCorrelation = jsonResponse("selCorrelation")
    
    ' Update the worksheet with correlations for 'Equity'
    Call UpdateCellsWithCorrelation(ws, selCorrelation, ColumnNameRow, startRow, LastContiguousRow, LastContiguousColumn, 3)

    ' Define the start row and column name row based on 'FX'
    Dim fxRow As Integer
    fxRow = ws.Columns(1).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    Dim FXmarker As Integer
    FXmarker = fxRow + 4
    
    Dim ColumnNameRow2 As Integer
    ColumnNameRow2 = fxRow + 3

    Dim LastContiguousRow2 As Integer
    LastContiguousRow2 = FXmarker
    
    Dim ColumnIndex2 As Integer
    ColumnIndex2 = 4
    
    While Not IsEmpty(ws.Cells(LastContiguousRow2 + 1, 1))
        LastContiguousRow2 = LastContiguousRow2 + 1
    Wend
    
    Dim LastContiguousColumn2 As Integer
    LastContiguousColumn2 = ColumnIndex2
    
    While Not IsEmpty(ws.Cells(ColumnNameRow2, LastContiguousColumn2 + 1))
        LastContiguousColumn2 = LastContiguousColumn2 + 1
    Wend

    ' Update the worksheet with correlations for 'FX'
    Call UpdateCellsWithCorrelation(ws, selCorrelation, ColumnNameRow2, FXmarker, LastContiguousRow2, LastContiguousColumn2, 4)
End Sub


