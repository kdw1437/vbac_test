Attribute VB_Name = "YieldCurve"
'YieldCurve값을 통화값에 맞춰 다이나믹하게 넣어주는 코드
Sub InputYieldCurves()
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
    baseURL = "http://localhost:8080/val/v1/YieldCurves/official?basedt="
    url = baseURL & dateParameter
    
    ' Parse the JSON response
    Dim jsonResponse As Object
    Set jsonResponse = GetJsonResponse(url)
    
    Dim YieldCurveRow As Integer
    YieldCurveRow = ws.Columns(1).Find(What:="Yield Curve", LookIn:=xlValues, LookAt:=xlPart).Row
    
    Dim CurrencyRow As Integer
    CurrencyRow = YieldCurveRow + 2
    
    'Dim CurrencyArray As Variant
    
    ' ... [earlier code remains the same]
    Dim LastCurrencyColumn As Integer
    Dim col As Integer
    col = 1 ' Assuming the first currency starts in column A (which is column 1)
    
    ' Loop through columns, jumping two at a time (to skip one column in between)
    Do
        ' Check if the next expected currency column is empty
        If IsEmpty(ws.Cells(CurrencyRow, col).Value) Then
            ' If it's empty, exit the loop and use the previous column as the last currency column
            LastCurrencyColumn = col - 2
            Exit Do
        Else
            ' If it's not empty, move to the next expected currency column
            col = col + 2
        End If
    Loop While col <= ws.Columns.Count 'Column끝까지 다 세아리기
    
    ' If no empty column is found, set the last currency column to the last checked column
    If LastCurrencyColumn = 0 Then
        LastCurrencyColumn = col - 2
    End If

    ' Assuming the currencies are in row CurrencyRow and start from column B
    Dim CurrencyColumn As Integer
    CurrencyColumn = 1 ' Column A
    
    
    ' Create a dictionary to hold currency column indexes
    Dim CurrencyDict As Object
    Set CurrencyDict = CreateObject("Scripting.Dictionary")
    
    ' Populate the dictionary with currency column indexes
    Dim i As Integer
    For i = CurrencyColumn To LastCurrencyColumn Step 2 'Increment by 2
        Dim currencyCode As String
        currencyCode = ws.Cells(CurrencyRow, i).Value
        CurrencyDict.Add currencyCode, i
    Next i
    
    ' Extract the data array from the JSON response
    Dim selYieldCurve As Collection
    Set selYieldCurve = jsonResponse("selYieldCurve")
    
    ' Variable to hold tenor and rate columns
' ... [previous code remains the same up to setting the CurrencyDict] ...
    Dim currencyCode2 As Variant
    ' Create a dictionary to hold the start and end rows for sorting
    Dim CurrencyRowsDict As Object
    Set CurrencyRowsDict = CreateObject("Scripting.Dictionary")
' Iterate over each currency in the dictionary
    For Each currencyCode2 In CurrencyDict.Keys
        ' Reset the TenorRow for each new currency
        Dim TenorRow As Integer
        TenorRow = YieldCurveRow + 4 ' Start 4 rows below the 'Yield Curve' header
    
        ' Find the columns for Tenor and Rate based on the currency
        TenorColumn = CurrencyDict(currencyCode2) ' Tenor is in the same column as the currency code
        RateColumn = TenorColumn + 1 ' Rate is one column to the right
        Dim item As Variant
        
            ' Keep track of the first row where data for this currency starts
        Dim FirstDataRow As Integer
        FirstDataRow = TenorRow

        ' Iterate through each entry in the JSON data
        For Each item In selYieldCurve
            ' Split the data string by '|'
            Dim dataParts As Variant
            dataParts = Split(item("data"), "|")
    
            ' Skip the header row in the JSON data
            If dataParts(0) = "DATA_ID" Then GoTo Continue
    
            ' Check if the current item's currency matches the dictionary currency
            If dataParts(0) = currencyCode2 Then
                ' Input the tenor and rate for the matching currency
                ws.Cells(TenorRow, TenorColumn).Value = dataParts(3) ' TENOR
                ws.Cells(TenorRow, RateColumn).Value = dataParts(4) ' RATE
                

    
                ' Move to the next row for the next tenor-rate pair
                TenorRow = TenorRow + 1
            End If
Continue:
        Next item
        
        ' After inputting all data for this currency, sort the data
        ' After inputting all data for this currency, sort the data
        If TenorRow > FirstDataRow Then ' Check if any data was input
            ' Calculate the number of rows to sort
            Dim NumRowsToSort As Integer
            NumRowsToSort = TenorRow - FirstDataRow
            
            ' Explicitly convert TenorColumn to Integer
            Dim TenorColumnInt As Integer
            TenorColumnInt = CInt(TenorColumn)
            
            ' Call the sorting subroutine with the converted TenorColumnInt
            SortTenorAndRate ws, FirstDataRow, TenorColumnInt, NumRowsToSort
        End If
    Next currencyCode2

End Sub



