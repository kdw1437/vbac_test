Attribute VB_Name = "InputClosePrices"
'ClosePrice�� �ش��ϴ� ���� �־��ִ� �ڵ�
Sub UpdateClosePrice()
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
    baseURL = "http://localhost:8080/val/v1/ClosePrices/official?basedt="
    url = baseURL & dateParameter

    ' Use the GetJsonResponse function to get the parsed JSON response
    Dim jsonResponse As Object
    Set jsonResponse = GetJsonResponse(url)
    ' ... [earlier code remains the same]

    ' Extract the data_get_1 array from the JSON response
    Dim selClosePrice As Collection
    Set selClosePrice = jsonResponse("selClosePrice")
    
    ' Find the row with 'Equity' in column A
    Dim equityRow As Integer
    equityRow = ws.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart).Row

    Dim startRow As Integer
    startRow = equityRow + 4
    
    Dim codeCell As Range
    Dim codeValue As String
    Dim i As Integer
    
    ' Loop through each code in the worksheet starting from startRow
    For Each codeCell In ws.Range("A" & startRow & ":A" & ws.Rows.Count).Cells 'ws.Rows.Count: ������ row����
        ' Stop if the cell is empty
        If IsEmpty(codeCell.Value) Then Exit For '����ִ� cell�� ������ loop�� ������.
        
        ' Get the code value from the cell
        codeValue = codeCell.Value
        
        ' Loop through each item in the selClosePrice collection
        For i = 1 To selClosePrice.Count
            Dim data As Variant
            data = selClosePrice(i)("data") 'JSON object�� key�� data�� value���� data�� ����
            
            ' Split the data string by '|'
            Dim dataParts As Variant
            dataParts = Split(data, "|")
            
            ' Check if the first part of the data (DATA_ID) matches the codeValue
            If dataParts(0) = codeValue Then
                ' If it matches, update the closed price in the next column
                codeCell.Offset(0, 1).Value = dataParts(3)
                Exit For 'Inner loop ������ ������ ���ؼ� �ʿ��ϴ�. codeValue�� �ش��ϴ� ���� ã�Ҵ�.
            End If
        Next i
    Next codeCell
    
    'MsgBox "Update complete!"
End Sub
    
    
