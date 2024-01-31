Attribute VB_Name = "PostYieldCurve"
'YieldCurve를 POST하는 코드
Sub PostYieldCurve()
    
    
    Dim i As Integer
    Dim baseDt As String
    Dim dataSetId As String
    Dim StartingPoint As String
    Dim dataId As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    Dim targetDate As Date
    ' Retrieve the base date and data set ID from the worksheet
    targetDate = Sheets("Market Data").Range("A2").Value
    
    baseDt = Format(targetDate, "yyyymmdd")
    dataSetId = Sheets("Market Data").Range("O2").Value
    StartingPoint = Sheets("Market Data").Range("P2").Value
        
    Dim Table1Point As Range
    Set Table1Point = Sheets("Market Data").Range(StartingPoint).Offset(3, 0)
    
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, Table1Point.Column).End(xlUp).Row
    
    ' Find the cell that contains "FX" after "Equity" table
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)

    
    Dim Table2Point As Range
    Set Table2Point = fxRow.Offset(3, 0)
    
    Dim YieldCurveRow As Range
    Set YieldCurveRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="Yield Curve", LookIn:=xlValues, LookAt:=xlWhole)
    'Debug.Print Table2Point.value

    Dim DATA_ID_Cell1 As Range
    Set DATA_ID_Cell1 = ws.Cells(YieldCurveRow.Row + 2, YieldCurveRow.Column)
    'Debug.Print DATA_ID_Cell1.value
    Dim DATA_ID_Cells() As Variant
    Dim colIndex As Long
    Dim currentCell As Range
    Dim cellCount As Integer
    
    Set currentCell = DATA_ID_Cell1
    cellCount = 0
    
    'Loop until an empty cell is found
    ' Loop until an empty cell is found
    Do
        ' Check if the current cell is empty
        If IsEmpty(currentCell.Value) Then
            Exit Do
        End If
        
        ' Resize the array and add the current cell
        cellCount = cellCount + 1
        ReDim Preserve DATA_ID_Cells(1 To cellCount)
        DATA_ID_Cells(cellCount) = currentCell.Value
        
        ' Move to the next cell 2 columns to the right
        Set currentCell = ws.Cells(currentCell.Row, currentCell.Column + 2)
    Loop
    Dim arraySize As Integer
    arraySize = UBound(DATA_ID_Cells)
    Dim InterestName As String
    Dim j As Integer
    Dim Tenor As Double
    Dim Rate As Double
    Dim RiskCode As String
    Dim DataString As String
    ' Initialize the DataString
    DataString = ""
    For i = 1 To arraySize
        InterestName = DATA_ID_Cells(i)
        j = 1
        Do While Not IsEmpty(ws.Cells(YieldCurveRow.Row + 3 + j, YieldCurveRow.Column + (i - 1) * 2))
            Tenor = ws.Cells(YieldCurveRow.Row + 3 + j, YieldCurveRow.Column + (i - 1) * 2).Value
            Rate = ws.Cells(YieldCurveRow.Row + 3 + j, YieldCurveRow.Column + (i - 1) * 2 + 1).Value
            RiskCode = Format(Tenor * 360, "00000")
            
            If Len(DataString) > 0 Then
                DataString = DataString & "&"
            End If
                DataString = DataString & "BASE_DT=" & baseDt & _
                             "&DATA_SET_ID=" & dataSetId & _
                             "&DATA_ID=" & InterestName & _
                             "&EXPR_VAL=" & Tenor & _
                             "&ERRT=" & Rate & _
                             "&RISK_FCTR_CODE=" & InterestName & "_" & RiskCode & _
                             "&OCR_DT=" & baseDt & _
                             "&PGM_ID=MANUALLY_INPUT" & _
                             "&WRKR_ID=HS" & _
                             "&WORK_TRIP=0.0.0.0"
                             
        'Increment j to move to the next cell
        j = j + 1
        Loop
    Next i
    
    'Debug.Print DataString
    
        ' Rest of the code for sending data remains the same
    Debug.Print DataString

    ' Encode the DataString for URL (x-www-form-urlencoded)
    DataString = URLEncode(DataString)
    
    
    ' The URL to send the request to
    Dim url As String
    url = "http://localhost:8080/val/extend_10"
    
    SendPostRequest DataString, url

End Sub

