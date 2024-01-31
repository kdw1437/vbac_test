Attribute VB_Name = "PostCorrelation1"
'������ Correlation Post�ڵ��Դϴ�.
Sub PostCorrelation1()

    
    
    Dim i As Integer
    Dim baseDt As String
    Dim dataSetId As String
    Dim StartingPoint As String
    Dim dataId As String
    Dim closePric As String
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
    
    ' Find the cell that contains "FX" after "Equity" table
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(ws.Rows.Count, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    Debug.Print fxRow.Value
    Debug.Print Table1Point.Value
    Debug.Print fxRow.Row
    Debug.Print Table1Point.Row
    
    Dim indexArray() As Variant
    
    Dim arraySize As Integer
    arraySize = fxRow.Row - Table1Point.Row - 2
    
    'Resize the array to the desired size
    ReDim indexArray(1 To arraySize)
    
    'Loop through the array to populate it
    For i = 1 To arraySize
        indexArray(i) = ws.Cells(Table1Point.Row + i, Table1Point.Column).Value
    Next i
    
    Dim j As Long
    Dim k As Long
    Dim combined_name As String
    Dim valueofcorrelation As Double
    'Dim correlationRow As Integer
    'Dim correlationColumn As Integer
    Dim DataString As String
    ' Initialize the DataString
    DataString = ""
    
    'j�� ���η� ����, k�� ���η� ����. ��� Ư�� j���� k�� �ϳ��� correlation �ٿ� �ֵ��� �ϱ�.
    'j���� ���� �� 2�̸�, k���� 1, 2�ΰ��� ���� ����. j�� k���� ���� ��쿡�� ���� ����
    'j���� 3�̸�, k���� 1, 2, 3�� ���� ���� ����.
    For j = LBound(indexArray) To UBound(indexArray)
        For k = LBound(indexArray) To UBound(indexArray)
            If Not (j = k Or j > k) Then
                combined_name = indexArray(j) & ":" & indexArray(k)
                valueofcorrelation = ws.Cells(Table1Point.Row + k, Table1Point.Column + j + 1).Value
                                ' Construct the string
                If Len(DataString) > 0 Then
                    DataString = DataString & "&"
                End If
                DataString = DataString & "BASE_DT=" & baseDt & _
                             "&DATA_SET_ID=" & dataSetId & _
                             "&DATA_ID=" & combined_name & _
                             "&CRLT_CFCN_MATX_ID=CORR" & _
                             "&TH01_DATA_ID=" & indexArray(j) & _
                             "&TH02_DATA_ID=" & indexArray(k) & _
                             "&CRLT_CFCN=" & valueofcorrelation & _
                             "&OCR_DT=" & baseDt & _
                             "&PGM_ID=MANUALLY_INPUT" & _
                             "&WRKR_ID=HS" & _
                             "&WORK_TRIP=0.0.0.0"
                
            End If
        Next k
    Next j
    Debug.Print DataString
    
            ' Encode the DataString for URL (x-www-form-urlencoded)
    DataString = URLEncode(DataString)


    ' The URL to send the request to
    Dim url As String
    url = "http://localhost:8080/val/extend_8"
    SendPostRequest DataString, url
    ' Open the HTTP request as a POST method

End Sub

