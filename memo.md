- Open connection
```vba
Function OpenDbConnection() As Object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' 🔁 Change this line for SQL Server or Access
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\To\Your\Database.accdb;"
    
    Set OpenDbConnection = conn
End Function
```

- Open conn for Sql server
```vba
Function OpenDbConnection() As Object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")

    ' Option 1: Windows Authentication (Integrated Security)
    conn.Open "Provider=SQLOLEDB;" & _
              "Data Source=YOUR_SERVER_NAME;" & _
              "Initial Catalog=YOUR_DATABASE_NAME;" & _
              "Integrated Security=SSPI;"

    ' Option 2: SQL Server Authentication (username/password)
    'conn.Open "Provider=SQLOLEDB;" & _
    '          "Data Source=YOUR_SERVER_NAME;" & _
    '          "Initial Catalog=YOUR_DATABASE_NAME;" & _
    '          "User ID=yourUser;" & _
    '          "Password=yourPassword;"

    Set OpenDbConnection = conn
End Function

```

- Execute query with header
```vba
Function ExecuteQueryWithHeaderAndRows(sql As String) As Object
    Dim conn As Object, rs As Object
    Dim i As Integer
    Dim headers As New Collection
    Dim rows As New Collection
    Dim row As Object
    Dim result As Object

    Set conn = OpenDbConnection()
    Set rs = CreateObject("ADODB.Recordset")

    rs.Open sql, conn, 1, 1 ' adOpenKeyset, adLockReadOnly

    ' Collect headers
    For i = 0 To rs.Fields.Count - 1
        headers.Add rs.Fields(i).Name
    Next i

    ' Collect rows
    Do While Not rs.EOF
        Set row = CreateObject("Scripting.Dictionary")
        For i = 0 To rs.Fields.Count - 1
            row.Add rs.Fields(i).Name, rs.Fields(i).Value
        Next i
        rows.Add row
        rs.MoveNext
    Loop

    rs.Close
    conn.Close

    ' Create result object with headers and rows
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "headers", headers
    result.Add "rows", rows

    Set ExecuteQueryWithHeaderAndRows = result
End Function
```
- Test
```
Sub TestResultWithHeaderAndRows()
    Dim result As Object
    Set result = ExecuteQueryWithHeaderAndRows("SELECT * FROM Customers")

    Dim header As Variant
    Debug.Print "Headers:"
    For Each header In result("headers")
        Debug.Print header
    Next header

    Debug.Print "Rows:"
    Dim row As Variant
    For Each row In result("rows")
        Debug.Print row("CustomerName") ' Use actual field name
    Next row
End Sub
```

- Write header to sheet
```vba
Sub WriteHeadersToSheet(result As Object, targetSheet As Worksheet, startRow As Long, startCol As Long)
    Dim c As Long
    Dim colName As Variant
    Dim headers As Collection

    Set headers = result("headers")

    c = 0
    For Each colName In headers
        targetSheet.Cells(startRow, startCol + c).Value = colName
        c = c + 1
    Next colName
End Sub

Sub WriteHeaders(headers As Collection, targetSheet As Worksheet, row As Long, col As Long)
    Dim i As Long
    For i = 1 To headers.Count
        targetSheet.Cells(row, col + i - 1).Value = headers(i)
    Next i
End Sub
```
```vba
Sub WriteRowsToSheet(result As Object, targetSheet As Worksheet, startRow As Long, startCol As Long)
    Dim r As Long, c As Long
    Dim colName As Variant
    Dim rowItem As Variant
    Dim headers As Collection
    Dim rows As Collection

    Set headers = result("headers")
    Set rows = result("rows")

    r = 0
    For Each rowItem In rows
        c = 0
        For Each colName In headers
            targetSheet.Cells(startRow + r, startCol + c).Value = rowItem(colName)
            c = c + 1
        Next colName
        r = r + 1
    Next rowItem
End Sub

Sub WriteRows(rows As Collection, headers As Collection, targetSheet As Worksheet, startRow As Long, startCol As Long)
    Dim r As Long, c As Long
    Dim rowItem As Variant
    Dim colName As Variant

    r = 0
    For Each rowItem In rows
        c = 0
        For Each colName In headers
            targetSheet.Cells(startRow + r, startCol + c).Value = rowItem(colName)
            c = c + 1
        Next colName
        r = r + 1
    Next rowItem
End Sub

```
- Insert empty rows before writing data
```
Sub InsertEmptyRows(targetSheet As Worksheet, startRow As Long, rowCount As Long)
    If rowCount <= 0 Then Exit Sub
    targetSheet.Rows(startRow & ":" & (startRow + rowCount - 1)).Insert Shift:=xlDown
End Sub
```
```
Sub WriteWithRowInsertion()
    Dim result As Object
    Set result = ExecuteQueryWithHeaderAndRows("SELECT TOP 10 * FROM Employees")

    Dim sheet As Worksheet: Set sheet = ThisWorkbook.Sheets("Sheet1")
    Dim startRow As Long: startRow = 5
    Dim startCol As Long: startCol = 2

    ' 1. Insert space for data
    InsertEmptyRows sheet, startRow, result("rows").Count

    ' 2. Write data into inserted space
    WriteRows result("rows"), result("headers"), sheet, startRow, startCol
End Sub
```

- Example
```
Sub TestModularOutput()
    Dim result As Object
    Set result = ExecuteQueryWithHeaderAndRows("SELECT TOP 10 * FROM Employees")

    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("Sheet1")

    ' Optional: write header
    Call
```

- Set style
```
Sub ApplyStyleFromTemplateRow(targetSheet As Worksheet, styleTemplateRow As Long, fromRow As Long, toRow As Long)
    Dim r As Long
    For r = fromRow To toRow
        targetSheet.Rows(styleTemplateRow).Copy
        targetSheet.Rows(r).PasteSpecial Paste:=xlPasteFormats
    Next r
    Application.CutCopyMode = False
End Sub
```
```
Call ApplyStyleFromTemplateRow(sheet, 3, 10, 10) ' only row 10
```


- Find row by cell value
```
Function FindMatchingRows(sheet As Worksheet, columnIndex As Long, searchValue As Variant, startRow As Long, endRow As Long) As Collection
    Dim matchedRows As New Collection
    Dim r As Long

    For r = startRow To endRow
        If sheet.Cells(r, columnIndex).Value = searchValue Then
            matchedRows.Add r ' Store the row number
        End If
    Next r

    Set FindMatchingRows = matchedRows
End Function
```
- Usage
```
Sub HighlightMatchingRows()
    Dim matches As Collection
    Set matches = FindMatchingRows(Sheets("Sheet1"), 2, "John", 5, 100)

    Dim r As Variant
    For Each r In matches
        Sheets("Sheet1").Rows(r).Interior.Color = RGB(255, 255, 150) ' Highlight yellow
    Next r
End Sub
```

```
Sub ApplyStyleToMatchingRows()
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("Sheet1")

    Dim searchColumn As Long: searchColumn = 2 ' Column B
    Dim searchValue As Variant: searchValue = "John"
    Dim startRow As Long: startRow = 5
    Dim endRow As Long: endRow = 100
    Dim templateRow As Long: templateRow = 3

    ' 1. Find matching rows
    Dim matches As Collection
    Set matches = FindMatchingRows(sheet, searchColumn, searchValue, startRow, endRow)

    ' 2. Apply style to each matched row
    Dim r As Variant
    For Each r In matches
        ApplyStyleFromTemplateRow sheet, templateRow, r, r
    Next r
End Sub

```
- FindMatchingRowsByCriteria
```
Function FindMatchingRowsByCriteria(sheet As Worksheet, _
                                     startRow As Long, endRow As Long, _
                                     criteria As Object) As Collection
    Dim matchedRows As New Collection
    Dim r As Long
    Dim key As Variant
    Dim isMatch As Boolean

    For r = startRow To endRow
        isMatch = True
        For Each key In criteria.Keys
            If sheet.Cells(r, key).Value <> criteria(key) Then
                isMatch = False
                Exit For
            End If
        Next key
        If isMatch Then matchedRows.Add r
    Next r

    Set FindMatchingRowsByCriteria = matchedRows
End Function

```
```
Sub ApplyStyleToMatchingRows_MultiCriteria()
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("Sheet1")

    Dim criteria As Object
    Set criteria = CreateObject("Scripting.Dictionary")
    criteria.Add 2, "John"     ' Column B
    criteria.Add 3, "Active"   ' Column C

    Dim matchedRows As Collection
    Set matchedRows = FindMatchingRowsByCriteria(sheet, 5, 100, criteria)

    Dim r As Variant
    For Each r In matchedRows
        ApplyStyleFromTemplateRow sheet, 3, r, r
    Next r
End Sub

```

- MergeColumnCells
```
Sub MergeColumnCells(sheet As Worksheet, columnIndex As Long, fromRow As Long, toRow As Long)
    If fromRow >= toRow Then Exit Sub
    sheet.Range(sheet.Cells(fromRow, columnIndex), sheet.Cells(toRow, columnIndex)).Merge
End Sub
```
```
Sub TestMerge()
    ' Merge cells in column B (2) from row 5 to row 10
    MergeColumnCells ThisWorkbook.Sheets("Sheet1"), 2, 5, 10
End Sub
```