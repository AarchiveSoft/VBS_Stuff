Public customerDict As Object
Public cellAddressDict As Object

Public Sub RentChecker()
    ' Initialize dictionaries
    Set customerDict = CreateObject("Scripting.Dictionary")
    Set cellAddressDict = CreateObject("Scripting.Dictionary")
    
    If Not IsDate(ActiveCell.Value) Then
        MsgBox "Falsche Zelle, bitte wähle eine Zelle mit einem Datum.", vbExclamation, "Ungültige Auswahl"
        Exit Sub
    End If

    Dim dateText As Date
    dateText = ActiveCell.Value

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, "Miete", vbTextCompare) > 0 Then
            ProcessWorkbook wb, dateText
        End If
    Next wb

    If customerDict.Count = 0 Then
        MsgBox "CustomerDict is not populated.", vbCritical
        Exit Sub
    End If

    SortCustomerDictByDate
    DisplayUserForm

    Set customerDict = Nothing
    Set cellAddressDict = Nothing
    Set wb = Nothing
End Sub

Private Sub ProcessWorkbook(ByRef wb As Workbook, ByVal dateText As Date)
    Dim sheet As Worksheet
    For Each sheet In wb.Sheets
        ProcessSheet sheet, dateText
    Next sheet
End Sub

Private Sub ProcessSheet(ByRef sheet As Worksheet, ByVal dateText As Date)
    Dim i As Long, j As Long
    Dim startCol As Long, lastCol As Long

    ' Check if the workbook name contains 'Profoto'
    If InStr(1, sheet.Parent.Name, "Profoto", vbTextCompare) > 0 Then
        ' New and fast way for 'Profoto' workbooks
        startCol = sheet.Range("D1").Column  ' Start checking from column D
        lastCol = sheet.Cells(1, Columns.Count).End(xlToLeft).Column ' Last relevant column

        For i = ActiveCell.Row To ActiveCell.Row + 11
            For j = startCol To lastCol
                If IsColumnRelevant(sheet.Cells(1, j)) Then
                    ProcessCell sheet.Cells(i, j), dateText, ActiveCell.Row
                End If
            Next j
        Next i
    Else
        ' Slower but more reliable way for other workbooks
        startCol = sheet.Range("G1").Column  ' Start checking from column G

        For i = ActiveCell.Row To ActiveCell.Row + 11
            For j = 1 To sheet.Columns.Count
                If j >= startCol And IsColumnRelevant(sheet.Cells(1, j)) Then
                    ProcessCell sheet.Cells(i, j), dateText, ActiveCell.Row
                End If
            Next j
        Next i
    End If
End Sub


Private Sub ProcessCell(ByRef cell As Range, ByVal initialDate As Date, ByVal startRow As Long)
    If IsColumnRelevant(cell) And IsValidEntry(cell.Value) Then
        ' Get the range for the entire entry
        Dim entryRange As Range
        Set entryRange = DetermineEntryRange(cell)
        Dim entryStartRow As Long
        entryStartRow = entryRange.Cells(1, 1).Row

        ' Calculate the date offset and update dateText based on entry start
        Dim dateOffset As Integer
        dateOffset = CalculateDateOffset(entryStartRow, startRow)
        Dim currentDate As Date
        currentDate = initialDate + dateOffset
        Dim currentDateText As String
        currentDateText = Format(currentDate, "dd.mm.yyyy")

        ' Determine the time slot based on entry start
        Dim slot As Variant
        slot = DetermineTimeSlot(entryStartRow, startRow)
        
        Dim info As String
        info = GatherEntryInfo(cell, cell.Interior.color)
        
        Dim customerNumber As String
        customerNumber = ExtractCustomerNumber(cell.Value)
        
        Dim customerName As String
        customerName = cell.Offset(1, 0).Value
        
        ' Extracting article data
        Dim articleName1 As String, articleName2 As String, articleName3 As String, articleNumber As String
        articleName1 = cell.Worksheet.Cells(3, cell.Column).Value
        articleName2 = cell.Worksheet.Cells(4, cell.Column).Value
        articleName3 = cell.Worksheet.Cells(5, cell.Column).Value
        articleNumber = cell.Worksheet.Cells(8, cell.Column).Value
    
        Dim commentString As String
        commentString = GetComment(entryRange)

        Dim brandName As String
        brandName = ExtractBrandName(cell.Worksheet.Parent.Name)

        StoreArticleData brandName, currentDateText, customerName, slot, info, _
                         commentString, customerNumber, articleName1, articleName2, articleName3, _
                         articleNumber, cell.Worksheet.Parent.Name, cell.Worksheet.Name, cell.Address
    End If
End Sub

Private Function IsColumnRelevant(ByVal cell As Range) As Boolean
    ' Get the first row cell in the same column as the passed cell
    Dim firstRowCell As Range
    Set firstRowCell = cell.Worksheet.Cells(1, cell.Column)

    ' Check if the first row cell's color is either yellow or green
    IsColumnRelevant = (firstRowCell.Interior.color = RGB(255, 255, 0)) Or (firstRowCell.Interior.color = RGB(0, 255, 0))
End Function

Private Function IsValidEntry(ByVal cellValue As String) As Boolean
    IsValidEntry = cellValue Like "[MmTtLlRr] #####" Or cellValue Like "[MmTtLlRr]#####"
End Function

Private Function DetermineEntryRange(ByVal startCell As Range) As Range
    Dim firstRow As Long, lastRow As Long
    firstRow = startCell.Row
    lastRow = startCell.Row

    ' Extend upwards
    Do While firstRow > 1 And startCell.Worksheet.Cells(firstRow - 1, startCell.Column).Interior.color = startCell.Interior.color
        firstRow = firstRow - 1
    Loop

    ' Extend downwards
    Do While startCell.Worksheet.Cells(lastRow + 1, startCell.Column).Interior.color = startCell.Interior.color
        lastRow = lastRow + 1
    Loop

    Set DetermineEntryRange = startCell.Worksheet.Range(startCell.Worksheet.Cells(firstRow, startCell.Column), startCell.Worksheet.Cells(lastRow, startCell.Column))
End Function

Private Function DetermineTimeSlot(ByVal currentRow As Long, ByVal startRow As Long) As Integer
    Dim rowWithinWeek As Integer
    Dim weekdayRows As Integer

    weekdayRows = 4 ' Number of rows for weekdays
    rowWithinWeek = (currentRow - startRow) Mod (weekdayRows * 5 + 2) ' +2 for the weekend single rows

    ' Determine time slot based on the row within the week
    If rowWithinWeek < weekdayRows * 5 Then ' It's a weekday
        DetermineTimeSlot = (rowWithinWeek Mod weekdayRows) + 1
    Else ' It's a weekend, no time slot
        DetermineTimeSlot = 0 ' Assuming 0 means no time slot for weekends
    End If
End Function

Private Function GatherEntryInfo(ByVal cell As Range, ByVal color As Long) As String
    Dim info As String, aboveCell As Range, belowCell As Range
    info = vbNewLine & vbTab & cell.Value

    ' Check above active cell, if same color, add text to info
    Set aboveCell = cell.Offset(-1, 0)
    Do While aboveCell.Interior.color = color
        If Trim(aboveCell.Value) <> "" Then
            info = info & vbNewLine & vbTab & aboveCell.Value
        End If
        Set aboveCell = aboveCell.Offset(-1, 0)
    Loop

    ' Same below active cell
    Set belowCell = cell.Offset(1, 0)
    Do While belowCell.Interior.color = color
        If Trim(belowCell.Value) <> "" Then
            info = info & vbNewLine & vbTab & belowCell.Value
        End If
        Set belowCell = belowCell.Offset(1, 0)
    Loop

    ' Filter out empty lines
    info = FilterOutEmptyLines(info)
    GatherEntryInfo = info
End Function

Private Function GetComment(entryRange As Range) As String
    Dim cell As Range
    For Each cell In entryRange
        If Not cell.Comment Is Nothing Then
            GetComment = cell.Comment.text
            Exit Function
        End If
    Next cell
    GetComment = "Kein Kommentar gefunden" & vbNewLine
End Function

Private Function ExtractCustomerNumber(ByVal cellValue As String) As String
    If cellValue Like "[MmTtLlRr]#####" Then
        ExtractCustomerNumber = Trim(Mid(cellValue, 2, 5))
    Else
        ExtractCustomerNumber = Trim(Mid(cellValue, InStr(cellValue, " ") + 1, 5))
    End If
End Function

Private Function GetArticleData(ByVal cell As Range) As String
    ' Assuming article names are in rows 3 to 5 and article number is in row 8 in the same column as the cell
    Dim articleName1 As String
    Dim articleName2 As String
    Dim articleName3 As String
    Dim articleNumber As String

    articleName1 = cell.Worksheet.Cells(3, cell.Column).Value
    articleName2 = cell.Worksheet.Cells(4, cell.Column).Value
    articleName3 = cell.Worksheet.Cells(5, cell.Column).Value
    articleNumber = cell.Worksheet.Cells(8, cell.Column).Value

    ' Concatenate article data into a single string
    GetArticleData = articleName1 & " " & articleName2 & " " & articleName3 & " (" & articleNumber & ")"
End Function

Private Function FilterOutEmptyLines(ByVal info As String) As String
    Dim lines() As String, line As Variant, filteredInfo As String
    lines = Split(info, vbLf)
    filteredInfo = ""

    For Each line In lines
        line = Trim(line)
        If line <> "" Then
            filteredInfo = filteredInfo & vbLf & line
        End If
    Next line
    FilterOutEmptyLines = filteredInfo
End Function

Private Function ExtractBrandName(ByVal workbookName As String) As String
    Dim startOfBrand As Integer
    startOfBrand = InStr(workbookName, "Miete") + Len("Miete") + 1
    Dim endOfBrand As Integer
    endOfBrand = InStrRev(workbookName, ".") - 1
    If endOfBrand <= 0 Then ' In case there's no dot in the name
        endOfBrand = Len(workbookName)
    End If
    ExtractBrandName = "Miete " & Trim(Mid(workbookName, startOfBrand, endOfBrand - startOfBrand + 1))
End Function

Private Sub StoreArticleData(ByVal brandName As String, ByVal dateText As String, _
                             ByVal customerName As String, ByVal slot As Variant, _
                             ByVal info As String, ByVal commentString As String, _
                             ByVal customerNumber As String, ByVal articleName1 As String, _
                             ByVal articleName2 As String, ByVal articleName3 As String, _
                             ByVal articleNumber As String, ByVal workbookName As String, _
                             ByVal sheetName As String, ByVal cellAddress As String)
    
    Dim articleInfo As String
    articleInfo = Chr(10) & Chr(9) & "- " & articleName1 & " " & articleName2 & " " & articleName3 & " (" & articleNumber & ")"

    Dim key As String
    key = brandName & " / " & dateText & " / " & customerName

    If Not customerDict.Exists(key) Then
        Dim fullEntry As String
        fullEntry = brandName & " " & vbNewLine & vbNewLine & vbTab & dateText & " " & vbNewLine & vbNewLine
        fullEntry = fullEntry & vbTab & "Zeitslot (1 - 4): " & slot & vbNewLine & vbNewLine & vbTab
        fullEntry = fullEntry & "Original Eintrag: " & info & vbNewLine & vbNewLine & vbTab & "Kommentar: " & commentString
        fullEntry = fullEntry & vbNewLine & vbTab & customerName & " (KN: " & customerNumber & ")"
        fullEntry = fullEntry & articleInfo
        customerDict.Add key, fullEntry
        cellAddressDict.Add key, workbookName & "," & sheetName & "," & cellAddress
    Else
        ' Append only new article data to existing entry
        customerDict(key) = customerDict(key) & articleInfo
        cellAddressDict(key) = cellAddressDict(key) & "," & workbookName & "," & cellAddress
    End If
End Sub

Private Function CalculateDateOffset(ByVal currentRow As Long, ByVal startRow As Long) As Integer
    Dim dateOffset As Integer
    Dim weekdayRows As Integer
    Dim weekendRows As Integer
    Dim totalRows As Integer
    Dim rowDifference As Integer

    weekdayRows = 4 ' Number of rows for weekdays
    weekendRows = 1 ' Number of rows for weekends (Saturday and Sunday)
    totalRows = weekdayRows * 5 + weekendRows * 2 ' Total rows for a full week

    rowDifference = currentRow - startRow ' Difference in rows from the start

    ' Calculate offset for weekdays
    If rowDifference < weekdayRows * 5 Then
        dateOffset = rowDifference \ weekdayRows
    Else ' Calculate offset for weekends and following weekdays
        dateOffset = 5 + (rowDifference - weekdayRows * 5) \ (weekendRows + weekdayRows)
    End If

    CalculateDateOffset = dateOffset
End Function

Private Sub SortCustomerDictByDate()
    Dim sortedDict As Object
    Set sortedDict = CreateObject("Scripting.Dictionary")

    Dim dictKeys() As Variant
    Dim dictLength As Integer
    dictLength = customerDict.Count
    ReDim dictKeys(1 To dictLength)

    ' Copy the keys to an array
    Dim i As Integer
    i = 1
    For Each key In customerDict.Keys
        dictKeys(i) = key
        i = i + 1
    Next key

    ' Sort the array of keys by date
    Dim j As Integer
    Dim tempKey As Variant
    For i = 1 To dictLength - 1
        For j = i + 1 To dictLength
            If ExtractDateFromKey(dictKeys(j)) < ExtractDateFromKey(dictKeys(i)) Then
                ' Swap the keys if they're out of order
                tempKey = dictKeys(i)
                dictKeys(i) = dictKeys(j)
                dictKeys(j) = tempKey
            End If
        Next j
    Next i

    ' Rebuild the dictionary with sorted keys
    For i = 1 To dictLength
        sortedDict.Add dictKeys(i), customerDict(dictKeys(i))
    Next i

    ' Replace the original customerDict with the sorted dictionary
    Set customerDict = sortedDict
End Sub

Private Function ExtractDateFromKey(ByVal key As String) As Date
    ' Assuming the date is the second part of the key
    Dim parts() As String
    parts = Split(key, " / ")
    If UBound(parts) >= 1 Then
        Dim dateParts() As String
        dateParts = Split(parts(1), ".")
        If UBound(dateParts) = 2 Then
            ' Extract day, month, and year
            Dim day As Integer, month As Integer, year As Integer
            day = CInt(dateParts(0))
            month = CInt(dateParts(1))
            year = CInt(dateParts(2))
            ExtractDateFromKey = DateSerial(year, month, day)
        Else
            ExtractDateFromKey = Date ' Default to current date if parsing fails
        End If
    Else
        ExtractDateFromKey = Date ' Default to current date if parsing fails
    End If
End Function

Private Sub DisplayUserForm()
    ' Check if customerDict is set
    If customerDict Is Nothing Then
        MsgBox "customerDict is not initialized.", vbCritical
        Exit Sub
    End If

    ' Check if customerDict has items
    If customerDict.Count = 0 Then
        MsgBox "customerDict is empty.", vbCritical
        Exit Sub
    End If

    ' Initialize and load data into UserForm1
    Dim uf As UserForm1
    Set uf = New UserForm1

    ' Ensure uf is properly initialized
    If Not uf Is Nothing Then
        uf.LoadData customerDict
        Set uf.cellAddressDict = cellAddressDict
        uf.Show vbModeless
    Else
        MsgBox "Failed to initialize UserForm1.", vbCritical
    End If
End Sub






