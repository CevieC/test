Option Explicit

Sub Create_Excel()
    Dim wsAGTA As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim dict As Object
    Dim key As String
    Dim rowsCollection As Collection
    Dim i As Long
    Dim outputFolder As String
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim dictKey As Variant
    Dim targetRow As Long
    Dim nameVal As String
    Dim fileName As String
    Dim filePath As String
    
    ' Logging
    Dim wsLog As Worksheet
    Dim logRow As Long
    
    On Error GoTo ErrHandler
    
    ' 🔹 Source sheet
    Set wsAGTA = ThisWorkbook.Sheets("AGTA")
    
    ' 🔹 Prepare Export_Log sheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Export_Log")
    On Error GoTo ErrHandler
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "Export_Log"
    End If
    
    wsLog.Cells.Clear
    wsLog.Range("A1:D1").Value = Array("GroupCode", "FileName", "FullPath", "RowCount")
    logRow = 2
    
    ' 🔹 Last row based on Column B
    lastRow = wsAGTA.Cells(wsAGTA.Rows.Count, "B").End(xlUp).Row
    
    ' 🔹 Dictionary for grouping
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' --- Build groups: only rows where Col P = "FA" AND Col AS = "A" ---
    For r = 2 To lastRow
        If wsAGTA.Cells(r, "P").Value = "FA" And wsAGTA.Cells(r, "AS").Value = "A" Then
            key = CStr(wsAGTA.Cells(r, "B").Value)
            
            If Len(key) > 0 Then
                If Not dict.Exists(key) Then
                    Set rowsCollection = New Collection
                    rowsCollection.Add r
                    Set dict(key) = rowsCollection
                Else
                    Set rowsCollection = dict(key)
                    rowsCollection.Add r
                End If
            End If
        End If
    Next r
    
    ' --- Output folder ---
    outputFolder = "C:\YOUR\OUTPUT\FOLDER\"   ' <<<< change this
    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder
    
    ' --- Speed-up ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' --- Loop groups ---
    For Each dictKey In dict.Keys
        
        Set rowsCollection = dict(dictKey)
        
        ' New workbook
        Set wbNew = Workbooks.Add(xlWBATWorksheet)
        Set wsNew = wbNew.Sheets(1)
        
        ' Write headers (F, N, J, M, T)
        targetRow = 1
        wsNew.Range("A1:E1").Value = Array(wsAGTA.Cells(1, "F").Value, _
                                           wsAGTA.Cells(1, "N").Value, _
                                           wsAGTA.Cells(1, "J").Value, _
                                           wsAGTA.Cells(1, "M").Value, _
                                           wsAGTA.Cells(1, "T").Value)
        
        ' Copy data rows
        For i = 1 To rowsCollection.Count
            r = rowsCollection(i)
            targetRow = targetRow + 1
            wsNew.Cells(targetRow, 1).Value = wsAGTA.Cells(r, "F").Value
            wsNew.Cells(targetRow, 2).Value = wsAGTA.Cells(r, "N").Value
            wsNew.Cells(targetRow, 3).Value = wsAGTA.Cells(r, "J").Value
            wsNew.Cells(targetRow, 4).Value = wsAGTA.Cells(r, "M").Value
            wsNew.Cells(targetRow, 5).Value = wsAGTA.Cells(r, "T").Value
        Next i
        
        ' Sort by Column F (col A)
        If targetRow > 1 Then
            With wsNew.Sort
                .SortFields.Clear
                .SortFields.Add Key:=wsNew.Range("A2:A" & targetRow), Order:=xlAscending
                .SetRange wsNew.Range("A1:E" & targetRow)
                .Header = xlYes
                .Apply
            End With
        End If
        
        ' 🎨 Format table in each output workbook
        Dim headerRange As Range, dataRange As Range
        Set headerRange = wsNew.Range("A1:E1")
        Set dataRange = wsNew.Range("A1:E" & targetRow)
        
        headerRange.Font.Bold = True
        headerRange.Interior.Color = RGB(242, 242, 242)
        
        With dataRange.Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
            .Weight = xlThin
        End With
        
        wsNew.Columns("A:E").AutoFit
        
        ' Filename
        nameVal = CStr(wsAGTA.Cells(rowsCollection(1), "M").Value)
        If Len(Trim$(nameVal)) = 0 Then nameVal = "Group_" & dictKey
        
        fileName = CleanFileName(nameVal & "_" & dictKey & ".xlsx")
        filePath = outputFolder & fileName
        
        wbNew.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
        wbNew.Close False
        
        ' --- Log to Export_Log ---
        ' RowCount = number of data rows (exclude header)
        Dim rowCount As Long
        rowCount = targetRow - 1
        
        wsLog.Cells(logRow, 1).Value = CStr(dictKey)   ' GroupCode
        wsLog.Cells(logRow, 2).Value = fileName       ' FileName
        wsLog.Cells(logRow, 3).Value = filePath       ' FullPath (will be hyperlinked)
        wsLog.Cells(logRow, 4).Value = rowCount       ' RowCount
        
        ' Make the FullPath cell a clickable hyperlink
        wsLog.Hyperlinks.Add Anchor:=wsLog.Cells(logRow, 3), _
                              Address:=filePath, _
                              TextToDisplay:=filePath
        
        logRow = logRow + 1
    Next dictKey
    
    ' --- Format Export_Log sheet ---
    With wsLog
        Dim lastLogRow As Long, lastLogCol As Long
        lastLogRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        lastLogCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        Dim logRange As Range
        Set logRange = .Range(.Cells(1, 1), .Cells(lastLogRow, lastLogCol))
        
        ' Header formatting
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Interior.Color = RGB(242, 242, 242)
        
        ' Borders
        With logRange.Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
            .Weight = xlThin
        End With
        
        ' Autofit all used columns
        .Columns("A:D").AutoFit
        
        ' Freeze header
        .Activate
        ActiveWindow.SplitRow = 1
        ActiveWindow.FreezePanes = True
    End With

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Error in Create_Excel: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close False
    GoTo CleanExit
End Sub

Private Function CleanFileName(ByVal fileName As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For Each ch In badChars
        fileName = Replace(fileName, ch, "_")
    Next ch
    
    fileName = Trim$(fileName)
    If fileName = "" Then fileName = "output"
    
    CleanFileName = fileName
End Function

