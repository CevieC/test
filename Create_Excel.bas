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
    
    ' 🔹 Prepare / create Export_Log sheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Export_Log")
    On Error GoTo ErrHandler
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "Export_Log"
    End If
    
    ' Clear previous log and set headers
    wsLog.Cells.Clear
    wsLog.Range("A1").Value = "GroupCode"
    wsLog.Range("B1").Value = "FileName"
    wsLog.Range("C1").Value = "FullPath"
    logRow = 2     ' first data row
    
    ' 🔹 Last row based on Column B
    lastRow = wsAGTA.Cells(wsAGTA.Rows.Count, "B").End(xlUp).Row
    
    ' 🔹 Dictionary for grouping
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' --- Build groups: only rows where Col P = "FA" AND Col AS = "A" ---
    For r = 2 To lastRow   ' assume row 1 is header
        If wsAGTA.Cells(r, "P").Value = "FA" And wsAGTA.Cells(r, "AS").Value = "A" Then
            key = CStr(wsAGTA.Cells(r, "B").Value)   ' group key: 6-digit number in Col B
            
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
    
    ' --- Output folder (🔧 replace this with your real path) ---
    outputFolder = "C:\YOUR\OUTPUT\FOLDER\"   ' <-- change this later
    
    ' Ensure trailing backslash
    If Right(outputFolder, 1) <> "\" Then
        outputFolder = outputFolder & "\"
    End If
    
    ' Create folder if it does not exist
    If Dir(outputFolder, vbDirectory) = "" Then
        MkDir outputFolder
    End If
    
    ' --- Speed up ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' --- Loop through each group (each unique value in Column B) ---
    For Each dictKey In dict.Keys
        Set rowsCollection = dict(dictKey)
        
        ' New workbook with 1 sheet
        Set wbNew = Workbooks.Add(xlWBATWorksheet)
        Set wsNew = wbNew.Sheets(1)
        
        ' Write headers (F, N, J, M, T in that order, as columns A:E)
        targetRow = 1
        wsNew.Cells(targetRow, 1).Value = wsAGTA.Cells(1, "F").Value
        wsNew.Cells(targetRow, 2).Value = wsAGTA.Cells(1, "N").Value
        wsNew.Cells(targetRow, 3).Value = wsAGTA.Cells(1, "J").Value
        wsNew.Cells(targetRow, 4).Value = wsAGTA.Cells(1, "M").Value
        wsNew.Cells(targetRow, 5).Value = wsAGTA.Cells(1, "T").Value
        
        ' Copy rows for this group (only F, N, J, M, T)
        For i = 1 To rowsCollection.Count
            r = rowsCollection(i)
            targetRow = targetRow + 1
            
            wsNew.Cells(targetRow, 1).Value = wsAGTA.Cells(r, "F").Value
            wsNew.Cells(targetRow, 2).Value = wsAGTA.Cells(r, "N").Value
            wsNew.Cells(targetRow, 3).Value = wsAGTA.Cells(r, "J").Value
            wsNew.Cells(targetRow, 4).Value = wsAGTA.Cells(r, "M").Value
            wsNew.Cells(targetRow, 5).Value = wsAGTA.Cells(r, "T").Value
        Next i
        
        ' 🔽 Sort data in the new file by Column F (which is Column A here)
        If targetRow > 1 Then   ' only sort if there is at least 1 data row
            With wsNew.Sort
                .SortFields.Clear
                .SortFields.Add Key:=wsNew.Range("A2:A" & targetRow), _
                                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .SetRange wsNew.Range("A1:E" & targetRow)
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .Apply
            End With
        End If
        
        ' 🎨 FORMAT THE TABLE
        
        Dim headerRange As Range, dataRange As Range
        
        Set headerRange = wsNew.Range("A1:E1")
        Set dataRange = wsNew.Range("A1:E" & targetRow)
        
        ' Bold header
        headerRange.Font.Bold = True
        
        ' Light grey fill
        headerRange.Interior.Color = RGB(242, 242, 242)
        
        ' Borders for entire table
        With dataRange.Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
            .Weight = xlThin
        End With
        
        ' Autofit columns
        wsNew.Columns("A:E").AutoFit
        
        ' Use name in Column M (from first row in the group) for filename
        nameVal = CStr(wsAGTA.Cells(rowsCollection(1), "M").Value)
        If Len(Trim$(nameVal)) = 0 Then
            nameVal = "Group_" & dictKey
        End If
        
        ' Clean file name and append the group key to avoid clashes
        fileName = CleanFileName(nameVal & "_" & dictKey & ".xlsx")
        filePath = outputFolder & fileName
        
        ' Save as .xlsx
        wbNew.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
        
        ' 🔹 Log this export in Export_Log
        wsLog.Cells(logRow, 1).Value = CStr(dictKey)   ' GroupCode
        wsLog.Cells(logRow, 2).Value = fileName       ' FileName
        wsLog.Cells(logRow, 3).Value = filePath       ' FullPath
        logRow = logRow + 1
    Next dictKey
    
CleanExit:
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Error in Draft_Email: " & Err.Description, vbExclamation, "Error"
    On Error Resume Next
    If Not wbNew Is Nothing Then
        wbNew.Close SaveChanges:=False
    End If
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
