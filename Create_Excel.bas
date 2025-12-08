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
    
    On Error GoTo ErrHandler
    
    ' 🔹 Source sheet with imported data
    Set wsAGTA = ThisWorkbook.Sheets("AGTA")
    
    ' 🔹 Last row based on Column B
    lastRow = wsAGTA.Cells(wsAGTA.Rows.Count, "B").End(xlUp).Row
    
    ' 🔹 Dictionary for grouping rows by Column B
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
        
        ' Use name in Column M (from first row in the group) for filename
        nameVal = CStr(wsAGTA.Cells(rowsCollection(1), "M").Value)
        If Len(Trim$(nameVal)) = 0 Then
            nameVal = "Group_" & dictKey
        End If
        
        ' Clean file name and append the group key to avoid clashes
        fileName = CleanFileName(nameVal & "_" & dictKey & ".xlsx")
        filePath = outputFolder & fileName
        
        ' Save as .xlsx and close
        wbNew.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
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
    Dim illegalChars As Variant
    Dim ch As Variant
    
    illegalChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For Each ch In illegalChars
        fileName = Replace(fileName, ch, "_")
    Next ch
    
    fileName = Trim$(fileName)
    If fileName = "" Then
        fileName = "output"
    End If
    
    CleanFileName = fileName
End Function

