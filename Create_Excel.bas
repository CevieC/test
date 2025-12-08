Option Explicit

Sub Draft_Email()
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
    
    ' 🔹 Source sheet
    Set wsAGTA = ThisWorkbook.Sheets("AGTA")
    
    lastRow = wsAGTA.Cells(wsAGTA.Rows.Count, "B").End(xlUp).Row
    
    ' 🔹 Dictionary for grouping
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' --- Build groups ---
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
    
    
    ' --- Output folder (placeholder) ---
    outputFolder = "C:\YOUR\OUTPUT\FOLDER\"   ' <-- update this
    
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
        
        ' Write headers
        targetRow = 1
        wsNew.Cells(1, 1).Value = wsAGTA.Cells(1, "F").Value
        wsNew.Cells(1, 2).Value = wsAGTA.Cells(1, "N").Value
        wsNew.Cells(1, 3).Value = wsAGTA.Cells(1, "J").Value
        wsNew.Cells(1, 4).Value = wsAGTA.Cells(1, "M").Value
        wsNew.Cells(1, 5).Value = wsAGTA.Cells(1, "T").Value
        
        ' Copy rows
        For i = 1 To rowsCollection.Count
            r = rowsCollection(i)
            targetRow = targetRow + 1
            
            wsNew.Cells(targetRow, 1).Value = wsAGTA.Cells(r, "F").Value
            wsNew.Cells(targetRow, 2).Value = wsAGTA.Cells(r, "N").Value
            wsNew.Cells(targetRow, 3).Value = wsAGTA.Cells(r, "J").Value
            wsNew.Cells(targetRow, 4).Value = wsAGTA.Cells(r, "M").Value
            wsNew.Cells(targetRow, 5).Value = wsAGTA.Cells(r, "T").Value
        Next i
        
        
        ' 🔽 Sort by Column F (col A)
        If targetRow > 1 Then
            With wsNew.Sort
                .SortFields.Clear
                .SortFields.Add Key:=wsNew.Range("A2:A" & targetRow), _
                                Order:=xlAscending
                .SetRange wsNew.Range("A1:E" & targetRow)
                .Header = xlYes
                .Apply
            End With
        End If
        
        
        ' 🎨 FORMAT THE TABLE -----------------------------------
        
        Dim headerRange As Range, dataRange As Range
        
        Set headerRange = wsNew.Range("A1:E1")
        Set dataRange = wsNew.Range("A1:E" & targetRow)
        
        ' Bold header
        headerRange.Font.Bold = True
        
        ' Light grey fill (Excel color index 15)
        headerRange.Interior.Color = RGB(242, 242, 242)
        
        ' Borders for entire table
        With dataRange.Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
            .Weight = xlThin
        End With
        
        ' Autofit columns
        wsNew.Columns("A:E").AutoFit
        
        ' --------------------------------------------------------
        
        
        ' Filename
        nameVal = CStr(wsAGTA.Cells(rowsCollection(1), "M").Value)
        If Len(Trim$(nameVal)) = 0 Then nameVal = "Group_" & dictKey
        
        fileName = CleanFileName(nameVal & "_" & dictKey & ".xlsx")
        filePath = outputFolder & fileName
        
        ' Save
        wbNew.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
        wbNew.Close False
    Next dictKey
    
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    

ErrHandler:
    MsgBox "Error in Draft_Email: " & Err.Description, vbExclamation
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

