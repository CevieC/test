Sub FilterAndCopyData()
    Dim wsI As Worksheet, wsS As Worksheet, wsT As Worksheet
    Dim lastIRow As Long, lastSRow As Long, lastTRow As Long
    Dim i As Long, j As Long, k As Long, row As Long
    Dim inputCode As String, aj As String, al As String
    Dim comCode As String, svcCode As String, com6 As String, svc6 As String
    Dim agentCode As String, agentName As String, unitCode As String
    Dim fileName As String, filePath As String, matchType As String
    Dim wb As Workbook, yearFolder As String, dateFolder As String
    Dim matchCom As Boolean, matchSvc As Boolean
    Dim filesCreated As Long, skippedCodes As String
    Dim labelId As String, labelName As String
    
    ' CONFIGURE YOUR SENSITIVITY LABEL HERE
    ' Option 1: If you know the exact label name your org uses
    labelName = "Confidential" ' Change to your org's label (e.g., "Internal", "Public", "Highly Confidential")
    labelId = "" ' Leave empty if you only know the name
    
    ' Option 2: If your IT admin provided a label GUID/ID
    ' labelId = "your-label-guid-here"
    ' labelName = "Your Label Name"
    
    Set wsI = ThisWorkbook.Worksheets("INPUT")
    Set wsS = ThisWorkbook.Worksheets("ROP_REPORT_DATA")
    Set wsT = ThisWorkbook.Worksheets("EXCEL_TEMPLATE")
    
    lastIRow = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
    lastSRow = wsS.Cells(wsS.Rows.Count, "Q").End(xlUp).Row
    filesCreated = 0
    skippedCodes = ""
    
    ' Process each input code separately
    For i = 2 To lastIRow
        inputCode = Trim(CStr(wsI.Cells(i, "A").Value))
        
        If inputCode = "" Then
            skippedCodes = skippedCodes & "Row " & i & ": Empty" & vbCrLf
            GoTo NextCode
        End If
        
        If Len(inputCode) <> 6 Then
            skippedCodes = skippedCodes & "Row " & i & ": " & inputCode & " (not 6 digits)" & vbCrLf
            GoTo NextCode
        End If
        
        ' Find agent of interest code and name first (scan source data)
        agentCode = ""
        agentName = ""
        For j = 2 To lastSRow
            If Trim(CStr(wsS.Cells(j, "Q").Value)) = "Inforce" Then
                aj = Trim(CStr(wsS.Cells(j, "AJ").Text))
                al = Trim(CStr(wsS.Cells(j, "AL").Text))
                If Len(aj) >= 6 And Right(aj, 6) = inputCode Then
                    agentCode = aj
                    agentName = Trim(CStr(wsS.Cells(j, "AK").Value))
                    Exit For
                ElseIf Len(al) >= 6 And Right(al, 6) = inputCode Then
                    agentCode = al
                    agentName = Trim(CStr(wsS.Cells(j, "AM").Value))
                    Exit For
                End If
            End If
        Next j
        
        If agentCode = "" Or Len(agentCode) < 6 Then
            skippedCodes = skippedCodes & "Row " & i & ": " & inputCode & " (no match in source)" & vbCrLf
            GoTo NextCode
        End If
        
        unitCode = Left(agentCode, 6)
        If agentName = "" Then agentName = "Agent"
        
        ' Clear template for this code
        lastTRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row
        If lastTRow > 1 Then wsT.Range("A2:I" & lastTRow).ClearContents
        wsT.Range("E:E").NumberFormat = "@"
        wsT.Range("G:G").NumberFormat = "@"
        
        ' Filter and copy data for this input code
        row = 2
        For j = 2 To lastSRow
            If Trim(CStr(wsS.Cells(j, "Q").Value)) = "Inforce" Then
                aj = Trim(CStr(wsS.Cells(j, "AJ").Text))
                al = Trim(CStr(wsS.Cells(j, "AL").Text))
                matchCom = (Len(aj) >= 6 And Right(aj, 6) = inputCode)
                matchSvc = (Len(al) >= 6 And Right(al, 6) = inputCode)
                
                If matchCom Or matchSvc Then
                    wsT.Cells(row, "A").Value = wsS.Cells(j, "B").Value
                    wsT.Cells(row, "B").Value = wsS.Cells(j, "E").Value
                    wsT.Cells(row, "C").Value = wsS.Cells(j, "J").Value
                    wsT.Cells(row, "D").Value = wsS.Cells(j, "Q").Value
                    wsT.Cells(row, "E").Value = CStr(wsS.Cells(j, "AJ").Text)
                    wsT.Cells(row, "F").Value = wsS.Cells(j, "AK").Value
                    wsT.Cells(row, "G").Value = CStr(wsS.Cells(j, "AL").Text)
                    wsT.Cells(row, "H").Value = wsS.Cells(j, "AM").Value
                    If matchCom And matchSvc Then
                        matchType = "Servicing & Commission"
                    ElseIf matchCom Then
                        matchType = "Commission Only"
                    Else
                        matchType = "Servicing Only"
                    End If
                    wsT.Cells(row, "I").Value = matchType
                    row = row + 1
                End If
            End If
        Next j
        
        If row = 2 Then
            skippedCodes = skippedCodes & "Row " & i & ": " & inputCode & " (no matching rows)" & vbCrLf
            GoTo NextCode
        End If
        
        lastTRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row
        
        ' Censor different unit agents
        For k = 2 To lastTRow
            comCode = Trim(CStr(wsT.Cells(k, "E").Text))
            svcCode = Trim(CStr(wsT.Cells(k, "G").Text))
            If Len(comCode) >= 6 And Len(svcCode) >= 6 And Left(comCode, 6) <> Left(svcCode, 6) Then
                matchType = Trim(CStr(wsT.Cells(k, "I").Value))
                com6 = Right(comCode, 6)
                svc6 = Right(svcCode, 6)
                
                If matchType = "Commission Only" Then
                    wsT.Cells(k, "G").Value = "WITH ANOTHER REP/FIRM"
                    wsT.Cells(k, "H").Value = "WITH ANOTHER REP/FIRM"
                ElseIf matchType = "Servicing Only" Then
                    wsT.Cells(k, "E").Value = "WITH ANOTHER REP/FIRM"
                    wsT.Cells(k, "F").Value = "WITH ANOTHER REP/FIRM"
                ElseIf matchType = "Servicing & Commission" Then
                    If com6 = inputCode Then
                        wsT.Cells(k, "G").Value = "WITH ANOTHER REP/FIRM"
                        wsT.Cells(k, "H").Value = "WITH ANOTHER REP/FIRM"
                    ElseIf svc6 = inputCode Then
                        wsT.Cells(k, "E").Value = "WITH ANOTHER REP/FIRM"
                        wsT.Cells(k, "F").Value = "WITH ANOTHER REP/FIRM"
                    End If
                End If
            End If
        Next k
        
        ' Create folders and save file
        yearFolder = ThisWorkbook.Path & "\" & Year(Now())
        dateFolder = yearFolder & "\" & Format(Now(), "DD-MMM")
        On Error Resume Next
        If Dir(yearFolder, vbDirectory) = "" Then MkDir yearFolder
        If Dir(dateFolder, vbDirectory) = "" Then MkDir dateFolder
        On Error GoTo 0
        
        fileName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(agentName, "/", "-"), "\", "-"), ":", "-"), "*", "-"), "?", "-"), """", "-"), "<", "-"), ">", "-"), "|", "-")
        fileName = fileName & "_" & inputCode & ".xlsx"
        filePath = dateFolder & "\" & fileName
        On Error Resume Next
        Kill filePath
        On Error GoTo 0
        
        Set wb = Workbooks.Add
        wsT.Copy Before:=wb.Sheets(1)
        Application.DisplayAlerts = False
        wb.Sheets(2).Delete
        
        ' Save file with password
        wb.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook, Password:=unitCode
        
        ' Apply sensitivity label after saving
        On Error Resume Next
        If labelId <> "" And labelName <> "" Then
            wb.SetSensitivityLabel LabelId:=labelId, LabelName:=labelName
        ElseIf labelName <> "" Then
            wb.SetSensitivityLabel LabelName:=labelName
        End If
        If Err.Number <> 0 Then
            Debug.Print "Warning: Could not set sensitivity label. Error: " & Err.Description
        End If
        On Error GoTo 0
        
        wb.Close SaveChanges:=True
        Application.DisplayAlerts = True
        filesCreated = filesCreated + 1
NextCode:
    Next i
    
    Dim msg As String
    msg = "Processing complete! " & filesCreated & " file(s) created."
    If skippedCodes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Skipped codes:" & vbCrLf & skippedCodes
    End If
    MsgBox msg
End Sub