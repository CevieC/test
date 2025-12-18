Sub Generate_Excel()
    Dim wsI As Worksheet, wsS As Worksheet, wsT As Worksheet
    Dim lastIRow As Long, lastSRow As Long, lastTRow As Long
    Dim i As Long, j As Long, k As Long, row As Long
    Dim code As String, aj As String, al As String
    Dim comCode As String, svcCode As String, com6 As String, svc6 As String
    Dim agentCode As String, agentName As String, unitCode As String
    Dim fileName As String, filePath As String, matchType As String
    Dim wb As Workbook, yearFolder As String, dateFolder As String
    Dim matchCom As Boolean, matchSvc As Boolean
    
    Set wsI = ThisWorkbook.Worksheets("INPUT")
    Set wsS = ThisWorkbook.Worksheets("ROP_REPORT_DATA")
    Set wsT = ThisWorkbook.Worksheets("EXCEL_TEMPLATE")
    
    ' Clear and format template
    lastTRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row
    If lastTRow > 1 Then wsT.Range("A2:I" & lastTRow).ClearContents
    wsT.Range("E:E").NumberFormat = "@"
    wsT.Range("G:G").NumberFormat = "@"
    
    lastIRow = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
    lastSRow = wsS.Cells(wsS.Rows.Count, "Q").End(xlUp).Row
    code = Trim(CStr(wsI.Cells(2, "A").Value))
    
    ' Filter and copy data
    row = 2
    For i = 2 To lastIRow
        agentCode = Trim(CStr(wsI.Cells(i, "A").Value))
        If agentCode = "" Or Len(agentCode) <> 6 Then GoTo NextInput
        For j = 2 To lastSRow
            If Trim(CStr(wsS.Cells(j, "Q").Value)) = "Inforce" Then
                aj = Trim(CStr(wsS.Cells(j, "AJ").Text))
                al = Trim(CStr(wsS.Cells(j, "AL").Text))
                matchCom = (Len(aj) >= 6 And Right(aj, 6) = agentCode)
                matchSvc = (Len(al) >= 6 And Right(al, 6) = agentCode)
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
NextInput:
    Next i
    
    ' Find agent info
    lastTRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row
    If lastTRow < 2 Then
        MsgBox "No data found to process."
        Exit Sub
    End If
    
    code = Trim(CStr(wsI.Cells(2, "A").Value))
    comCode = Trim(CStr(wsT.Cells(2, "E").Text))
    svcCode = Trim(CStr(wsT.Cells(2, "G").Text))
    
    If Len(comCode) >= 6 And Right(comCode, 6) = code Then
        agentCode = comCode
        agentName = Trim(CStr(wsT.Cells(2, "F").Value))
    ElseIf Len(svcCode) >= 6 And Right(svcCode, 6) = code Then
        agentCode = svcCode
        agentName = Trim(CStr(wsT.Cells(2, "H").Value))
    End If
    
    If Len(agentCode) < 6 Then
        MsgBox "Could not find agent code."
        Exit Sub
    End If
    
    unitCode = Left(agentCode, 6)
    If agentName = "" Then agentName = "Agent"
    
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
                If com6 = code Then
                    wsT.Cells(k, "G").Value = "WITH ANOTHER REP/FIRM"
                    wsT.Cells(k, "H").Value = "WITH ANOTHER REP/FIRM"
                ElseIf svc6 = code Then
                    wsT.Cells(k, "E").Value = "WITH ANOTHER REP/FIRM"
                    wsT.Cells(k, "F").Value = "WITH ANOTHER REP/FIRM"
                End If
            End If
        End If
    Next k
    
    ' Create folders and save
    yearFolder = ThisWorkbook.Path & "\" & Year(Now())
    dateFolder = yearFolder & "\" & Format(Now(), "DD-MMM")
    On Error Resume Next
    If Dir(yearFolder, vbDirectory) = "" Then MkDir yearFolder
    If Dir(dateFolder, vbDirectory) = "" Then MkDir dateFolder
    On Error GoTo 0
    
    fileName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(agentName, "/", "-"), "\", "-"), ":", "-"), "*", "-"), "?", "-"), """", "-"), "<", "-"), ">", "-"), "|", "-")
    fileName = fileName & "_" & code & ".xlsx"
    filePath = dateFolder & "\" & fileName
    On Error Resume Next
    Kill filePath
    On Error GoTo 0
    
    Set wb = Workbooks.Add
    wsT.Copy Before:=wb.Sheets(1)
    Application.DisplayAlerts = False
    wb.Sheets(2).Delete
    wb.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook, Password:=unitCode
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    MsgBox "Complete! " & (row - 2) & " rows copied." & vbCrLf & "File: " & fileName & vbCrLf & "Path: " & filePath & vbCrLf & "Password: " & unitCode
End Sub