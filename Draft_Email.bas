Option Explicit

Sub Create_Email()
    Dim outputFolder As String
    Dim wsMap As Worksheet
    Dim lastRow As Long, r As Long
    Dim mapDict As Object
    Dim code As String
    Dim toEmail As String, ccEmail As String
    Dim entry As Variant
    
    Dim olApp As Object
    Dim olMail As Object
    
    Dim fileName As String
    Dim fullPath As String
    Dim groupCode As String
    
    On Error GoTo ErrHandler
    
    ' 🔹 Folder where Draft_Email saved the files
    outputFolder = "C:\YOUR\OUTPUT\FOLDER\"   ' <-- change to the SAME path as in Draft_Email
    
    If Right(outputFolder, 1) <> "\" Then
        outputFolder = outputFolder & "\"
    End If
    
    ' 🔹 Mapping sheet
    Set wsMap = ThisWorkbook.Sheets("Firm_Email_Data")
    
    ' 🔹 Build dictionary: Code -> (To, CC)
    Set mapDict = CreateObject("Scripting.Dictionary")
    
    lastRow = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow   ' assume header row at 1
        code = Trim$(CStr(wsMap.Cells(r, "A").Value))
        If Len(code) > 0 Then
            If Not mapDict.Exists(code) Then
                toEmail = Trim$(CStr(wsMap.Cells(r, "C").Value))
                ccEmail = Trim$(CStr(wsMap.Cells(r, "D").Value))
                mapDict.Add code, Array(toEmail, ccEmail)
            End If
        End If
    Next r
    
    ' 🔹 Get or create Outlook instance (late binding)
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrHandler
    
    If olApp Is Nothing Then
        MsgBox "Unable to start Outlook. Email drafts will not be created.", vbExclamation, "Outlook Error"
        Exit Sub
    End If
    
    ' 🔹 Loop through all .xlsx files in the folder
    fileName = Dir(outputFolder & "*.xlsx")
    
    Do While Len(fileName) > 0
        fullPath = outputFolder & fileName
        
        ' Extract group code from file name (part after last "_" before ".xlsx")
        groupCode = GetGroupCodeFromFileName(fileName)
        
        ' Default: empty fields
        toEmail = ""
        ccEmail = ""
        
        ' If we have a mapping, fill To / CC
        If Len(groupCode) > 0 Then
            If mapDict.Exists(groupCode) Then
                entry = mapDict(groupCode)
                toEmail = entry(0)
                ccEmail = entry(1)
            End If
        End If
        
        ' 🔹 Create email draft
        Set olMail = olApp.CreateItem(0)   ' 0 = olMailItem
        
        With olMail
            .To = toEmail          ' may be empty if no mapping
            .CC = ccEmail          ' may be empty if no mapping
            .Subject = ""          ' leave blank as requested
            .Body = ""             ' leave blank as requested
            .Attachments.Add fullPath
            .Save                  ' save to Drafts
        End With
        
        ' Next file
        fileName = Dir()
    Loop
    
    MsgBox "Draft emails created for all files in the folder.", vbInformation, "Done"
    
    Exit Sub

ErrHandler:
    MsgBox "Error in create_email: " & Err.Description, vbExclamation, "Error"
End Sub

Private Function GetGroupCodeFromFileName(ByVal fileName As String) As String
    Dim posUnderscore As Long
    Dim posDot As Long
    Dim code As String
    
    posUnderscore = InStrRev(fileName, "_")
    posDot = InStrRev(fileName, ".")
    
    If posUnderscore > 0 And posDot > posUnderscore Then
        code = Mid$(fileName, posUnderscore + 1, posDot - posUnderscore - 1)
    Else
        code = ""
    End If
    
    GetGroupCodeFromFileName = code
End Function
