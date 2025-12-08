Option Explicit

Sub Import_AGTA()
    Dim wbThis As Workbook
    Dim wbSource As Workbook
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet
    Dim fd As FileDialog
    Dim filePath As String
    Dim defaultPath As String
    
    ' 🔧 Set your default folder here
    defaultPath = "C:\YOUR\FOLDER\PATH\"   ' <-- Replace this with desired folder
    
    On Error GoTo ErrHandler
    
    Set wbThis = ThisWorkbook
    Set wsTarget = ActiveSheet   ' Or change to a fixed sheet if needed
    
    ' --- File Picker ---
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select workbook to import"
        .InitialFileName = defaultPath   ' Open in your preferred folder
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = False
        
        If .Show <> -1 Then Exit Sub  ' User cancelled
        filePath = .SelectedItems(1)
    End With
    
    ' --- Optimize for speed ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' --- Open the source workbook ---
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
    
    ' --- Always use the FIRST sheet ---
    Set wsSource = wbSource.Sheets(1)
    
    ' --- Clear existing data in target sheet ---
    wsTarget.Cells.Clear
    
    ' --- Copy full used range ---
    wsSource.UsedRange.Copy Destination:=wsTarget.Range("A1")
    
    ' --- Close source workbook ---
    wbSource.Close SaveChanges:=False
    
CleanExit:
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Error during import: " & Err.Description, vbExclamation, "Import Failed"
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    GoTo CleanExit
End Sub
