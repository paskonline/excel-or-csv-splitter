Sub SplitIntoCSV()
    Dim SourceSheet As Worksheet
    Dim NewBook As Workbook
    Dim RowCount As Long
    Dim i As Long, StartRow As Long, EndRow As Long
    Dim ChunkSize As Long
    Dim FolderPath As String
    
    ' SETTINGS
    Set SourceSheet = ThisWorkbook.Sheets(1) ' Assumes data is in the first sheet
    ChunkSize = 999 ' Rows per CSV file
    FolderPath = Application.ActiveWorkbook.Path & "\SplitCSVFiles\"
    
    ' Create folder if it doesn't exist
    If Dir(FolderPath, vbDirectory) = "" Then MkDir FolderPath
    
    ' Get total rows used in Column A
    RowCount = SourceSheet.Cells(SourceSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Loop setup
    StartRow = 2 ' Start at row 2 (assuming Row 1 is headers)
    i = 1
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False ' REQUIRED for CSV to prevent "Save changes?" popups
    
    Do While StartRow <= RowCount
        EndRow = StartRow + ChunkSize - 1
        
        ' Create new workbook
        Set NewBook = Workbooks.Add
        
        ' Copy Headers
        SourceSheet.Rows(1).Copy Destination:=NewBook.Sheets(1).Rows(1)
        
        ' Copy Data Chunk
        ' Check if EndRow exceeds total rows to prevent copying empty space
        If EndRow > RowCount Then EndRow = RowCount
        
        SourceSheet.Rows(StartRow & ":" & EndRow).Copy Destination:=NewBook.Sheets(1).Rows(2)
        
        ' Save as CSV
        ' FileFormat:=xlCSV is the key command here
        NewBook.SaveAs Filename:=FolderPath & "Part_" & i & ".csv", FileFormat:=xlCSV
        
        ' Close
        NewBook.Close SaveChanges:=False
        
        ' Increment
        StartRow = EndRow + 1
        i = i + 1
    Loop
    
    ' Restore settings
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Done! " & (i - 1) & " CSV files saved in 'SplitCSVFiles' folder."
End Sub
