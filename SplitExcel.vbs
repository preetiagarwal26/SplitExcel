'sub routine to split xls in smaller files
Sub SplitExcel()
  Dim workbook As Workbook
  Dim CurrSheet As Worksheet
  Dim TotalColumns As Integer 
  Dim RangeToCopy As Range
  Dim RangeOfHeader As Range        'data (range) of header row 
  Dim WorkbookCounter As Integer 
  Dim TotalRows                    'total rows including header in new files? 
  
  Application.ScreenUpdating = False 
 
  'Initialize data 
  Set CurrSheet = ThisWorkbook.ActiveSheet
  TotalColumns = CurrSheet.UsedRange.Columns.Count
  WorkbookCounter = 1 
  TotalRows = 1000                   'rows per file 
 
  'Copy the data of the first row (header) 
  Set RangeOfHeader = CurrSheet.Range(CurrSheet.Cells(1, 1), CurrSheet.Cells(1, TotalColumns))
  
  For p = 2 To CurrSheet.UsedRange.Rows.Count Step TotalRows - 1 
    Set workbook = Workbooks.Add
   
  'Paste the header row in new file 
    RangeOfHeader.Copy workbook.Sheets(1).Range("A1")
   
  'Paste the chunk of rows for this file 
    Set RangeToCopy = CurrSheet.Range(CurrSheet.Cells(p, 1), CurrSheet.Cells(p + TotalRows - 2, TotalColumns))
    RangeToCopy.Copy workbook.Sheets(1).Range("A2")
  
  'Save the new workbook, and close it 
    workbook.SaveAs ThisWorkbook.Path & "\file " & WorkbookCounter
    workbook.Close
   
  'Increment file counter 
    WorkbookCounter = WorkbookCounter + 1 
  Next p

  Application.ScreenUpdating = True 
  Set workbook = Nothing
End Sub