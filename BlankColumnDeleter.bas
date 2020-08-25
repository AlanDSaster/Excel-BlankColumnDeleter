Attribute VB_Name = "BlankColumnDeleter"
Option Explicit

Sub DeleteBlankColumnsInSelection()

    DeleteBlankColumnsInRange Selection
    
End Sub

Sub DeleteBlankColumnsInActiveWorksheet()
    
    DeleteBlankColumnsInWorksheet ActiveSheet
    
End Sub
Sub DeleteBlankColumnsInActiveWorkbook()
    
    DeleteBlankColumnsInWorkbook ActiveWorkbook
    
End Sub

Sub DeleteBlankColumnsInWorkbook(wb As Workbook)
    
    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
        
        DeleteBlankColumnsInWorksheet ws
    
    Next ws
    
End Sub

Sub DeleteBlankColumnsInWorksheet(ws As Worksheet)
    
    DeleteBlankColumnsInRange ws.UsedRange
    
End Sub

Sub DeleteBlankColumnsInRange(tr As Range)
    
    Dim col As Range
    Dim nonEmptyCells As Long
    Dim columnIterator As Long
    Dim i As Long
    
    Set tr = Intersect(tr, tr.Parent.UsedRange)
    
    columnIterator = tr.Columns.Count
    
    Do While columnIterator > 0
        
        Set col = tr.Columns(columnIterator).EntireColumn
        
        i = columnIterator
        Application.StatusBar = "Processing - wb: " & col.Parent.Parent.Name & ", sheet: " & col.Parent.Index & "/" & col.Parent.Parent.Sheets.Count & ", cell: " & i & "/" & tr.Columns.Count
        
        DeleteColumnIfBlank col
        
        columnIterator = columnIterator - 1
        
    Loop
    
    Application.StatusBar = False
    
End Sub

Sub DeleteColumnIfBlank(col As Range)
    
    Dim nonEmptyCells As Long
    nonEmptyCells = WorksheetFunction.CountA(col)
    
    If nonEmptyCells = 1 Then col.EntireColumn.Delete
    
End Sub
