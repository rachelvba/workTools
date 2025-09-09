Attribute VB_Name = "DataImportExport"
'*******************************************************************************
' Data Import/Export Module
' Description: Template for handling data import and export operations
' Author: Finance Team
' Date: September 9, 2025
'*******************************************************************************

Option Explicit

'*******************************************************************************
' Public Functions for Data Import
'*******************************************************************************
Public Function ImportCSVFile(filePath As String, targetWorksheet As Worksheet, _
                            Optional startRow As Long = 1, _
                            Optional startColumn As Long = 1, _
                            Optional hasHeader As Boolean = True) As Boolean
    ' Imports data from a CSV file into the specified worksheet
    ' Parameters:
    '   filePath - Full path to the CSV file
    '   targetWorksheet - Worksheet to import data into
    '   startRow - Starting row for import (default: 1)
    '   startColumn - Starting column for import (default: 1)
    '   hasHeader - Whether the CSV has a header row (default: True)
    
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim dataLine As String
    Dim dataValues() As String
    Dim row As Long
    Dim col As Long
    Dim i As Long
    
    ' Get the next available file number
    fileNum = FreeFile
    
    ' Open the CSV file
    Open filePath For Input As #fileNum
    
    row = startRow
    
    ' Process the file line by line
    Do Until EOF(fileNum)
        Line Input #fileNum, dataLine
        dataValues = Split(dataLine, ",")
        
        For i = 0 To UBound(dataValues)
            col = startColumn + i
            targetWorksheet.Cells(row, col).Value = Trim(dataValues(i))
        Next i
        
        row = row + 1
    Loop
    
    Close #fileNum
    ImportCSVFile = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ImportCSVFile = False
End Function

'*******************************************************************************
' Public Functions for Data Export
'*******************************************************************************
Public Function ExportRangeToCSV(sourceRange As Range, filePath As String, _
                               Optional includeHeaders As Boolean = True) As Boolean
    ' Exports a range to a CSV file
    ' Parameters:
    '   sourceRange - Range to export
    '   filePath - Full path for the CSV file to be created
    '   includeHeaders - Whether to include headers (default: True)
    
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim dataLine As String
    Dim row As Long
    Dim col As Long
    Dim startRow As Long
    
    ' Get the next available file number
    fileNum = FreeFile
    
    ' Open/create the CSV file
    Open filePath For Output As #fileNum
    
    ' Determine the starting row
    If includeHeaders Then
        startRow = 1
    Else
        startRow = 2
    End If
    
    ' Process each row in the range
    For row = startRow To sourceRange.Rows.Count
        dataLine = ""
        
        ' Process each column in the row
        For col = 1 To sourceRange.Columns.Count
            ' Add comma between values (but not before the first value)
            If col > 1 Then dataLine = dataLine & ","
            
            ' Add the cell value
            If IsNumeric(sourceRange.Cells(row, col).Value) Then
                ' Format numeric values
                dataLine = dataLine & sourceRange.Cells(row, col).Value
            Else
                ' Format text values (add quotes and escape any existing quotes)
                dataLine = dataLine & """" & Replace(sourceRange.Cells(row, col).Value, """", """""") & """"
            End If
        Next col
        
        ' Write the line to the file
        Print #fileNum, dataLine
    Next row
    
    Close #fileNum
    ExportRangeToCSV = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ExportRangeToCSV = False
End Function
