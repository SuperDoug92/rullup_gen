Attribute VB_Name = "Module2"
Option Base 1
Option Explicit
Sub GenerateTarget()
    Call AppFalse
    
    Dim GenObj As Object
    Set GenObj = ActiveWorkbook
    Dim bSource As String ' variable to store source workbook address
    Dim tSourceColCnt As Integer ' variable to store number of columns in source
    Dim BunoCol As Integer ' variable to store which column BUNOs are in
    Dim j As Integer ' loop counter
    bSource = ThisWorkbook.Sheets(1).Cells(2, 2)
    Workbooks.Open Filename:=bSource, ReadOnly:=True ' open source document
    Dim wbSource As Object
    Set wbSource = ActiveWorkbook 'set source document as object
    
    tSourceColCnt = FindLastColSource
    For j = 1 To tSourceColCnt ' find buno column by header
        If Trim(Cells(1, j)) = "BUNO" Then
            BunoCol = j
        End If
    Next
    
    Dim x As Integer
    
    Dim DateCol As Integer
    For x = 1 To tSourceColCnt
        If Trim(Cells(1, x)) = "Flight Date" Then
            DateCol = x
        End If
    Next
    
    Dim i As Integer ' loop counter
    Dim tSourceRowCnt As Integer ' variable which stores number of rows in source
    Dim v As Integer ' unique BUNO counter
    Dim cCount As Integer ' size of BUNO Array
    ReDim Barray(1) As Long ' BUNO Array
    tSourceRowCnt = FindLastRow
    v = 1
   
    For i = 2 To tSourceRowCnt ' create array of all Buno numbers
        If Left(Trim(Cells(i, BunoCol)), 6) <> "Grand " And _
        Left(Trim(Cells(i, BunoCol)), 6) = Left(Trim(Cells(i + 1, BunoCol)), 6) Then ' if not last row and the BUNOs match
            If v > UBound(Barray) And v > 1 Then  ' if we are looping through a new BUNO
                ReDim Preserve Barray(v)
                Barray(v) = Left(Trim(Cells(i, BunoCol)), 6)
            Else:
                ReDim Preserve Barray(v)      ' for the first entry
                Barray(v) = Left(Trim(Cells(i, BunoCol)), 6)
            End If
        End If
        If Left(Trim(Cells(i, BunoCol)), 6) <> Left(Trim(Cells(i + 1, BunoCol)), 6) Then
            v = v + 1
        End If
    Next
    
    cCount = UBound(Barray) ' find size of array (option base 1)
    
    Application.Workbooks.Add ' Open new workbook
    Dim TestSum As Object
    Set TestSum = ActiveWorkbook
    GenObj.Sheets(2).UsedRange.Copy
    TestSum.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths
    TestSum.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
    
    Dim z As Integer
    Dim q As Integer
    
    For z = 1 To (2 * cCount) - 2 ' expand table on flight count side
        Range(Cells(1, 3), Cells(1, 4)).EntireColumn.Select
        Selection.Copy
        Range(Cells(1, z + 2), Cells(1, z + 3)).EntireColumn.Insert
        z = z + 1
    Next
    For q = 1 To 2 * cCount - 2 ' expand table on flight hour side
        Range(Cells(1, 2 * cCount + 4), Cells(1, 2 * cCount + 5)).EntireColumn.Select
        Selection.Copy
        Range(Cells(1, q + 2 * cCount + 5), Cells(1, q + 2 * cCount + 6)).EntireColumn.Insert
        q = q + 1
    Next
    
    ' Write titles into table headers
    Dim f As Integer
    Dim n As Integer
    n = 1
    For f = 1 To cCount
        Cells(3, n + 2) = CStr(Barray(f)) + " Project"
        Cells(3, n + 3) = CStr(Barray(f)) + " Other"
        Cells(3, n + 2).Font.Bold = True
        Cells(3, n + 3).Font.Bold = True
        Cells(3, n + 2 * cCount + 3) = CStr(Barray(f)) + " Project"
        Cells(3, n + 2 * cCount + 4) = CStr(Barray(f)) + " Other"
        Cells(3, n + 2 * cCount + 3).Font.Bold = True
        Cells(3, n + 2 * cCount + 4).Font.Bold = True
        n = n + 2
    Next
    
    'insert sum formulas across row
    Dim a As Integer
    a = 2 * cCount
    a = CStr(a)
    
    Dim y As Integer
    For y = 1 To 13
        Cells(y + 3, 2 * cCount + 3).Formula = "=Sum(RC[" & -a & "]:RC[-1])"
        Cells(y + 3, 4 * cCount + 4).Formula = "=Sum(RC[" & -a & "]:RC[-1])"
        Cells(y + 3, 2 * cCount + 3).Font.Bold = True
        Cells(y + 3, 4 * cCount + 4).Font.Bold = True
    Next
    
    x = 0
    n = 1
    y = 0
    ReDim YArray(1) As Integer
    
    For y = 1950 To 2011
        For x = 2 To tSourceRowCnt
            If Trim(Right(wbSource.Sheets(1).Cells(x, DateCol), 4)) = y Then
                ReDim Preserve YArray(n)
                YArray(n) = y
                n = n + 1
                Exit For
            End If
        Next
    Next
    
    Dim NumYears As Integer
    NumYears = UBound(YArray)
    
    Cells(2, 1) = "FY" & CStr(YArray(1))
    x = 0
    Dim tRow As Integer
    tRow = 18
    For x = YArray(2) To YArray(NumYears)
        Range(Cells(2, 1), Cells(16, 1)).EntireRow.Copy
        Cells(tRow, 1).EntireRow.Insert
        Cells(tRow, 1) = "FY " & CStr(x)
        Cells(tRow, cCount * 4 + 5) = "FY " & CStr(x)
        Cells(tRow + 14, cCount * 4 + 5) = "FY " & Right(CStr(x), 2) & " Avg"
        tRow = tRow + 16
    Next
    
    Dim Row As Integer
    Row = FindLastRow
    'add single instance cell fills
    Cells(Row + 1, 2).Select
    Selection = "Grand Total"
    Selection.ColumnWidth = 7
    Selection.RowHeight = 29
    Selection.WrapText = True
    Selection.Interior.Color = RGB(204, 204, 255)
    Range(Cells(Row + 1, 2), Cells(Row + 1, 4 * cCount + 4)).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
    Dim ac As Integer ' array coumn variable
    Dim ar As Integer ' array row variable
    ac = 1
    ar = 1
    ReDim Total(x, cCount * 4 + 2) As Variant ' Subtotal array
    
    Dim r As Integer
    r = 1
    
    For r = 1 To Row ' fill array with all subtotals
        If Cells(r, 2) = "Total" Then
            For ac = 1 To cCount * 4 + 2
                Total(ar, ac) = Cells(r, ac + 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            Next
            ar = ar + 1
        End If
    Next
    
    ar = 1
    Dim c As Integer ' column in template variable
    Dim FormHold As String ' dummy variable, holds formuls for grand total cell until used
    
    For c = 1 To cCount * 4 + 2 ' cycle through columns in template
        FormHold = "="
        For ar = 1 To NumYears ' cycle through array
            If ar <> 1 Then ' if not first entry
                FormHold = FormHold & "+"
            End If
            FormHold = FormHold & Total(ar, c)
        Next
        Cells(Row + 1, c + 2) = FormHold ' insert formula
        Cells(Row + 1, c + 2).Font.Bold = True
    Next
    
    
    Call AppTrue
End Sub

