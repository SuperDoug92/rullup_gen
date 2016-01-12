Attribute VB_Name = "Common"
Public pSource As String
Public pTarget As String
Public pSourceObj As Workbook
Public pTargetObj As Workbook
Public pGenObj As Workbook
Public pInvCtr As Integer

Option Explicit

' This section turns on high overhead operations
Sub AppTrue()
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

' This section increases performance by turning off high overhead operations
Sub AppFalse()
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

End Sub

' Main processing loop. Loops through all files in "Source" folder updating the "Target".
Sub ProcessIt()

    Set pGenObj = Workbooks(ActiveWorkbook.Name)   ' Set Rollup generator object
    pInvCtr = 14                                   ' Used to report bad Invalid BUNOs
    Range("C15:D9999").Select                      ' Clear Previous runs data
    Selection.Delete Shift:=xlUp                   ' Clear Previous runs data
    
    Dim PFltCnt As Double
    Dim PHrCnt As Double
    Dim OFltCnt As Double
    Dim OHrCnt As Double
    Dim tSourceRowCnt As Integer
    Dim tSourceColCnt As Integer
    Dim x As Integer
    
    PFltCnt = 0   ' Initialize project flight counter
    PHrCnt = 0    ' Initialize project hour accumulator
    OFltCnt = 0   ' Initialize other flight counter
    OHrCnt = 0    ' Initialize other hour accumulator
    
    Call AppFalse ' turn off high overhead processes
      
    pSource = ThisWorkbook.Sheets(1).Cells(2, 2) ' taken off of Sheet1 - Controlled by user input
    pTarget = ThisWorkbook.Sheets(1).Cells(3, 2) ' taken off of Sheet1 - Controlled by user input
    
    Workbooks.Open Filename:=pSource, ReadOnly:=True  ' Open Source
    Set pSourceObj = Workbooks(ActiveWorkbook.Name)   ' Set source object
    tSourceRowCnt = FindLastRow   ' determine number of rows
    tSourceColCnt = FindLastColSource 'determine number of columns
    
    Dim BunoCol As Integer
    Dim DateCol As Integer
    Dim ProjHrsCol As Integer
    Dim OthHrsCol As Integer

    For x = 1 To tSourceColCnt
        Select Case Trim(Cells(1, x))
            Case Is = "BUNO"
                BunoCol = x
            Case Is = "Flight Date"
                DateCol = x
            Case Is = "Project Hours"
                ProjHrsCol = x
            Case Is = "Other Hours"
                OthHrsCol = x
        End Select
    Next
        
    Workbooks.Open Filename:=pTarget, ReadOnly:=False  ' Open Target
    Set pTargetObj = Workbooks(ActiveWorkbook.Name)    ' Set target object
    
    pSourceObj.Activate   ' focus on Source
           
    For x = 2 To tSourceRowCnt   ' Skip headings and work top down through source file
        If Mid(Cells(x, DateCol), 8, 4) <> "" Then    ' Check if date cell is empty
            If IsNumeric(Trim(Cells(x, ProjHrsCol))) = True Then
                If CDec(Cells(x, ProjHrsCol)) > 0 Then      ' Check if entry is project flight
                PFltCnt = PFltCnt + 1    ' increment flight count
                PHrCnt = PHrCnt + CDec(Cells(x, ProjHrsCol))   ' sum project hours
                End If
            End If
            If IsNumeric(Trim(Cells(x, OthHrsCol))) = True Then
                If CDec(Cells(x, OthHrsCol)) > 0 Then ' Check if entry is other flight
                OFltCnt = OFltCnt + 1    ' increment flight count
                OHrCnt = OHrCnt + CDec(Cells(x, OthHrsCol))    ' sum other hours
                End If
            End If
        End If
        
        If Left(Trim(Cells(x, BunoCol)), 6) = Left(Trim(Cells(x + 1, BunoCol)), 6) Then  ' Still same BUNO
            If Len(Trim(Mid(Cells(x, DateCol), 4, 3))) > 0 Then   ' Valid value in the date field
                If Mid(Cells(x, DateCol), 4, 3) <> Mid(Cells(x + 1, DateCol), 4, 3) Then    ' When month changes
                    Call SetTarget(Cells(x, BunoCol), Cells(x, DateCol), PFltCnt, PHrCnt, OFltCnt, OHrCnt)  ' go find position in Target and update
                    PFltCnt = 0  ' reset project counter
                    PHrCnt = 0   ' reset project accumulator
                    OFltCnt = 0  ' reset other counter
                    OHrCnt = 0   ' reset other accumulator
                End If
            End If
        End If
    Next x
    
    Call AppTrue   ' turn on high overhead processes
    'Call SaveFile  ' saves populated workbook
    'pSourceObj.Close SaveChanges:=False 'closes source without saving
    'pTargetObj.Close SaveChanges:=False 'closes target without saving
    
      
    'pGenObj.Activate   ' Set Focus back to Roll Up Generator
     
    MsgBox ("Completed")   ' Done
  
End Sub

Sub SetTarget(tBuno, tFltDate, PFltCnt, PHrCnt, OFltCnt, OHrCnt)

    Dim tFY As Integer
    Dim x, y, tX, tY As Integer
    Dim tTargetRowCnt As Integer
    Dim tTargetColCnt As Integer

    tFY = GetFiscal(Mid(tFltDate, 4, 3), Mid(tFltDate, 8, 4))  ' determine FY base on date passed in

    pTargetObj.Activate             ' Focus on target
    tTargetRowCnt = FindLastRow     ' determine number of rows
        
    For x = 2 To tTargetRowCnt      ' Skip headings and work top down through target file
        If Left(Trim(Cells(x, 1)), 2) = "FY" Then       ' If you hit an FY section
            If Right(Trim(Cells(x, 1)), 4) = tFY Then   ' If the FYs are the same
                tTargetColCnt = FindLastCol(x + 1)      ' Determine columns before subtotal
                For y = x To tTargetRowCnt              ' Loop through Dates
                    If Mid(tFltDate, 4, 3) = Trim(Cells(y, 2)) Then   ' looking for Month from date passed in
                       Exit For  ' if found exit
                    End If
                Next y
                tX = y   ' Row for the matching FY and Month
                For y = 3 To tTargetColCnt   ' Looking across headings for correct BUNO
                    If tBuno = Left(Trim(Cells(x + 1, y)), 6) Then  ' Looking for BUNO passed in
                       Exit For  ' if found exit
                    End If
                Next y
                tY = y   ' Column for Matching BUNO
                Exit For
            End If
        End If
    Next x
    
    If tX < tTargetRowCnt And tY < tTargetColCnt Then  ' If not at End of Col or Row (valid Value)
        Cells(tX, tY) = PFltCnt       ' Set Project flight count to Target
        Cells(tX, tY + tTargetColCnt - 1) = PHrCnt ' Set Project hour count to Target
        Cells(tX, tY + 1) = OFltCnt   ' Set Other flight count to Target
        Cells(tX, tY + tTargetColCnt) = OHrCnt  ' Set Other hour count to Target
    Else
        pGenObj.Activate        ' Error did not find correct BUNO and / or FY
        pInvCtr = pInvCtr + 1
        Cells(pInvCtr, 3) = tBuno
        Cells(pInvCtr, 4) = tFY
    End If
    
    pSourceObj.Activate  ' Focus on Source to get next section of data

End Sub

Sub SaveFile()
    pTargetObj.Activate 'Activate target file
    
    Dim SaveName As String
    Dim Length As Integer
    Dim a As String
    Dim i As Integer
    
    
    Length = Len(Trim(ActiveWorkbook.Name))
    a = Left(ActiveWorkbook.Name, Length - 5) 'remove file type
    
    i = vbNo
    
    Do While i = vbNo
        SaveName = Application.GetSaveAsFilename(a & "- complete.xlsx", FileFilter:="Excel Files (*.xls*),*.xls*", Title:="Save As")
        Select Case SaveName
            Case Is = "False"
                ' They pressed Cancel
                i = MsgBox("Are you sure you do not want to save? File will be closed without saving changes", vbYesNo, "Save Result")
            Case Is = pTarget
                MsgBox "Error: Cannot overwrite template"
            Case Else
                ActiveWorkbook.SaveAs (SaveName) 'save file
                i = vbYes
        End Select
    Loop
End Sub
    
' Find last row on sheet
Function FindLastRow() As Integer
    If WorksheetFunction.CountA(Cells) > 0 Then
        FindLastRow = Cells.Find(What:="*", After:=[A1], _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
         
End Function

'Find last Column on Source
Function FindLastColSource()
    If WorksheetFunction.CountA(Cells) > 0 Then
        'Search for any entry, by searching backwards by Columns.
        FindLastColSource = Cells.Find(What:="*", After:=[A1], _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
End Function

' Find last Column on sheet
Function FindLastCol(RowNumber) As Integer
    Dim y As Integer
    For y = 3 To 500   'check through columns, 500 is arbitrary but larger then any expected value
        If Trim(Cells(RowNumber, y)) = "Total Sorties" Then
        FindLastCol = y - 1
        Exit For
        End If
    Next y
 
End Function

' Determine FY base on date
Function GetFiscal(tmonth, tyear) As Integer
    If tmonth = "OCT" _
       Or tmonth = "NOV" _
       Or tmonth = "DEC" Then
        GetFiscal = tyear + 1
        Else: GetFiscal = tyear
    End If
    
End Function
