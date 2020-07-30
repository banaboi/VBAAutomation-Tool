Sub CLEARCONTENTS()
Worksheets("Search").Range("AM11:AM1000").Interior.Color = xlNone
Worksheets("Search").Range("a11:am354").CLEARCONTENTS
End Sub

Sub pleasedontcrash()
Dim cur_row As Integer, lastrow As Integer, x As Integer, p As Integer, mydte As Date

cur_row = 2
p = 11
mydte = Date
lastrow = Worksheets("Sheet2").Cells(Rows.Count, 1).End(xlUp).Row


Dim dates(0 To 4) As Date
    
While cur_row <= lastrow
    dates(0) = Worksheets("Sheet2").Cells(cur_row, 6).value
    dates(1) = Worksheets("Sheet2").Cells(cur_row, 12).value
    dates(2) = Worksheets("Sheet2").Cells(cur_row, 18).value
    dates(3) = Worksheets("Sheet2").Cells(cur_row, 24).value
    dates(4) = Worksheets("Sheet2").Cells(cur_row, 30).value
    x = 0
        While x <= 4
            If IsEmpty(dates(x)) = True Then
                x = x + 1
            ElseIf Application.WorksheetFunction.Days(dates(x), mydte) <= 94 And Application.WorksheetFunction.Days(dates(x), mydte) > 0 Then
           
                Worksheets("Sheet2").Rows(cur_row).Copy
                Worksheets("Search").Rows(p).Insert
                p = p + 1
                x = 69
    
    
            Else
                x = x + 1
            End If
       
      
        Wend
    
    cur_row = cur_row + 1
   
Wend
       
Application.ScreenUpdating = False



End Sub
