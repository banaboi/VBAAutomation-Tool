Sub UpdateMonthlyData()
Call CLEARCONTENTS2
Call Open_copy_and_paste
Call changecolumnwidth
Call changerowwidth
Call hidecolumns
Call Port_ColourCOde
End Sub



Sub Open_copy_and_paste()

Dim sourceSheet As Worksheet
Dim destSheet As Worksheet

'' copy from the source
' THE FILE PATH MUST BE CHANGED IN ACCORDANCE WITH THE LOCATION ON THE FILE
Workbooks.Open Filename:="D:\Qantas\training report.xlsx"
Set sourceSheet = Sheets(1)
sourceSheet.Activate
sourceSheet.Range("A2:AZ350").Select
Selection.Copy

'' paste to the destination
Workbooks("Training Automation Projectxlsm.xlsm").Activate
Set destSheet = Worksheets("Sheet2")
destSheet.Activate
destSheet.Cells(1, 1).Select
destSheet.Paste

'' save & close
Workbooks("training report.xlsx").Close

End Sub

Sub changecolumnwidth()

Worksheets("Sheet2").Activate
Columns("A:AZ").ColumnWidth = 18

End Sub

Sub changerowwidth()
Worksheets("Sheet2").Activate
Rows("1:351").RowHeight = 15

End Sub

Sub hidecolumns()

Worksheets("Sheet2").Activate

Worksheets("Sheet2").Range("D:E").Columns.Hidden = True
Worksheets("Sheet2").Range("h:h").Columns.Hidden = True
Worksheets("Sheet2").Range("j:k").Columns.Hidden = True
Worksheets("Sheet2").Range("n:n").Columns.Hidden = True
Worksheets("Sheet2").Range("p:q").Columns.Hidden = True
Worksheets("Sheet2").Range("t:t").Columns.Hidden = True
Worksheets("Sheet2").Range("v:w").Columns.Hidden = True
Worksheets("Sheet2").Range("z:z").Columns.Hidden = True
Worksheets("Sheet2").Range("ab:ac").Columns.Hidden = True
Worksheets("Sheet2").Range("af:af").Columns.Hidden = True
Worksheets("Sheet2").Range("ag:al").Columns.Hidden = True
Worksheets("Sheet2").Range("an:az").Columns.Hidden = True

Worksheets("Sheet2").Range("g:g").Columns.Hidden = True
Worksheets("Sheet2").Range("m:m").Columns.Hidden = True
Worksheets("Sheet2").Range("s:s").Columns.Hidden = True
Worksheets("Sheet2").Range("y:y").Columns.Hidden = True
Worksheets("Sheet2").Range("ae:ae").Columns.Hidden = True

End Sub

Sub CLEARCONTENTS2()
Worksheets("Sheet2").Range("a2:az354").CLEARCONTENTS
Worksheets("Sheet2").Range("AM11:AM1000").Interior.Color = xlNone
End Sub

Sub Port_ColourCOde()
 
 Dim lastrow As Integer
    lastrow = Worksheets("Sheet2").Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
        On Error Resume Next
        If Cells(i, 39) = "TMWDTB Tamworth Domestic Terminal" Then
            Cells(i, 39).Interior.Color = RGB(255, 192, 0)
        End If
        If Cells(i, 39) = "TMW Tamworth (Iata Airport)" Then
            Cells(i, 39).Interior.Color = RGB(255, 192, 0)
        End If
        If Cells(i, 39) = "CBRHGR Qantas" Then
            Cells(i, 39).Interior.Color = RGB(0, 255, 0)
        End If
        If Cells(i, 39) = "BNEDTB Brisbane Domestic Terminal" Then
            Cells(i, 39).Interior.Color = RGB(255, 139, 139)
        End If
        If Cells(i, 39) = "BNEQFL" Then
            Cells(i, 39).Interior.Color = RGB(255, 139, 139)
        End If
        If Cells(i, 39) = "CNSDTB Cairns Domestic Terminal" Then
            Cells(i, 39).Interior.Color = RGB(255, 139, 139)
        End If
        If Cells(i, 39) = "AKLDTB Auckland Domestic Terminal" Then
            Cells(i, 39).Interior.Color = RGB(255, 139, 139)
        End If
        If Cells(i, 39) = "AKLITB Auckland International Terminal" Then
            Cells(i, 39).Interior.Color = RGB(255, 139, 139)
        End If
        If Cells(i, 39) = "SYD085 Sydney Hangar 85 Mascot" Then
            Cells(i, 39).Interior.Color = RGB(29, 196, 255)
        End If
        If Cells(i, 39) = "SYDDTB Sydney Domestic Terminal" Then
            Cells(i, 39).Interior.Color = RGB(29, 196, 255)
        End If
        If Cells(i, 39) = "SYD133 Sydney Hangar 133 Mascot" Then
            Cells(i, 39).Interior.Color = RGB(29, 196, 255)
        End If
        If Cells(i, 39) = "SYDQCB Qantas Centre B Wing" Then
            Cells(i, 39).Interior.Color = RGB(29, 196, 255)
        End If
        If Cells(i, 39) = "PMR - PALMERSTON NORTH" Then
            Cells(i, 39).Interior.Color = RGB(255, 139, 139)
        End If
        If Cells(i, 39) = "MELDTB Melbourne Domestic Terminal" Then
            Cells(i, 39).Interior.Color = RGB(174, 119, 215)
        End If
        If Cells(i, 39) = "MEL Melbourne (Iata Airport)" Then
            Cells(i, 39).Interior.Color = RGB(174, 119, 215)
        End If
        If Cells(i, 39) = "NPE - NAPIER" Then
            Cells(i, 39).Interior.Color = RGB(192, 192, 192)
        End If
        If Cells(i, 39) = "WEI - WEIPA" Then
            Cells(i, 39).Interior.Color = RGB(204, 255, 255)
        End If
        If Cells(i, 39) = "ADL - Adelaide" Then
            Cells(i, 39).Interior.Color = RGB(128, 128, 0)
        End If
        If Cells(i, 39) = "CBRDTB Canberra Domestic Terminal" Then
            Cells(i, 39).Interior.Color = RGB(0, 255, 0)
        End If
        If Cells(i, 39) = "SYDQCA Qantas Centre A Wing" Then
            Cells(i, 39).Interior.Color = RGB(29, 196, 255)
        End If
        If Cells(i, 39) = "SYD131 Sydney Hangar 131 Mascot" Then
            Cells(i, 39).Interior.Color = RGB(29, 196, 255)
        End If
        If Cells(i, 39) = "SYDQCC Qantas Centre C Wing" Then
            Cells(i, 39).Interior.Color = RGB(29, 196, 255)
        End If
        If Cells(i, 39) = "SYD191 Qantas Jetbase Mascot Syd" Then
            Cells(i, 39).Interior.Color = RGB(29, 196, 255)
        End If
            If Cells(i, 39) = "BNE Brisbane (Iata Airport)" Then
            Cells(i, 39).Interior.Color = RGB(255, 139, 139)
        End If
        Next i
        
End Sub
