Attribute VB_Name = "Module1"
Sub ticker()
    For Each ws In Worksheets
    Dim lastRow As Long
    Dim x As Integer
    
    x = 2
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    MsgBox ("Total rows are " & lastRow)
    
    For i = 2 To lastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            x = x + 1
            
            Cells(x, 9).Value = Cells(i, 1).Value
       
        End If
    
    Next i
    Next ws
    
    
End Sub
Sub yearlyChange()
Attribute yearlyChange.VB_ProcData.VB_Invoke_Func = " \n14"
    
    For Each ws In Worksheets
    Dim lastRow As Long
    Dim x As Integer
    Dim yearOpen As Double
    Dim yearClose As Double
    
    x = 2
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    MsgBox ("Total rows are " & lastRow)
    
        For i = 2 To lastRow
        
        If Cells(i, 2).Value = 20160101 Then
            yearOpen = Cells(i, 3).Value
        End If
        
        If Cells(i, 2).Value = 20161230 Then
            yearClose = Cells(i, 6).Value
            x = x + 1
            Cells(x, 10).Value = yearClose - yearOpen
        End If
    Next i
    Next ws
    
End Sub
Sub yearlyPercent()

    For Each ws In Worksheets
    Dim lastRow As Long
    Dim x As Integer
    Dim yearOpen As Double
    Dim yearClose As Double
    
    x = 2
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    MsgBox ("Total rows are " & lastRow)
    
        For i = 2 To lastRow
        
        If Cells(i, 2).Value = 20140101 Then
            yearOpen = Cells(i, 3).Value
        End If
        
        If Cells(i, 2).Value = 20141230 Then
            yearClose = Cells(i, 6).Value
            x = x + 1
            Cells(x, 11).Value = (yearClose - yearOpen) / yearOpen * 100
        End If
    Next i
    Next ws
    
End Sub

Sub yearlyVolume()


    For Each ws In Worksheets
    Dim lastRow As Long
    Dim x As Integer
    Dim vol As Double
    
    
        x = 2
        vol = 0
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        MsgBox ("Total rows are " & lastRow)
    
        For i = 2 To lastRow
            vol = vol + Cells(i, 7)
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            x = x + 1
                Cells(x, 12).Value = vol
                vol = 0
                
        End If
    Next i
    Next ws
    
End Sub

Sub table()

    For Each ws In Worksheets
    Dim lastRow As Long
    Dim x As Integer
    Dim percentChange As Double
    
        x = 2
        percent = 0
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        MsgBox ("percent change")
        
        For i = 2 To lastRow
            percentChange = percentChange + Cells(i, 11)
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            x = x + 1
                Cells(x, 11).Value = percentChange
                percentChange = 0
    
End Sub

