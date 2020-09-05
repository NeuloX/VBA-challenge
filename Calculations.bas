Attribute VB_Name = "Calculations"
Option Explicit

'worksheet controllers
Dim rowcounter As Long
Dim colcounter As Long
Dim sheetnum As Integer

'date verification
Dim firstpricedate As String
Dim lastpricedate As String

Dim totalstock As Variant
Dim curtick As String
'increment
Dim dcount As Integer
'prices
Dim openprice As Double
Dim closeprice As Double
Dim opencount As Long
Dim closecount As Long

'ranges
Dim rowrange As Long
Dim colrange As Long
'bonus
Dim permax As Double
Dim permin As Double
Dim stockmax As Variant


Sub wallstreet()

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False

dcount = 2
totalstock = 0
'loop through the sheets
For sheetnum = 1 To Worksheets.Count
    Worksheets(sheetnum).Activate
    'set range
    rowrange = Worksheets(sheetnum).Cells(Rows.Count, 1).End(xlUp).Row
    opencount = 2
        
    'set titles
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "YearlyChange"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
    'loop through rows
    For rowcounter = 2 To rowrange
        If Cells(rowcounter, 1) = Cells(rowcounter + 1, 1) Then
            'arrange date to readable date
            firstpricedate = Left(Cells(rowcounter, 2), 4) & "/" & _
                        Right(Left(Cells(rowcounter, 2), 6), 2) & "/" & _
                        Right(Cells(rowcounter, 2), 2)
            
            'set current ticker
            curtick = Cells(rowcounter, 1)
            'set total stock
            totalstock = totalstock + Cells(rowcounter, 7)
            'assign to cells
            Cells(dcount, 9) = curtick
            Cells(dcount, 12) = totalstock
            opencount = rowcounter
        ElseIf Cells(rowcounter, 1) <> Cells(rowcounter + 1, 1) Then
            'get the first price
            opencount = (opencount + closecount) - (rowcounter - 2)
            'bypass 1st price error
            If opencount = 1 Then
                opencount = opencount + 1
            End If
            'get open and close prices
            openprice = Cells(opencount, 3)
            closeprice = Cells(rowcounter, 6)
            'next open price incrementation
            closecount = rowcounter
            'set yearly change value
            Cells(dcount, 10) = closeprice - openprice
            'set percent change
            'bypass zero values value
            If openprice = 0 And closeprice = 0 Then
                Cells(dcount, 11) = 0
            ElseIf openprice = 0 And closeprice > 0 Then
                Cells(dcount, 11) = 100
            Else
                Cells(dcount, 11) = Round(((closeprice / openprice) * 100) - 100, 2)
            End If
            'set new total stock
            totalstock = totalstock + Cells(rowcounter, 7)
            Cells(dcount, 12) = totalstock
            
            
            'go to next line for new ticker
            dcount = dcount + 1
            curtick = Cells(rowcounter + 1, 1)
            totalstock = 0
        End If
    
    Next rowcounter
    'reset  for next sheet
            dcount = 2
            opencount = 0
            closecount = 0
    'set formatting and visual formatting

    Visuals.conditional_format
    Calculations.bonus
    Visuals.format
Next sheetnum



Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub


Sub bonus()
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False


'set titles
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest % increase"
Cells(3, 15) = "Greatest % decrease"
Cells(4, 15) = "Greatest total volume"

'get the greatest % increase range
rowrange = Cells(Rows.Count, 11).End(xlUp).Row
Range(Cells(2, 11), Cells(rowrange, 11)).Select
permax = Application.WorksheetFunction.Max(Selection)
permin = Application.WorksheetFunction.Min(Selection)
'set the values
Cells(2, 17) = permax
Cells(3, 17) = permin
'reselect for total
Range(Cells(2, 12), Cells(rowrange, 12)).Select
stockmax = Application.WorksheetFunction.Max(Selection)
Cells(4, 17) = stockmax

'get the tickers
For rowcounter = 2 To rowrange
    If Cells(rowcounter, 11) = permax Then
        Cells(2, 16) = Cells(rowcounter, 9)
    End If
    If Cells(rowcounter, 11) = permin Then
        Cells(3, 16) = Cells(rowcounter, 9)
    End If
    If Cells(rowcounter, 12) = stockmax Then
        Cells(4, 16) = Cells(rowcounter, 9)
    End If
Next rowcounter



Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
