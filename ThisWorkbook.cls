VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub makeMeSomeMoney()

Dim tickerName As String
Dim open_ As Single
Dim volume As Double
Dim tickerCount As Long
Dim greatestIncrease As Single
Dim greatestDecrease As Single
Dim greatestVolume As Double
Dim greatestIncreaseTicker As String
Dim greatestDecreaseTicker As String
Dim greatestVolumeTicker As String
Dim firstBool As Boolean
Dim i As Long
Dim sheet_ As Integer

'Probably makes it go faster kind of.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'Loops through sheets.
For sheet_ = 1 To 3

    'Sets approprite worksheet active from loop.
    Worksheets(sheet_).Activate
    
    With ActiveSheet
    
        'Some simple formating for those of us who are not savages.
        Range("A1:Q1").Font.Bold = True
        Range("J2").EntireColumn.NumberFormat = "$0.00"
        Range("K2").EntireColumn.NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
    
        'Time to name things.
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        Range("O2") = "Greatest % Increase"
        Range("O3") = "Greatest % Decrease"
        Range("O4") = "Greatest Total Volume"
    
        'initializes variables for the start of the sheet.
        tickerName = vbNullString
        tickerCount = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        firstBool = True
    
        'Loops rows from 2 to bottom row.
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row + 1
    
            'Check for first iteration.
            If tickerName <> Cells(i, 1) And firstBool = True Then
            
                'Initialze new set of variables from the first ticker value and turn off the flag for first iteration.
                Cells(tickerCount, 9) = Cells(i, 1)
                tickerName = Cells(i, 1)
                open_ = Cells(i, 3)
                volume = Cells(i, 7)
                firstBool = False
           
            'n'th iteration of new ticker names.
            ElseIf tickerName <> Cells(i, 1) Then
                Cells(tickerCount, 10) = Cells(i - 1, 6) - open_
                Cells(tickerCount, 11) = Cells(tickerCount, 10) / open_
                Cells(tickerCount, 12) = volume
    
                'Format yearly change cells with color to highlight those failures.
                If Cells(tickerCount, 10) < 0 Then
                    Cells(tickerCount, 10).Interior.Color = RGB(255, 0, 0)
                Else
                    Cells(tickerCount, 10).Interior.Color = RGB(0, 255, 0) 'The winners.
                End If
    
                'Track the highest and lowest changes.
                If Cells(tickerCount, 11) > greatestIncrease Then
                    greatestIncrease = Cells(tickerCount, 11)
                    greatestIncreaseTicker = Cells(i - 1, 1)
                End If
                    
                If Cells(tickerCount, 11) < greatestDecrease Then
                    greatestDecrease = Cells(tickerCount, 11)
                    greatestDecreaseTicker = Cells(i - 1, 1)
                End If
    
                If volume > greatestVolume Then
                    greatestVolume = volume
                    greatestVolumeTicker = Cells(i - 1, 1)
                End If
    
                'Initialze new set of variables from the next ticker value.
                tickerName = Cells(i, 1)
                open_ = Cells(i, 3)
                volume = Cells(i, 7)
                tickerCount = tickerCount + 1
                Cells(tickerCount, 9) = Cells(i, 1)
                
            'Not a new ticker value? Get dem digets.
            Else
                volume = volume + Cells(i, 7)
            End If
        Next
    
        'DO THE FINAL THING!!!
        Range("P2") = greatestIncreaseTicker
        Range("P3") = greatestDecreaseTicker
        Range("P4") = greatestVolumeTicker
        Range("Q2") = greatestIncrease
        Range("Q3") = greatestDecrease
        Range("Q4") = greatestVolume
        
        'Because I could not auto fit at the top... I tried :(
        Range("I1").EntireColumn.AutoFit
        Range("J1").EntireColumn.AutoFit
        Range("K1").EntireColumn.AutoFit
        Range("L1").EntireColumn.AutoFit
        Range("O1").EntireColumn.AutoFit
        Range("P1").EntireColumn.AutoFit
        Range("Q1").EntireColumn.AutoFit
    End With
Next

'Yeah...be fast and stuff.
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
