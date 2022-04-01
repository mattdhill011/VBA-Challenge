Attribute VB_Name = "Module1"
Sub StockLoop()

    'First we set up the headers for the output data table
    'I figure that I could just set this up on the spreadsheet but I dont want to do it for every sheet and I like having it here so I can keep track of what number the cells are
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Range("O7").Value = "Percent Completed:"
    
    Dim currentTicker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim stockVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
        
    'We are going to use this to keep track of what row we are going to output our data to
    Dim tickerLine As Integer
    
    'We also want to figure out how many rows there are so we know how much to run the loop
    Dim numRows As Long
    
    
    'Here we set some initial values, the tickerLine starts at 2 to go under the header, everything else at 0
    tickerLine = 2
    openPrice = 0
    stockVolume = 0
    
    'Now we figure out how many rows there are, I found this function that goes to the bottom of a range and outputs the row that it is
    numRows = Range("A1").End(xlDown).Row
      
      
    'This chunk of code is for setting up the bonus
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
        
    
      
    
    Dim i As Long
    For i = 2 To numRows
        'I wanted a to add a progress bar so I know how long it will take and to know that it hasnt stopped working on me
        Range("P7").Value = FormatPercent((i / numRows), 0)
    
        'Keep a running total of the stock volume, we reset it to 0 everytime we meet a new ticker
        stockVolume = stockVolume + Cells(i, 7).Value
    
        'Checks to see if openPice has been found yet, it should be 0 at the start of the loop and everytime we meet a new ticker
        'If it is the first time then we take the openPrice and keep it until we need it at the end of the ticker
        If (openPrice = 0) Then
            openPrice = Cells(i, 3).Value
            
        End If
        
        'While we iterate through the data we check if the next cell down has the same ticker name, if it doesnt then that means we have reached the end of this set under that ticker name
        'So now we run all our calculations and reset the variables that we will need for the next set
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
        
            'We take the name of the ticker and slot it into the cell at the tickerLine row
            Cells(tickerLine, 9).Value = Cells(i, 1).Value
            
            'Since this is the end of the set of the ticker (and the end of the year) we take the closing price
            closePrice = Cells(i, 6).Value
            
            'Calculates the yearly change and enters it into the line
            yearlyChange = closePrice - openPrice
            Cells(tickerLine, 10).Value = yearlyChange
            
            'Calculate the percent change between the closePrice and the openPrice
            percentChange = yearlyChange / openPrice
            Cells(tickerLine, 11).Value = FormatPercent(percentChange)
            
            'Output the current total stock volume
            Cells(tickerLine, 12).Value = stockVolume
            
            
            'Now we do the bonus: first checking if the percent change is higher than the greatest, or lower than the lowest
            If (percentChange > greatestIncrease) Then
                greatestIncrease = percentChange
                Cells(2, 16).Value = Cells(i, 1).Value
                Cells(2, 17).Value = FormatPercent(greatestIncrease)
                
            ElseIf (percentChange < greatestDecrease) Then
                greatestDecrease = percentChange
                Cells(3, 16).Value = Cells(i, 1).Value
                Cells(3, 17).Value = FormatPercent(greatestDecrease)
            End If
            
            'Then we check if the stock volume is larger than the current largest
            If (stockVolume > greatestVolume) Then
                greatestVolume = stockVolume
                Cells(4, 16).Value = Cells(i, 1).Value
                Cells(4, 17).Value = greatestVolume
            End If
            
            
                                    
            'Advances the row for the output data table, so next time we do this its a row lower
            tickerLine = tickerLine + 1
            
            'sets open price back to 0 so the first conditional will trigger in the next loop
            openPrice = 0
            
            'And set stock volume to 0 se we can count up again for the next loop
            stockVolume = 0
        
        End If
    Next i
    
    'Now that everything is done, all thats left is to set up all the formating
    'First we start with the yearly change percent, so we create a new range variable
    Dim yearlyChangeRange As Range
    
    'I want to define a range for all of column J, so I use numRows to make a string that will be equal to J2:J<numRows> and set the range to that
    Dim rangeValue As String
    rangeValue = "J2:J" & numRows
    Set yearlyChangeRange = Range(rangeValue)
    
    
    'Remove prior formating
    yearlyChangeRange.FormatConditions.Delete
    
    'Set the color of the cell to red if it is less than 0
    yearlyChangeRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    yearlyChangeRange.FormatConditions(1).Interior.ColorIndex = 3
    
    'Set the color of the cell to green if it is greater than 0
    yearlyChangeRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    yearlyChangeRange.FormatConditions(2).Interior.ColorIndex = 4
    
    'Change the progress bar to green
    Range("P7").Interior.ColorIndex = 4
    Range("P7").Value = "Done!"
    
                     
End Sub


'I just added this so I could clear everything and run it again for testing
Sub Clear()
    Dim numRows As Long
    Dim rangeValue As String
    
    numRows = Range("A1").End(xlDown).Row
    rangeValue = "Q" & numRows
    
    Range("I1", rangeValue).Clear
    
End Sub



