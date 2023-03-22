Attribute VB_Name = "Module1"
Sub ticker():
        
  
For Each ws In Worksheets
     
       ' Initialzing the variables that will be needed to store values
       
        Dim ticker As String
        Dim initialRow As Long
        Dim rowVariable As Long
        Dim stockVolume As Double
        Dim openPrice As Double
        Dim closePrice As Double
        Dim yearlyChange As Double
        Dim percent As Double
        Dim totalVolume As Double
        Dim lastRow As Long
        
        ' Creating all the columns to store our printed results
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ' Intializing these variables with a default value
        
         initialRow = 2
         rowVariable = 2
         stockVolume = 0
         increase = 0
         decrease = 0
         totalVolume = 0
       
       ' Determine the Last Row
       
       lastRow = ws.Range("a1").CurrentRegion.Rows.Count
        
        For i = 2 To lastRow

       ' This is the function to determine the total stock volume for each ticker
            
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            
        ' This is to check the ticker and print them out
            
             If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & rowVariable).Value = ticker
                
                ' This will prnt the results of the total stock volume in the column
                
                ws.Range("L" & rowVariable).Value = stockVolume
                
                ' Reseting the set variable
                
                stockVolume = 0
                
                ' Defining the open and closing prices
                
                openPrice = ws.Range("C" & initialRow)
                closePrice = ws.Range("F" & i)
                
                ' Creating the yearly change values and printing them
                
                yearlyChange = closePrice - openPrice
                ws.Range("J" & rowVariable).Value = yearlyChange

                ' Creating the function to find percent change
                
                If openPrice = 0 Then
                    percent = 0
                Else
                    openPrice = ws.Range("C" & initialRow)
                    percent = yearlyChange / openPrice
                End If
                
                ' Creating the percentage values and printing them accordingly
                percent = Round(percent * 100, 2)
                ws.Range("K" & rowVariable).Value = "%" & percent

                ' Creating the condition to change the color based on the values in the Yearly Change column
                
                If ws.Range("J" & rowVariable).Value < 0 Then
                    ws.Range("J" & rowVariable).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & rowVariable).Interior.ColorIndex = 4
                End If
            
                ' Adding 1 so it can properly loop and iterate
                
                rowVariable = rowVariable + 1
                initalRow = i + 1
                End If
                
                
            Next i

            ' To revalue the lastRow variable for these next functions
            
            lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' Initialzing the beginning of the new loop
            
            For i = 2 To lastRow
            
            ' Looking through the K column to find the greatest value
            
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                
            ' Prints out the value and the ticker associated with it in the right cells
            
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If
             
            ' Runs through the K value again to find the least percent change
            
                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                
             ' Prints out the value and the ticker associated with it in the correct cells
             
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

             ' Runs through the whole L column to find the greatest volume value
             
                 If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                 
             ' Prints the largest volume and the associated ticker with it
             
                     ws.Range("Q4").Value = ws.Range("L" & i).Value
                     ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
            
             ' Goes to the next workheetto run through all the loops procedurely
            
           Next ws

End Sub




