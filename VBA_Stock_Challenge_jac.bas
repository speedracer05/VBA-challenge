Attribute VB_Name = "Module1"
' Create script that loops through 1 yr of stocks; output the following.
    ' Ticker symbol - DONE
    ' Yearly change from open to closing price of that year
    ' Total stock volume
    ' Conditional formatting, positive change in green, negative in red
    
    ' * Bonus return stock with
        ' "Greatest % increase",
        ' "Greatest % decrease" and
        ' "Greatest total volume"
        
        
    ' Loop through all worksheets; For Each ws In Worksheets
    ' Set Summary Table as integer: Dim Summary_Table as Integer
    ' Summary_Table = 2MS Visual Basic and try to import
    ' Set lastrow as integer: Dim lastrow as integer
    ' Find lastrow: lastrow = cells(row.count, 1).end(xlUp).Row
    ' Loop through all ticker symbols: For r = 2 to lastrow
    ' Check if the ticker symbol is the same: If cells(r + 1, 1).value <> cells(r, 1) Then
    ' Set the Ticker_Symbol: Ticker_Symbol = Cells(r, 1).value
    ' Check if Open_Price
    ' Close_Price
    ' Price_Delta
    ' Percent_Delta

Sub stock_analysis()
    
    ' Set initial variable for ticker symbol
    Dim Ticker_Symbol As String
    
  
    ' Variables for moderate solution
    Dim Stock_Volume_Total As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Close_Price = 0
    Dim Delta_Price As Double
    Delta_Price = 0
    Dim Delta_Percent As Double
    Delta_Percent = 0
        
    ' Keep track of location for each ticker symbol in the summary table
    Dim Ticker_Summary_Row As Integer
    Ticker_Summary_Row = 2
    ' MsgBox (Open_Price)
        
    
    ' ADD HEADERS TO TICKER SUMMARY
    ' --------------------------------------------

    ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
    ' Dim WorksheetName As String
    
    ' Determine the Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Grabbed the WorksheetName
    ' WorksheetName = ws.Name
    ' MsgBox WorksheetName

    ' Add the word Ticker to the First Column Header
    Cells(1, 9).Value = "Ticker"

    ' Add the word Yearly Change to the First Column Header
    Cells(1, 10).Value = "Yearly Change"

    ' Add the word Percent Change to the First Column Header
    Cells(1, 11).Value = "Percent Change"

    ' Add the word Total Stock Volume to the First Column Header
    Cells(1, 12).Value = "Total Stock Volume"

    'Auto fit column based on column content
    Range("I:L").EntireColumn.ColumnWidth = 16
    
    
    ' Set initial Open Price. Subsequent Open Price will be set in the forLoop
    Open_Price = Cells(2, 3)
    
    ' LOOP THROUGH ALL TICKERS
    ' --------------------------------------------
        
   
    For r = 2 To LastRow
      ' Check if if the next Ticker Symbol is the same, if it is not...
      If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then

        ' Set the Ticker Symbol
        Ticker_Symbol = Cells(r, 1).Value
        
        
        ' Add to the Stock Volume Total
        Stock_Volume_Total = Stock_Volume_Total + Cells(r, 7).Value
        
        ' Set the ticker Close price
        Close_Price = Cells(r, 6).Value
        'MsgBox (Close_Price)
        
        ' Calculate the yearly change (Delta Price) between Close Price and initial Open Price
        Delta_Price = Close_Price - Open_Price
        MsgBox (Delta_Price)
        
        ' Calculate the Delta Percent for the ticker
        Delta_Percent = Delta_Price / Open_Price
        
        ' Print the Ticker Symbol in the Ticker Summary Row
        Range("I" & Ticker_Summary_Row).Value = Ticker_Symbol
        
        ' Print the Yearly Change in the Ticker Summary Row
        Range("J" & Ticker_Summary_Row).Value = Delta_Price
        
        ' Print the Delta Percent in the Ticker Summary Row
        Range("K" & Ticker_Summary_Row).Value = Delta_Percent
        
        ' Print the Stock Volume Total in the Ticker Summary Row
        Range("L" & Ticker_Summary_Row).Value = Stock_Volume_Total
        
        ' Add one row to the Ticker Summary Row
        Ticker_Summary_Row = Ticker_Summary_Row + 1
        
        ' Reset the Stock Volume Total
        Stock_Volume_Total = 0
        Open_Price = 0
        Close_Price = 0
        Delta_Price = 0
        Delta_Percent = 0
        ' Set Open Price to move to the next Ticker Symbol Open Price
        Open_Price = Cells(r + 1, 3).Value
            
    ' If the cell immediately following a row is the same Ticker Symbol...
    Else
        
        ' Set initial ticker Open Price
        'Open_Price = Cells(2, 3).Value
        ' MsgBox (Open_Price)
        
        ' Add to Ticker Volume Total
        Stock_Volume_Total = Stock_Volume_Total + Cells(r, 7).Value
        
        ' Set the ticker Close Price
        ' Close_Price = Cells(r, 6).Value
          
        ' Set the Delta Price
        ' Delta_Price = Close_Price - Open_Price
               
        ' Print the Delta Price in the Ticker Summary Row
        ' Range("J" & Ticker_Summary_Row).Value = Delta_Price

        ' Set Open_Price to zero
        ' Open_Price = 0

        
      End If

    Next r
             
End Sub
