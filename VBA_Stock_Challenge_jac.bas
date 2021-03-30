Attribute VB_Name = "Module11"
' Create script that loops through 1 yr of stocks; output the following.
    ' Ticker symbol - DONE
    ' Yearly change from open to closing price of that year
    ' Total stock volume
    ' Conditional formatting, positive change in green, negative in red
' *Bonus return stock with
    ' "Greatest % increase",
    ' "Greatest % decrease" and
    ' "Greatest total volume"

' Force declaration of all variables to mitigate generation of errors due to undeclared variables
'Option Explicit

Sub stock_analysis()
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
            
    'Dim ws As Worksheet
    'Dim starting_ws As Worksheet
    'Set starting_ws = ActiveSheet 'sets reference of the starting active worksheet

    ' Loop through all worksheets in this workbook
    'For Each ws In ThisWorkbook.Worksheets
    'ws.Activate
    
        ' Set initial variable for ticker symbol
        Dim Ticker_Symbol As String
        Ticker_Symbol = " "

        ' Keep track of location for each ticker symbol in the summary table
        Dim Ticker_Summary_Row As Integer
        Ticker_Summary_Row = 2
    
        ' Variables for moderate solution
        Dim r As Long
        Dim LastRow As Long
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Delta_Price As Double
        Delta_Price = 0
        Dim Delta_Percent As Double
        Delta_Percent = 0
        Dim Stock_Volume_Total As Double
        Stock_Volume_Total = 0

        ' Set initial Open Price. Subsequent Open Price will be set in the forLoop of each worksheet
        Open_Price = Cells(2, 3)

        ' ****Variable for Bonus Solution
        ' *Dim Percent_Increase As Double

        ' --------------------------------------------
        ' ADD HEADERS TO TICKER SUMMARY AND RESIZE COLUMNS
        ' --------------------------------------------

        ' Determine the Last Row in column 1
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        ' Add the word Ticker to the first column header
        Cells(1, 9).Value = "Ticker"

        ' Add the word Yearly Change to the second column header
        Cells(1, 10).Value = "Yearly Change"

        ' Add the word Percent Change to the third column header
        Cells(1, 11).Value = "Percent Change"

        ' Add the word Total Stock Volume to the fourth column header
        Cells(1, 12).Value = "Total Stock Volume"

        'Set column width for columns I:L
        Range("I:L").EntireColumn.ColumnWidth = 16

        ' --------------------------------------------
        ' ****BONUS SOLUTION -- ADD COLUMN HEADERS, ROW DESCRIPTORS AND FORMAT
        ' --------------------------------------------
        ' ****Add the word Ticker to column O
        ' ws.Cells(1, 15).Value = "Ticker"

        ' ****Add the word Value to column P
        ' ws.Cells(1, 16).Value = "Value"
    
        ' ****Add "Greatest % Increase"
        ' ws.Cells(2, 14).Value = "Greatest % Increase"

        ' ****Add "Greatest % Decrease"
        ' ws.Cells(3, 14).Value = "Greatest % Decrease"

        ' ****Add "Greatest Total Volume"
        ' ws.Cells(4, 14).Value = "Greatest Total Volume"

        '****Set column width for column N
        ' ws.Range("N:N").EntireColumn.ColumnWidth = 20

        ' --------------------------------------------
        ' LOOP THROUGH ALL TICKERS
        ' --------------------------------------------
        ' Loop through from the start of row 2 to the last row that has a value in the last cell.
        For r = 2 To LastRow
            ' Check if if the next Ticker Symbol is the same, if it is not...
            If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then

                ' Set the Ticker Symbol
                Ticker_Symbol = Cells(r, 1).Value

                ' Set the ticker Close price
                Close_Price = Cells(r, 6).Value
                'MsgBox (Close_Price)

                ' Add the stock volume to the Stock Volume Total
                Stock_Volume_Total = Stock_Volume_Total + Cells(r, 7).Value

                ' Calculate the yearly change (Delta Price) between last ticker Close Price and initial Open Price
                Delta_Price = Close_Price - Open_Price
                ' MsgBox (Delta_Price)
                
                ' Check for divide by zero error. I lost 5 hours going through this DANG thing,
                ' checking the code, then finally looking at the last sheet! Uggh!
                    If Open_Price <> 0 Then
                
                    ' Calculate the Delta Percent for the ticker
                    Delta_Percent = Delta_Price / Open_Price
                
                Else
                
                    ' Mitigate program crashing due to a run-time error 6 overflow
                    MsgBox ("For " & Ticker_Symbol & ", there is an error due to " & "Open Price =" & Open_Price & ". Fix the field(s), or remove the worksheet and save the workbook.")
                    
                    End If
                    
                ' Print the Ticker Symbol in the Ticker Summary Row
                Range("I" & Ticker_Summary_Row).Value = Ticker_Symbol

                ' Print the Yearly Change in the Ticker Summary Row
                Range("J" & Ticker_Summary_Row).Value = Delta_Price

                    ' Check Delta Price in the  Ticker Summary Row if it is >= 0, color it green, if it is not...
                    If Delta_Price > 0 Then
                    Range("J" & Ticker_Summary_Row).Interior.ColorIndex = 4

                    'Otherwise color it red
                Else
                    Range("J" & Ticker_Summary_Row).Interior.ColorIndex = 3

                    End If

                ' Print the Delta Percent in the Ticker Summary Row
                Range("K" & Ticker_Summary_Row).Value = Delta_Percent

                ' Change number format to percent, with two decimal places.
                Range("K" & Ticker_Summary_Row).NumberFormat = "0.00%"

                ' Print the Stock Volume Total in the Ticker Summary Row
                Range("L" & Ticker_Summary_Row).Value = Stock_Volume_Total

                ' Change number format to use thousands seperator.
                Range("L" & Ticker_Summary_Row).NumberFormat = "#,##0"

                ' Add one row to the Ticker Summary Row
                Ticker_Summary_Row = Ticker_Summary_Row + 1

               ' Reset moderate solution variables to zero for the next ticker symbol
                Open_Price = 0
                Close_Price = 0
                Delta_Price = 0
                Delta_Percent = 0
                Stock_Volume_Total = 0

            ' --------------------------------------------
            ' ****BONUS SOLUTION -- FIND GREATEST % INCREASE AND DECREASE OF STOCK
            ' --------------------------------------------
            ' Nothing here yet -- only placeholder
            
            
            ' Set Open Price to move to the next Ticker Symbol Open Price
            Open_Price = Cells(r + 1, 3).Value

            ' If the cell immediately following a row is the same Ticker Symbol...
            Else

                ' Add to Ticker Volume Total
                Stock_Volume_Total = Stock_Volume_Total + Cells(r, 7).Value

            End If

        Next r

    'Next ' Moves on to next worksheet in the workbook

    'starting_ws.Activate 'activate the initial active worksheet

End Sub
