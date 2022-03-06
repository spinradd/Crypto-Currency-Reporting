Attribute VB_Name = "Transaction_Summary"
Sub CreateTxnTables()

''' this macro does it all, it will repopulate the entire workbook with updated tables
''' unfortunately, this can get slow as it will do all the currencies and all the calculations every time
''' This macro will:
        ''' look through transaction table and get a list of all the currencies entered
        ''' it will then delete all the old sheets and replace with new sheets, each populated with the
        ''' coin's income history, sale data, sale summary data, and income summary data
        ''' after wards, the macro will look at each coin's yearly totals and create a new sheet with
        ''' the yearly gain/loss and income totals
        
''' run once a year!

    ' set sheets to variables
    Set txn_sht = Worksheets("Transaction")
    
    SortByDate "Transaction", "Transaction_tbl", "Date"
    
    ThisWorkbook.PrecisionAsDisplayed = True
    
    ' set var to main Transactions table
    Set txn_tbl = txn_sht.ListObjects("Transaction_tbl")
    Set TxnType = txn_tbl.ListColumns("type").DataBodyRange
    Set TxnTicker = txn_tbl.ListColumns("Ticker").DataBodyRange
    Set TxnUnits = txn_tbl.ListColumns("Transacted Units").DataBodyRange
    Set TxnPrice = txn_tbl.ListColumns("Transacted Price (per unit)").DataBodyRange
    Set TxnType = txn_tbl.ListColumns("type").DataBodyRange
    Set TxnDate = txn_tbl.ListColumns("Date").DataBodyRange
    Set TxnFees = txn_tbl.ListColumns("Fees").DataBodyRange
    
    ' create array
    Dim tickerarr() As Variant
    
    ' get unique tickers in array
    tickerarr = GetArray("Transaction", "Transaction_tbl", "ticker")
    
    ' make sure array exists and table was not empty
    tickerarr = DoesArrayExist(tickerarr)
    
    Application.DisplayAlerts = False
    
    ' for each crypto currency in transactions table (income)...
    For Each tick In tickerarr
    
    
   
            ' create sheetname
            sheetnm = UCase(tick) & "_txn"
            sheetexist = False
            
            'if sheet already exists, delete it
            For Each Sheet In ThisWorkbook.Worksheets
            If Sheet.Name = sheetnm Then
                Sheet.Delete
                sheetexists = False
            End If
            Next Sheet
            
            ' make a new sheet for the currency
            If sheetexist = False Then
                Make_New_Sheet_Txn CStr(tick)
            End If
            
            ' assign vars to income and sale table (Table 1/2) of current ticker
            Dim tick_tbl As ListObject
            Set tick_tbl = Worksheets(sheetnm).ListObjects(LCase(tick) & "_income_txn")
            Set TickDate = tick_tbl.ListColumns("Date of Buy/Income").DataBodyRange
            Set Tick_p_coin = tick_tbl.ListColumns("Price/Coin").DataBodyRange
            Set Tick_c_gain = tick_tbl.ListColumns("Coins Gained").DataBodyRange
            Set Tick_v_gain = tick_tbl.ListColumns("Value Gained").DataBodyRange
            Set Tick_type = tick_tbl.ListColumns("Buy or Income").DataBodyRange
            
            ' if table range is less thn one add a row to eliminate bug
            If tick_tbl.Range.Rows.Count < 1 Then
                tick_tbl.ListRows.Add
            End If
            
            
            ' set totals variables
            Current_Coin_Total = 0
            Total_Gain = 0
            Total_Loss = 0
            Short_Gain = 0
            Short_Loss = 0
            Long_Gain = 0
            Long_Loss = 0
            row_count = 0
            Buy_count = 0
            ticker_tbl_count = 0
        
        ' for each row in transaction table
        For Each ticker In TxnTicker
        
        
        
            ' next row
            row_count = row_count + 1
            
            ' if ticker is this loop's selected ticker and the type is buy or income...
            If tick = ticker And (TxnType(row_count, 1).value = "Buy" Or TxnType(row_count, 1).value = "Income") Then
            
                
                ' add to the row count of the currency's transaction table
                ticker_tbl_count = ticker_tbl_count + 1
                
                ' copy and paste type into currency's transaction table
                Tick_type(ticker_tbl_count, 1).value = TxnType(row_count, 1).value
                TickDate(ticker_tbl_count, 1).value = TxnDate(row_count, 1).value
        
                ' if there are no fees for the transaction...
                If TxnFees(row_count, 1).value = 0 Then
                    
                    ' the price of the coin will be
                    Tick_p_coin(ticker_tbl_count, 1).value = TxnPrice(row_count, 1).value
                
            
                ' if there are fees...
                Else
                    ' if the quantity of the coin is greater than 1...
                    If (TxnUnits(row_count, 1).value > 1 And TxnFees(row_count, 1).value > 1) _
                        Or (TxnUnits(row_count, 1).value > 1 And TxnFees(row_count, 1).value < 1) Then
                            
                            ' add the fees to the price of the coin by dividing
                            Tick_p_coin(ticker_tbl_count, 1).value = TxnPrice(row_count, 1).value + (TxnFees(row_count, 1).value / TxnUnits(row_count, 1).value)
                    
                    ' if the quantity of the coin is less than one...
                    ElseIf (TxnUnits(row_count, 1).value < 1 And TxnFees(row_count, 1).value > 1) _
                        Or (TxnUnits(row_count, 1).value < 1 And TxnFees(row_count, 1).value < 1) Then
                            
                            ' add fees to the price of the coin by multiplcation
                            Tick_p_coin(ticker_tbl_count, 1).value = TxnPrice(row_count, 1).value + (TxnFees(row_count, 1).value * TxnUnits(row_count, 1).value)
                    
                    End If
                End If
                
                ' place coin's gain into table
                Tick_c_gain(ticker_tbl_count, 1).value = TxnUnits(row_count, 1).value
                Tick_v_gain(ticker_tbl_count, 1).value = (TxnPrice(row_count, 1).value) * TxnUnits(row_count, 1).value
                
                
                'If ticker_tbl_count = 1 Then
                 '   'Tick_cumulative(ticker_tbl_count, 1).value = 0
                  ' Else
                '    'Tick_cumulative(ticker_tbl_count, 1).value = tick_tbl.ListColumns("Cumulative Coins").DataBodyRange(ticker_tbl_count - 1, 1).value
                'End If
                
            End If
        Next ticker
        
        
        ' for each row in transaction table (sale)...
        row_count = 0
        For Each ticker In TxnTicker
            
            row_count = row_count + 1
            
            ' if designated ticker and a sale or fee value
            If tick = ticker And (TxnType(row_count, 1).value = "Sell" Or TxnType(row_count, 1).value = "Fee") Then
                
                ' get sale data as variables
                sell_date = TxnDate(row_count, 1).value
                Price_Of_Sale = TxnPrice(row_count, 1).value
                Coin_Sale_Amount = TxnUnits(row_count, 1).value
                
                ' calculate gains and losses for sale
                Calculate.Calc_GainsLosses CDbl(Coin_Sale_Amount), CDbl(Price_Of_Sale), CDate(sell_date), UCase(ticker)
            
            End If
            
        Next ticker
        
        ' after evaluating all the sales for one ticker, update other tables
        
        UpdateSaleSummary CStr(tick)
        Calc_Income CStr(tick)
        Calculate_Summary CStr(tick)
    
    
    Next tick
    
    ' create sheet name for portfolio sheet
    sheetnm = "Portfolio_Summary"
    
    ' if this sheet alrready exists delete it
    For Each Sheet In ThisWorkbook.Worksheets
    If Sheet.Name = sheetnm Then
        Sheet.Delete
    End If
    Next Sheet
    
    ' create new sheet, set it to a variable
    Sheets.Add after:=ThisWorkbook.Worksheets("Control Center")
    ActiveSheet.Name = sheetnm
    Set portfolio_sht = Worksheets(sheetnm)
    
    ' length of rows needed for table = number of crypto currencies
    Length = UBound(tickerarr) - LBound(tickerarr) + 1
                
    ' min/max year from transaction table (years f activity)
    min_year = Year(Application.WorksheetFunction.Min(Worksheets("Transaction").ListObjects("Transaction_tbl").ListColumns("Date").DataBodyRange))
    max_year = Year(Application.WorksheetFunction.Max(Worksheets("Transaction").ListObjects("Transaction_tbl").ListColumns("Date").DataBodyRange))
    
    'for each year in range...
    year_count = 0
    For yr = min_year To max_year
        ' go to next year
        year_count = year_count + 1
        
        ' create table for each year
        ' set top left corner of table to range, set bottom right corner of table to range
        Set Top_Left = portfolio_sht.Range("A2").Offset(0, (year_count - 1) * 9)
        Set Bottom_Right = Top_Left.Offset(Length, 5)
        ' set table range
        table_range = Top_Left.Address & ":" & Bottom_Right.Address
        
        ' create table
        Set Year_tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(table_range), , xlYes)
        
        ' name table
        Year_tbl.Name = yr & "_tbl"
        
        ' create listcolumn names for table
        Year_tbl.HeaderRowRange(1, 1) = "Coin"
        Year_tbl.HeaderRowRange(1, 2) = "Mined/Staked Income"
        Year_tbl.HeaderRowRange(1, 3) = "Long Gain"
        Year_tbl.HeaderRowRange(1, 4) = "Long Loss"
        Year_tbl.HeaderRowRange(1, 5) = "Short Gain"
        Year_tbl.HeaderRowRange(1, 6) = "Short Loss"
        Year_tbl.HeaderRowRange(1, 7) = "Realized Gain Loss"
        Year_tbl.HeaderRowRange(1, 8) = "Holdings by EOY"
        
        ' put year on top of table
        Year_tbl.HeaderRowRange(1, 1).Offset(-1, 0).value = yr
        
        ' center year label across rows
        Set YearRng = Range(Year_tbl.HeaderRowRange(1, 1).Offset(-1, 0).Address & ":" & Year_tbl.HeaderRowRange(1, 5).Offset(-1, 0).Address)
        YearRng.HorizontalAlignment = xlCenterAcrossSelection
       
        
        ' add row to table
        If Year_tbl.Range.Rows.Count < 1 Then
                Year_tbl.ListRows.Add
        End If
        
    
        ' for each cell in the year table...
        ticker_count = 0
        For Each cell In Year_tbl.ListColumns("Coin").DataBodyRange
        
            ' place a new currencies ticker in row, it will be that currencies row for the year
            cell.value = tickerarr(ticker_count)
            ticker_count = ticker_count + 1
           
           ' set variable to that coin's sheet's summary table
           Set coin_tbl = Worksheets(cell.value & "_txn").ListObjects(LCase(CStr(cell.value)) & "_year_sum_tbl_txn")
           
           ' assume starting year is 2000
           year_val = 2000
           row_count = 0
           
           ' assign year value to the first year in the coin's sheet's yearly summary table
           ' get to current loops year
           Do While year_val <> yr
                row_count = row_count + 1
                year_val = coin_tbl.ListColumns("Year").DataBodyRange(row_count, 1).value
                
                If year_val = Empty Then
                    Exit Do
                End If
           Loop
           
           ' if there is no year, there hasn't been any income/sales for this year - set everything to zero
           If year_val <> Empty Then
                Year_tbl.ListColumns("Mined/Staked Income").DataBodyRange(ticker_count, 1).value = coin_tbl.ListColumns("Income").DataBodyRange(row_count, 1).value
                Year_tbl.ListColumns("Long Gain").DataBodyRange(ticker_count, 1).value = coin_tbl.ListColumns("Long Gain").DataBodyRange(row_count, 1).value
                Year_tbl.ListColumns("Long Loss").DataBodyRange(ticker_count, 1).value = coin_tbl.ListColumns("Long Loss").DataBodyRange(row_count, 1).value
                Year_tbl.ListColumns("Short Gain").DataBodyRange(ticker_count, 1).value = coin_tbl.ListColumns("Short Gain").DataBodyRange(row_count, 1).value
                Year_tbl.ListColumns("Short Loss").DataBodyRange(ticker_count, 1).value = coin_tbl.ListColumns("Short Loss").DataBodyRange(row_count, 1).value
            Else
                Year_tbl.ListColumns("Mined/Staked Income").DataBodyRange(ticker_count, 1).value = 0
                Year_tbl.ListColumns("Long Gain").DataBodyRange(ticker_count, 1).value = 0
                Year_tbl.ListColumns("Long Loss").DataBodyRange(ticker_count, 1).value = 0
                Year_tbl.ListColumns("Short Gain").DataBodyRange(ticker_count, 1).value = 0
                Year_tbl.ListColumns("Short Loss").DataBodyRange(ticker_count, 1).value = 0
          End If
        Next
        
        ' for each cell in the portfolio's year summary table...
        row_count = 0
        For Each cell In Year_tbl.ListColumns("Realized Gain Loss").DataBodyRange
            
            row_count = row_count + 1
            
            ' the total loss/profit is equal to th sums of the short/long gain/loss columns
            cell.value = Year_tbl.ListColumns("Long Gain").DataBodyRange(row_count, 1).value + _
                Year_tbl.ListColumns("Long Loss").DataBodyRange(row_count, 1).value + _
                Year_tbl.ListColumns("Short Gain").DataBodyRange(row_count, 1).value + _
                Year_tbl.ListColumns("Short Loss").DataBodyRange(row_count, 1).value
        Next
        
        ' for each row in portfolio's year table...
        row_count = 0
        For Each cell In Year_tbl.ListColumns("Holdings by EOY").DataBodyRange
            
            row_count = row_count + 1
            income_count = 0
            income_sum = 0
            Year_sum = 0
            
            ' get the coin's quantity
            currency_val = Year_tbl.ListColumns("Coin").DataBodyRange(row_count, 1).value
            
            ' set the variable to the income table from the currency's individual sheet
            Set currency_gain_tbl = Worksheets(currency_val & "_txn").ListObjects(currency_val & "_income_txn")
            
            ' for each year in that currency's individual summary table...
            For Each year_cell In currency_gain_tbl.ListColumns("Date of Buy/Income").DataBodyRange
                
                income_count = income_count + 1
                
                ' if that rows year equals this year's loop, add to the income
                If Year(year_cell) <= yr Then
                    income_sum = income_sum + currency_gain_tbl.ListColumns("Coins Gained").DataBodyRange(income_count, 1).value
                
                End If
            Next
            
           ' set variable to currency's individual sale summary table
           Set currency_sell_tbl = Worksheets(currency_val & "_txn").ListObjects(currency_val & "_summary_tbl_txn")
           sell_count = 0
           sell_sum = 0
           
           ' for each year in sale summary, if the year of the row equals this loop's year, add the coin quantity
           ' to the coins sold total
           For Each year_cell In currency_sell_tbl.ListColumns("Date of Sale").DataBodyRange
                sell_count = sell_count + 1
                
                If Year(year_cell) <= yr Then
                    sell_sum = sell_sum + currency_sell_tbl.ListColumns("Coins Sold (#)").DataBodyRange(sell_count, 1).value
                
                End If
           Next year_cell
            
        
        ' place EOY holding total in appropriate cell
        Year_tbl.ListColumns("Holdings by EOY").DataBodyRange(row_count, 1).value = income_sum - sell_sum
            
        Next cell
        
        ' format table
        With Year_tbl
        .DataBodyRange.HorizontalAlignment = xlRight
        Year_tbl.ListColumns("Mined/Staked Income").DataBodyRange.NumberFormat = "$#,##0.00_);($#,##0.00)"
        Year_tbl.ListColumns("Long Gain").DataBodyRange.NumberFormat = "$#,##0.00_);($#,##0.00)"
        Year_tbl.ListColumns("Long Loss").DataBodyRange.NumberFormat = "$#,##0.00_);($#,##0.00)"
        Year_tbl.ListColumns("Short Gain").DataBodyRange.NumberFormat = "$#,##0.00_);($#,##0.00)"
        Year_tbl.ListColumns("Short Loss").DataBodyRange.NumberFormat = "$#,##0.00_);($#,##0.00)"
        Year_tbl.ListColumns("Realized Gain Loss").DataBodyRange.NumberFormat = "$#,##0.00_);($#,##0.00)"
        Year_tbl.ListColumns("Holdings by EOY").DataBodyRange.NumberFormat = "#####.#####"
        End With
        
    Next yr
    
    
End Sub

Sub Make_New_Sheet_Txn(ticker As String)
    
    
    ThisWorkbook.PrecisionAsDisplayed = True
    
    ' name for sheet
    sheetnm = ticker + "_txn"

    ' if sheet exists, delete it
    For Each Sheet In ThisWorkbook.Worksheets
    If Sheet.Name = sheetnm Then
        MsgBox "Sheet already exists"
        Exit Sub
    End If
    Next Sheet
    
    ' create new sheet, set to variable
    Sheets.Add after:=ThisWorkbook.Worksheets("Control Center")
    ActiveSheet.Name = sheetnm
    Set Tnd_sheet = ThisWorkbook.Worksheets(sheetnm)
    
    
    ' for four different tables...
    For x = 1 To 4
    
    Select Case x
        ' previous length equals the space of the previous table plus an extra column or two as a buffer, increases as x
        ' increases because the tables are located horizontally from each other
        Case 1
            table_name = LCase(ticker) & "_income_txn"
            table_length = 14
            previous_length = 14
            
            'initialize top left cell for table start
            Set Top_Left = Worksheets(sheetnm).Range("A4")
             
             'specify last cell of table (= 2 * number of tag columns)
            Set Bottom_Right = Top_Left.Offset(1, table_length)
        Case 2
            table_name = LCase(ticker) & "_summary_tbl_txn"
            table_length = 7
            
            'initialize top left cell for table start
            Set Top_Left = Worksheets(sheetnm).Range("A4").Offset(0, previous_length + 3)
            
            'specify last cell of table (= 2 * number of tag columns)
            Set Bottom_Right = Top_Left.Offset(1, table_length)
            previous_length = previous_length + table_length + 2
        Case 3
            table_name = LCase(ticker) & "_yearly_income_txn"
            table_length = 2
            
            'initialize top left cell for table start
            Set Top_Left = Worksheets(sheetnm).Range("A4").Offset(0, previous_length + 3)
            
            'specify last cell of table (= 2 * number of tag columns)
            Set Bottom_Right = Top_Left.Offset(1, table_length)
            previous_length = previous_length + table_length + 2
        Case 4
            table_name = LCase(ticker) & "_year_sum_tbl_txn"
            table_length = 5
            
            'initialize top left cell for table start
            Set Top_Left = Worksheets(sheetnm).Range("A4").Offset(0, previous_length + 3)
            
            'specify last cell of table (= 2 * number of tag columns)
            Set Bottom_Right = Top_Left.Offset(1, table_length)
            previous_length = previous_length + table_length + 2
        
        End Select
       
   'place table addresses into a string
   table_range = Top_Left.Address & ":" & Bottom_Right.Address
   
  'assume table does not exist
  does_table_exist = False
    
    'check if table exists (redundant) new sheet is deleted every time
    For Each ListObj In Sheets(sheetnm).ListObjects
        
        If ListObj.Name = table_name Then
            ListObj.Range.ClearFormats
            ListObj.Range.ClearContents
            does_table_exist = False
        End If
        
    Next ListObj
    
    'if table doesn't exist, make a new one
    If Not (does_table_exist) Then
        Set MainTagTbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(table_range), , xlYes)
        MainTagTbl.Name = table_name
    End If
    
    
    ' formatting changes based on what table we are creating
    Select Case x
        Case 1
            Set Main_Tbl = Worksheets(sheetnm).ListObjects(LCase(ticker) & "_income_txn")
            Main_Tbl.HeaderRowRange(1, 1) = "Date of Buy/Income"
            Main_Tbl.HeaderRowRange(1, 2) = "Buy or Income"
            Main_Tbl.HeaderRowRange(1, 3) = "Price/Coin"
            Main_Tbl.HeaderRowRange(1, 4) = "Coins Gained"
            Main_Tbl.HeaderRowRange(1, 5) = "Value Gained"
            Main_Tbl.HeaderRowRange(1, 6) = "Coins Sold (#)"
            Main_Tbl.HeaderRowRange(1, 7) = "Price Sold At"
            Main_Tbl.HeaderRowRange(1, 8) = "Realized Gain/Loss"
            Main_Tbl.HeaderRowRange(1, 9) = "Sale Number"
            Main_Tbl.HeaderRowRange(1, 10) = "Date of Sale"
            Main_Tbl.HeaderRowRange(1, 11) = ">1 Year?"
            Main_Tbl.HeaderRowRange(1, 12) = "Total Sale Gain"
            Main_Tbl.HeaderRowRange(1, 13) = "% Gain Above 1 Year"
            Main_Tbl.HeaderRowRange(1, 14) = "Total Sale Loss"
            Main_Tbl.HeaderRowRange(1, 15) = "% Loss Above 1 Year"
            
            ' make headers for tables, center across columns
            Main_Tbl.HeaderRowRange(1, 1).Offset(-1, 0).value = "Txn Data"
            Set TxnRange = Range(Main_Tbl.HeaderRowRange(1, 1).Offset(-1, 0).Address & ":" & Main_Tbl.HeaderRowRange(1, 5).Offset(-1, 0).Address)
            TxnRange.HorizontalAlignment = xlCenterAcrossSelection
            
            ' make headers for tables, center across columns
            Main_Tbl.HeaderRowRange(1, 6).Offset(-1, 0).value = "Sale Data"
            Set SaleRng = Range(Main_Tbl.HeaderRowRange(1, 6).Offset(-1, 0).Address & ":" & Main_Tbl.HeaderRowRange(1, 15).Offset(-1, 0).Address)
            SaleRng.HorizontalAlignment = xlCenterAcrossSelection
            
            SaleRng.Offset(1, 0).Interior.Color = RGB(100, 50, 0)
            
        Case 2
            Set Sale_tbl = Worksheets(sheetnm).ListObjects(LCase(ticker) & "_summary_tbl_txn")
            Sale_tbl.HeaderRowRange(1, 1) = "Sale Number:"
            Sale_tbl.HeaderRowRange(1, 2) = "Date of Sale"
            Sale_tbl.HeaderRowRange(1, 3) = "Coins Sold (#)"
            Sale_tbl.HeaderRowRange(1, 4) = "Sell Price:"
            Sale_tbl.HeaderRowRange(1, 5) = "Gain:"
            Sale_tbl.HeaderRowRange(1, 6) = "% Gain Above 1 Year"
            Sale_tbl.HeaderRowRange(1, 7) = "Loss:"
            Sale_tbl.HeaderRowRange(1, 8) = "% Loss Above 1 Year"
            
            ' make headers for tables, center across columns
            Sale_tbl.HeaderRowRange(1, 1).Offset(-1, 0).value = "Sale Summaries"
            Set Sale_rng = Range(Sale_tbl.HeaderRowRange(1, 1).Offset(-1, 0).Address & ":" & Sale_tbl.HeaderRowRange(1, 8).Offset(-1, 0).Address)
            Sale_rng.HorizontalAlignment = xlCenterAcrossSelection
            
        Case 3
            Set income_tbl = Worksheets(sheetnm).ListObjects(LCase(ticker) & "_yearly_income_txn")
            income_tbl.HeaderRowRange(1, 1) = "Year"
            income_tbl.HeaderRowRange(1, 2) = "Coins Total"
            income_tbl.HeaderRowRange(1, 3) = "Income Earned"
            
            ' make headers for tables, center across columns
            income_tbl.HeaderRowRange(1, 1).Offset(-1, 0).value = "Income Summaries"
            Set income_rng = Range(income_tbl.HeaderRowRange(1, 1).Offset(-1, 0).Address & ":" & income_tbl.HeaderRowRange(1, 3).Offset(-1, 0).Address)
            income_rng.HorizontalAlignment = xlCenterAcrossSelection
            
        Case 4
            Set year_sum_table = Worksheets(sheetnm).ListObjects(LCase(ticker) & "_year_sum_tbl_txn")
            year_sum_table.HeaderRowRange(1, 1) = "Year"
            year_sum_table.HeaderRowRange(1, 2) = "Income"
            year_sum_table.HeaderRowRange(1, 3) = "Short Gain"
            year_sum_table.HeaderRowRange(1, 4) = "Long Gain"
            year_sum_table.HeaderRowRange(1, 5) = "Short Loss"
            year_sum_table.HeaderRowRange(1, 6) = "Long Loss"
            
            ' make headers for tables, center across columns
            year_sum_table.HeaderRowRange(1, 1).Offset(-1, 0).value = "Year Summaries"
            Set year_rng = Range(year_sum_table.HeaderRowRange(1, 1).Offset(-1, 0).Address & ":" & year_sum_table.HeaderRowRange(1, 6).Offset(-1, 0).Address)
            year_rng.HorizontalAlignment = xlCenterAcrossSelection
            
        End Select
Next

' format
For Each obj In Worksheets(sheetnm).ListObjects
    obj.HeaderRowRange.WrapText = True
Next

Functions_M.SortWorksheetsAlphabetially
ThisWorkbook.Worksheets("Transaction").Move before:=ThisWorkbook.Worksheets(1)
ThisWorkbook.Worksheets("Control Center").Move after:=ThisWorkbook.Worksheets(1)
ThisWorkbook.Worksheets("Portfolio_Summary").Move after:=ThisWorkbook.Worksheets(2)

End Sub
