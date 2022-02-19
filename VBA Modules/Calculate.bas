Attribute VB_Name = "Calculate"


Sub Calc_GainsLosses(Amount_Sold As Double, Price_Sold As Double, date_sale As Date, ticker As String)
' This sub will take the transactions from Transaction sheet and calculate stats for each sale
    ' stats such as:
        ' the gain/loss value of a sell
        ' percentage of short/long gain/loss
        
ThisWorkbook.PrecisionAsDisplayed = True

' assign worksheet and listobject to vars
Dim coin_sht As Worksheet
Dim coin_tbl As ListObject

' set workbook/worksheet to var, assign crypto ticker string to var
Set Workbook = ThisWorkbook
Dim u_ticker As String
u_ticker = CStr(ticker)
Set coin_sht = Workbook.Worksheets(u_ticker & "_txn")


' change case of ticker for naming new list objects
l_ticker = LCase(ticker)

' set variable to income and sale table (Table 1/2)
Set coin_tbl = coin_sht.ListObjects(l_ticker & "_income_txn")
Set buy_date_col = coin_tbl.ListColumns("Date of Buy/Income").DataBodyRange
Set txn_type = coin_tbl.ListColumns("Buy or Income").DataBodyRange
Set p_coin = coin_tbl.ListColumns("Price/Coin").DataBodyRange
Set coins_gained = coin_tbl.ListColumns("Coins Gained").DataBodyRange
Set value_gained_col = coin_tbl.ListColumns("Value Gained").DataBodyRange
Set coins_sold = coin_tbl.ListColumns("Coins Sold (#)").DataBodyRange
Set p_sold = coin_tbl.ListColumns("Price Sold At").DataBodyRange
Set gl_col = coin_tbl.ListColumns("Realized Gain/Loss").DataBodyRange
Set sale_col = coin_tbl.ListColumns("Sale Number").DataBodyRange
Set sale_date_col = coin_tbl.ListColumns("Date of Sale").DataBodyRange
Set sl_col = coin_tbl.ListColumns(">1 Year?").DataBodyRange
Set gain_col = coin_tbl.ListColumns("Total Sale Gain").DataBodyRange
Set gain_year = coin_tbl.ListColumns("% Gain Above 1 Year").DataBodyRange
Set loss_col = coin_tbl.ListColumns("Total Sale Loss").DataBodyRange
Set loss_year = coin_tbl.ListColumns("% Loss Above 1 Year").DataBodyRange

' saled number starts at 0.
' get number of current sale and add one, this is the next sale ID number
' round sale ID to whole number
Next_Sale_Num = Application.WorksheetFunction.Max(coin_tbl.ListColumns("Sale Number").DataBodyRange) + 1
Next_Sale_Num = Application.WorksheetFunction.Round(Next_Sale_Num, 2)


' initialize sale totals
Current_Sale_Total = 0

' initialize the row where the sale ends
End_Sale_Row = 0

' initialize row count
row_count = 0

' initialize currencies total gain and total loss
Total_Gain = 0
Total_Loss = 0

' initialize long gain and loss variables
greater_year_gain = 0
greater_year_loss = 0

' for each coin or fraction of coin involved in sale
For Each cell In coins_sold
    
    ' go to next row
    row_count = 1 + row_count
    
        ' if coins_sold cell is empty that means the next quantity of coins sold in this sale
        ' will come from this rows coin_gained quantity. Evaluate whether it will be included, options:
        ' in this sale, the next sale, or no sale (end of sale)
        If coins_sold(row_count, 1).value = Empty Then
        
            ' if the total sale including this row's coin quantity is less than the amount sold in this current sale,
            ' then this row will be included in the sale
            If Current_Sale_Total + coins_gained(row_count, 1).value <= Amount_Sold Then
               
               ' since row is part of current sale, add to total
               Current_Sale_Total = Current_Sale_Total + coins_gained(row_count, 1).value
               
               ' get variable for ro's coin quantity
               Individual_Sale_Amount = coins_gained(row_count, 1).value
               
               ' get variable for coin's original dollar value
               Original_Value_of_Buy = value_gained_col(row_count, 1).value
               
               ' place price of sale under sale section
               p_sold(row_count, 1) = Price_Sold 'price of sale
               
               ' how much money the sale of this coin gave (loss or gain)
               realized_num = (Price_Sold * Individual_Sale_Amount) - Original_Value_of_Buy
               
               ' place realized num in gain/loss column under sale section
               gl_col(row_count, 1) = CCur(realized_num)
               
               ' place sale number in row under sale section
               sale_col(row_count, 1) = Next_Sale_Num
               
               ' place count amount in row under sale section
               coins_sold(row_count, 1) = Amount_Sold
               
               ' place date of sale under sale section
               sale_date_col(row_count, 1).value = date_sale
               
               ' logic to determine of long or short gain/loss
               If DateDiff("d", date_sale, buy_date_col(row_count, 1).value) > 365 Then
                    sl_col(row_count, 1).value = "Yes"
                    
                        ' add total gain/loss to appropriate total
                        If realized_num > 0 Then
                            greater_year_gain = greater_year_gain + realized_num
                            Else
                            greater_year_loss = greater_year_loss + realized_num
                        End If
                Else
                    sl_col(row_count, 1).value = "No"
                End If
                
                ' if value is gain, put under gain column, if loss loss column
                If gl_col(row_count, 1).value > 0 Then
                    Total_Gain = Application.WorksheetFunction.Round(Total_Gain + gl_col(row_count, 1).value, 2)
                ElseIf gl_col(row_count, 1) < 0 Then
                    Total_Loss = Application.WorksheetFunction.Round(Total_Loss + gl_col(row_count, 1).value, 2)
                End If
            
            ' if the cumulative total = the total coin quantity sale, then stop loop and begin to tidy and format
            ElseIf Current_Sale_Total = Amount_Sold Then
               GoTo OutSideLoop
               
            ' if the current total + the row's coin total is greater than the sale's total, insert new row
            ' place leftover coin_gained quantity in that row
            ElseIf Current_Sale_Total + coins_gained(row_count, 1).value > Amount_Sold Then
            
               
                
                ' if the row is row 1,
                If row_count = 1 Then
                
                    ' add extra row
                    coin_tbl.ListRows.Add
                    
                    ' take row 1's transaction date and copy into new row
                    
                        ' get buy date and input
                    buy_date_col(row_count + 1, 1).value = buy_date_col(row_count, 1).value
                    
                    ' copy and paste transaction type into new row
                    txn_type(row_count + 1, 1).value = txn_type(row_count, 1).value
                
                    ' amount of coins from current row to make sale whole and input
                    coins_gained(row_count + 1, 1).value = coins_gained(row_count, 1).value
                    
                    ' copy and paste price of coin into new retained coin row
                    p_coin(row_count + 1, 1).value = p_coin(row_count, 1).value
                    
                    
                    ' calculate and input value gained from sale
                    value_gained_col(row_count + 1, 1).value = value_gained_col(row_count, 1).value
        
                    
                Else
                    ' make new row add extra row where row count is
                    coin_tbl.ListRows.Add row_count
                End If
            
                ' section for copyig some coins information and calculating others
                    ' for example, you bought 3 BTC over 3 transactions at a price of 10$/BTC, and later sold 2.5 at 20$/BTC
                    ' in the context of these tables, you will sell 1 BTC from the first row/transaction, 1 BTC from the 2nd row
                    ' but you will have .5 to sell in the last. You will need someway to divy up your holdings to
                    ' retain the .5 you did not sell.
                    ' This code will give us that extra row and split the difference among the original row.
                    ' Now instead of having 3 rows of transactions with 1 BTC each, you will have 4:
                    ' row 1 = 1 BTC | Buy: $10 | Buy Date: 1/1/2000 | Sell: $20: | Sell Date: 1/5/2000 | Gain/Loss: $10 | Cumulative Sold: 1
                    ' row 2 = 1 BTC | Buy: $10 | Buy Date: 1/2/2000 | Sell: $20: | Sell Date: 1/5/2000 | Gain/Loss: $10 | Cumulative Sold: 2
                    'row 3 = .5 BTC | Buy: $10 | Buy Date: 1/3/2000 | Sell: $20: | Sell Date: 1/5/2000 | Gain/Loss: $5  | Cumulative Sold: 2.5
                    'row 4 = .5 BTC | Buy: $10 | Buy Date: 1/3/2000 |
                    ' The last two rows have coins that were cultivated from the same transaction, but since they were only partly sold
                    ' the coins will be split amongs two entries to satisfy the total sale and allocate retained coins
             
                ' get buy date and input
                buy_date_col(row_count, 1).value = buy_date_col(row_count + 1, 1).value
                
                ' amount of coins from current row to make sale whole and input
                coins_gained(row_count, 1).value = Amount_Sold - Current_Sale_Total
                
                ' calculate extra coin retained and place in new empty row and input
                coins_gained(row_count + 1, 1).value = coins_gained(row_count + 1, 1).value - coins_gained(row_count, 1).value
                
                ' copy and paste price of coin into new retained coin row
                p_coin(row_count, 1).value = p_coin(row_count + 1, 1).value
                
                
                ' calculate and input value gained from sale
                value_gained_col(row_count, 1).value = coins_gained(row_count, 1).value * p_coin(row_count, 1).value
                
                ' copy and paste transaction type into new row
                txn_type(row_count, 1).value = txn_type(row_count + 1, 1).value
                
                ' gather the quantity of the coin sold
                Individual_Sale_Amount = coins_gained(row_count, 1).value
                
                ' calculate the value of the retained (leftover) coin
                value_gained_col(row_count + 1, 1).value = p_coin(row_count + 1, 1).value * coins_gained(row_count + 1, 1).value
                
                ' get variable for the original price of the coin from purchase/mine
                Original_Value_of_Buy = value_gained_col(row_count, 1).value
            
                ' get price of sale and input in sale section
                p_sold(row_count, 1) = Price_Sold
                
                ' get realized gain/loss of sale
                realized_num = (Price_Sold * Individual_Sale_Amount) - Original_Value_of_Buy
                
                ' place realized gain/loss in column in sale section
                gl_col(row_count, 1) = CCur(realized_num)
                
                ' place sale number in sale section
                sale_col(row_count, 1) = Next_Sale_Num
                
                ' place coins sold quantity in sale section
                coins_sold(row_count, 1) = Amount_Sold
                
                ' place sale date in sale section
                sale_date_col(row_count, 1).value = date_sale
                
                ' evaluate whether row's coins were a long or short gain/loss
                If DateDiff("d", date_sale, buy_date_col(row_count, 1).value) > 365 Then
                    sl_col(row_count, 1).value = "Yes"
                        If realized_num > 0 Then
                            greater_year_gain = greater_year_gain + realized_num
                            Else
                            greater_year_loss = greater_year_loss + realized_num
                        End If
                Else
                    sl_col(row_count, 1).value = "No"
                End If
                
                ' if gain add to total, is loss add to totoal
                If gl_col(row_count, 1).value > 0 Then
                    Total_Gain = Application.WorksheetFunction.Round(Total_Gain + gl_col(row_count, 1).value, 2)
                ElseIf gl_col(row_count, 1) < 0 Then
                    Total_Loss = Application.WorksheetFunction.Round(Total_Loss + gl_col(row_count, 1).value, 2)
                End If
                
                GoTo OutSideLoop
                
            End If
        End If
Next cell

' once done with the sale...
OutSideLoop:

' get long gain percentage of total gain
If Total_Gain > 0 Then
greater_year_gain = FormatPercent(greater_year_gain / Total_Gain)
Else
greater_year_gain = FormatPercent(0)
End If

' get long short percentage of total gain
If Total_Loss < 0 Then
greater_year_loss = FormatPercent(greater_year_loss / Total_Loss)
Else
greater_year_loss = FormatPercent(0)
End If

' for each row of sale, input total long/short gain/loss
row_count = 0
For Each cell In gain_col
    row_count = row_count + 1
        If cell.value = Empty And sale_col(row_count, 1).value = Next_Sale_Num Then
            cell.value = Total_Gain
            loss_col(row_count, 1).value = Total_Loss
            gain_year(row_count, 1).value = greater_year_gain
            loss_year(row_count, 1).value = greater_year_loss
        End If
Next

' format table
With coin_tbl
    .DataBodyRange.HorizontalAlignment = xlRight
    buy_date_col.NumberFormat = "mm/dd/yyyy"
    sale_date_col.NumberFormat = "mm/dd/yyyy"
    coins_gained.NumberFormat = "#####.#####"
    coins_sold.NumberFormat = "#####.######"
    p_sold.NumberFormat = "$#,##0.00_);($#,##0.00)"
    gl_col.NumberFormat = "$#,##0.00_);($#,##0.00)"
    loss_col.NumberFormat = "$#,##0.00_);($#,##0.00)"
    gain_col.NumberFormat = "$#,##0.00_);($#,##0.00)"
    gl_col.NumberFormat = "$#,##0.00_);($#,##0.00)"
    p_sold.NumberFormat = "$#,##0.00_);($#,##0.00)"
    value_gained_col.NumberFormat = "$#,##0.00_);($#,##0.00)"
End With



End Sub

Sub UpdateSaleSummary(ticker As String)
''' create summary of sales for coin in seperate table on same sheet

ThisWorkbook.PrecisionAsDisplayed = True

' store ticker as variable for naming conventions
Dim u_ticker As String
u_ticker = CStr(ticker)
l_ticker = LCase(ticker)

' set file as variable
Set crypt_file = ThisWorkbook

' set worksheets as variables
Dim coin_sht As Worksheet
Set Workbook = ThisWorkbook
Set coin_sht = Workbook.Worksheets(u_ticker & "_txn")

' set tables as variables
Dim coin_tbl As ListObject

' set vars to income and sale table (Table 1/2)
Set coin_tbl = coin_sht.ListObjects(l_ticker & "_income_txn")
Set buy_date_col = coin_tbl.ListColumns("Date of Buy/Income").DataBodyRange
Set txn_type = coin_tbl.ListColumns("Buy or Income").DataBodyRange
Set p_coin = coin_tbl.ListColumns("Price/Coin").DataBodyRange
Set coins_gained = coin_tbl.ListColumns("Coins Gained").DataBodyRange
Set value_gained_col = coin_tbl.ListColumns("Value Gained").DataBodyRange
Set coins_sold = coin_tbl.ListColumns("Coins Sold (#)").DataBodyRange
Set p_sold = coin_tbl.ListColumns("Price Sold At").DataBodyRange
Set gl_col = coin_tbl.ListColumns("Realized Gain/Loss").DataBodyRange
Set sale_col = coin_tbl.ListColumns("Sale Number").DataBodyRange
Set sale_date_col = coin_tbl.ListColumns("Date of Sale").DataBodyRange
Set sl_col = coin_tbl.ListColumns(">1 Year?").DataBodyRange
Set gain_col = coin_tbl.ListColumns("Total Sale Gain").DataBodyRange
Set gain_year = coin_tbl.ListColumns("% Gain Above 1 Year").DataBodyRange
Set loss_col = coin_tbl.ListColumns("Total Sale Loss").DataBodyRange
Set loss_year = coin_tbl.ListColumns("% Loss Above 1 Year").DataBodyRange

' set vars to new sale summary table (Table 3)
Dim coin_summary_tbl As ListObject
Set coin_summary_tbl = coin_sht.ListObjects(l_ticker & "_summary_tbl_txn")
Set sum_sale_num = coin_summary_tbl.ListColumns("Sale Number:").DataBodyRange
Set sum_sale_date = coin_summary_tbl.ListColumns("Date of Sale").DataBodyRange
Set sum_sale_gain = coin_summary_tbl.ListColumns("Gain:").DataBodyRange
Set sum_sale_loss = coin_summary_tbl.ListColumns("Loss:").DataBodyRange
Set sum_sale_coin = coin_summary_tbl.ListColumns("Coins Sold (#)").DataBodyRange
Set sum_gain_year = coin_summary_tbl.ListColumns("% Gain Above 1 Year").DataBodyRange
Set sum_loss_year = coin_summary_tbl.ListColumns("% Loss Above 1 Year").DataBodyRange
Set sum_coin_price = coin_summary_tbl.ListColumns("Sell Price:").DataBodyRange


' with the new table...
With coin_summary_tbl
        
        'Check If any data exists in the table, if there is clear the data
        If Not .DataBodyRange Is Nothing Then
            
            'Clear Content from the table
            
            .DataBodyRange.ClearContents
        End If
        
    End With


' intialize counting variables
coin_row_count = 0
CurrentSale = 1
summary_row_count = 1

' for each cell in the sale column of the main sale table
For Each cell In sale_col
    coin_row_count = coin_row_count + 1
    
    ' if first row of main table, copy sale # 1 info to summary sale table
    If coin_row_count = 1 Then
         sum_sale_num(summary_row_count, 1).value = cell.value
         sum_sale_date(summary_row_count, 1).value = sale_date_col(coin_row_count, 1).value
         sum_sale_gain(summary_row_count, 1).value = gain_col(coin_row_count, 1).value
         sum_sale_loss(summary_row_count, 1).value = loss_col(coin_row_count, 1).value
         'sum_sale_total(summary_row_count, 1).value = sale_date_col(coin_row_count, 1).value
         sum_sale_coin(summary_row_count, 1).value = coins_sold(coin_row_count, 1).value
         sum_gain_year(summary_row_count, 1).value = gain_year(coin_row_count, 1).value
         sum_loss_year(summary_row_count, 1).value = loss_year(coin_row_count, 1).value
         sum_coin_price(summary_row_count, 1).value = p_sold(coin_row_count, 1).value
         summary_row_count = summary_row_count + 1
    
    ' if not row one, see if the sale of current row is equal to the previous sale in the summary sale table,
    ' if they are not the same it is a new sale, copy and paste the sale data to the summary sale table
    ElseIf sum_sale_num(summary_row_count - 1, 1).value <> cell.value And cell <> blank Then
         coin_summary_tbl.ListRows.Add
         sum_sale_num(summary_row_count, 1).value = cell
         sum_sale_date(summary_row_count, 1).value = sale_date_col(coin_row_count, 1).value
         sum_sale_gain(summary_row_count, 1).value = gain_col(coin_row_count, 1).value
         sum_sale_loss(summary_row_count, 1).value = loss_col(coin_row_count, 1).value
         'sum_sale_total(summary_row_count, 1).Value = sale_date_col(coin_row_count, 1).Value
         sum_sale_coin(summary_row_count, 1).value = coins_sold(coin_row_count, 1).value
         sum_gain_year(summary_row_count, 1).value = gain_year(coin_row_count, 1).value
         sum_loss_year(summary_row_count, 1).value = loss_year(coin_row_count, 1).value
         sum_coin_price(summary_row_count, 1).value = p_sold(coin_row_count, 1).value
         
         summary_row_count = summary_row_count + 1
    End If
Next cell

' format table
With coin_summary_tbl
    .DataBodyRange.HorizontalAlignment = xlRight
    sum_sale_loss.NumberFormat = "$#,##0.00_);($#,##0.00)"
    sum_sale_gain.NumberFormat = "$#,##0.00_);($#,##0.00)"
    sum_sale_date.NumberFormat = "mm/dd/yyyy"
End With
    
End Sub

Sub Calc_Income(ticker As String)
' summarize earned income per coin for each year of activity

ThisWorkbook.PrecisionAsDisplayed = True


' initialize ticker for naming conventions
Dim u_ticker As String
u_ticker = CStr(ticker)
l_ticker = LCase(ticker)

' set variable to workbook
Set Workbook = ThisWorkbook

' set variable to worksheet
Dim coin_sht As Worksheet
Set coin_sht = Workbook.Worksheets(u_ticker & "_txn")

' set vars to income and sale table (Tables 1/2)
Dim coin_tbl As ListObject
Set coin_tbl = coin_sht.ListObjects(l_ticker & "_income_txn")
Set buy_date_col = coin_tbl.ListColumns("Date of Buy/Income").DataBodyRange
Set txn_type = coin_tbl.ListColumns("Buy or Income").DataBodyRange
Set p_coin = coin_tbl.ListColumns("Price/Coin").DataBodyRange
Set coins_gained = coin_tbl.ListColumns("Coins Gained").DataBodyRange
Set value_gained_col = coin_tbl.ListColumns("Value Gained").DataBodyRange
Set coins_sold = coin_tbl.ListColumns("Coins Sold (#)").DataBodyRange
Set p_sold = coin_tbl.ListColumns("Price Sold At").DataBodyRange
Set gl_col = coin_tbl.ListColumns("Realized Gain/Loss").DataBodyRange
Set sale_col = coin_tbl.ListColumns("Sale Number").DataBodyRange
Set sale_date_col = coin_tbl.ListColumns("Date of Sale").DataBodyRange
Set sl_col = coin_tbl.ListColumns(">1 Year?").DataBodyRange
Set gain_col = coin_tbl.ListColumns("Total Sale Gain").DataBodyRange
Set gain_year = coin_tbl.ListColumns("% Gain Above 1 Year").DataBodyRange
Set loss_col = coin_tbl.ListColumns("Total Sale Loss").DataBodyRange
Set loss_year = coin_tbl.ListColumns("% Loss Above 1 Year").DataBodyRange

' set variable to new income table
Set income_tbl = coin_sht.ListObjects(l_ticker & "_yearly_income_txn")

Dim yeararr() As Variant

' for each year from 2021 to current year...
row_count = 0

' evaluate earliest year when income as earned
min_year = Year(Application.WorksheetFunction.Min(buy_date_col))

' from year 1 of activity to present year
For x = min_year To Year(Date)
    
    row_count = row_count + 1

    ' add row if the number of rows in income table are less than the number of years in range (2021-Present)
    If row_count > income_tbl.DataBodyRange.Rows.Count Then
    income_tbl.ListRows.Add
    End If
    
    ' assign year value to row
    income_tbl.ListColumns("Year").DataBodyRange(row_count, 1).value = x
Next x

income_row_count = 0
coin_row_count = 0

' for each year...
For Each cell In income_tbl.ListColumns("year").DataBodyRange

    ' get next row
    income_row_count = income_row_count + 1
    
    ' initialize annual sums
    year_total = 0
    coin_total = 0
    
    ' for each date in the main coin table
    For Each datum In buy_date_col
    
        ' next row
        coin_row_count = coin_row_count + 1
        
            ' if year is current year in loop and transaction is income (bought or mined)...
            If Year(datum) = cell.value And txn_type(coin_row_count, 1).value = "Income" Then
                
                ' add to the yearly totals of bought/mined crypto
                year_total = year_total + value_gained_col(coin_row_count, 1).value
                coin_total = coin_total + coins_gained(coin_row_count, 1).value
            End If
    Next datum
    
    ' if there was no income for the coin in this year, set to 0, otherwise place the totals
    If coin_total = 0 Then
        income_tbl.ListColumns("Coins Total").DataBodyRange(income_row_count, 1).value = 0
    Else
        income_tbl.ListColumns("Coins Total").DataBodyRange(income_row_count, 1).value = coin_total
        income_tbl.ListColumns("Coins Total").DataBodyRange(income_row_count, 1).NumberFormat = "####.###"
    End If
    
    
    income_tbl.ListColumns("Income Earned").DataBodyRange(income_row_count, 1).value = year_total
    

Next cell

' format table
With income_tbl.DataBodyRange
    .HorizontalAlignment = xlRight
    income_tbl.ListColumns("Income Earned").DataBodyRange.NumberFormat = "$#,##0.00_);($#,##0.00)"
End With
End Sub

Sub Calculate_Summary(ticker As String)
''' calculates total summary from income table for each coin

ThisWorkbook.PrecisionAsDisplayed = True

' initialize ticker as variable for namng convention
u_ticker = ticker
l_ticker = LCase(ticker)
Dim coin_sht As Worksheet
Dim coin_tbl As ListObject

' set variable to workbook
Set Workbook = ThisWorkbook

' set variable to worksheet and year summary table
Set coin_sht = Workbook.Worksheets(u_ticker & "_txn")

' set var to year summary table (Table 5)
Set yearly_tbl = coin_sht.ListObjects(l_ticker & "_year_sum_tbl_txn")
Set year_col = yearly_tbl.ListColumns("Year").DataBodyRange
Set income_col = yearly_tbl.ListColumns("Income").DataBodyRange
Set income_short_g = yearly_tbl.ListColumns("Short Gain").DataBodyRange
Set income_long_g = yearly_tbl.ListColumns("Long Gain").DataBodyRange
Set income_short_l = yearly_tbl.ListColumns("Short Loss").DataBodyRange
Set income_long_l = yearly_tbl.ListColumns("Long Loss").DataBodyRange

' set variable to income summary table (Table 4)
Set income_tbl = coin_sht.ListObjects(l_ticker & "_yearly_income_txn")
Set income_year = income_tbl.ListColumns("Year").DataBodyRange
Set income_total = income_tbl.ListColumns("Income Earned").DataBodyRange

' set variable to sale summary table (Table 3)
Set Sale_tbl = coin_sht.ListObjects(l_ticker & "_summary_tbl_txn")
Set sum_sale_num = Sale_tbl.ListColumns("Sale Number:").DataBodyRange
Set sum_sale_date = Sale_tbl.ListColumns("Date of Sale").DataBodyRange
Set sum_sale_gain = Sale_tbl.ListColumns("Gain:").DataBodyRange
Set sum_sale_loss = Sale_tbl.ListColumns("Loss:").DataBodyRange
'Set sum_sale_total = Sale_tbl.ListColumns("Total:").DataBodyRange
Set sum_sale_coin = Sale_tbl.ListColumns("Coins Sold (#)").DataBodyRange
Set sum_gain_year = Sale_tbl.ListColumns("% Gain Above 1 Year").DataBodyRange
Set sum_loss_year = Sale_tbl.ListColumns("% Loss Above 1 Year").DataBodyRange
Set sum_coin_price = Sale_tbl.ListColumns("Sell Price:").DataBodyRange

' intitialize counter variables
year_count = 0
row_count = 0

' for each year in income table


' get minimum and maximum dates where income was earned
min_year_income = Application.WorksheetFunction.Min(income_year)
max_year_income = Application.WorksheetFunction.Max(income_year)

' get minimum and maximum dates where sales were made
min_year_sale = Year(Application.WorksheetFunction.Min(sum_sale_date))
max_year_sale = Year(Application.WorksheetFunction.Max(sum_sale_date))

' evaluate earliest and latest year
If min_year_income <= min_year_sale Then
    min_year = min_year_income
    Else
    min_year = min_year_income
End If

If max_year_income <= max_year_sale Then
    max_year = max_year_income
    Else
    max_year = max_year_income
End If

' for all years of activity
For x = min_year To max_year
    row_count = row_count + 1
    
    ' if row count is less than the number of years in range, add rows
    If row_count > yearly_tbl.DataBodyRange.Rows.Count Then
    yearly_tbl.ListRows.Add
    End If
    year_col(row_count, 1).value = x

    ' initialize summary totals
    year_long_gain = 0
    year_long_loss = 0
    year_short_gain = 0
    year_short_loss = 0
    sale_count = 0
    income_count = 0
    year_income_total = 0
    year_summary = x
    year_count = year_count + 1

    ' for each sale in summary sale table, add to yearly sale summary
    For Each Sale In sum_sale_date
        sale_count = sale_count + 1
        If Year(Sale) = year_summary Then
    
            year_long_gain = year_long_gain _
                            + (sum_sale_gain(sale_count, 1).value * sum_gain_year(sale_count, 1).value)
            year_short_gain = year_short_gain _
                            + (sum_sale_gain(sale_count, 1).value * (1 - sum_gain_year(sale_count, 1).value))
            year_long_loss = year_long_loss _
                            + (sum_sale_loss(sale_count, 1).value * sum_loss_year(sale_count, 1).value)
            year_short_loss = year_short_loss _
                            + (sum_sale_loss(sale_count, 1).value * (1 - sum_loss_year(sale_count, 1).value))
        End If
    Next
    
    ' for each year in income table, add incomes for annual total
    For Each income_yr In income_year
        income_count = income_count + 1
        If income_yr = year_summary Then
        year_income_total = year_income_total + income_total(income_count, 1).value
        End If
    Next
    
    ' place acquired values in table
    income_col(year_count, 1).value = year_income_total
    income_short_g(year_count, 1).value = year_short_gain
    income_short_l(year_count, 1).value = year_short_loss
    income_long_g(year_count, 1).value = year_long_gain
    income_long_l(year_count, 1).value = year_long_loss
    
Next x

With yearly_tbl
    .DataBodyRange.HorizontalAlignment = xlRight
    income_short_g.NumberFormat = "$#,##0.00_);($#,##0.00)"
    income_short_l.NumberFormat = "$#,##0.00_);($#,##0.00)"
    income_long_g.NumberFormat = "$#,##0.00_);($#,##0.00)"
    income_long_l.NumberFormat = "$#,##0.00_);($#,##0.00)"
End With

End Sub

Sub delete_sheets()

For Each Sheet In ThisWorkbook.Worksheets
    DisplayAlerts = False
    Debug.Print (Sheet.Name)
    If Right(Sheet.Name, 4) = "_txn" Then
        Sheet.Delete
    End If
Next Sheet

End Sub


