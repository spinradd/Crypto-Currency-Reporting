# Crypto-Currency-Reporting
A VBA application to help you keep track of your crypto currency holdings, income, and gains/losses

## Basics
The application transforms an excel table of your crypto currency transactions (buys, sales, and instances of generated income) and runs a report 
that summarizes your total long/short gain/loss, your mined/staked income, and your EOY holdings.

The application uses several macros linked together to accomplish this.

### Purpose
This sheet was made to help make organizing your taxes easier. As crypto is normally bought and sold in fractional amount, I found it difficult to keep track of what
my gain/loss really was, as the crypto I sold was bought over multiple dates at multiple prices. Not to mention some sells would be long or short depending on when a specific
quantity of crypto was aquired. 

Rather than spend days every tax season trying to keep ym head above water, I developed this application. As it is a VBA application, it has its limits.
Therefore, this program can be best utilized by the amatuer investor. 

I am currently developing a python equivalent that will be much more suited to the
advanced investor with many different assets.

## Setup
You will need to enable scripting runtime to use this report.
Add the developer icon to your ribbon in excel (in Options). From then, open the VBA module. Then go to tools >> references >> and make sure
"Visual Basic for Applications" and "Scripting Runtime" are both checked.

## Use
To use, you will fill out the main transaction table like you would any other ledger of transactions. You'll need the date of transaction, the type,
the currency, th quantity, and the value. More on this below.

This workbook makes a few assumptions on how you fill out your main transaction table, where you will log buys, sales, and earned income:
- Your transactions are accurate
- You do not sell more than you own
- There are no negative values
- You have only transactions of "Buy," "Sell," and "Income
  - currently there is no feature for tracking use of crypto as payment, nor gifts

## Formatting
Because excel listobjects are identified by names, you should keep several things as they appear **exactly**, unless you are a VBA wiz and want to change the column headers
throughout the main macros in the report (there aren't that many).

The transaction table ("transaction_tbl") should be located on the "Transaction" sheet, with at least column headers (as datatypes/"values") _exactly_ of:
<pre>
Date (UTC)  |	           Type               | Ticker  | Transacted Units | Transacted Price (per unit)

date-type   | "Buy", Sell", "Fee," or "Income"  | string  |  float/int/cur   | float/int/cur

</pre>

Other than that, the rest of the formatting/column headers of resulting tables are auto generated.

## Report Process
Each crypto currency will recieve their own summaries sheet with 5 total charts located horizontally within the page:
- Income/Buy Table (1 & 2)
  - this table itemizes your income and buy transactions of the currency of the sheet.
  - these rows will not match your main Transaction table because other macros split up the quantities in order to calculate gains/loss
- Sale Table (3)
  - the sale table is a mechanism to calculate the gain/loss of your sales
  - as the sale macro progresses, each registered sale in the Transaction table needs to be filled 
  by the quantities found in the Income/Buy Table (1 & 2). If the Income/Buy Table's row is included in the current sale, that quantities gain/loss is calculated, 
  and following a system of logic any excess quantity of that currency is divided and placed in an adjacent row consisting of the same transaction settings.
      - _a full description following this iterative process is located on the "Control Center" sheet in the main file_
- Sale Summaries Table (4)
  - this table summarizes the sales calculated in the Sale Table (3) for easy viewing
- Income Summaries Table (5)
  - this table summarizes the income earned
- Year Summaries Table (6)
  - this table displays yearly summaries of the total long/short gain/loss, income of all the active years.

Each table is automatically named something when created as these naming coventions are required for other macros throughout the reporting process.

Lastly, one more sheet ("Portfolio_Summary") is produced. On this sheet are tables that summarize your long/short gain/loss and income for each different
currency across one year: Table 6. There will be as many tables on this sheet as there are active years in your Transaction sheet.

## Calculations
The gains and losses from sales are calculated using FIFO. Where your earliest aquired assets are sold first to satisfy the sale. Although not the most complicated and
effective way to calculate gains/losses, it is enough for the scale of this program.

This program also 

