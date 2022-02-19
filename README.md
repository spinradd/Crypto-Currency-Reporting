# Crypto-Reporting
A VBA application to help you keep track of your cryptocurrency holdings, income, and gains/losses

## Basics
The application transforms an excel table of your cryptocurrency transactions (buys, sales, fees, and instances of generated income) and runs a report 
that summarizes your total long/short gain/loss, your mined/staked income, and your EOY holdings.

The application uses several macros linked together to accomplish this.

### Purpose
This sheet was made to help make organizing your taxes easier. As crypto is normally bought and sold in fractional amounts, I found it difficult to keep track of what
my gain/loss really was. For example, a large bought quantity of crypto is sold over multiple dates at multiple prices. Not to mention some sales would be long or short depending on when a specific quantity of crypto was aquired. 

Rather than spend days every tax season trying to keep my head above water, I developed this application. As it is a VBA application, it has its limits.
Therefore, this program can be best utilized by the amatuer investor. 

I am currently developing a python equivalent that will be much more suited to the
advanced investor who can expectedly have thousands of transactions over numerous currencies.

## Setup
You will need to enable scripting runtime to use this program's reporting capabilities.

Add the developer icon to your ribbon in excel (in File >> Settings >> Options). From then, open the VBA module. Then go to tools >> references >> and make sure
"Visual Basic for Applications" and "Scripting Runtime" are both checked.

Alternatively, you could download the VBA modules seperately from this repository and make your file yourself. If you do this, you'll need to keep the
naming conventions of the excel listobjects, sheet, and table headers the same as  **exactly** as they are in the original. To perfectly copy the original file, 

The transaction table ("transaction_tbl") should be an excel table, or a lstobject, located on the "Transaction" sheet, with at least column headers (as datatypes/"values") _exactly_ of:
<pre>
Date       |	             Type               | Ticker  | Transacted Units | Transacted Price (per unit) | Fees

date-type   | "Buy", Sell", "Fee," or "Income"  | string  |  float/int/cur   |        float/int/cur        | float/int/cur

</pre>

Additionally, the VBA module names and subroutines should remain the names as you downloaded them, unless you'd want to refactor your code, hich is difficult in VBA

Other than that, the rest of the formatting/column headers of resulting tables are auto generated.

Again, VBA code relies on names to set variables for the reports to run. If you are a VBA wiz and want to change the names of the objects
throughout the main macros in the report (there aren't that many), then feel free.

## Use
To use, you will fill out the main transaction table on the "Transaction" sheet like you would any other ledger of transactions. 
To run the report effectivel, you'll need the date of transaction, the type, the currency, the quantity, and the value. More on this below under **Setup**.

This workbook makes a few assumptions on how you fill out your main transaction table, where you will log buys, sales, fees, and earned income:
- Your transactions are accurate
- You do not sell more than you own
- There are no negative values
- You have only transactions of "Buy," "Sell," "Income," and "Fee"
  - currently there are no features for logging use of crypto as payment, nor gifts

## Report Process

The reports can be run by running the "CreateTxnTables" macro in the "Transaction_Summary" module. Or, by clicking the "Run Portfolio Tracker" button on the
"Control Center" sheet.

Once run, each cryptocurrency will recieve their own summaries sheet with 5 total charts located horizontally within the page:
- Income/Buy Table (1)
  - Table 1 itemizes your income and buy transactions.
  - The rows of Table 1 will not match your main Transaction table because other macros split up the quantities in order to calculate gains/loss. More on this below.
- Sale Table (2)
  - Table 2 shows the stats of your itemized income/buy when sold.
  - the sale table is a mechanism to calculate the gain/loss of your sales
  - as the sale macro progresses, each itemized income/buy is evaluated to see if that quantity has been sold.
   If the Income/Buy Table's transaction is included in the current sale, that crypto's quantities of gain/loss are calculated. Next a system of logic evaluates
   if there is any excess quantity from that itemized row that wasn't sold. In that case, the currency is divided and placed in an adjacent row consisting of the same transaction settings.
      - _a full description following this iterative process is located on the "Control Center" sheet in the main file_
- Sale Summaries Table (4)
  - this table summarizes the sales calculated in the Sale Table (3) for easy viewing
- Income Summaries Table (5)
  - this table summarizes the income earned (not buys)
- Year Summaries Table (6)
  - this table displays yearly summaries of the total long/short gain/loss, income of all the active years.

Each table is automatically named something when created as these naming coventions are required for other macros throughout the reporting process.

Lastly, one more sheet ("Portfolio_Summary") is produced. On this sheet are tables that summarize your long/short gain/loss and income for each different
currency across one year: Table 6. There will be as many tables on this sheet as there are active years in your Transaction table.

## Calculations
The gains and losses from sales are calculated using FIFO. Where your earliest aquired assets are sold first to satisfy a sale. Although not the most complicated and
effective way to calculate gains/losses for taxes, it is enough for the scale of this program.

This program also calculates fees in two different ways:
1. Fiat fees are added to the cost basis of the currency. Much like you would see in traditional accounting methods.
2. Fees paid in crypto are calculated as individual _sales_, subject to gains and losses. This is the most conservative way to address this aspect of accounting, which
is appropriate considering crypto is always evolving.

