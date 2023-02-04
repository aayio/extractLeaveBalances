# extractLeaveBalances.py

The text extracted from a WA Health PDF payslip contains the following block with your leave balances:

```
...

Leave Type Balance Calculated
ANNUAL LEAVE       1.23 W
MED PRACT AL ADDIT LVE      99.00 H
LONG SERVICE LEAVE       0.00 W
PROF DEV LV ACCRUING       3.02 H
SICK LEAVE - FULL PAY     111.11 H
TOIL PUBLIC HOLIDAY       0.00 H

...
```

**extractLeaveBalances.py** will process all PDFs in the working directory and generate an Excel spreadsheet (`output.xlsx`) with balances for each leave type for each pay period. The period end date is also recorded for convenience.

Don't feed it anything other than payslips.
