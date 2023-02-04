#! python3
# extractLeaveBalances.py - Extract leave balances from WA Health PDF payslips.

import PyPDF2, os, re, openpyxl

from pathlib import Path

# Define Period End Date and Period Number strings

periodEndDateStr = 'Period End Date'

periodNumberStr = 'Period Number'

# Define leave types

leaveTypes = ['ANNUAL LEAVE',
'MED PRACT AL ADDIT LVE',
'LONG SERVICE LEAVE',
'PROF DEV LV ACCRUING',
'SICK LEAVE - FULL PAY',
'TOIL PUBLIC HOLIDAY']

# Create an excel spreadsheet with the relevant column headings

leaveBalancesSpreadsheet = openpyxl.Workbook()

sheet = leaveBalancesSpreadsheet.active

sheet.title = 'Leave Balances'

headings = [periodEndDateStr, periodNumberStr] + leaveTypes

headingRowNum = 1

for columnNum in range(1, len(headings) + 1): # need to do + 1 to include the last one
    headingsIndex = columnNum - 1
    sheet.cell(row=headingRowNum, column=columnNum).value = headings[headingsIndex]

# Define regexes

periodNumberRegex = re.compile(r'(?<=Period Number:)\s*[0-9]{3}')

periodEndDateRegex = re.compile(r'(?<=Period End Date:)\s*[0-9]{2}-[0-9]{2}-[0-9]{4}')

leaveBalancesTextRegex = re.compile(r'(?<=Leave Type Balance Calculated).*(?=Leave balances)', re.S)

leaveTypeRegexes = {}

for leaveType in leaveTypes:
    leaveTypeRegex = re.compile(r'(?<=' + leaveType + r')\s+[0-9]+\.[0-9]+')

    leaveTypeRegexes[leaveType] = leaveTypeRegex

# Define other parameters

pageContainingLeaveBalancesIndex = 0

rowToWrite = 2 # first row to write

# Iterate over PDFs in the working directory

for payslipPathObject in Path.cwd().glob('*.pdf'):
    # Create empty dict to store data

    extractedData = {}

    # Open the payslip file and extract text

    payslipObject = open(payslipPathObject, 'rb') # 'rb' means 'read binary' mode

    pdfReader = PyPDF2.PdfReader(payslipObject)

    pageObj = pdfReader.pages[pageContainingLeaveBalancesIndex]

    pageText = pageObj.extract_text()

    # Extract payslip period number

    periodNumberMatchObjects = periodNumberRegex.search(pageText)

    periodNumber = periodNumberMatchObjects.group().strip()

    # Extract payslip date as YYYY-MM-DD (international format)

    periodEndDateMatchObjects = periodEndDateRegex.search(pageText)

    periodEndDateYear = periodEndDateMatchObjects.group()[-4:]

    periodEndDateMonth = periodEndDateMatchObjects.group()[-7:-5]

    periodEndDateDay = periodEndDateMatchObjects.group()[-10:-8]

    periodEndDateInternationalFormat = '-'.join([periodEndDateYear, periodEndDateMonth, periodEndDateDay])

    # Extract just the leave balances part of the text

    leaveBalancesTextMatchObjects = leaveBalancesTextRegex.search(pageText)

    leaveBalancesText = leaveBalancesTextMatchObjects.group()

    # Loop over leaveTypes and extract leave balances

    for leaveType, leaveTypeRegex in leaveTypeRegexes.items():
        leaveBalanceMatchObjects = leaveTypeRegex.search(leaveBalancesText)

        leaveBalance = leaveBalanceMatchObjects.group().strip()

        extractedData[leaveType] = leaveBalance

    # Write data to spreadsheet

    periodEndDateCell = sheet.cell(row=rowToWrite, column=1)

    periodEndDateCell.value = periodEndDateInternationalFormat

    periodNumberCell = sheet.cell(row=rowToWrite, column=2)

    periodNumberCell.value = periodNumber
    periodNumberCell.data_type = 'n'

    for columnNum in range(3, sheet.max_column + 1):
        heading = sheet.cell(row=1, column=columnNum).value
        if heading in extractedData:
            leaveBalanceCell = sheet.cell(row=rowToWrite, column=columnNum)
            leaveBalanceCell.value = extractedData[heading]
            leaveBalanceCell.data_type = 'n'
    
    # Close the file

    payslipObject.close()

    # Increase rowToWrite

    rowToWrite += 1

# Save the spreadsheet

leaveBalancesSpreadsheet.save('output.xlsx')
