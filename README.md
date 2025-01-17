**Documentation**
Purpose
This Google Apps Script:

Searches for emails from Zoro (matching specific date ranges and subjects).
Extracts relevant order data (Order Number, SKU, Order Date, Tax, etc.).
Appends that data to a specific Google Sheet.
Removes any duplicate rows in that sheet.
Formats the sheet (fonts, alignment, etc.).

**Environment / Requirements**
Google Apps Script environment (via Google Sheets or the Apps Script Editor).
Gmail access permissions (the script queries and reads emails).
The Google Spreadsheet where the data is to be appended must exist, and the script must have edit privileges to it.

**File Location**
Typically, you can add this script by going to Extensions → Apps Script in your Google Spreadsheet, or in the dedicated Google Apps Script editor.

**How to Run the Script**
Open your Google Apps Script project that is bound to your spreadsheet OR create a new standalone Google Apps Script project.
Copy and paste the entire code into the Script Editor.

**Update:**
spreadsheetUrl with the URL of your actual Google Sheet.
sheet.getSheetByName("All Zoro Order's (Don't change)") with your target sheet name if needed.
startDate and endDate as desired (currently set to fetch orders between Feb 1, 2024, and Aug 2, 2024).
Save the script.
In the Apps Script editor, click Run → Run function → select processZoroOrdersFor2024.
You may be asked to grant permissions for accessing your Gmail. Grant the necessary permissions.
Once it runs, it will search for emails matching the defined query, parse the data, and append it to your spreadsheet.

**Parameters**
maxEmails (Number): The maximum number of email threads to process (to prevent processing an excessively large batch in one run).

**Functions Breakdown**

processZoroOrdersFor2024(maxEmails)

Main function that searches Gmail for Zoro order emails between specific dates.
Extracts order information and appends it to the specified Google Sheet.
Calls removeDuplicates and formatSheet after processing.
extractOrderDataFromBody(body)

Parses the email body for details like Order Number, Order Date, Promo Amount, Tax, etc.
Returns an object with each extracted field.
extractStateFromAddress(body)

Extracts the U.S. state abbreviation from the “Shipping Address” block.
removeDuplicates(sheet)

Removes duplicate rows in columns B through I in the sheet.
formatSheet(sheet)

Applies font styling, size, and alignment to the data range.
Replaces any “0” in the Tax column (Column F) with a blank.
extractLastLineBefore(text, marker)

Utility function to get the line of text that appears right before a given marker.
extractFullNumberSequence(text)

Utility function to extract and combine all numeric sequences into one continuous string.
extractSKUFromBody(body)

Matches the “Zoro #:” line in the email body to retrieve the SKU.
Error Handling / Logging
The script uses Logger.log to display messages about missing sheets, extracted order data, etc.
If no emails are found, it logs “No emails found for October 2024.” (You can adjust this message as needed.)
Security Considerations
Ensure only authorized users have access to the Sheet and the Apps Script.
The script references the user’s Gmail. Only run this within an account you trust.
