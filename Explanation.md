
**Explanation / Description of the Code**


function processZoroOrdersFor2024(maxEmails) {
    // The Google Sheets URL where order data will be stored
    const spreadsheetUrl = "https://docs.google.com/spreadsheets/d/1aNMSPkBjmUUgZNO45KXtMd88ddeBYJbqDxOknaEVooM/edit?gid=1697100617#gid=1697100617";
    
    // Open the spreadsheet by its URL
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    
    // Retrieve the specific sheet by name
    const sheet = spreadsheet.getSheetByName("All Zoro Order's (Don't change)");
    if (!sheet) {
        Logger.log("Sheet 'All Zoro Order's (Don't change)' not found!");
        return; // Stop execution if sheet not found
        }

    // Define the date range for searching Zoro orders
    const startDate = new Date("2024-2-1");  // Feb 1, 2024
    const endDate = new Date("2024-8-2");    // Aug 2, 2024

    // Format the dates in YYYY/MM/DD format for the Gmail search query
    const formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const formattedEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

    // Build the Gmail query
    // Looks for emails from zoro@e.zoro.com with subject containing "Has Been Received"
    // and within the specified date range
    const query = `from:zoro@e.zoro.com subject:"Has Been Received" after:${formattedStartDate} before:${formattedEndDate}`;

    // Search Gmail threads based on this query
    const threads = GmailApp.search(query);

    let count = 0;
    for (const thread of threads) {
        // Only process up to maxEmails
        if (count >= maxEmails) break;

        // Get all messages in the thread
        const messages = thread.getMessages();
        for (const message of messages) {
            // Extract plain text body
            const body = message.getPlainBody();
            // Extract order info using the helper function
            const orderData = extractOrderDataFromBody(body);

            // Log the extracted data for debugging
            Logger.log(`Extracted Order Data: ${JSON.stringify(orderData)}`);

            // Append data to the Google Sheet in a new row
            sheet.appendRow([
                "",                     // Column A is left blank
                orderData.orderNumber, // B: WB_Order#
                orderData.sku,         // C: SKU
                orderData.orderDate,   // D: Order_Date
                orderData.promoAmount, // E: Coupon
                orderData.endingIn,    // F: Payment_Card_Used
                orderData.tax,         // G: Tax
                orderData.orderTotal,  // H: Order_Total
                orderData.shippingSequence, // I: Zip_Code_(Duplicates_Check)
                orderData.state        // J: US_State_(For_Tax_Purpose)
            ]);
        }
        count++;
    }

    // If no messages were found, log a message
    if (count === 0) {
        Logger.log("No emails found for October 2024.");
    } else {
        // Remove duplicates from the sheet
        removeDuplicates(sheet);
        // Apply formatting after data is appended
        formatSheet(sheet);
    }
    }

**What happens here?**

We define start and end dates for 2024.
We construct a Gmail query to find emails from Zoro within that date range.
For each email found (up to maxEmails), we parse its body to extract order data.
We then append that data to the Google Sheet, followed by a routine to remove duplicates and format the sheet.
//


function extractOrderDataFromBody(body) {
    const orderDetails = {};

    // 1. Order Number
    const orderNumberMatch = body.match(/\*Order Number:\*\s*(\S+)/);
    orderDetails.orderNumber = orderNumberMatch ? orderNumberMatch[1] : "";

    // 2. Order Date
    const orderDateMatch = body.match(/\*Order Date:\*\s*(\d+\/\d+\/\d+)/);
    orderDetails.orderDate = orderDateMatch ? orderDateMatch[1] : "";

    // 3. Promo/Coupon Amount
    const promoAmountMatch = body.match(/Promo applied: -\$(\d+\.\d{2})/);
    orderDetails.promoAmount = promoAmountMatch ? promoAmountMatch[1] : "";

    // 4. Payment Card Info (or default to "AB ECOM LLC - Net 30")
    const endingInMatch = body.match(/Ending in\s*\*\*+([\d]+)/);
    orderDetails.endingIn = endingInMatch ? endingInMatch[1] : "AB ECOM LLC - Net 30";

    // 5. Tax
    const taxMatch = body.match(/Tax\s*\$(\d+\.\d{2})/);
    orderDetails.tax = taxMatch ? taxMatch[1] : "";

    // 6. Order Total
    const orderTotalMatch = body.match(/Order Total\s*\$(\d+\.\d{2})/);
    orderDetails.orderTotal = orderTotalMatch ? orderTotalMatch[1] : "";

    // 7. Shipping ZIP (from the line before "Delivery Method:")
    const lastAddressLine = extractLastLineBefore(body, "Delivery Method:");
    orderDetails.shippingSequence = extractFullNumberSequence(lastAddressLine);

    // 8. State
    orderDetails.state = extractStateFromAddress(body);

    // 9. SKU
    const skuMatch = body.match(/Zoro\s#:\s(\S+)/);
    orderDetails.sku = skuMatch ? skuMatch[1] : "";

    return orderDetails;
}

**Explanation:**
This function looks for specific patterns in the email text to capture data. The code uses regular expressions (body.match(...)) to find items like “Order Number,” “Order Date,” “Tax,” etc. Each matched value is stored in an object, which we then pass back to the main function.
//
function extractStateFromAddress(body) {
    // Finds the address block, then captures the two-letter state code before the ZIP code
    const addressMatch = body.match(/Shipping Address:[\s\S]*?\n.*\s([A-Z]{2})(?=\s\d{5})/);
    return addressMatch ? addressMatch[1] : "";
}

**Explanation:**

We assume the “Shipping Address:” section has a line that ends with two-letter state code, followed by a five-digit ZIP.
The captured group ([A-Z]{2}) is the state abbreviation.
//
function removeDuplicates(sheet) {
    const range = sheet.getRange("B:I");
    const values = range.getValues();

    const uniqueRows = [];
    const uniqueSet = new Set();

    for (const row of values) {
        // Build a hash string representing the row
        const hash = row.join("-");
        if (!uniqueSet.has(hash)) {
            uniqueRows.push(row);
            uniqueSet.add(hash);
        }
    }
    // Clear content in the original range
    range.clearContent();
    // Write back only unique rows
    sheet.getRange(1, 2, uniqueRows.length, 8).setValues(uniqueRows);
    }

**Explanation:**

Gets all data from columns B to I.
Creates a Set of row “hashes” to identify duplicates.
Clears old data, then rewrites only unique rows.
//
function formatSheet(sheet) {
    const range = sheet.getRange("B:I");
    range.setFontFamily("Verdana");
    range.setFontSize(9);
    range.setHorizontalAlignment("center");

    // Replace '0' with empty in the Tax column (Column F)
    sheet.getRange("F2:F").createTextFinder("0").replaceAllWith("");
    }

**Explanation:**

Applies formatting to columns B through I.
Any “0” values in the tax column (F) are replaced with an empty string.
//
function extractLastLineBefore(text, marker) {
    const parts = text.split('\n');
    let secondLastLine = '';
    let lastSeenLine = '';

    for (const line of parts) {
        if (line.includes(marker)) {
            break;
        }
        secondLastLine = lastSeenLine;
        lastSeenLine = line.trim();
    }
    return secondLastLine;
    }
  
**Explanation:**

Splits the email body into lines.
Iterates over each line until it encounters the marker text, e.g. “Delivery Method:”.
Returns the line immediately before the marker line. This is often where we find shipping ZIP details.
//
function extractFullNumberSequence(text) {
    const match = text.match(/[0-9]+/g);
    return match ? match.join('') : '';
    }
    
**Explanation:**

Finds all numeric sequences in the text and concatenates them into one string.
Useful for extracting ZIP codes or any numeric data from partial text.
//
function extractSKUFromBody(body) {
    // Regular expression to match "Zoro #: GXXXXXXX"
    const skuMatch = body.match(/Zoro\s#:\s(\S+)/);
    return skuMatch ? skuMatch[1] : "SKU not found";
    }

**Explanation:**

Another **helper** function to specifically capture the SKU in the format “Zoro #: ABC1234”.
Returns “SKU not found” if nothing matches.
