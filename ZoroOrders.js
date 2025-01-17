function processZoroOrdersFor2024(maxEmails) {
    const spreadsheetUrl = "https://docs.google.com/spreadsheets/d/1aNMSPkBjmUUgZNO45KXtMd88ddeBYJbqDxOknaEVooM/edit?gid=1697100617#gid=1697100617";
    const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    const sheet = spreadsheet.getSheetByName("All Zoro Order's (Don't change)");
    if (!sheet) {
        Logger.log("Sheet 'All Zoro Order's (Don't change)' not found!");
        return;
    }

    const startDate = new Date("2024-2-1");
    const endDate = new Date("2024-7-26");
    const formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const formattedEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const query = `from:zoro@e.zoro.com subject:"Has Been Received" after:${formattedStartDate} before:${formattedEndDate}`;
    const threads = GmailApp.search(query);

    let count = 0;
    for (const thread of threads) {
        if (count >= maxEmails) break;
        const messages = thread.getMessages();
        for (const message of messages) {
    const body = message.getPlainBody();
    const orderData = extractOrderDataFromBody(body);
    Logger.log(`Extracted Order Data: ${JSON.stringify(orderData)}`);
    sheet.appendRow([
        "", // Leave column A empty
        orderData.orderNumber,    // WB_Order#
        orderData.sku,            // SKU
        orderData.orderDate,      // Order_Date
        orderData.promoAmount,    // Coupon
        orderData.endingIn,       // Payment_Card_Used
        orderData.tax,            // Tax
        orderData.orderTotal,     // Order_Total
        orderData.shippingSequence, // Zip_Code_(Duplicates_Check)
        orderData.state           // US_State_(For_Tax_Purpose)
      ]);
    }

        count++;
    }

    if (count === 0) {
        Logger.log("No emails found for October 2024.");
    } else {
        removeDuplicates(sheet);
        formatSheet(sheet);
    }
}

function extractOrderDataFromBody(body) {
    const orderDetails = {};

    const orderNumberMatch = body.match(/\*Order Number:\*\s*(\S+)/);
    orderDetails.orderNumber = orderNumberMatch ? orderNumberMatch[1] : "";

    const orderDateMatch = body.match(/\*Order Date:\*\s*(\d+\/\d+\/\d+)/);
    orderDetails.orderDate = orderDateMatch ? orderDateMatch[1] : "";

    const promoAmountMatch = body.match(/Promo applied: -\$(\d+\.\d{2})/);
    orderDetails.promoAmount = promoAmountMatch ? promoAmountMatch[1] : "";

    const endingInMatch = body.match(/Ending in\s*\*\*+([\d]+)/);
    orderDetails.endingIn = endingInMatch ? endingInMatch[1] : "AB ECOM LLC - Net 30";

    const taxMatch = body.match(/Tax\s*\$(\d+\.\d{2})/);
    orderDetails.tax = taxMatch ? taxMatch[1] : "";

    const orderTotalMatch = body.match(/Order Total\s*\$(\d+\.\d{2})/);
    orderDetails.orderTotal = orderTotalMatch ? orderTotalMatch[1] : "";

    const lastAddressLine = extractLastLineBefore(body, "Delivery Method:");
    orderDetails.shippingSequence = extractFullNumberSequence(lastAddressLine);

    orderDetails.state = extractStateFromAddress(body); // Extract state for tax purposes

    // Extract SKU
    const skuMatch = body.match(/Zoro\s#:\s(\S+)/);
    orderDetails.sku = skuMatch ? skuMatch[1] : "";

    return orderDetails;
}

function extractStateFromAddress(body) {
    const addressMatch = body.match(/Shipping Address:[\s\S]*?\n.*\s([A-Z]{2})(?=\s\d{5})/);
    return addressMatch ? addressMatch[1] : "";
}

function removeDuplicates(sheet) {
    const range = sheet.getRange("B:I");
    const values = range.getValues();
    const uniqueRows = [];
    const uniqueSet = new Set();
    for (const row of values) {
        const hash = row.join("-");
        if (!uniqueSet.has(hash)) {
            uniqueRows.push(row);
            uniqueSet.add(hash);
        }
    }
    range.clearContent();
    sheet.getRange(1, 2, uniqueRows.length, 8).setValues(uniqueRows);
}

function formatSheet(sheet) {
    const range = sheet.getRange("B:I");
    range.setFontFamily("Verdana");
    range.setFontSize(9);
    range.setHorizontalAlignment("center");

    // Replace '0' with empty in Tax column (Column F)
    sheet.getRange("F2:F").createTextFinder("0").replaceAllWith("");
}

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

function extractFullNumberSequence(text) {
    const match = text.match(/[0-9]+/g);
    return match ? match.join('') : '';
}

// Removed incorrect function call to testExtractState

// Replace '0' with empty in Tax column (Column F)

function extractSKUFromBody(body) {
    // Regular expression to match "Zoro #: GXXXXXXX"
    const skuMatch = body.match(/Zoro\s#:\s(\S+)/);
    return skuMatch ? skuMatch[1] : "SKU not found";
}

