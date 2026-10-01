# Fursat-e-Barkat - Secure Google Sheets Webhook Backend

This setup includes **built-in server-side fraud protection**, **price tampering detection**, **duplicate UTR checks**, and **automated calculation integrity**.

---

## 1. Create the Google Sheet

1. Go to [Google Sheets](https://sheets.new) and create a new spreadsheet.
2. Name it **`Fursat-e-Barkat Registrations`**.
3. In **Row 1**, set these exact column headers:
   - **A1:** `Timestamp (IST)`
   - **B1:** `Verification Status`
   - **C1:** `Full Name`
   - **D1:** `Email`
   - **E1:** `WhatsApp Number`
   - **F1:** `Tickets Booked`
   - **G1:** `Paid Tickets`
   - **H1:** `Free Tickets`
   - **I1:** `Expected Amount (₹)`
   - **J1:** `Client Submitted Amount (₹)`
   - **K1:** `UTR / Transaction ID`
   - **L1:** `Security Notes`

---

## 2. Add the Fraud-Protected Google Apps Script

1. In Google Sheets, click **Extensions** > **Apps Script**.
2. Replace all code in `Code.gs` with this hardened script:

```javascript
/**
 * Fursat-e-Barkat - Secure Registration Webhook
 * Features:
 *  - Server-side price recalculation (prevents browser tampering)
 *  - Duplicate UTR check (prevents reusing same payment receipt)
 *  - Automated status tagging (PENDING, DUPLICATE_UTR, PRICE_MISMATCH)
 *  - Lock service to prevent concurrent write collisions
 */

var BASE_PRICE = 189;

function doPost(e) {
  var lock = LockService.getScriptLock();
  // Wait up to 10 seconds for other operations to finish
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getActiveSheet();

    // 1. Sanitize & extract inputs
    var name = (e.parameter.Name || "").toString().trim();
    var email = (e.parameter.Email || "").toString().trim();
    var phone = (e.parameter.Phone || "").toString().trim();
    var utr = (e.parameter.UTR || "").toString().trim().toUpperCase();
    var submittedAmount = parseFloat(e.parameter.TotalAmount) || 0;
    var rawTickets = parseInt(e.parameter.Tickets, 10);
    var tickets = (isNaN(rawTickets) || rawTickets < 1) ? 1 : rawTickets;

    // 2. SERVER-SIDE RECALCULATION (Prevents client-side price hacks)
    var freeTickets = Math.floor(tickets / 5);
    var paidTickets = tickets - freeTickets;
    var expectedAmount = paidTickets * BASE_PRICE;

    // 3. Security checks
    var securityNotes = [];
    var status = "PENDING_VERIFICATION";

    // Check A: Price Tampering
    if (submittedAmount !== expectedAmount) {
      status = "⚠️ FLAGGED: PRICE_MISMATCH";
      securityNotes.push("Submitted ₹" + submittedAmount + " but expected ₹" + expectedAmount);
    }

    // Check B: Duplicate UTR Detection (Scan column K)
    var data = sheet.getDataRange().getValues();
    var isDuplicateUtr = false;
    for (var i = 1; i < data.length; i++) {
      var existingUtr = (data[i][10] || "").toString().trim().toUpperCase();
      if (existingUtr && existingUtr === utr) {
        isDuplicateUtr = true;
        break;
      }
    }

    if (isDuplicateUtr) {
      status = "🚨 FLAGGED: DUPLICATE_UTR";
      securityNotes.push("UTR " + utr + " was already used in a previous booking!");
    }

    if (!isDuplicateUtr && submittedAmount === expectedAmount) {
      securityNotes.push("Valid calculation. Awaiting manual bank/UPI credit verification.");
    }

    // 4. IST Timestamp
    var timestampIST = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");

    // 5. Append verified row to Google Sheet
    sheet.appendRow([
      timestampIST,
      status,
      name,
      email,
      phone,
      tickets,
      paidTickets,
      freeTickets,
      expectedAmount,
      submittedAmount,
      utr,
      securityNotes.join(" | ")
    ]);

    // Optional: Highlight flagged rows in light red
    var lastRow = sheet.getLastRow();
    if (status.indexOf("FLAGGED") !== -1) {
      sheet.getRange(lastRow, 1, 1, 12).setBackground("#fde8e8");
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "success",
        verificationStatus: status,
        expectedAmount: expectedAmount
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
```

---

## 3. Deploy as Web App

1. In Apps Script, click **Deploy** > **New deployment**.
2. Click the gear icon (**⚙️**) > **Web app**.
3. Settings:
   - **Description:** `Fursat-e-Barkat Secure Webhook`
   - **Execute as:** `Me (your email)`
   - **Who has access:** `Anyone`
4. Click **Deploy** and authorize the script.
5. Copy the **Web App URL** (`https://script.google.com/macros/s/.../exec`).

---

## 4. Paste URL into `index.html`

Open `index.html` and update line ~540:

```javascript
const GOOGLE_SCRIPT_WEBHOOK_URL = "https://script.google.com/macros/s/YOUR_DEPLOYED_URL_HERE/exec";
```

---

## 5. How You Verify Bookings as the Organizer

1. **User completes booking on the website**:
   - They pay via UPI (`mainakroy273@oksbi`), submit form, and are redirected to your WhatsApp.
2. **You receive their WhatsApp message**:
   - Message has Name, Email, Phone, Number of Tickets, Amount Paid, and the 12-digit UTR.
3. **Double Verification (Takes 5 seconds)**:
   - Open your **SBI / GPay / UPI app** and search the 12-digit UTR to confirm credit of the exact ₹ amount.
   - Check your **Google Sheet** — verified entries show `PENDING_VERIFICATION` in green/neutral, while any tampered amount or reused UTR is highlighted in **red** with `🚨 FLAGGED`.
4. **Issue Ticket**: Once verified in your bank account, send them their digital ticket / entry QR on WhatsApp.
