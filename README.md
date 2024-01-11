# SendGrid Email Automation with Google Apps Script

## Overview

This Google Apps Script automates the process of sending emails using the SendGrid API. The email includes multiple recipients (To, CC, BCC) and is structured with a customizable subject, free-text header, and a tabularized data section. The primary use case is to send emails containing table-formatted data to specified recipients.

## Usage Instructions

1. **Set Up SendGrid API Key and Sender Email**

    Before using the script, make sure to include your SendGrid API key and the sender's email address.

    ```javascript
    // Include your SendGrid API key and sender email address
    const sendGridApiKey = 'YOUR_SENDGRID_API_KEY';
    const senderEmail = 'your.email@example.com';
    ```

2. **Create Email Button**

    In your Google Sheets, create a button linked to the `sendDataInEmail` function using the Apps Script trigger.

3. **Configure Email Settings in 'Sheet1'**

    - **Recipients (To, CC, BCC):** Specify the email addresses in the corresponding cells (B15, B16, B17).
    - **Email Subject (B19):** Set the desired subject for the email.
    - **Header Text (B20):** Enter the text for the free-text header.
    - **Data Range (B21):** Define the range of data in the source sheet.

4. **Click the Email Button**

    When you click the configured button, the script will prompt SendGrid to send an email to the specified recipients. The email body will contain a free-text header followed by a table-formatted section with the specified data.

## Code Explanation

- `sendDataInEmail()`: Collects user input, fetches data from the source sheet, constructs the email body, and triggers `sendEmailUsingSendGrid`.

- `sendEmailUsingSendGrid(payload)`: Sends the email using the SendGrid API. The payload includes recipients, subject, sender, and the HTML-formatted email body.

## Customization

Feel free to customize the script based on your specific needs. You can modify the HTML structure, add additional functionalities, or enhance the email formatting as required.

## Important Notes

- Ensure your SendGrid API key and sender email are kept secure.
- Make sure the required permissions are granted for the script to access and modify spreadsheets.
