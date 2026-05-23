/*
 * Demo / sanitized Google Apps Script tool for freelancer invoice reminders.
 
 * This script collects active freelancer emails from a Google Sheet,
 * calculates the current invoice period, and sends an HTML reminder email.
 */

const INVOICE_REMINDER_CONFIG = {

  START_ROW: 1,
  END_ROW: 152,

  STATUS_COLUMN: 4,
  EXCLUDED_STATUSES: ['Reserve', 'Inactive'],
  
  EMAIL_COLUMN: 5,
  SEND_COLUMN: 23,

  WORKING_DAYS_ROW: 2,
};

function sendInvoiceReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const freelancersSheet = ss.getSheetByName('Freelancers');
  const calculationsSheet = ss.getSheetByName('Calculations');

  if (!freelancersSheet || !calculationsSheet) 
    {
      throw new Error(`Required sheets not found`);
    }

  const freelancerEmails = getActiveFreelancerEmails(freelancersSheet);
  const invoicePeriod = getInvoicePeriod();
  const workingDays = getWorkingDays(calculationsSheet);

  const subject = `Reminder: Invoice for ${invoicePeriod.currentMonthName} ${invoicePeriod.year}`;
  const htmlBody = buildInvoiceReminderBody(invoicePeriod, workingDays);

  GmailApp.sendEmail(
    'manager@example.com',
    subject,
    'Requires HTML',
    {
      htmlBody: htmlBody,
      cc: 'teamlead@example.com',
      bcc: freelancerEmails.join(',')
    }
  );
}


function getActiveFreelancerEmails(sheet) {
  const emails = [];

  for (let row = INVOICE_REMINDER_CONFIG.START_ROW; row < INVOICE_REMINDER_CONFIG.END_ROW; row++) 
    {
      const shouldSend = sheet.getRange(row, INVOICE_REMINDER_CONFIG.SEND_COLUMN).getValue();
      const status = sheet.getRange(row, INVOICE_REMINDER_CONFIG.STATUS_COLUMN).getValue();
      const email = sheet.getRange(row, INVOICE_REMINDER_CONFIG.EMAIL_COLUMN).getValue();

      const isExcludedStatus = INVOICE_REMINDER_CONFIG.EXCLUDED_STATUSES.includes(status);

      if (shouldSend === true && !isExcludedStatus && email) 
        {
          emails.push(email);
        }
    }

  return emails;
}


function getInvoicePeriod() {
  const today = new Date();

  const currentMonthName = today.toLocaleString('en-EN', { month: 'long' });
  const year = Utilities.formatDate(today, 'UTC+4', 'YYYY');

  const previousMonthDate = new Date(today);
  previousMonthDate.setMonth(previousMonthDate.getMonth() - 1);

  const previousMonthName = previousMonthDate.toLocaleString('en-EN', { month: 'long' });

  return {
    currentMonthName,
    previousMonthName,
    year
  };
}


function getWorkingDays(sheet) {
  const currentMonth = new Date().getMonth() + 1;

  return sheet.getRange(INVOICE_REMINDER_CONFIG.WORKING_DAYS_ROW, currentMonth + 1, 1, 1).getValue();
}

//HTML part
function buildInvoiceReminderBody(invoicePeriod, workingDays) {
  return `
    Hello!<br><br>

    I hope you are doing well.<br><br>

    Let me remind you that you need to prepare your invoice for the period from
    ${invoicePeriod.previousMonthName} 16 to ${invoicePeriod.currentMonthName} 15 (included)
    and send a PDF version by ${invoicePeriod.currentMonthName} 24.<br>

    Please note that there were ${workingDays} working days in total during this period.<br><br>

    If you require any assistance, feel free to contact me.<br><br>

    Best regards,<br>
    --<br>
    <p style="line-height: 1">
      <font size="2" color="black" face="Roboto, sans-serif">
        <strong>Localization Team</strong><br>
        Localization Coordinator
      </font><br>
      <font size="2">
        <a href="mailto:manager@example.com" style="color: #4970bf">
          manager@example.com
        </a>
      </font>
    </p>
  `;
}