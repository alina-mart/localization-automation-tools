/*
 * Demo / sanitized Google Apps Script tool for monthly schedule emails.
 
 * This script prepares a formatted schedule table in Google Sheets,
 * converts it into an HTML table while preserving visual formatting,
 * collects recipient emails, and sends the schedule by email.
 */

const SCHEDULE_EMAIL_CONFIG = {

  SOURCE_TABLE_START_ROW: 1,
  SOURCE_TABLE_START_COLUMN: 10,
  SOURCE_TABLE_COLUMN_COUNT: 7,

  EMAIL_TABLE_START_ROW: 3,
  EMAIL_TABLE_START_COLUMN: 1,
  EMAIL_TABLE_MAX_ROWS: 15,
  EMAIL_TABLE_COLUMN_COUNT: 7,

  START_ROW: 1,
  END_ROW: 152,

  STATUS_COLUMN: 4,
  EXCLUDED_STATUSES: ['Reserve', 'Inactive'],
  EMAIL_COLUMN: 5,
  SEND_COLUMN: 23,
};

function sendMonthlySchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceTableSheet = ss.getSheetByName('Generated Tables');
  const emailTableSheet = ss.getSheetByName('Schedule Email');
  const freelancersSheet = ss.getSheetByName('Freelancers');

  if (!sourceTableSheet || !emailTableSheet || !freelancersSheet) 
    {
      throw new Error(`Required sheets not found`);
    }


  prepareScheduleTable(sourceTableSheet, emailTableSheet);

  const dateInfo = getNextMonthInfo();
  const freelancerEmails = getScheduleRecipientEmails(freelancersSheet);

  const scheduleRange = emailTableSheet.getDataRange();
  const htmlTable = getHtmlTable(scheduleRange);

  const subject = `Working schedule for ${dateInfo.monthName} ${dateInfo.year}`;
  const htmlBody = buildScheduleEmailBody(dateInfo.monthName, htmlTable);

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

//Copies the generated schedule into the email-ready sheet, preserving background colors and values
function prepareScheduleTable(sourceSheet, targetSheet) {
  const lastRow = sourceSheet.getLastRow();

  const sourceColorRange = sourceSheet.getRange(1, 1, lastRow, SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_COLUMN_COUNT);

  const backgrounds = sourceColorRange.getBackgrounds();

  const targetClearRange = targetSheet.getRange(
    SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_START_ROW,
    SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_START_COLUMN,
    SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_MAX_ROWS,
    SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_COLUMN_COUNT);

  targetClearRange.clearContent();
  targetClearRange.clearFormat();
  targetClearRange.setBorder(true, true, true, true, true, true);

  for (let row = 0; row < backgrounds.length; row++) 
    {
      for (let col = 0; col < backgrounds[row].length; col++) 
        {
          targetSheet.getRange(
          SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_START_ROW + row,
          SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_START_COLUMN + col).setBackground(backgrounds[row][col]);
        }
    }

  sourceSheet.getRange(
      SCHEDULE_EMAIL_CONFIG.SOURCE_TABLE_START_ROW,
      SCHEDULE_EMAIL_CONFIG.SOURCE_TABLE_START_COLUMN,
      lastRow,
      SCHEDULE_EMAIL_CONFIG.SOURCE_TABLE_COLUMN_COUNT).copyTo(targetSheet.getRange(
        SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_START_ROW,
        SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_START_COLUMN,
        lastRow,
        SCHEDULE_EMAIL_CONFIG.EMAIL_TABLE_COLUMN_COUNT), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}


function getNextMonthInfo() {
  const today = new Date();
  today.setMonth(today.getMonth() + 1);

  return {
    monthName: today.toLocaleString('en-EN', { month: 'long' }),
    year: Utilities.formatDate(today, 'UTC+4', 'YYYY')
  };
}


function getScheduleRecipientEmails(sheet) {
  const emails = [];

  for (let row = SCHEDULE_EMAIL_CONFIG.START_ROW; row < SCHEDULE_EMAIL_CONFIG.END_ROW; row++) 
    {
      const shouldSend = sheet.getRange(row, SCHEDULE_EMAIL_CONFIG.SEND_COLUMN).getValue();
      const status = sheet.getRange(row, SCHEDULE_EMAIL_CONFIG.STATUS_COLUMN).getValue();
      const email = sheet.getRange(row, SCHEDULE_EMAIL_CONFIG.EMAIL_COLUMN).getValue();

      const isExcludedStatus = SCHEDULE_EMAIL_CONFIG.EXCLUDED_STATUSES.includes(status);

      if (shouldSend === true && !isExcludedStatus && email) 
        {
          emails.push(email);
        }
    }

  return emails;
}

// the HTML part
function buildScheduleEmailBody(monthName, htmlTable) {
  return `
    Hello,<br><br>

    Please find our working schedule for ${monthName}.<br><br>

    <div>${htmlTable}</div>

    * red cells are non-working days<br>
    ** yellow cells are working Saturdays. The team works on such days,
    but no translation tasks are sent on these days.<br><br>

    Make sure to inform your responsible manager about your availability for the month.<br><br>

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

//Builds an HTML table from a Google Sheets range while preserving cell values, background colors, font styles, alignments and merged cells
function getHtmlTable(range) {
  const sheet = range.getSheet();

  const startRow = range.getRow();
  const startCol = range.getColumn();
  const lastRow = range.getLastRow();
  const lastCol = range.getLastColumn();

  const data = range.getDisplayValues();

  const fontColors = range.getFontColors();
  const backgrounds = range.getBackgrounds();
  const fontSizes = range.getFontSizes();
  const fontWeights = range.getFontWeights();
  const horizontalAlignments = range.getHorizontalAlignments();
  const verticalAlignments = range.getVerticalAlignments();

  const colWidths = [];

  for (let col = startCol; col <= lastCol; col++) 
    {
      colWidths.push(sheet.getColumnWidth(col));
    }

  const rowHeights = [];

  for (let row = startRow; row <= lastRow; row++) 
    {
      rowHeights.push(sheet.getRowHeight(row));
    }

  const tableFormat = 'style="border:1px solid black;border-collapse:collapse;text-align:center" border="1" cellpadding="5"';
  const html = [`<table ${tableFormat}>`];

  for (let col = 0; col < colWidths.length; col++) 
    {
      html.push(`<col width="${colWidths[col]}">`);
    }

  for (let row = 0; row < data.length; row++) 
    {
      html.push(`<tr height="${rowHeights[row]}">`);

      for (let col = 0; col < data[row].length; col++) 
        {
          const currentCell = sheet.getRange(startRow + row, startCol + col);

          if (currentCell.isPartOfMerge() && data[row][col] !== '') 
            {
              const colspan = getMergedCellColspan(sheet, data, startRow, startCol, row, col);
              const cellHtml = buildTableCellHtml(
                data[row][col],
                fontColors[row][col],
                fontSizes[row][col],
                fontWeights[row][col],
                backgrounds[row][col],
                horizontalAlignments[row][col],
                verticalAlignments[row][col],
                colspan);

                html.push(cellHtml);
            } else if (currentCell.isPartOfMerge() && data[row][col] === '') 
              {
                 // Skip empty merged cells
              } else 
                {
                  const cellHtml = buildTableCellHtml(
                    data[row][col],
                    fontColors[row][col],
                    fontSizes[row][col],
                    fontWeights[row][col],
                    backgrounds[row][col],
                    horizontalAlignments[row][col],
                    verticalAlignments[row][col],
                    null);

                  html.push(cellHtml);
                }
        }

    html.push('</tr>');
  }

  html.push('</table>');

  return html.join('');
}


function getMergedCellColspan(sheet, data, startRow, startCol, row, col) {
  let colspan = 1;

  for (let colIndex = col + 1; colIndex < data[row].length; colIndex++) 
    {
      const nextCell = sheet.getRange(startRow + row, startCol + colIndex);

      if (nextCell.isPartOfMerge() && data[row][colIndex] === '') 
        {
          colspan++;
        } 
      else 
        {
           break;
        }
    }

  return colspan;
}

function buildTableCellHtml(
  cellText,
  fontColor,
  fontSize,
  fontWeight,
  background,
  horizontalAlignment,
  verticalAlignment,
  colspan
) {
  const formattedText = formatCellText(cellText);

  const style = [
    `color: ${fontColor}`,
    `font-size: ${fontSize}`,
    `font-weight: ${fontWeight}`,
    `background-color: ${background}`,
    `text-align: ${horizontalAlignment}`,
    `vertical-align: ${verticalAlignment}`
  ].join('; ');

  const colspanAttribute = colspan ? ` colspan="${colspan}"` : '';

  return `<td style="${style}"${colspanAttribute}>${formattedText}</td>`;
}

//Formats dates before inserting them into the HTML table
function formatCellText(cellText) {
  if (cellText instanceof Date) {
    return Utilities.formatDate(
      cellText,
      'UTC+4',
      'M/d/yyyy'
    );
  }

  return cellText;
}