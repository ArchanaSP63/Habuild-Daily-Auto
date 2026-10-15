
import { test,expect } from '@playwright/test' ;
const ExcelJS = require('exceljs');

const INPUT_FILE = './input.xlsx';
const OUTPUT_FILE = './output.xlsx';
const LOGIN_URL = 'https://your-website.com/login';

// Replace these example selectors with your actual page selectors.
const SELECTORS = {
  username: '#username',
  currentPassword: '#currentPassword',
  newPassword: '#newPassword',
  confirmPassword: '#confirmPassword',
  loginButton: 'button[type="submit"]',
};

// Optional: set this to your expected destination URL.
// Example: '**/dashboard'
const EXPECTED_URL = null;

async function main() {
  const inputWorkbook = new ExcelJS.Workbook();
  await inputWorkbook.xlsx.readFile(INPUT_FILE);

  const inputSheet = inputWorkbook.getWorksheet('Sheet1');

  if (!inputSheet) {
    throw new Error('Worksheet "Sheet1" was not found.');
  }

  // Read column headers from the first row.
  const headers = inputSheet.getRow(1).values
    .slice(1)
    .map(value => String(value ?? '').trim());

  const requiredHeaders = [
    'username',
    'current_password',
    'new_password',
    'confirm_password',
  ];

  for (const header of requiredHeaders) {
    if (!headers.includes(header)) {
      throw new Error(`Missing Excel column: ${header}`);
    }
  }

  const column = Object.fromEntries(
    headers.map((header, index) => [header, index + 1])
  );

  const outputWorkbook = new ExcelJS.Workbook();
  const outputSheet = outputWorkbook.addWorksheet('Results');

  outputSheet.addRow([
    'Username',
    'Status',
    'Final URL',
    'Error Message',
    'Execution Time (seconds)',
  ]);

  const browser = await chromium.launch({
    headless: false,
  });

  try {
    for (let rowNumber = 2;
         rowNumber <= inputSheet.rowCount;
         rowNumber++) {

      const row = inputSheet.getRow(rowNumber);

      const getValue = header =>
        String(row.getCell(column[header]).text ?? '').trim();

      const username = getValue('username');
      const currentPassword = getValue('current_password');
      const newPassword = getValue('new_password');
      const confirmPassword = getValue('confirm_password');

      if (!username) continue;

      const context = await browser.newContext();
      const page = await context.newPage();
      const start = Date.now();

      let status = 'FAILED';
      let errorMessage = '';
      let finalUrl = '';

      try {
        // Step 1: Open the login page.
        await page.goto(LOGIN_URL, {
          waitUntil: 'domcontentloaded',
        });

        // Step 2: Fill in all four fields.
        await page.locator(SELECTORS.username)
          .fill(username);

        await page.locator(SELECTORS.currentPassword)
          .fill(currentPassword);

        await page.locator(SELECTORS.newPassword)
          .fill(newPassword);

        await page.locator(SELECTORS.confirmPassword)
          .fill(confirmPassword);

        // Step 3: Click the login / password-update button.
        const oldUrl = page.url();

        await page.locator(SELECTORS.loginButton).click();

        // Step 4: Wait for the expected destination or
        // a URL change. A URL change alone isn't proof of success.
        if (EXPECTED_URL) {
          await page.waitForURL(EXPECTED_URL, {
            timeout: 15000,
          });
        } else {
          try {
            await page.waitForURL(
              url => url.href !== oldUrl,
              { timeout: 10000 }
            );
          } catch {
            // The page may update without changing its URL.
          }
        }

        finalUrl = page.url();

        // Step 5: Verify the actual success condition.
        // Replace this example with your site's success message
        // or destination-page element.
        const successMessage = page.getByText(
          'Password updated successfully',
          { exact: false }
        );

        const successVisible = await successMessage
          .isVisible()
          .catch(() => false);

        if (successVisible) {
          status = 'SUCCESS';
        } else if (finalUrl !== oldUrl) {
          status = 'REDIRECTED - VERIFY SUCCESS';
        } else {
          status = 'CHECK REQUIRED';
        }

      } catch (error) {
        errorMessage = error.message;
        finalUrl = page.url();
      } finally {
        const elapsed = (
          (Date.now() - start) / 1000
        ).toFixed(2);

        // Step 6: Record the outcome in the output workbook.
        outputSheet.addRow([
          username,
          status,
          finalUrl,
          errorMessage,
          Number(elapsed),
        ]);

        // Save progress after each account.
        await outputWorkbook.xlsx.writeFile(OUTPUT_FILE);

        // Close this account's browser context.
        await context.close();

        console.log(`${username}: ${status}`);
      }
    }
  } finally {
    await browser.close();
  }

  console.log(`Results saved to ${OUTPUT_FILE}`);
}

main().catch(error => {
  console.error('Automation failed:', error.message);
  process.exitCode = 1;
});
