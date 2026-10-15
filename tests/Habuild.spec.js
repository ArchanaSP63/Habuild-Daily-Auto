
const { test, expect } = require('@playwright/test');
const ExcelJS = require('exceljs');
const fs = require('fs');

test('YouTube member URL', async ({ page }) => {
    const START_URL = 'https://u.habuild.in/y/sandeeppatil54';

    // Open the member URL
    await page.goto(START_URL, {
        waitUntil: 'domcontentloaded'
    });

    // Wait until the browser redirects to YouTube
    await page.waitForURL(
        url =>
            url.hostname === 'youtube.com' ||
            url.hostname.endsWith('.youtube.com') ||
            url.hostname === 'youtu.be' ||
            url.hostname.endsWith('.youtu.be'),
        { timeout: 60000 }
    );

    // Capture the final YouTube URL
    const youtubeUrl = page.url();
    console.log('YouTube URL:', youtubeUrl);

    // Open the existing Excel file, or create it on first run
    const file = 'youtube-links.xlsx';
    const workbook = new ExcelJS.Workbook();

    if (fs.existsSync(file)) {
        await workbook.xlsx.readFile(file);
    }

    let sheet = workbook.getWorksheet('YouTube Links');

    if (!sheet) {
        sheet = workbook.addWorksheet('YouTube Links');
    }

    if (sheet.rowCount === 0) {
        sheet.addRow(['Date', 'Time', 'YouTube URL']);
    }

    const now = new Date();

    sheet.addRow([
        now.toLocaleDateString('en-CA', {
            timeZone: 'America/Edmonton'
        }),
        now.toLocaleTimeString('en-CA', {
            timeZone: 'America/Edmonton'
        }),
        { text: youtubeUrl, hyperlink: youtubeUrl }
    ]);

    await workbook.xlsx.writeFile(file);

    console.log('YouTube URL saved to Excel.');
});
