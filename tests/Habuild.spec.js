import {test,expect} from '@playwright/test';
const ExcelJS = require('exceljs');
const { DateTime } = require('luxon');
const { chromium } = require('playwright');

test('YouTube member URL',async()=>
{
    const TIME_ZONE = 'America/Edmonton';
    // Calculate the next 2:20 PM in Calgary
    let targetTime = DateTime.now()
    .setZone(TIME_ZONE)
    .set({ hour: 16, minute: 20, second: 0, millisecond: 0 });

    if (targetTime <= DateTime.now().setZone(TIME_ZONE))
    {
        targetTime = targetTime.plus({ days: 1 });
    }

    console.log('Scheduled time:', targetTime.toString());

    // Wait until the scheduled time
    const waitTime = targetTime.toMillis() - Date.now();

    console.log(`Waiting ${Math.ceil(waitTime / 1000)} seconds...`);

    await new Promise(resolve => setTimeout(resolve, waitTime));
    // Launch browser
    const browser = await chromium.launch({ headless: false });
    const page = await browser.newPage();    
    const START_URL = "https://u.habuild.in/y/sandeeppatil54";
    
    //await page.waitForLoadState('networkidle');
    
    try {
    // Open the specified URL at the scheduled time
    await page.goto(START_URL);

    // Wait for redirect to YouTube
    await page.waitForURL(
        url => /(^|\.)youtube\.com$|(^|\.)youtu\.be$/.test(url.hostname),
        { timeout: 60000 }
    );

    // Copy the final YouTube URL
    const youtubeUrl = page.url();

    console.log('YouTube URL:', youtubeUrl);

    // Save the URL to Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('YouTube Links');

    sheet.addRow(['Execution Time (Calgary)', 'YouTube URL']);

    sheet.addRow([
        DateTime.now().setZone(TIME_ZONE)
            .toFormat('yyyy-MM-dd HH:mm:ss'),
        {
            text: youtubeUrl,
            hyperlink: youtubeUrl
        }
    ]);

    await workbook.xlsx.writeFile('youtube-links.xlsx');

    console.log('URL saved to youtube-links.xlsx');

} catch (error) {
    console.error('Automation failed:', error.message);
} finally {
    await browser.close();
}
}

)