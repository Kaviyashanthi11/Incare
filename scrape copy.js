import puppeteer from 'puppeteer';
import xlsx from 'xlsx';
import fs from 'fs';

// Run the script
const inputFile = 'NPI.xlsx';

scrapeNPIData(inputFile)
    .then(() => console.log('🎉 Process completed!'))
    .catch(error => console.error('❌ Error:', error.message));

async function scrapeNPIData(inputFile) {
    let browser;
    const processedNPIs = new Set(); // Track already updated NPIs

    try {
        if (!fs.existsSync(inputFile)) throw new Error(`Input file '${inputFile}' does not exist.`);

        const workbook = xlsx.readFile(inputFile);
        const sheetName = workbook.SheetNames[0];
        let sheet = workbook.Sheets[sheetName];
        let data = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        const headerRow = data[0]; // Extract the header row
        const npiIndex = headerRow.indexOf("NPI");
        const npiTypeIndex = 1; // Second column for "NPI Type"

        if (npiIndex === -1) throw new Error("Column 'NPI' not found in the Excel file.");

        // Track NPIs already updated
        for (let i = 1; i < data.length; i++) {
            if (data[i][npiTypeIndex] && !data[i][npiTypeIndex].startsWith("Error:")) {
                processedNPIs.add(data[i][npiIndex]);
            }
        }

        browser = await puppeteer.launch({
            headless: false,
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized'],
            defaultViewport: null
        });

        const page = await browser.newPage();
        const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

        for (let i = 1; i < data.length; i++) {
            const npiValue = data[i][npiIndex]?.toString().trim();

            if (!npiValue) {
                console.warn(`Skipping row ${i} due to missing NPI.`);
                data[i][npiTypeIndex] = "Missing NPI";
                continue;
            }

            // Skip if NPI is already updated
            if (processedNPIs.has(npiValue)) {
                console.log(`⚠️ Alert: NPI ${npiValue} is already updated. Skipping row ${i}.`);
                continue;
            }

            try {
                console.log(`Processing row ${i}/${data.length - 1}: NPI = ${npiValue}`);

                // Navigate to NPI Registry
                await page.goto('https://npiregistry.cms.hhs.gov/', { waitUntil: 'domcontentloaded' });

                // Ensure the NPI search field is visible
                await page.waitForSelector('#npiNumber', { timeout: 15000 });

                // Clear previous input and enter new NPI
                await page.evaluate(() => {
                    const input = document.querySelector('#npiNumber');
                    if (input) input.value = '';
                });
                await page.type('#npiNumber', npiValue, { delay: 200 });

                // Click the Search Button & wait for results
                await Promise.all([
                    page.click('#formButtons > button.btn.btn-primary.active'),
                    page.waitForNavigation({ waitUntil: 'networkidle2' })
                ]);

                // Wait for the NPI Type selector
                const npiTypeSelector = 'body > app-root > main > div > app-provider-view > div.jumbotron > table > tbody > tr:nth-child(3) > td:nth-child(2)';
                await page.waitForSelector(npiTypeSelector, { timeout: 15000 });

                // Extract NPI Type
                const npiType = await page.evaluate(() => {
                    const pageText = document.body.innerText;
                    const match = pageText.match(/NPI-\d+\s+\w+/i);
                    return match ? match[0].trim() : 'Not Found';
                });

                data[i][npiTypeIndex] = npiType; // Update the second column
                processedNPIs.add(npiValue);
                console.log(`✅ Success: NPI ${npiValue} → NPI Type: ${npiType}`);

                // Random delay (2-3s) before next request
                await delay(Math.random() * 1000 + 2000);

            } catch (error) {
                console.error(`❌ Error on row ${i}: ${error.message}`);
                data[i][npiTypeIndex] = `Error: ${error.message}`;
            }
        }

        // Convert updated data back to sheet
if (data[0][npiTypeIndex] !== "NPI Type") {
    data[0][npiTypeIndex] = "NPI Type";
}

// Convert updated data back to sheet
const newSheet = xlsx.utils.aoa_to_sheet(data);

// **Set column widths**
newSheet["!cols"] = [
    { wch: 11 }, // NPI column width
    { wch: 20 }  // NPI Type column width
];

// **Apply center alignment to headers**
const headerRange = xlsx.utils.decode_range(newSheet["!ref"]);
for (let C = headerRange.s.c; C <= headerRange.e.c; C++) {
    const cellAddress = xlsx.utils.encode_cell({ r: 0, c: C });
    if (newSheet[cellAddress]) {
        newSheet[cellAddress].s = {
            alignment: { horizontal: "center", vertical: "center" },
            font: { bold: true }
        };
    }
}

// Update the workbook
workbook.Sheets[sheetName] = newSheet;

// Save the updated Excel file
xlsx.writeFile(workbook, inputFile);
console.log('✅ Scraping completed successfully!');


    } catch (error) {
        console.error('🚨 Fatal error:', error.message);
    } finally {
        if (browser) {
            console.log('Closing browser...');
            await browser.close();
        }
    }
}
