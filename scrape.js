import puppeteer from 'puppeteer';
import xlsx from 'xlsx';
import fs from 'fs';

const inputFile = 'NPI.xlsx';
scrapeNPIData(inputFile)
    .then(() => console.log('🎉 Process completed!'))
    .catch(error => console.error('❌ Error:', error.message));

    async function scrapeNPIData(inputFile) {
        let browser;
    
        try {
            const workbook = xlsx.readFile(inputFile);
            const sheetName = workbook.SheetNames[0];
            let sheet = workbook.Sheets[sheetName];
            let data = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    
            const headerRow = data[0];
            const npiIndex = headerRow.indexOf("NPI");
            const npiTypeIndex = 1;
    
            if (npiIndex === -1) throw new Error("Column 'NPI' not found.");
    
            browser = await puppeteer.launch({ headless: true });
            const page = await browser.newPage();
            await page.setDefaultNavigationTimeout(60000);
    
            for (let i = 1; i < data.length; i++) {
                const npiValue = data[i][npiIndex]?.toString().trim();
                if (!npiValue) {
                    data[i][npiTypeIndex] = "Missing NPI";
                    continue;
                }
    
                try {
                    await page.goto('https://npiregistry.cms.hhs.gov/', { waitUntil: 'domcontentloaded' });
                    await page.waitForSelector('#npiNumber', { timeout: 5000 });
    
                    await page.type('#npiNumber', npiValue, { delay: 100 });
    
                    await Promise.all([
                        page.click('#formButtons > button.btn.btn-primary.active'),
                        page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 })
                    ]);
    
                    const npiType = await page.evaluate(() => {
                        const match = document.body.innerText.match(/NPI-\d+\s+\w+/i);
                        return match ? match[0].trim() : 'Not Found';
                    });
    
                    data[i][npiTypeIndex] = npiType;
                } catch (error) {
                    console.error(`Error processing NPI ${npiValue}:`, error.message);
                    data[i][npiTypeIndex] = `Error: ${error.message}`;
                }
            }
    
            await browser.close();
    
            if (data[0][npiTypeIndex] !== "NPI Type") {
                data[0][npiTypeIndex] = "NPI Type";
            }
    
            // Ensure 'scraped' directory exists
            const scrapedDir = path.join(__dirname, 'scraped');
            if (!fs.existsSync(scrapedDir)) fs.mkdirSync(scrapedDir);
    
            const outputFileName = `NPI_${Date.now()}.xlsx`;
            const outputFilePath = path.join(scrapedDir, outputFileName);
    
            const newSheet = xlsx.utils.aoa_to_sheet(data);
            workbook.Sheets[sheetName] = newSheet;
            xlsx.writeFile(workbook, outputFilePath);
    
            return `scraped/${outputFileName}`; // Corrected path
    
        } catch (error) {
            if (browser) await browser.close();
            throw error;
        }
    }
    