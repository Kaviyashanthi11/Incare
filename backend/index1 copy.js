// import express from 'express';
// import multer from 'multer';
// import cors from 'cors';
// import xlsx from 'xlsx';
// import ExcelJS from 'exceljs';
// import puppeteer from 'puppeteer';
// import fs from 'fs';
// import path from 'path';
// import { fileURLToPath } from 'url'; // To handle __dirname in ES modules
// import { dirname } from 'path';

// const app = express();
// const port = 5000;

// app.use(cors());
// app.use(express.json());

// const storage = multer.memoryStorage();
// const upload = multer({ storage });

// // Store processed data temporarily
// const processedFiles = new Map();

// // Get __dirname in ES modules
// const __filename = fileURLToPath(import.meta.url);
// const __dirname = dirname(__filename);

// // VPN Connectivity Check Function
// import https from 'https';
// import fetch from 'node-fetch';

// async function loginToSystem(page) {
//     try {
//         console.log('Navigating to login page...');
//         await page.goto('https://maicllejgit6695ivkapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp', { waitUntil: 'load', timeout: 60000 });

//         // Enter username
//         console.log('Entering username...');
//         await page.waitForSelector('#doctorID', { timeout: 5000 });
//         await page.type('#doctorID', 'rajasekarp', { delay: 50 });

//         // Click Next
//         console.log('Clicking Next after username...');
//         await page.click('input[type="submit"]');
//         await page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 60000 });

//         // Enter password
//         console.log('Entering password...');
//         await page.waitForSelector('#passwordField', { timeout: 5000 });
//         await page.type('#passwordField', 'Bright@2025', { delay: 50 });

//         // Click Next after entering password
//         console.log('Clicking Next after password...');
//         await page.click('input[type="submit"]');
//         await page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 60000 });

//         console.log('Login successful!');

//         console.log('Waiting for dropdown to be visible...');
//         await page.waitForSelector('.lookup-toogle.lookup-toogle-1366 #jellybean-panelLink66', { visible: true, timeout: 10000 }); // Wait until the dropdown button is visible
        
//         console.log('Clicking dropdown...');
//         await page.click('.lookup-toogle.lookup-toogle-1366 #jellybean-panelLink66');  // Click the dropdown button
        
//         // Wait for the page to load or any JavaScript actions to complete (if necessary)
//         await page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 10000 });
        
//         console.log('Waiting for dropdown options to be visible...');
//         await page.waitForSelector('#jellybean-panelUl11', { visible: true, timeout: 10000 }); // Wait for the dropdown menu
        
//         const isDropdownVisible = await page.$eval('#jellybean-panelUl11', (dropdown) => {
//             return window.getComputedStyle(dropdown).display !== 'none' && dropdown.clientHeight > 0;
//         });
        
//         if (isDropdownVisible) {
//             console.log('Dropdown is visible, selecting option...');
//             await page.waitForSelector('#jellybean-panelUl11 #jellybean-panelLink67', { visible: true, timeout: 30000 });  // Wait for the specific option to be visible
//             await page.click('#jellybean-panelUl11 #jellybean-panelLink67');  // Select the option
//             console.log('Dropdown option selected!');
//         } else {
//             console.error('Dropdown is not visible or not active.');
//         }
        
        

//     } catch (err) {
//         console.error("Login or navigation failed:", err.message);

//         // Take a screenshot for debugging
//         await page.screenshot({ path: 'error-screenshot.png' });

//         // Log the page content for further inspection
//         console.log(await page.content());

//         throw new Error("Login process or navigation encountered an issue. Please check credentials or website accessibility.");
//     }
// }

// async function processAccount(page, accountNumber, notes) {
//     // Implement your account processing logic here
// }

// // Endpoint to process data and return status
// app.post("/process", upload.single("file"), async (req, res) => {
//     if (!req.file || !req.body.originalPath) {
//         return res.status(400).json({ error: "No file or original path provided" });
//     }

//     let browser;
//     try {
//         const workbookXLSX = xlsx.read(req.file.buffer, { type: "buffer", cellDates: true ,
//             cellNF: false,  // Reduce memory overhead
//             bookVBA: false  // Skip VBA storage
//         });

//         if (workbookXLSX.Workbook?.WBProps?.Encrypted) {
//             throw new Error("This file is password protected");
//         }

//         const sheetXLSX = workbookXLSX.Sheets[workbookXLSX.SheetNames[0]];
//         const originalData = xlsx.utils.sheet_to_json(sheetXLSX, {
//             header: 1,
//             raw: false,
//             dateNF: 'mm/dd/yyyy'
//         });

//         if (!originalData.length) {
//             throw new Error("Empty file or missing headers.");
//         }
//         const headers = originalData[0].map(h => h?.toString().trim());
//         const columnMap = {};
        
//         // Check if "Pt Account #" exists in headers
//         if (!headers.includes("Pt Account #")) {
//             throw new Error("'Pt Account #' Column is not defined.");
//         }
        
//         // Map headers to their index positions
//         headers.forEach((header, index) => {
//             columnMap[header] = index;
//         });

//         // Check if Result column exists, if not, add it
//         if (!columnMap.hasOwnProperty("Result")) {
//             columnMap["Result"] = headers.length;
//             headers.push("Result");
//             // Add the Result column to all existing rows
//             for (let i = 1; i < originalData.length; i++) {
//                 if (originalData[i]) {
//                     originalData[i].push("");
//                 }
//             }
//         }

//         let newRowsToProcess = 0;
//         const processedAccounts = new Set();
        
//         let missingCallNotes = 0;  // Track rows missing Call Notes

//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             const accountNumberRaw = row[columnMap["Pt Account #"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
//             if (!accountNumberRaw || status === "done") continue;  // Skip processed rows
            
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             processedAccounts.add(accountNumber);
//             newRowsToProcess++;
        
//             // Check if Call Notes is missing
//             const callNotes = row[columnMap["Call Notes"]] || "";
//             if (!callNotes.trim()) {
//                 missingCallNotes++;
//             }
//         }
        
//         // ✅ Check if all remaining rows are missing Call Notes
//         if (newRowsToProcess > 0 && newRowsToProcess === missingCallNotes) {
//             return res.json({
//                 success: false,
//                 message: "Notes not found in Excel.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // ✅ Check if all rows are already updated
//         if (newRowsToProcess === 0) {
//             return res.json({
//                 success: true,
//                 message: "All rows were already updated.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // Launch Puppeteer (browser is only launched if VPN is connected)
//         browser = await puppeteer.launch({
//             headless: false,
//             args: ["--ignore-certificate-errors",
//                 '--no-sandbox',
//                 '--disable-setuid-sandbox',
//                 '--disable-dev-shm-usage',
//                 '--reduce-security-for-testing'
//             ],
//             defaultViewport: null
//         });
//         const page = await browser.newPage();
//         await loginToSystem(page);
        
//         const results = { successful: [], failed: [] };
        
//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             const accountNumberRaw = row[columnMap["Pt Account #"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
//             // Skip row if there's no account number or it's already processed.
//             if (!accountNumberRaw || status === "done") continue;
            
//             // Skip row if the "Call Notes" field is empty.
//             const callNotes = row[columnMap["Call Notes"]] || "";
//             if (!callNotes.toString().trim()) continue;
            
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             console.log(`Processing account: ${accountNumber}`);
        
//             const success = await processAccount(page, accountNumber, callNotes);
        
//             if (success) {
//                 results.successful.push(accountNumber);
//                 originalData[i][columnMap["Result"]] = "Done";
//             } else {
//                 results.failed.push(accountNumber);
//             }
//         }
        
//         // Convert data back to an ExcelJS workbook
//         const workbook = new ExcelJS.Workbook();
//         const sheet = workbook.addWorksheet("Sheet1");

//         // Write headers with green background
//         const headerRow = sheet.addRow(headers);
//         headerRow.eachCell((cell) => {
//             cell.fill = {
//                 type: "pattern",
//                 pattern: "solid",
//                 fgColor: { argb: "FFFF00" }
//             };
//             cell.font = { bold: true };
//         });

//         // Write rows
//         originalData.slice(1).forEach((row) => sheet.addRow(row));

//         // Save to buffer
//         const buffer = await workbook.xlsx.writeBuffer();

//         const filePath = path.join(__dirname, `updated_file_${Date.now()}.xlsx`);
//         fs.writeFileSync(filePath, buffer);

//         res.json({
//             success: true,
//             message: "File processed successfully",
//             filePath,
//             results
//         });
//     } catch (error) {
//         console.error('Error during processing:', error.message);
//         res.status(500).json({ error: 'An error occurred during processing.' });
//     } finally {
//         if (browser) {
//             await browser.close();
//         }
//     }
// });

// // Start Express server
// app.listen(port, () => {
//     console.log(`Server is running on port ${port}`);
// });


// import express from 'express';
// import multer from 'multer';
// import cors from 'cors';
// import xlsx from 'xlsx';
// import ExcelJS from 'exceljs';
// import puppeteer from 'puppeteer';
// import fs from 'fs';
// import path from 'path';
// import { Readable } from 'stream';

// const app = express();
// const port = 5000;

// app.use(cors());
// app.use(express.json());

// const storage = multer.memoryStorage();
// const upload = multer({ storage });

  

// // Store processed data temporarily
// const processedFiles = new Map();

// // VPN Connectivity Check Function
// import https from 'https';
// import fetch from 'node-fetch';
// async function loginToSystem(page) {
//     try {
//         console.log('Navigating to login page...');
//         await page.goto('https://maicllejgit6695ivkapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp', { waitUntil: 'load', timeout: 60000 });

//         // Enter username
//         console.log('Entering username...');
//         await page.waitForSelector('#doctorID', { timeout: 5000 });
//         await page.type('#doctorID', 'rajasekarp', { delay: 50 });

//         // Click Next
//         console.log('Clicking Next after username...');
//         await page.click('input[type="submit"]');
//         await page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 120000 });

//         // Enter password
//         console.log('Entering password...');
//         await page.waitForSelector('#passwordField', { timeout: 5000 });
//         await page.type('#passwordField', 'Bright@2025', { delay: 50 });

//         // Click Next after entering password
//         console.log('Clicking Next after password...');
//         await page.click('input[type="submit"]');
//         await page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 60000 });

//         console.log('Login successful!');

//         // Wait for the dropdown toggle to be visible and clickable
//         console.log('Waiting for dropdown toggle to be visible...');
//         await page.waitForSelector('#jellybean-panelLink66', { visible: true, timeout: 10000 });

//         // Click the dropdown toggle to open it
//         console.log('Clicking dropdown toggle...');
//         await page.click('#jellybean-panelLink66');

//         // Add a small delay using Promise-based delay
//         await new Promise(resolve => setTimeout(resolve, 500)); // Wait for 500 milliseconds

      

//         // Wait for the Patient Lookup option to be visible within the dropdown
//         console.log('Waiting for Patient Lookup option...');
//         await page.waitForSelector('#jellybean-panelLink67', { visible: true, timeout: 20000 });

//         // Click the Patient Lookup option
//         console.log('Clicking Patient Lookup option...');
//         await page.evaluate(() => {
//             const element = document.querySelector('#jellybean-panelLink67');
//             if (element) {
//                 element.click();
//             } else {
//                 console.error('Could not find the Patient Lookup element');
//             }
//         });

//         console.log('Patient Lookup option selected!');

//         // Wait for the page to navigate after selecting the option
//         await page.waitForNavigation({ timeout: 10000 }).catch(err => {
//             console.log('No navigation occurred after clicking Patient Lookup, which might be normal');
//         });

//     } catch (err) {
//         console.error("Process failed:", err.message);
//         await page.screenshot({ path: 'error-screenshot.png' });
//         throw new Error("Process encountered an issue: " + err.message);
//     }
// }

// async function processAccount(page, accountNumber, callNotes) {
//     try {
//         console.log('Waiting for the dropdown to appear...');
//         await page.waitForSelector('#patient-lookup-screen-detview > div > div:nth-child(1) > div:nth-child(1) > div > div > div > div.col-sm-5.nopadding.nameselect > div > div > span.ng-binding', { visible: true });

//         console.log('Clicking the dropdown to open options...');
//         await page.click('#patient-lookup-screen-detview > div > div:nth-child(1) > div:nth-child(1) > div > div > div > div.col-sm-5.nopadding.nameselect > div > div > span.ng-binding');

//         console.log('Waiting for the dropdown options to load...');
//         await new Promise(resolve => setTimeout(resolve, 1000));

//         console.log('Selecting "Acct No (MRN)" from the dropdown...');
//         await page.evaluate(() => {
//             let options = document.querySelectorAll('span.ng-binding');
//             options.forEach(option => {
//                 if (option.innerText.trim() === 'Acct No (MRN)') {
//                     option.click();
//                 }
//             });
//         });

//         console.log('Waiting for the #searchText field...');
//         await page.waitForSelector('#searchText', { visible: true });

//         console.log(`Entering account number: ${accountNumber} in #searchText field`);
//         await page.type('#searchText', accountNumber, { delay: 100 });
//         await new Promise(resolve => setTimeout(resolve, 3000));

//         console.log('Clicking the Patient Demographics button...');
//         await page.waitForSelector('#patientSearchBtn12', { visible: true });
//         await page.click('#patientSearchBtn12');
//         await new Promise(resolve => setTimeout(resolve, 8000));

//         console.log('Waiting for the Notes field...');
//         await page.waitForSelector('#patientdemographics-AddInformationIpt32', { visible: true });

//         // Read existing text from the Notes field
//         const existingNotes = await page.evaluate(() => {
//             const notesField = document.querySelector('#patientdemographics-AddInformationIpt32');
//             if (!notesField) {
//                 console.log('Notes field not found!');
//                 return '';
//             }

//             console.log('Notes field detected:', notesField);

//             // Check if the field is an input or textarea
//             if (notesField.tagName.toLowerCase() === 'input' || notesField.tagName.toLowerCase() === 'textarea') {
//                 return notesField.value || '';
//             }

//             // If it's a div or span, use innerText
//             return notesField.innerText || notesField.textContent || '';
//         });

//         console.log(`Existing Notes: ${existingNotes}`);

//         // Prepend the new call notes before the existing text
//         const updatedNotes = `${callNotes}\n${existingNotes || ''}`.trim();

//         console.log(`Entering updated call notes: ${updatedNotes}`);

//         // Set the new value without clearing existing text
//         await page.evaluate((selector, text) => {
//             let notesField = document.querySelector(selector);
//             if (notesField) {
//                 // For input/textarea fields, set the value directly
//                 if (notesField.tagName.toLowerCase() === 'input' || notesField.tagName.toLowerCase() === 'textarea') {
//                     notesField.value = text;
//                 } else {
//                     // For other fields like div/span, set innerText
//                     notesField.innerText = text;
//                 }
//                 // Dispatch an 'input' event to trigger UI update if it's an input field
//                 notesField.dispatchEvent(new Event('input', { bubbles: true }));
//             }
//         }, '#patientdemographics-AddInformationIpt32', updatedNotes);

//         console.log('Clicking the Cancel button...');
//         await page.waitForSelector('#patient-demographicsBtn57', { visible: true });
//         await page.click('#patient-demographicsBtn57');

//         console.log('Process completed for account:', accountNumber);
//         return true;
//     } catch (error) {
//         console.error(`Error processing account ${accountNumber}:`, error);
//         return false;
//     }
// }


// // Endpoint to process data and return status
// app.post("/process", upload.single("file"), async (req, res) => {
//     if (!req.file || !req.body.originalPath) {
//         return res.status(400).json({ error: "No file or original path provided" });
//     }

//     let browser;
//     try {
//         const workbookXLSX = xlsx.read(req.file.buffer, { type: "buffer", cellDates: true ,
//             cellNF: false,  // Reduce memory overhead
//             bookVBA: false  // Skip VBA storage
//         });

//         if (workbookXLSX.Workbook?.WBProps?.Encrypted) {
//             throw new Error("This file is password protected");
//         }

//         const sheetXLSX = workbookXLSX.Sheets[workbookXLSX.SheetNames[0]];
//         const originalData = xlsx.utils.sheet_to_json(sheetXLSX, {
//             header: 1,
//             raw: false,
//             dateNF: 'mm/dd/yyyy'
//         });

//         if (!originalData.length) {
//             throw new Error("Empty file or missing headers.");
//         }
//         const headers = originalData[0].map(h => h?.toString().trim());
//         const columnMap = {};
        
//         // Check if "Acct No" exists in headers
//         console.log(headers);
//         if (!headers.includes("Acct No")) {
//             throw new Error("'Acct No' Column is not defined.");
//         }
        
//         // Map headers to their index positions
//         headers.forEach((header, index) => {
//             columnMap[header] = index;
//         });
        
        
//         // Check if Result column exists, if not, add it
//         if (!columnMap.hasOwnProperty("Result")) {
//             columnMap["Result"] = headers.length;
//             headers.push("Result");
//             // Add the Result column to all existing rows
//             for (let i = 1; i < originalData.length; i++) {
//                 if (originalData[i]) {
//                     originalData[i].push("");
//                 }
//             }
//         }
        
//         let newRowsToProcess = 0;
//         const processedAccounts = new Set();
        
//         let missingCallNotes = 0;  // Track rows missing Notes

//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             const accountNumberRaw = row[columnMap["Acct No"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
//             if (!accountNumberRaw || status === "done") continue;  // Skip processed rows
            
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             processedAccounts.add(accountNumber);
//             newRowsToProcess++;
        
//             // Check if Notes is missing
//             const callNotes = row[columnMap["Notes"]] || "";
//             if (!callNotes.trim()) {
//                 missingCallNotes++;
//             }
//         }
        
//         // ✅ Check if all remaining rows are missing Notes
//         if (newRowsToProcess > 0 && newRowsToProcess === missingCallNotes) {
//             return res.json({
//                 success: false,
//                 message: "Notes not found in Excel.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // ✅ Check if all rows are already updated
//         if (newRowsToProcess === 0) {
//             return res.json({
//                 success: true,
//                 message: "All rows were already updated.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
        
//         // Launch Puppeteer (browser is only launched if VPN is connected)
//         browser = await puppeteer.launch({
//             headless: false,
//             args: ["--ignore-certificate-errors",
//                 '--no-sandbox',
//                 '--disable-setuid-sandbox',
//                 '--disable-dev-shm-usage',
//                 '--reduce-security-for-testing'
//             ],
//             defaultViewport: null
//         });
//         const page = await browser.newPage();
//         await loginToSystem(page);
        
//         const results = { successful: [], failed: [] };
        
//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             const accountNumberRaw = row[columnMap["Acct No"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
//             // Skip row if there's no account number or it's already processed.
//             if (!accountNumberRaw || status === "done") continue;
            
//             // Skip row if the "Notes" field is empty.
//             const callNotes = row[columnMap["Notes"]] || "";
//             if (!callNotes.toString().trim()) continue;
            
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             console.log(`Processing account: ${accountNumber}`);
        
//             const success = await processAccount(page, accountNumber, callNotes);
        
//             if (success) {
//                 results.successful.push(accountNumber);
//                 originalData[i][columnMap["Result"]] = "Done";
//             } else {
//                 results.failed.push(accountNumber);
//             }
//         }
        
//         // Convert data back to an ExcelJS workbook
//         const workbook = new ExcelJS.Workbook();
//         const sheet = workbook.addWorksheet("Sheet1");

//         // Write headers with green background
//         const headerRow = sheet.addRow(headers);
//         headerRow.eachCell((cell) => {
//             cell.fill = {
//                 type: "pattern",
//                 pattern: "solid",
//                 fgColor: { argb: "FFFF00" }
//             };
//             cell.font = { bold: true, color: { argb: "black" }, size: 10 };
//             cell.alignment = { horizontal: "center", vertical: "middle" };
//             cell.border = {
//                 top: { style: "thin" },
//                 left: { style: "thin" },
//                 bottom: { style: "thin" },
//                 right: { style: "thin" }
//             };
//         });

//         // Write data back to the sheet
// // In your column widths array, add width for EPS Comments if needed
// const columnWidths = [10, 16, 20, 10, 5, 25, 28, 8, 27, 11, 16, 15, 12, 25, 12, 80, 80, 26, 12, 9, 9];

// // When setting up columns, ensure EPS Comments has wrap text enabled
// sheet.columns = columnWidths.map((width, index) => ({
//     width,
//     style: (index === columnMap["Notes"] || index === columnMap["EPS Comments"]) 
//         ? { alignment: { wrapText: true } } 
//         : {},
// }));

// // When formatting cells in each row, add special handling for EPS Comments
// originalData.slice(1).forEach((row) => {
//     const hasData = row.some(cell => cell && cell.toString().trim() !== "");
//     const rowData = headers.map(header => row[columnMap[header]]);
    
//     if (hasData) {
//         const newRow = sheet.addRow(rowData);
//         newRow.height = 150;

//         headers.forEach((header, colIndex) => {
//             const cell = newRow.getCell(colIndex + 1);

//             // Apply border to all sides
//             cell.border = {
//                 top: { style: "thin", color: { argb: "000000" } },
//                 left: { style: "thin", color: { argb: "000000" } },
//                 bottom: { style: "thin", color: { argb: "000000" } },
//                 right: { style: "thin", color: { argb: "000000" } }
//             };

//             // Apply default font size
//             cell.font = { size: 10 };

//             if (header === "Acct No" && cell.value) {
//                 cell.fill = {
//                     type: "pattern",
//                     pattern: "solid",
//                     fgColor: { argb: "FFFF00" } // Yellow
//                 };
//                 cell.alignment = { horizontal: "center", vertical: "bottom" };
//             }
        
//             // Formatting for "Notes"
//             if (header === "Notes" && cell.value) {
//                 cell.fill = {
//                     type: "pattern",
//                     pattern: "solid",
//                     fgColor: { argb: "FFFF00" } // Yellow
//                 };
//                 cell.alignment = { wrapText: true, vertical: "bottom" };
//             }

//             // Formatting for "EPS Comments"
//             if (header === "EPS Comments" && cell.value) {
//                 cell.fill = {
//                     type: "pattern",
//                     pattern: "solid",
//                     fgColor: { argb: "FFFFFF" } // White
//                 };
//                 cell.alignment = { wrapText: true, vertical: "bottom" };
//             }
//         });
//     }
// });

//         // Generate a unique file ID
//         const fileId = Date.now().toString();

//         // Create a buffer of the Excel file and store it
//         const fileBuffer = await workbook.xlsx.writeBuffer();
//         processedFiles.set(fileId, {
//             buffer: fileBuffer,
//             filename: "updated_file.xlsx"
//         });

//         // Return the results and the file ID for later download
//         res.json({
//             success: true,
//             message: `${newRowsToProcess} rows processed successfully.`,
//             results,
//             fileId
//         });

//     } catch (error) {
//         console.error("Error:", error);
//         res.status(500).json({ error: error.message });
//     } finally {
//         if (browser) {
//             await new Promise(resolve => setTimeout(resolve, 1000));
//             await browser.close();
//             console.log("Browser closed");
//         }
//     }
// });


// // Separate endpoint to download the processed file
// app.get("/download/:fileId", (req, res) => {
//     const fileId = req.params.fileId;
//     const fileData = processedFiles.get(fileId);
    
//     if (!fileData) {
//         return res.status(404).json({ error: "File not found" });
//     }
    
//     // Set the response headers
//     res.setHeader('Content-Disposition', `attachment; filename="${fileData.filename}"`);
//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//     res.setHeader('Content-Length', fileData.buffer.length);
    
//     // Send the file buffer
//     res.send(fileData.buffer);
    
//     // Optional: Remove the file from memory after some time
//     setTimeout(() => {
//         processedFiles.delete(fileId);
//         console.log(`File ${fileId} removed from memory`);
//     }, 30 * 60 * 1000); // Remove after 30 minutes
// });

// app.listen(port, () => console.log(`Server running on http://localhost:${port}`));

import express from 'express';
import multer from 'multer';
import cors from 'cors';
import xlsx from 'xlsx';
import ExcelJS from 'exceljs';
import puppeteer from 'puppeteer';
import fs from 'fs';
import path from 'path';
import { Readable } from 'stream';

const app = express();
const port = 5000;

app.use(cors());
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage });

  

// Store processed data temporarily
const processedFiles = new Map();

// VPN Connectivity Check Function
import https from 'https';
import fetch from 'node-fetch';
async function loginToSystem(page) {
    try {
        console.log('Navigating to login page...');
        await page.goto('https://maicllejgit6695ivkapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp', { waitUntil: 'load', timeout: 60000 });

        // Enter username
        console.log('Entering username...');
        await page.waitForSelector('#doctorID', { timeout: 5000 });
        await page.type('#doctorID', 'rajasekarp', { delay: 50 });

        // Click Next
        console.log('Clicking Next after username...');
        await page.click('input[type="submit"]');
        await page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 120000 });

        // Enter password
        console.log('Entering password...');
        await page.waitForSelector('#passwordField', { timeout: 5000 });
        await page.type('#passwordField', 'Bright@2025', { delay: 50 });

        // Click Next after entering password
        console.log('Clicking Next after password...');
        await page.click('input[type="submit"]');
        await page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 60000 });

        console.log('Login successful!');

        // Wait for the dropdown toggle to be visible and clickable
        console.log('Waiting for dropdown toggle to be visible...');
        await page.waitForSelector('#jellybean-panelLink66', { visible: true, timeout: 10000 });

        // Click the dropdown toggle to open it
        console.log('Clicking dropdown toggle...');
        await page.click('#jellybean-panelLink66');

        // Add a small delay using Promise-based delay
        await new Promise(resolve => setTimeout(resolve, 500)); // Wait for 500 milliseconds

      

        // Wait for the Patient Lookup option to be visible within the dropdown
        console.log('Waiting for Patient Lookup option...');
        await page.waitForSelector('#jellybean-panelLink67', { visible: true, timeout: 20000 });

        // Click the Patient Lookup option
        console.log('Clicking Patient Lookup option...');
        await page.evaluate(() => {
            const element = document.querySelector('#jellybean-panelLink67');
            if (element) {
                element.click();
            } else {
                console.error('Could not find the Patient Lookup element');
            }
        });

        console.log('Patient Lookup option selected!');

        // Wait for the page to navigate after selecting the option
        await page.waitForNavigation({ timeout: 10000 }).catch(err => {
            console.log('No navigation occurred after clicking Patient Lookup, which might be normal');
        });

    } catch (err) {
        console.error("Process failed:", err.message);
        await page.screenshot({ path: 'error-screenshot.png' });
        throw new Error("Process encountered an issue: " + err.message);
    }
}

async function processAccount(page, accountNumber, callNotes) {
    try {
        console.log('Waiting for the dropdown to appear...');
        await page.waitForSelector('#patient-lookup-screen-detview > div > div:nth-child(1) > div:nth-child(1) > div > div > div > div.col-sm-5.nopadding.nameselect > div > div > span.ng-binding', { visible: true });

        console.log('Clicking the dropdown to open options...');
        await page.click('#patient-lookup-screen-detview > div > div:nth-child(1) > div:nth-child(1) > div > div > div > div.col-sm-5.nopadding.nameselect > div > div > span.ng-binding');

        console.log('Waiting for the dropdown options to load...');
        await new Promise(resolve => setTimeout(resolve, 1000));

        console.log('Selecting "Acct No (MRN)" from the dropdown...');
        await page.evaluate(() => {
            let options = document.querySelectorAll('span.ng-binding');
            options.forEach(option => {
                if (option.innerText.trim() === 'Acct No (MRN)') {
                    option.click();
                }
            });
        });

        console.log('Waiting for the #searchText field...');
        await page.waitForSelector('#searchText', { visible: true });

        console.log(`Entering account number: ${accountNumber} in #searchText field`);
        await page.type('#searchText', accountNumber, { delay: 100 });
        await new Promise(resolve => setTimeout(resolve, 3000));

        console.log('Clicking the Patient Demographics button...');
        await page.waitForSelector('#patientSearchBtn12', { visible: true });
        await page.click('#patientSearchBtn12');
        await new Promise(resolve => setTimeout(resolve, 8000));

        console.log('Waiting for the Notes field...');
        await page.waitForSelector('#patientdemographics-AddInformationIpt32', { visible: true });

        // Read existing text from the Notes field
        const existingNotes = await page.evaluate(() => {
            const notesField = document.querySelector('#patientdemographics-AddInformationIpt32');
            if (!notesField) {
                console.log('Notes field not found!');
                return '';
            }

            console.log('Notes field detected:', notesField);

            // Check if the field is an input or textarea
            if (notesField.tagName.toLowerCase() === 'input' || notesField.tagName.toLowerCase() === 'textarea') {
                return notesField.value || '';
            }

            // If it's a div or span, use innerText
            return notesField.innerText || notesField.textContent || '';
        });

        console.log(`Existing Notes: ${existingNotes}`);

        // Prepend the new call notes before the existing text
        const updatedNotes = `${callNotes}\n${existingNotes || ''}`.trim();

        console.log(`Entering updated call notes: ${updatedNotes}`);

        // Set the new value without clearing existing text
        await page.evaluate((selector, text) => {
            let notesField = document.querySelector(selector);
            if (notesField) {
                // For input/textarea fields, set the value directly
                if (notesField.tagName.toLowerCase() === 'input' || notesField.tagName.toLowerCase() === 'textarea') {
                    notesField.value = text;
                } else {
                    // For other fields like div/span, set innerText
                    notesField.innerText = text;
                }
                // Dispatch an 'input' event to trigger UI update if it's an input field
                notesField.dispatchEvent(new Event('input', { bubbles: true }));
            }
        }, '#patientdemographics-AddInformationIpt32', updatedNotes);

        console.log('Waiting for the Save button...');
        await page.waitForSelector('#patient-demographicsBtn56', { visible: true });

        console.log('Clicking the Save button...');
        await page.click('#patient-demographicsBtn56');
        await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for save action to complete
        console.log('Process completed for account:', accountNumber);
        return true;
    } catch (error) {
        console.error(`Error processing account ${accountNumber}:`, error);
        return false;
    }
}


// Endpoint to process data and return status
app.post("/process", upload.single("file"), async (req, res) => {
    if (!req.file || !req.body.originalPath) {
        return res.status(400).json({ error: "No file or original path provided" });
    }

    let browser;
    try {
        const workbookXLSX = xlsx.read(req.file.buffer, { type: "buffer", cellDates: true ,
            cellNF: false,  // Reduce memory overhead
            bookVBA: false  // Skip VBA storage
        });

        if (workbookXLSX.Workbook?.WBProps?.Encrypted) {
            throw new Error("This file is password protected");
        }

        const sheetXLSX = workbookXLSX.Sheets[workbookXLSX.SheetNames[0]];
        const originalData = xlsx.utils.sheet_to_json(sheetXLSX, {
            header: 1,
            raw: false,
            dateNF: 'mm/dd/yyyy'
        });

        if (!originalData.length) {
            throw new Error("Empty file or missing headers.");
        }
        const headers = originalData[0].map(h => h?.toString().trim());
        const columnMap = {};
        
        // Check if "Acct No" exists in headers
        console.log(headers);
        if (!headers.includes("Acct No")) {
            throw new Error("'Acct No' Column is not defined.");
        }
        
        // Map headers to their index positions
        headers.forEach((header, index) => {
            columnMap[header] = index;
        });
        
        
        // Check if Result column exists, if not, add it
        if (!columnMap.hasOwnProperty("Result")) {
            columnMap["Result"] = headers.length;
            headers.push("Result");
            // Add the Result column to all existing rows
            for (let i = 1; i < originalData.length; i++) {
                if (originalData[i]) {
                    originalData[i].push("");
                }
            }
        }
        
        let newRowsToProcess = 0;
        const processedAccounts = new Set();
        
        let missingCallNotes = 0;  // Track rows missing Notes

        for (let i = 1; i < originalData.length; i++) {
            const row = originalData[i];
            const accountNumberRaw = row[columnMap["Acct No"]] || "";
            const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
            if (!accountNumberRaw || status === "done") continue;  // Skip processed rows
            
            const accountNumber = accountNumberRaw.split("-")[0].trim();
            processedAccounts.add(accountNumber);
            newRowsToProcess++;
        
            // Check if Notes is missing
            const callNotes = row[columnMap["Notes"]] || "";
            if (!callNotes.trim()) {
                missingCallNotes++;
            }
        }
        
        // ✅ Check if all remaining rows are missing Notes
        if (newRowsToProcess > 0 && newRowsToProcess === missingCallNotes) {
            return res.json({
                success: false,
                message: "Notes not found in Excel.",
                results: { successful: [], failed: [] }
            });
        }
        
        // ✅ Check if all rows are already updated
        if (newRowsToProcess === 0) {
            return res.json({
                success: true,
                message: "All rows were already updated.",
                results: { successful: [], failed: [] }
            });
        }
        
        
        // Launch Puppeteer (browser is only launched if VPN is connected)
        browser = await puppeteer.launch({
            headless: false,
            args: ["--ignore-certificate-errors",
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--reduce-security-for-testing'
            ],
            defaultViewport: null
        });
        const page = await browser.newPage();
        await loginToSystem(page);
        
        const results = { successful: [], failed: [] };
        
        for (let i = 1; i < originalData.length; i++) {
            const row = originalData[i];
            const accountNumberRaw = row[columnMap["Acct No"]] || "";
            const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
            // Skip row if there's no account number or it's already processed.
            if (!accountNumberRaw || status === "done") continue;
            
            // Skip row if the "Notes" field is empty.
            const callNotes = row[columnMap["Notes"]] || "";
            if (!callNotes.toString().trim()) continue;
            
            const accountNumber = accountNumberRaw.split("-")[0].trim();
            console.log(`Processing account: ${accountNumber}`);
        
            const success = await processAccount(page, accountNumber, callNotes);
        
            if (success) {
                results.successful.push(accountNumber);
                originalData[i][columnMap["Result"]] = "Done";
            } else {
                results.failed.push(accountNumber);
            }
        }
        
        // Convert data back to an ExcelJS workbook
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("Sheet1");

        // Write headers with green background
        const headerRow = sheet.addRow(headers);
        headerRow.eachCell((cell) => {
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFFF00" }
            };
            cell.font = { bold: true, color: { argb: "black" }, size: 10 };
            cell.alignment = { horizontal: "center", vertical: "middle" };
            cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" }
            };
        });

        // Write data back to the sheet
// In your column widths array, add width for EPS Comments if needed
const columnWidths = [10, 16, 20, 10, 5, 25, 28, 8, 27, 11, 16, 15, 12, 25, 12, 80, 80, 26, 12, 9, 9];

// When setting up columns, ensure EPS Comments has wrap text enabled
sheet.columns = columnWidths.map((width, index) => ({
    width,
    style: (index === columnMap["Notes"] || index === columnMap["EPS Comments"]) 
        ? { alignment: { wrapText: true } } 
        : {},
}));

// When formatting cells in each row, add special handling for EPS Comments
originalData.slice(1).forEach((row) => {
    const hasData = row.some(cell => cell && cell.toString().trim() !== "");
    const rowData = headers.map(header => row[columnMap[header]]);
    
    if (hasData) {
        const newRow = sheet.addRow(rowData);
        newRow.height = 150;

        headers.forEach((header, colIndex) => {
            const cell = newRow.getCell(colIndex + 1);

            // Apply border to all sides
            cell.border = {
                top: { style: "thin", color: { argb: "000000" } },
                left: { style: "thin", color: { argb: "000000" } },
                bottom: { style: "thin", color: { argb: "000000" } },
                right: { style: "thin", color: { argb: "000000" } }
            };

            // Apply default font size
            cell.font = { size: 10 };

            if (header === "Acct No" && cell.value) {
                cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: "FFFF00" } // Yellow
                };
                cell.alignment = { horizontal: "center", vertical: "bottom" };
            }
        
            // Formatting for "Notes"
            if (header === "Notes" && cell.value) {
                cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: "FFFF00" } // Yellow
                };
                cell.alignment = { wrapText: true, vertical: "bottom" };
            }

            // Formatting for "EPS Comments"
            if (header === "EPS Comments" && cell.value) {
                cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: "FFFFFF" } // White
                };
                cell.alignment = { wrapText: true, vertical: "bottom" };
            }
        });
    }
});

        // Generate a unique file ID
        const fileId = Date.now().toString();

        // Create a buffer of the Excel file and store it
        const fileBuffer = await workbook.xlsx.writeBuffer();
        processedFiles.set(fileId, {
            buffer: fileBuffer,
            filename: "updated_file.xlsx"
        });

        // Return the results and the file ID for later download
        res.json({
            success: true,
            message: `${newRowsToProcess} rows processed successfully.`,
            results,
            fileId
        });

    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({ error: error.message });
    } finally {
        if (browser) {
            await new Promise(resolve => setTimeout(resolve, 1000));
            await browser.close();
            console.log("Browser closed");
        }
    }
});


// Separate endpoint to download the processed file
app.get("/download/:fileId", (req, res) => {
    const fileId = req.params.fileId;
    const fileData = processedFiles.get(fileId);
    
    if (!fileData) {
        return res.status(404).json({ error: "File not found" });
    }
    
    // Set the response headers
    res.setHeader('Content-Disposition', `attachment; filename="${fileData.filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Length', fileData.buffer.length);
    
    // Send the file buffer
    res.send(fileData.buffer);
    
    // Optional: Remove the file from memory after some time
    setTimeout(() => {
        processedFiles.delete(fileId);
        console.log(`File ${fileId} removed from memory`);
    }, 30 * 60 * 1000); // Remove after 30 minutes
});

app.listen(port, () => console.log(`Server running on http://localhost:${port}`));
