// import express from 'express';
// import multer from 'multer';
// import cors from 'cors';
// import xlsx from 'xlsx';
// import ExcelJS from 'exceljs';
// import puppeteer from 'puppeteer';
// import fs from 'fs';
// import path from 'path';
// import { Readable } from 'stream';
// import https from 'https';
// import fetch from 'node-fetch';

// const app = express();
// const port = 5000;

// app.use(cors());
// app.use(express.json());

// const storage = multer.memoryStorage();
// const upload = multer({ storage });

// // Store processed data temporarily
// const processedFiles = new Map();

// // Private Access Connectivity Check Function
// async function isPrivateAccessConnected() {
//     try {
//         const controller = new AbortController();
//         const timeoutId = setTimeout(() => controller.abort(), 5000);

//         // Create an HTTPS agent that disables certificate validation
//         const agent = new https.Agent({ rejectUnauthorized: false });

//         const response = await fetch('https://maicllejgit6695ivkapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp', {
//             method: 'HEAD',
//             signal: controller.signal,
//             agent
//         });
//         clearTimeout(timeoutId);
//         return response.ok;
//     } catch (error) {
//         console.error("Private Access connectivity check failed:", error.message);
//         return false;
//     }
// }

// async function continueWithLoggedInSession() {
//     try {
//         const browser = await puppeteer.connect({
//             browserURL: 'http://localhost:9222', // Connect to the existing Chrome instance
//             defaultViewport: null,
//             timeout: 60000, // Increased timeout
//             headless: false, // Open in non-headless mode for debugging
//         });

//         // Get all open pages in the browser
//         const pages = await browser.pages();
        
//         // Find an existing page with the desired URL
//         let sessionPage = pages.find(p => p.url().includes('scheduling/resourceSchedule.jsp')); 

//         // If an existing session page is found, use it instead of opening a new one
//         if (sessionPage) {
//             console.log('Existing session found. Continuing with the current session...');
//             await new Promise(resolve => setTimeout(resolve, 500)); // Wait for 500 milliseconds
//             // Wait for the dropdown toggle to be visible
//             console.log('Waiting for three-line menu...');
//         await sessionPage.waitForSelector('#JellyBeanCountCntrl > div.fav-block > div', { visible: true, timeout: 10000 });

//         console.log('Clicking three-line menu...');
//         await sessionPage.click('#JellyBeanCountCntrl > div.fav-block > div');

//         // Wait for the Billing option inside the menu list
//         console.log('Waiting for Billing menu...');
//         await sessionPage.waitForSelector('#Billing_menu', { visible: true, timeout: 10000 });

//         // Scroll into view before clicking
//         console.log('Scrolling to Billing menu...');
//         await sessionPage.evaluate(() => {
//             const element = document.querySelector('#Billing_menu');
//             if (element) {
//                 element.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         await new Promise(resolve => setTimeout(resolve, 1000));
// // Wait for smooth scrolling effect

//         console.log('Clicking Billing menu...');
//         await sessionPage.click('#Billing_menu');

//         // Wait for Claims menu to appear
//         console.log('Waiting for Claims option...');
//         await sessionPage.waitForSelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)', { visible: true, timeout: 10000 });

//         // Scroll to Claims menu before clicking
//         console.log('Scrolling to Claims option...');
//         await sessionPage.evaluate(() => {
//             const element = document.querySelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');
//             if (element) {
//                 element.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         await new Promise(resolve => setTimeout(resolve, 1000));


//         console.log('Clicking Claims option...');
//         await sessionPage.click('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');

//         console.log('Successfully navigated to Claims!');

//             return sessionPage; // Return the existing session page
//         } else {
//             console.log('No existing session found. Skipping the process...');
//             // Simulate an alert on the browser if no session is found
//             await browser.pages().then(async (pages) => {
//                 const firstPage = pages[0]; // Assuming there's at least one open page
//                 await firstPage.evaluate(() => {
//                     alert('Session not found!');
//                 });
//             });
//             return null;  // No page is returned if no session is found
//         }
//     } catch (error) {
//         console.error("Error in continueWithLoggedInSession:", error);
//         throw error;
//     }
// }

// async function processAccount(newPage, accountNumber, callNotes, copayAmount) {
//     try {
//         // Set From Date to 01/01/2023
//         const fromDate = '01/01/2023';
//         console.log(`Setting From Date to: ${fromDate}`);
//         await newPage.waitForSelector('#fromdate', { visible: true });
//         await newPage.click('#fromdate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#fromdate', fromDate, { delay: 50 });

//         // Set To Date to current date in MM/DD/YYYY format
//         const now = new Date();
//         const toDate = `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
//         console.log(`Setting To Date to: ${toDate}`);
//         await newPage.waitForSelector('#todate', { visible: true });
//         await newPage.click('#todate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#todate', toDate, { delay: 50 });

//         // Select "Encounter/Claim" option from the dropdown
//         console.log('Selecting Claim Status "Encounter/Claim"...');
//         await newPage.waitForSelector('#claimStatusCodeId', { visible: true });

//         // Get the value of the 2nd option (Encounter/Claim)
//         const optionValue = await newPage.$eval('#claimStatusCodeId > option:nth-child(2)', el => el.value);
//         await newPage.select('#claimStatusCodeId', optionValue);

//         // Click the Claim Lookup button
//         console.log('Clicking the Claim Lookup button...');
//         await newPage.waitForSelector('#claimLookupBtn10', { visible: true });
//         await newPage.click('#claimLookupBtn10');
//         await new Promise(resolve => setTimeout(resolve, 3000));

//         // Click Create New Claim
//         console.log('Selecting Create New Claim option...');
//         await newPage.waitForSelector('#claimLookupBtn11', { visible: true });
//         await newPage.click('#claimLookupBtn11');
//         await new Promise(resolve => setTimeout(resolve, 3000));

//     } catch (error) {
//         console.error('Error processing account:', error);
//     }
// }


// // Endpoint to process data and return status
// app.post("/process", upload.single("file"), async (req, res) => {
//     if (!req.file || !req.body.originalPath) {
//         return res.status(400).json({ error: "No file or original path provided" });
//     }
    
//     const privateAccessConnected = await isPrivateAccessConnected();
//     if (!privateAccessConnected) {
//         return res.status(500).json({
//             error: "Private Access is not connected. Please connect to the Private Access"
//         });
//     }
    
//     let newPage = null;
    
//     try {
//         const workbookXLSX = xlsx.read(req.file.buffer, { 
//             type: "buffer", 
//             cellDates: true,
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
//             if (!row) continue; // Skip undefined rows
            
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
        
//         // Check if all remaining rows are missing Notes
//         if (newRowsToProcess > 0 && newRowsToProcess === missingCallNotes) {
//             return res.json({
//                 success: false,
//                 message: "Notes not found in Excel.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // Check if all rows are already updated
//         if (newRowsToProcess === 0) {
//             return res.json({
//                 success: true,
//                 message: "All rows were already updated.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // Get the active page for automation
//         newPage = await continueWithLoggedInSession();
        
//         const results = { successful: [], failed: [] };
        
//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue; // Skip undefined rows
            
//             const accountNumberRaw = row[columnMap["Acct No"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
//             // Skip row if there's no account number or it's already processed.
//             if (!accountNumberRaw || status === "done") continue;
            
//             // Skip row if the "Notes" field is empty.
//             const callNotes = row[columnMap["Notes"]] || "";
//             if (!callNotes.toString().trim()) continue;
            
//             // Get copay amount if column exists
//             let copayAmount = null;
//             if (columnMap.hasOwnProperty("Specialist Copay")) {
//                 copayAmount = row[columnMap["Specialist Copay"]] || "";
//                 // Skip copay update if it's empty or already has $ symbol
//                 if (!copayAmount.toString().trim() || copayAmount.toString().trim().startsWith('$')) {
//                     console.log(`Skipping copay update for account ${accountNumberRaw}: ${copayAmount ? 'Already has $ symbol' : 'Empty copay'}`);
//                 } else {
//                     console.log(`Will update copay for account ${accountNumberRaw}: ${copayAmount}`);
//                 }
//             }
            
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             console.log(`Processing account: ${accountNumber}`);
        
//             // Pass the newPage to processAccount
//             const success = await processAccount(newPage, accountNumber, callNotes, copayAmount);
        
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

//         // Write headers with orange background
//         const headerRow = sheet.addRow(headers);
//         headerRow.height = 42;
//         headerRow.eachCell((cell) => {
//             cell.fill = {
//                 type: "pattern",
//                 pattern: "solid",
//                 fgColor: { argb: "FFA500" }
//             };
//             cell.font = { bold: true, color: { argb: "000000" }, size: 10 };
//             cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
//             cell.border = {
//                 top: { style: "thin" },
//                 left: { style: "thin" },
//                 bottom: { style: "thin" },
//                 right: { style: "thin" }
//             };
//         });

//         // In your column widths array, add width for EPS Comments if needed
//         const columnWidthsWithSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 15, 50, 20];
//         const columnWidthsWithoutSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 50, 20];
        
//         // Function to check if Specialist Copay column exists
//         const hasSpecialistCopay = sheet.getRow(1).values.includes("Specialist Copay");
        
//         // Select appropriate column width array
//         const columnWidths = hasSpecialistCopay ? columnWidthsWithSpecialistCopay : columnWidthsWithoutSpecialistCopay;
        
//         // Set column widths dynamically
//         sheet.columns = columnWidths.map((width, index) => ({
//             width,
//             style: (index === columnMap["Notes"] || index === columnMap["EPS Comments"]) 
//                 ? { alignment: { wrapText: true } } 
//                 : {},
//         }));

//         // When formatting cells in each row, add special handling for EPS Comments
//         originalData.slice(1).forEach((row) => {
//             if (!row) return; // Skip undefined rows
            
//             const hasData = row.some(cell => cell && cell.toString().trim() !== "");
//             if (!hasData) return; // Skip empty rows
            
//             const rowData = headers.map(header => row[columnMap[header]] || "");
            
//             const newRow = sheet.addRow(rowData);
//             newRow.height = 42;

//             headers.forEach((header, colIndex) => {
//                 const cell = newRow.getCell(colIndex + 1);

//                 // Apply border to all sides
//                 cell.border = {
//                     top: { style: "thin", color: { argb: "000000" } },
//                     left: { style: "thin", color: { argb: "000000" } },
//                     bottom: { style: "thin", color: { argb: "000000" } },
//                     right: { style: "thin", color: { argb: "000000" } }
//                 };

//                 // Apply default font size
//                 cell.font = { size: 10 };

//                 if (header === "Acct No" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFF00" } // Yellow
//                     };
//                     cell.alignment = { horizontal: "center", vertical: "bottom" };
//                 }
            
//                 // Formatting for "Notes"
//                 if (header === "Notes" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFF00" } // Yellow
//                     };
//                     cell.alignment = { wrapText: true, vertical: "bottom" };
//                 }
//             });
//         });

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
//         if (newPage) {
//             await new Promise(resolve => setTimeout(resolve, 1000));
//             await newPage.close();
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



// import express from 'express';
// import multer from 'multer';
// import cors from 'cors';
// import xlsx from 'xlsx';
// import ExcelJS from 'exceljs';
// import puppeteer from 'puppeteer';
// import fs from 'fs';
// import path from 'path';
// import { Readable } from 'stream';
// import https from 'https';
// import fetch from 'node-fetch';

// const app = express();
// const port = 5000;

// app.use(cors());
// app.use(express.json());

// const storage = multer.memoryStorage();
// const upload = multer({ storage });

// // Store processed data temporarily
// const processedFiles = new Map();

// // Private Access Connectivity Check Function
// async function isPrivateAccessConnected() {
//     try {
//         const controller = new AbortController();
//         const timeoutId = setTimeout(() => controller.abort(), 5000);

//         // Create an HTTPS agent that disables certificate validation
//         const agent = new https.Agent({ rejectUnauthorized: false });

//         const response = await fetch('https://maicllejgit6695ivkapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp', {
//             method: 'HEAD',
//             signal: controller.signal,
//             agent
//         });
//         clearTimeout(timeoutId);
//         return response.ok;
//     } catch (error) {
//         console.error("Private Access connectivity check failed:", error.message);
//         return false;
//     }
// }

// async function continueWithLoggedInSession() {
//     try {
//         const browser = await puppeteer.connect({
//             browserURL: 'http://localhost:9222', // Connect to the existing Chrome instance
//             defaultViewport: null,
//             timeout: 60000, // Increased timeout
//             headless: false, // Open in non-headless mode for debugging
//         });

//         // Get all open pages in the browser
//         const pages = await browser.pages();
        
//         // Find an existing page with the desired URL
//         let sessionPage = pages.find(p => p.url().includes('scheduling/resourceSchedule.jsp')); 

//         // If an existing session page is found, use it instead of opening a new one
//         if (sessionPage) {
//             console.log('Existing session found. Continuing with the current session...');
//             await new Promise(resolve => setTimeout(resolve, 500)); // Wait for 500 milliseconds
//             // Wait for the dropdown toggle to be visible
//             console.log('Waiting for three-line menu...');
//         await sessionPage.waitForSelector('#JellyBeanCountCntrl > div.fav-block > div', { visible: true, timeout: 10000 });

//         console.log('Clicking three-line menu...');
//         await sessionPage.click('#JellyBeanCountCntrl > div.fav-block > div');

//         // Wait for the Billing option inside the menu list
//         console.log('Waiting for Billing menu...');
//         await sessionPage.waitForSelector('#Billing_menu', { visible: true, timeout: 10000 });

//         // Scroll into view before clicking
//         console.log('Scrolling to Billing menu...');
//         await sessionPage.evaluate(() => {
//             const element = document.querySelector('#Billing_menu');
//             if (element) {
//                 element.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         await new Promise(resolve => setTimeout(resolve, 1000));
// // Wait for smooth scrolling effect

//         console.log('Clicking Billing menu...');
//         await sessionPage.click('#Billing_menu');

//         // Wait for Claims menu to appear
//         console.log('Waiting for Claims option...');
//         await sessionPage.waitForSelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)', { visible: true, timeout: 10000 });

//         // Scroll to Claims menu before clicking
//         console.log('Scrolling to Claims option...');
//         await sessionPage.evaluate(() => {
//             const element = document.querySelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');
//             if (element) {
//                 element.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         await new Promise(resolve => setTimeout(resolve, 1000));


//         console.log('Clicking Claims option...');
//         await sessionPage.click('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');

//         console.log('Successfully navigated to Claims!');

//             return sessionPage; // Return the existing session page
//         } else {
//             console.log('No existing session found. Skipping the process...');
//             // Simulate an alert on the browser if no session is found
//             await browser.pages().then(async (pages) => {
//                 const firstPage = pages[0]; // Assuming there's at least one open page
//                 await firstPage.evaluate(() => {
//                     alert('Session not found!');
//                 });
//             });
//             return null;  // No page is returned if no session is found
//         }
//     } catch (error) {
//         console.error("Error in continueWithLoggedInSession:", error);
//         throw error;
//     }
// }
// async function processAccount(newPage, accountNumber, provider, facilityName, billingDate) {
//     try {
//         // Validate required fields
//         if (!provider || !facilityName || !billingDate || !accountNumber) {
//             console.error('Missing required fields:', { provider, facilityName, billingDate, accountNumber });
//             return false;
//         }

//         // Set From Date to 01/01/2023
//         const fromDate = '01/01/2023';
//         console.log(`Setting From Date to: ${fromDate}`);
//         await newPage.waitForSelector('#fromdate', { visible: true, timeout: 5000 });
//         await newPage.click('#fromdate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#fromdate', fromDate, { delay: 50 });

//         // Set To Date to current date in MM/DD/YYYY format
//         const now = new Date();
//         const toDate = `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
//         console.log(`Setting To Date to: ${toDate}`);
//         await newPage.waitForSelector('#todate', { visible: true, timeout: 5000 });
//         await newPage.click('#todate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#todate', toDate, { delay: 50 });

//         // Select "Encounter/Claim" option from the dropdown
//         console.log('Selecting Claim Status "Encounter/Claim"...');
//         await newPage.waitForSelector('#claimStatusCodeId', { visible: true, timeout: 5000 });
//         const optionValue = await newPage.$eval('#claimStatusCodeId > option:nth-child(2)', el => el.value);
//         await newPage.select('#claimStatusCodeId', optionValue);

//         // Click the Claim Lookup button
//         console.log('Clicking the Claim Lookup button...');
//         await newPage.waitForSelector('#claimLookupBtn10', { visible: true, timeout: 5000 });
//         await newPage.click('#claimLookupBtn10');
//         await new Promise(resolve => setTimeout(resolve, 3000));

//         // Click Create New Claim
//         console.log('Selecting Create New Claim option...');
//         await newPage.waitForSelector('#claimLookupBtn11', { visible: true, timeout: 5000 });
//         await newPage.click('#claimLookupBtn11');
//         await new Promise(resolve => setTimeout(resolve, 3000));

//         console.log(`Filling in Resource with Provider value: ${provider}`);

//        // Fill Provider and Resource fields
//         await newPage.waitForSelector("input[id^='provider-lookupIpt1']", { visible: true, timeout: 5000 });

//         // Clear the existing value
//         await newPage.click("input[id^='provider-lookupIpt1']", { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type("input[id^='provider-lookupIpt1']", provider, { delay: 30 });

//         await newPage.waitForSelector("input[id='resource-lookupIpt1']", { visible: true, timeout: 8000 });

//         // Clear the existing value
//         await newPage.click("input[id='resource-lookupIpt1']", { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type("input[id='resource-lookupIpt1']", provider, { delay: 30 });


//         // Fill in Facility Name
//         console.log(`Filling in Facility: ${facilityName}`);
//         await newPage.waitForSelector("input[id^='facility-lookupIpt1-']", { visible: true, timeout: 5000 });
//         await newPage.type("input[id^='facility-lookupIpt1-']", facilityName, { delay: 30 });

//         // Set Service Date (from Billing Date)
//         console.log(`Setting Billing Date: ${billingDate}`);
//         await newPage.waitForSelector('#claimserviceDate', { visible: true, timeout: 5000 });
//         await newPage.click('#claimserviceDate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#claimserviceDate', billingDate, { delay: 50 });

//         // Click Patient Search Button
//         console.log('Opening Patient Search...');
//         await newPage.waitForSelector('#searchPtBtn', { visible: true, timeout: 5000 });
//         await newPage.click('#searchPtBtn');

//         // Click Account Number (4th Option)
//         console.log('Selecting Account Number option...');
//         await newPage.waitForSelector('#patient-lookupLink8', { visible: true, timeout: 5000 });
//         await newPage.click('#patient-lookupLink8');

//         // Enter ECW# into Account Number field
//         console.log(`Entering Account Number: ${accountNumber}`);
//         await newPage.waitForSelector("input[id='ptLookup']", { visible: true, timeout: 5000 });
//         await newPage.type("input[id='ptLookup']", accountNumber, { delay: 50 });

//         // Final Step - Click OK Button
//         console.log('Clicking OK to create claim...');
//         await newPage.waitForSelector('#createClaimBtn', { visible: true, timeout: 5000 });
//         await newPage.click('#createClaimBtn');
        
//         return true; // Return success
//     } catch (error) {
//         console.error('Error processing account:', error);
//         return false; // Return failure
//     }
// }

// // Endpoint to process data and return status
// app.post("/process", upload.single("file"), async (req, res) => {
//     if (!req.file || !req.body.originalPath) {
//         return res.status(400).json({ error: "No file or original path provided" });
//     }
    
//     const privateAccessConnected = await isPrivateAccessConnected();
//     if (!privateAccessConnected) {
//         return res.status(500).json({
//             error: "Private Access is not connected. Please connect to the Private Access"
//         });
//     }
    
//     let newPage = null;
    
//     try {
//         const workbookXLSX = xlsx.read(req.file.buffer, { 
//             type: "buffer", 
//             cellDates: true,
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
        
//         // Check if "ECW #" exists in headers
//         console.log(headers);
//         if (!headers.includes("ECW #")) {
//             throw new Error("'ECW #' Column is not defined.");
//         }
        
//         // Map headers to their index positions
//         headers.forEach((header, index) => {
//             columnMap[header] = index;
//         });
        
//         // Check if required columns exist
//         const requiredColumns = ["ECW #", "Provider", "Facility Name", "Billing Date"];
//         for (const column of requiredColumns) {
//             if (!headers.includes(column)) {
//                 throw new Error(`Required column '${column}' is missing from the spreadsheet.`);
//             }
//         }
        
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
//             if (!row) continue; // Skip undefined rows
            
//             const accountNumberRaw = row[columnMap["ECW #"]] || "";
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
        
//         // Check if all remaining rows are missing Notes
//         if (newRowsToProcess > 0 && newRowsToProcess === missingCallNotes) {
//             return res.json({
//                 success: false,
//                 message: "Notes not found in Excel.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // Check if all rows are already updated
//         if (newRowsToProcess === 0) {
//             return res.json({
//                 success: true,
//                 message: "All rows were already updated.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // Get the active page for automation
//         newPage = await continueWithLoggedInSession();
        
//         const results = { successful: [], failed: [] };
        
//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue;
        
//             const accountNumberRaw = row[columnMap["ECW #"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
//             if (!accountNumberRaw || status === "done") continue;
        
//             const provider = row[columnMap["Provider"]] || "";
//             const facilityName = row[columnMap["Facility Name"]] || "";
//             const billingDate = row[columnMap["Billing Date"]] || "";
        
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             console.log(`Processing account: ${accountNumber}`);
            
//             // Validate data before processing
//             if (!provider) console.log('Provider value is missing or undefined');
//             if (!facilityName) console.log('Facility Name is missing or undefined');
//             if (!billingDate) console.log('Billing Date is missing or undefined');
        
//             const success = await processAccount(
//                 newPage,
//                 accountNumber,
//                 provider,
//                 facilityName,
//                 billingDate
//             );
        
//             if (success) {
//                 results.successful.push(accountNumber);
//                 originalData[i][columnMap["Result"]] = "Done";
//             } else {
//                 results.failed.push(accountNumber);
//                 originalData[i][columnMap["Result"]] = "Failed";
//             }
//         }
        
//         // Convert data back to an ExcelJS workbook
//         const workbook = new ExcelJS.Workbook();
//         const sheet = workbook.addWorksheet("Sheet1");

//         // Write headers with orange background
//         const headerRow = sheet.addRow(headers);
//         headerRow.height = 42;
//         headerRow.eachCell((cell) => {
//             cell.fill = {
//                 type: "pattern",
//                 pattern: "solid",
//                 fgColor: { argb: "FFA500" }
//             };
//             cell.font = { bold: true, color: { argb: "000000" }, size: 10 };
//             cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
//             cell.border = {
//                 top: { style: "thin" },
//                 left: { style: "thin" },
//                 bottom: { style: "thin" },
//                 right: { style: "thin" }
//             };
//         });

//         // In your column widths array, add width for EPS Comments if needed
//         const columnWidthsWithSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 15, 50, 20];
//         const columnWidthsWithoutSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 50, 20];
        
//         // Function to check if Specialist Copay column exists
//         const hasSpecialistCopay = headers.includes("Specialist Copay");
        
//         // Select appropriate column width array
//         const columnWidths = hasSpecialistCopay ? columnWidthsWithSpecialistCopay : columnWidthsWithoutSpecialistCopay;
        
//         // Set column widths dynamically based on actual number of columns
//         sheet.columns = headers.map((header, index) => {
//             // Use predefined width if available, otherwise use default width
//             const width = index < columnWidths.length ? columnWidths[index] : 15;
            
//             return {
//                 width,
//                 style: (header === "Notes" || header === "EPS Comments") 
//                     ? { alignment: { wrapText: true } } 
//                     : {},
//             };
//         });

//         // When formatting cells in each row, add special handling for EPS Comments
//         originalData.slice(1).forEach((row) => {
//             if (!row) return; // Skip undefined rows
            
//             const hasData = row.some(cell => cell && cell.toString().trim() !== "");
//             if (!hasData) return; // Skip empty rows
            
//             const rowData = headers.map(header => row[columnMap[header]] || "");
            
//             const newRow = sheet.addRow(rowData);
//             newRow.height = 42;

//             headers.forEach((header, colIndex) => {
//                 const cell = newRow.getCell(colIndex + 1);

//                 // Apply border to all sides
//                 cell.border = {
//                     top: { style: "thin", color: { argb: "000000" } },
//                     left: { style: "thin", color: { argb: "000000" } },
//                     bottom: { style: "thin", color: { argb: "000000" } },
//                     right: { style: "thin", color: { argb: "000000" } }
//                 };

//                 // Apply default font size
//                 cell.font = { size: 10 };

//                 if (header === "ECW #" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFF00" } // Yellow
//                     };
//                     cell.alignment = { horizontal: "center", vertical: "bottom" };
//                 }
            
//                 // Formatting for "Notes"
//                 if (header === "Notes" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFF00" } // Yellow
//                     };
//                     cell.alignment = { wrapText: true, vertical: "bottom" };
//                 }
                
//                 // Formatting for "Result"
//                 if (header === "Result" && cell.value) {
//                     const value = cell.value.toString().toLowerCase();
//                     if (value === "done") {
//                         cell.fill = {
//                             type: "pattern",
//                             pattern: "solid",
//                             fgColor: { argb: "92D050" } // Green
//                         };
//                     } else if (value === "failed") {
//                         cell.fill = {
//                             type: "pattern",
//                             pattern: "solid",
//                             fgColor: { argb: "FF0000" } // Red
//                         };
//                     }
//                     cell.alignment = { horizontal: "center", vertical: "middle" };
//                 }
//             });
//         });

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
//             message: `${results.successful.length} rows processed successfully. ${results.failed.length} rows failed.`,
//             results,
//             fileId
//         });

//     } catch (error) {
//         console.error("Error:", error);
//         res.status(500).json({ error: error.message });
//     } finally {
//         if (newPage) {
//             await new Promise(resolve => setTimeout(resolve, 1000));
//             await newPage.close();
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



// import express from 'express';
// import multer from 'multer';
// import cors from 'cors';
// import xlsx from 'xlsx';
// import ExcelJS from 'exceljs';
// import puppeteer from 'puppeteer';
// import fs from 'fs';
// import path from 'path';
// import { Readable } from 'stream';
// import https from 'https';
// import fetch from 'node-fetch';

// const app = express();
// const port = 5000;

// app.use(cors());
// app.use(express.json());

// const storage = multer.memoryStorage();
// const upload = multer({ storage });

// // Store processed data temporarily
// const processedFiles = new Map();

// // Private Access Connectivity Check Function
// async function isPrivateAccessConnected() {
//     try {
//         const controller = new AbortController();
//         const timeoutId = setTimeout(() => controller.abort(), 5000);

//         // Create an HTTPS agent that disables certificate validation
//         const agent = new https.Agent({ rejectUnauthorized: false });

//         const response = await fetch('https://maicllejgit6695ivkapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp', {
//             method: 'HEAD',
//             signal: controller.signal,
//             agent
//         });
//         clearTimeout(timeoutId);
//         return response.ok;
//     } catch (error) {
//         console.error("Private Access connectivity check failed:", error.message);
//         return false;
//     }
// }

// async function continueWithLoggedInSession() {
//     try {
//         const browser = await puppeteer.connect({
//             browserURL: 'http://localhost:9222', // Connect to the existing Chrome instance
//             defaultViewport: null,
//             timeout: 60000, // Increased timeout
//             headless: false, // Open in non-headless mode for debugging
//         });

//         // Get all open pages in the browser
//         const pages = await browser.pages();
        
//         // Find an existing page with the desired URL
//         let sessionPage = pages.find(p => p.url().includes('scheduling/resourceSchedule.jsp')); 

//         // If an existing session page is found, use it instead of opening a new one
//         if (sessionPage) {
//             console.log('Existing session found. Continuing with the current session...');
//             await new Promise(resolve => setTimeout(resolve, 500)); // Wait for 500 milliseconds
//             // Wait for the dropdown toggle to be visible
//             console.log('Waiting for three-line menu...');
//         await sessionPage.waitForSelector('#JellyBeanCountCntrl > div.fav-block > div', { visible: true, timeout: 10000 });

//         console.log('Clicking three-line menu...');
//         await sessionPage.click('#JellyBeanCountCntrl > div.fav-block > div');

//         // Wait for the Billing option inside the menu list
//         console.log('Waiting for Billing menu...');
//         await sessionPage.waitForSelector('#Billing_menu', { visible: true, timeout: 10000 });

//         // Scroll into view before clicking
//         console.log('Scrolling to Billing menu...');
//         await sessionPage.evaluate(() => {
//             const element = document.querySelector('#Billing_menu');
//             if (element) {
//                 element.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         await new Promise(resolve => setTimeout(resolve, 1000));
// // Wait for smooth scrolling effect

//         console.log('Clicking Billing menu...');
//         await sessionPage.click('#Billing_menu');

//         // Wait for Claims menu to appear
//         console.log('Waiting for Claims option...');
//         await sessionPage.waitForSelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)', { visible: true, timeout: 10000 });

//         // Scroll to Claims menu before clicking
//         console.log('Scrolling to Claims option...');
//         await sessionPage.evaluate(() => {
//             const element = document.querySelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');
//             if (element) {
//                 element.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         await new Promise(resolve => setTimeout(resolve, 1000));


//         console.log('Clicking Claims option...');
//         await sessionPage.click('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');

//         console.log('Successfully navigated to Claims!');

//             return sessionPage; // Return the existing session page
//         } else {
//             console.log('No existing session found. Skipping the process...');
//             // Simulate an alert on the browser if no session is found
//             await browser.pages().then(async (pages) => {
//                 const firstPage = pages[0]; // Assuming there's at least one open page
//                 await firstPage.evaluate(() => {
//                     alert('Session not found!');
//                 });
//             });
//             return null;  // No page is returned if no session is found
//         }
//     } catch (error) {
//         console.error("Error in continueWithLoggedInSession:", error);
//         throw error;
//     }
// }
// async function processAccount(newPage, accountNumber, provider, facilityName, billingDate) {
//     try {
//         // Validate required fields
//         if (!provider || !facilityName || !billingDate || !accountNumber) {
//             console.error('Missing required fields:', { provider, facilityName, billingDate, accountNumber });
//             return false;
//         }

//         // Set From Date to 01/01/2023
//         const fromDate = '01/01/2023';
//         console.log(`Setting From Date to: ${fromDate}`);
//         await newPage.waitForSelector('#fromdate', { visible: true, timeout: 5000 });
//         await newPage.click('#fromdate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#fromdate', fromDate, { delay: 50 });

//         // Set To Date to current date in MM/DD/YYYY format
//         const now = new Date();
//         const toDate = `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
//         console.log(`Setting To Date to: ${toDate}`);
//         await newPage.waitForSelector('#todate', { visible: true, timeout: 5000 });
//         await newPage.click('#todate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#todate', toDate, { delay: 50 });

//         // Select "Encounter/Claim" option from the dropdown
//         console.log('Selecting Claim Status "Encounter/Claim"...');
//         await newPage.waitForSelector('#claimStatusCodeId', { visible: true, timeout: 5000 });
//         const optionValue = await newPage.$eval('#claimStatusCodeId > option:nth-child(2)', el => el.value);
//         await newPage.select('#claimStatusCodeId', optionValue);

//         // Click the Claim Lookup button
//         console.log('Clicking the Claim Lookup button...');
//         await newPage.waitForSelector('#claimLookupBtn10', { visible: true, timeout: 5000 });
//         await newPage.click('#claimLookupBtn10');
//         await new Promise(resolve => setTimeout(resolve, 3000));

//         // Click Create New Claim
//         console.log('Selecting Create New Claim option...');
//         await newPage.waitForSelector('#claimLookupBtn11', { visible: true, timeout: 5000 });
//         await newPage.click('#claimLookupBtn11');
//         await new Promise(resolve => setTimeout(resolve, 3000));

//         console.log(`Filling in Resource with Provider value: ${provider}`);

//         // Fill Provider Input
//        // Fill Provider dropdown using exact match
// await newPage.waitForSelector("input[id^='provider-lookupIpt1']", { visible: true, timeout: 5000 });
// await newPage.click("input[id^='provider-lookupIpt1']", { clickCount: 3 });
// await newPage.keyboard.press('Backspace');

// await newPage.type("input[id^='provider-lookupIpt1']", provider.slice(0, 5), { delay: 30 });

// await newPage.waitForSelector("#provider-lookupUl1 li", { visible: true, timeout: 5000 });

// await newPage.evaluate((providerName) => {
//     const options = document.querySelectorAll("#provider-lookupUl1 li");
//     for (let opt of options) {
//         if (opt.innerText.trim() === providerName.trim()) {
//             opt.click();
//             break;
//         }
//     }
// }, provider);


//         // Clear and Fill Resource Input
//         await newPage.waitForSelector("input[id='resource-lookupIpt1']", { visible: true, timeout: 8000 });
//         await newPage.click("input[id='resource-lookupIpt1']", { clickCount: 3 }); // Click to select the text
//         await newPage.keyboard.press('Backspace'); // Clear the existing text
//         await new Promise(resolve => setTimeout(resolve, 200));

//        // Step 1: Clear field using evaluate
// await newPage.evaluate(() => {
//     const input = document.querySelector("#resource-lookupIpt1");
//     if (input) {
//         input.value = "";
//         input.dispatchEvent(new Event('input', { bubbles: true })); // ensure change triggers
//     }
// });

// // Step 2: Type the partial input
// await newPage.type("#resource-lookupIpt1", provider.slice(0, 5), { delay: 30 });

// // Step 3: Wait for dropdown to show
// await newPage.waitForSelector("#resource-lookupUl1 li", { visible: true, timeout: 5000 });

// // Step 4: Select full match
// await newPage.evaluate((providerName) => {
//     const options = document.querySelectorAll("#resource-lookupUl1 li");
//     for (let opt of options) {
//         if (opt.innerText.trim() === providerName.trim()) {
//             opt.click();
//             break;
//         }
//     }
// }, provider);


//         // Fill in Facility Name
//         console.log(`Filling in Facility: ${facilityName}`);
//         await newPage.waitForSelector("input[id^='facility-lookupIpt1-']", { visible: true, timeout: 5000 });
//         await newPage.type("input[id^='facility-lookupIpt1-']", facilityName, { delay: 50 });

//         // Set Service Date (from Billing Date)
//         console.log(`Setting Billing Date: ${billingDate}`);
//         await newPage.waitForSelector('#claimserviceDate', { visible: true, timeout: 5000 });
//         await newPage.click('#claimserviceDate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#claimserviceDate', billingDate, { delay: 50 });

//         // Click Patient Search Button
//         console.log('Opening Patient Search...');
//         await newPage.waitForSelector('#searchPtBtn', { visible: true, timeout: 5000 });
//         await newPage.click('#searchPtBtn');

//         // Click Account Number (4th Option)
//         console.log('Selecting Account Number option...');
//         await newPage.waitForSelector('#patient-lookupLink8', { visible: true, timeout: 5000 });
//         await newPage.click('#patient-lookupLink8');

//         // Enter ECW# into Account Number field
//         console.log(`Entering Account Number: ${accountNumber}`);
//         await newPage.waitForSelector("input[id='ptLookup']", { visible: true, timeout: 5000 });
//         await newPage.type("input[id='ptLookup']", accountNumber, { delay: 50 });

//         // Final Step - Click OK Button
//         console.log('Clicking OK to create claim...');
//         await newPage.waitForSelector('#createClaimBtn', { visible: true, timeout: 5000 });
//         await newPage.click('#createClaimBtn');
        
//         return true; // Return success
//     } catch (error) {
//         console.error('Error processing account:', error);
//         return false; // Return failure
//     }
// }

// // Endpoint to process data and return status
// app.post("/process", upload.single("file"), async (req, res) => {
//     if (!req.file || !req.body.originalPath) {
//         return res.status(400).json({ error: "No file or original path provided" });
//     }
    
//     const privateAccessConnected = await isPrivateAccessConnected();
//     if (!privateAccessConnected) {
//         return res.status(500).json({
//             error: "Private Access is not connected. Please connect to the Private Access"
//         });
//     }
    
//     let newPage = null;
    
//     try {
//         const workbookXLSX = xlsx.read(req.file.buffer, { 
//             type: "buffer", 
//             cellDates: true,
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
        
//         // Check if "ECW #" exists in headers
//         console.log(headers);
//         if (!headers.includes("ECW #")) {
//             throw new Error("'ECW #' Column is not defined.");
//         }
        
//         // Map headers to their index positions
//         headers.forEach((header, index) => {
//             columnMap[header] = index;
//         });
        
//         // Check if required columns exist
//         const requiredColumns = ["ECW #", "Provider", "Facility Name", "Billing Date"];
//         for (const column of requiredColumns) {
//             if (!headers.includes(column)) {
//                 throw new Error(`Required column '${column}' is missing from the spreadsheet.`);
//             }
//         }
        
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
//             if (!row) continue; // Skip undefined rows
            
//             const accountNumberRaw = row[columnMap["ECW #"]] || "";
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
        
//         // Check if all remaining rows are missing Notes
//         if (newRowsToProcess > 0 && newRowsToProcess === missingCallNotes) {
//             return res.json({
//                 success: false,
//                 message: "Notes not found in Excel.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // Check if all rows are already updated
//         if (newRowsToProcess === 0) {
//             return res.json({
//                 success: true,
//                 message: "All rows were already updated.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // Get the active page for automation
//         newPage = await continueWithLoggedInSession();
        
//         const results = { successful: [], failed: [] };
        
//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue;
        
//             const accountNumberRaw = row[columnMap["ECW #"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
//             if (!accountNumberRaw || status === "done") continue;
        
//             const provider = row[columnMap["Provider"]] || "";
//             const facilityName = row[columnMap["Facility Name"]] || "";
//             const billingDate = row[columnMap["Billing Date"]] || "";
        
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             console.log(`Processing account: ${accountNumber}`);
            
//             // Validate data before processing
//             if (!provider) console.log('Provider value is missing or undefined');
//             if (!facilityName) console.log('Facility Name is missing or undefined');
//             if (!billingDate) console.log('Billing Date is missing or undefined');
        
//             const success = await processAccount(
//                 newPage,
//                 accountNumber,
//                 provider,
//                 facilityName,
//                 billingDate
//             );
        
//             if (success) {
//                 results.successful.push(accountNumber);
//                 originalData[i][columnMap["Result"]] = "Done";
//             } else {
//                 results.failed.push(accountNumber);
//                 originalData[i][columnMap["Result"]] = "Failed";
//             }
//         }
        
//         // Convert data back to an ExcelJS workbook
//         const workbook = new ExcelJS.Workbook();
//         const sheet = workbook.addWorksheet("Sheet1");

//         // Write headers with orange background
//         const headerRow = sheet.addRow(headers);
//         headerRow.height = 42;
//         headerRow.eachCell((cell) => {
//             cell.fill = {
//                 type: "pattern",
//                 pattern: "solid",
//                 fgColor: { argb: "FFA500" }
//             };
//             cell.font = { bold: true, color: { argb: "000000" }, size: 10 };
//             cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
//             cell.border = {
//                 top: { style: "thin" },
//                 left: { style: "thin" },
//                 bottom: { style: "thin" },
//                 right: { style: "thin" }
//             };
//         });

//         // In your column widths array, add width for EPS Comments if needed
//         const columnWidthsWithSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 15, 50, 20];
//         const columnWidthsWithoutSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 50, 20];
        
//         // Function to check if Specialist Copay column exists
//         const hasSpecialistCopay = headers.includes("Specialist Copay");
        
//         // Select appropriate column width array
//         const columnWidths = hasSpecialistCopay ? columnWidthsWithSpecialistCopay : columnWidthsWithoutSpecialistCopay;
        
//         // Set column widths dynamically based on actual number of columns
//         sheet.columns = headers.map((header, index) => {
//             // Use predefined width if available, otherwise use default width
//             const width = index < columnWidths.length ? columnWidths[index] : 15;
            
//             return {
//                 width,
//                 style: (header === "Notes" || header === "EPS Comments") 
//                     ? { alignment: { wrapText: true } } 
//                     : {},
//             };
//         });

//         // When formatting cells in each row, add special handling for EPS Comments
//         originalData.slice(1).forEach((row) => {
//             if (!row) return; // Skip undefined rows
            
//             const hasData = row.some(cell => cell && cell.toString().trim() !== "");
//             if (!hasData) return; // Skip empty rows
            
//             const rowData = headers.map(header => row[columnMap[header]] || "");
            
//             const newRow = sheet.addRow(rowData);
//             newRow.height = 42;

//             headers.forEach((header, colIndex) => {
//                 const cell = newRow.getCell(colIndex + 1);

//                 // Apply border to all sides
//                 cell.border = {
//                     top: { style: "thin", color: { argb: "000000" } },
//                     left: { style: "thin", color: { argb: "000000" } },
//                     bottom: { style: "thin", color: { argb: "000000" } },
//                     right: { style: "thin", color: { argb: "000000" } }
//                 };

//                 // Apply default font size
//                 cell.font = { size: 10 };

//                 if (header === "ECW #" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFF00" } // Yellow
//                     };
//                     cell.alignment = { horizontal: "center", vertical: "bottom" };
//                 }
            
//                 // Formatting for "Notes"
//                 if (header === "Notes" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFF00" } // Yellow
//                     };
//                     cell.alignment = { wrapText: true, vertical: "bottom" };
//                 }
                
//                 // Formatting for "Result"
//                 if (header === "Result" && cell.value) {
//                     const value = cell.value.toString().toLowerCase();
//                     if (value === "done") {
//                         cell.fill = {
//                             type: "pattern",
//                             pattern: "solid",
//                             fgColor: { argb: "92D050" } // Green
//                         };
//                     } else if (value === "failed") {
//                         cell.fill = {
//                             type: "pattern",
//                             pattern: "solid",
//                             fgColor: { argb: "FF0000" } // Red
//                         };
//                     }
//                     cell.alignment = { horizontal: "center", vertical: "middle" };
//                 }
//             });
//         });

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
//             message: `${results.successful.length} rows processed successfully. ${results.failed.length} rows failed.`,
//             results,
//             fileId
//         });

//     } catch (error) {
//         console.error("Error:", error);
//         res.status(500).json({ error: error.message });
//     } finally {
//         if (newPage) {
//             await new Promise(resolve => setTimeout(resolve, 1000));
//             await newPage.close();
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



// import express from 'express';
// import multer from 'multer';
// import cors from 'cors';
// import xlsx from 'xlsx';
// import ExcelJS from 'exceljs';
// import puppeteer from 'puppeteer';
// import fs from 'fs';
// import path from 'path';
// import { Readable } from 'stream';
// import https from 'https';
// import fetch from 'node-fetch';

// const app = express();
// const port = 5000;

// app.use(cors());
// app.use(express.json());

// const storage = multer.memoryStorage();
// const upload = multer({ storage });

// // Store processed data temporarily
// const processedFiles = new Map();

// // Private Access Connectivity Check Function
// async function isPrivateAccessConnected() {
//     try {
//         const controller = new AbortController();
//         const timeoutId = setTimeout(() => controller.abort(), 5000);

//         // Create an HTTPS agent that disables certificate validation
//         const agent = new https.Agent({ rejectUnauthorized: false });

//         const response = await fetch('https://maicllejgit6695ivkapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp', {
//             method: 'HEAD',
//             signal: controller.signal,
//             agent
//         });
//         clearTimeout(timeoutId);
//         return response.ok;
//     } catch (error) {
//         console.error("Private Access connectivity check failed:", error.message);
//         return false;
//     }
// }

// async function continueWithLoggedInSession() {
//     try {
//         const browser = await puppeteer.connect({
//             browserURL: 'http://localhost:9222', // Connect to the existing Chrome instance
//             defaultViewport: null,
//             timeout: 120000, // Increased timeout
//             headless: false, // Open in non-headless mode for debugging
//         });

//         // Get all open pages in the browser
//         const pages = await browser.pages();
        
//         // Find an existing page with the desired URL
//         let sessionPage = pages.find(p => p.url().includes('scheduling/resourceSchedule.jsp')); 

//         // If an existing session page is found, use it instead of opening a new one
//         if (sessionPage) {
//             console.log('Existing session found. Continuing with the current session...');
//             await new Promise(resolve => setTimeout(resolve, 500)); // Wait for 500 milliseconds
//             // Wait for the dropdown toggle to be visible
//             console.log('Waiting for three-line menu...');
//         await sessionPage.waitForSelector('#JellyBeanCountCntrl > div.fav-block > div', { visible: true, timeout: 10000 });

//         console.log('Clicking three-line menu...');
//         await sessionPage.click('#JellyBeanCountCntrl > div.fav-block > div');

//         // Wait for the Billing option inside the menu list
//         console.log('Waiting for Billing menu...');
//         await sessionPage.waitForSelector('#Billing_menu', { visible: true, timeout: 10000 });

//         // Scroll into view before clicking
//         console.log('Scrolling to Billing menu...');
//         await sessionPage.evaluate(() => {
//             const element = document.querySelector('#Billing_menu');
//             if (element) {
//                 element.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         await new Promise(resolve => setTimeout(resolve, 1000));
// // Wait for smooth scrolling effect

//         console.log('Clicking Billing menu...');
//         await sessionPage.click('#Billing_menu');

//         // Wait for Claims menu to appear
//         console.log('Waiting for Claims option...');
//         await sessionPage.waitForSelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)', { visible: true, timeout: 10000 });

//         // Scroll to Claims menu before clicking
//         console.log('Scrolling to Claims option...');
//         await sessionPage.evaluate(() => {
//             const element = document.querySelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');
//             if (element) {
//                 element.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         await new Promise(resolve => setTimeout(resolve, 1000));


//         console.log('Clicking Claims option...');
//         await sessionPage.click('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');

//         console.log('Successfully navigated to Claims!');

//             return sessionPage; // Return the existing session page
//         } else {
//             console.log('No existing session found. Skipping the process...');
//             // Simulate an alert on the browser if no session is found
//             await browser.pages().then(async (pages) => {
//                 const firstPage = pages[0]; // Assuming there's at least one open page
//                 await firstPage.evaluate(() => {
//                     alert('Session not found!');
//                 });
//             });
//             return null;  // No page is returned if no session is found
//         }
//     } catch (error) {
//         console.error("Error in continueWithLoggedInSession:", error);
//         throw error;
//     }
// }
// async function processAccount(newPage, accountNumber, provider, facilityName, billingDate, diagnosisCodes) {
//     try {
//         // Validate required fields
//         if (!provider || !facilityName || !billingDate || !accountNumber) {
//             console.error('Missing required fields:', { provider, facilityName, billingDate, accountNumber });
//             return false;
//         }
//         // Set From Date to 01/01/2023
//         const fromDate = '01/01/2023';
//         console.log(`Setting From Date to: ${fromDate}`);
//         await newPage.waitForSelector('#fromdate', { visible: true, timeout: 5000 });
//         await newPage.click('#fromdate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#fromdate', fromDate, { delay: 50 });

//         // Set To Date to current date in MM/DD/YYYY format
//         const now = new Date();
//         const toDate = `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
//         console.log(`Setting To Date to: ${toDate}`);
//         await newPage.waitForSelector('#todate', { visible: true, timeout: 5000 });
//         await newPage.click('#todate', { clickCount: 3 });
//         await newPage.keyboard.press('Backspace');
//         await newPage.type('#todate', toDate, { delay: 50 });

//         // Select "Encounter/Claim" option from the dropdown
//         console.log('Selecting Claim Status "Encounter/Claim"...');
//         await newPage.waitForSelector('#claimStatusCodeId', { visible: true, timeout: 5000 });
//         const optionValue = await newPage.$eval('#claimStatusCodeId > option:nth-child(2)', el => el.value);
//         await newPage.select('#claimStatusCodeId', optionValue);

//         // Click the Claim Lookup button
//         console.log('Clicking the Claim Lookup button...');
//         await newPage.waitForSelector('#claimLookupBtn10', { visible: true, timeout: 5000 });
//         await newPage.click('#claimLookupBtn10');
//         await new Promise(resolve => setTimeout(resolve, 3000));

//         // Click Create New Claim
//         console.log('Selecting Create New Claim option...');
//         await newPage.waitForSelector('#claimLookupBtn11', { visible: true, timeout: 5000 });
//         await newPage.click('#claimLookupBtn11');
//         await new Promise(resolve => setTimeout(resolve, 3000));

//         console.log(`Filling in Resource with Provider value: ${provider}`);

// // Extract the second part of the provider name (last name)
// const providerLastName = provider.split(' ').slice(-1).join(' ');
// console.log(`Using provider's last name: ${providerLastName}`);

// // Fill Provider Input - Using only the last name (second part of the provider name)
// await newPage.waitForSelector("input[id^='provider-lookupIpt1']", { visible: true, timeout: 5000 });
// await newPage.click("input[id^='provider-lookupIpt1']", { clickCount: 3 });
// await newPage.keyboard.press('Backspace');
// await newPage.type("input[id^='provider-lookupIpt1']", providerLastName, { delay: 30 });

// await newPage.waitForSelector("#provider-lookupUl1 li", { visible: true, timeout: 5000 });

// await newPage.evaluate((providerLastName) => {
//     const options = document.querySelectorAll("#provider-lookupUl1 li");
//     for (let opt of options) {
//         if (opt.innerText.trim().toLowerCase() === providerLastName.trim().toLowerCase()) {
//             opt.click();
//             break;
//         }
//     }
// }, providerLastName);

// // Clear and Fill Resource Input - Using only the last name (second part of the resource name)
// // Clear and Fill Resource Input
// await newPage.waitForSelector("#resource-lookupIpt1", { visible: true, timeout: 8000 });

// // First, clear the field completely
// await newPage.click("#resource-lookupIpt1", { clickCount: 3 });
// await newPage.keyboard.press('Backspace');

// // Use evaluate to ensure the field is completely cleared
// await newPage.evaluate(() => {
//     const input = document.querySelector("#resource-lookupIpt1");
//     input.value = "";
//     // Trigger input event to ensure any listeners know the input changed
//     input.dispatchEvent(new Event('input', { bubbles: true }));
// });

// // Type the provider's last name
// await newPage.type("#resource-lookupIpt1", providerLastName, { delay: 50 });

// // Wait for the dropdown to appear and contain options
// await newPage.waitForSelector("#resource-lookupUl1 li.list-item-filter.ng-scope", { visible: true, timeout: 5000 });

// // Get all dropdown items
// const dropdownItems = await newPage.$$("#resource-lookupUl1 li.list-item-filter.ng-scope");

// let optionFound = false;

// // Loop through the dropdown items and click the matching option
// for (let item of dropdownItems) {
//     const itemText = await newPage.evaluate(el => el.innerText.trim(), item);
//     if (itemText.toLowerCase().includes(providerLastName.toLowerCase())) {
//         await item.click();
//         optionFound = true;
//         break;
//     }
// }

// if (!optionFound) {
//     console.error('No matching resource option found');
// }

// // Wait for the selection to be applied
// await new Promise(resolve => setTimeout(resolve, 1000));

// // Optionally, you can add additional waits or checks here if needed.
// // Fill in Facility Name
// const inputSelector = "input.createClaimFacilityCls";

// // Wait for the input to be visible
// await newPage.waitForSelector(inputSelector, { visible: true, timeout: 10000 });

// // Clear the field first in case there's existing text
// await newPage.evaluate((selector) => {
//     document.querySelector(selector).value = '';
// }, inputSelector);

// // Type the facility name into the input field
// await newPage.type(inputSelector, facilityName, { delay: 50 });

// // Wait briefly to allow the application to process the input
// await new Promise(resolve => setTimeout(resolve, 1000));

// // Try to find the dropdown using multiple possible selectors
// try {
//     // First try the specific ID
//     await newPage.waitForSelector("#facility-lookupUl1", { visible: true, timeout: 2000 });
// } catch (e) {
//     // If that fails, try a more general selector for dropdowns
//     try {
//         await newPage.waitForSelector("ul.dropdown-menu", { visible: true, timeout: 2000 });
//     } catch (err) {
//         // If no visible dropdown, press arrow down to try to trigger it
//         await newPage.keyboard.press('ArrowDown');
//         await new Promise(resolve => setTimeout(resolve, 500));
//     }
// }

// // Try to select an option using multiple approaches
// try {
//     // First try to find and click an option with JavaScript
//     const found = await newPage.evaluate((facilityName) => {
//         // Try multiple selectors for dropdown options
//         const selectors = [
//             "#facility-lookupUl1 li", 
//             "ul.dropdown-menu li",
//             ".dropdown-menu li",
//             "li[role='option']"
//         ];
        
//         let options = [];
//         // Try each selector until we find options
//         for (const selector of selectors) {
//             options = document.querySelectorAll(selector);
//             if (options.length > 0) break;
//         }
        
//         // If we found options, try to find a match and click it
//         if (options.length > 0) {
//             for (let opt of options) {
//                 if (opt.textContent.trim().toLowerCase().includes(facilityName.trim().toLowerCase())) {
//                     opt.click();
//                     return true;
//                 }
//             }
//         }
//         return false;
//     }, facilityName);
    
//     // If JavaScript selection didn't work, use keyboard navigation
//     if (!found) {
//         // Press arrow down to select first option
//         await newPage.keyboard.press('ArrowDown');
//         await new Promise(resolve => setTimeout(resolve, 300));
//         // Press Enter to confirm selection
//         await newPage.keyboard.press('Enter');
//     }
// } catch (e) {
//     // If all else fails, try keyboard navigation
//     await newPage.keyboard.press('ArrowDown');
//     await new Promise(resolve => setTimeout(resolve, 300));
//     await newPage.keyboard.press('Enter');
// }

// // Wait to ensure the selection is processed
// await new Promise(resolve => setTimeout(resolve, 1000));
//         // Set Service Date (from Billing Date)
//         console.log(`Setting Billing Date: ${billingDate}`);
//         await newPage.click('#claimserviceDate', { clickCount: 3 });
//         await newPage.focus('#claimserviceDate');
//         await newPage.keyboard.down('Control');
//         await newPage.keyboard.press('A');
//         await newPage.keyboard.up('Control');
//         await newPage.keyboard.press('Backspace');
        
//         // Double check it's cleared
//         await newPage.evaluate(() => {
//           document.querySelector('#claimserviceDate').value = '';
//         });
        
//         await newPage.type('#claimserviceDate', billingDate, { delay: 50 });
//         await newPage.keyboard.press('Tab');
//         const ainputSelector = "input.patientFilterForClaim";

//         // Step 1: Wait for input and type
//         await newPage.waitForSelector(ainputSelector, { visible: true, timeout: 10000 });
//         await newPage.evaluate(selector => {
//             const input = document.querySelector(selector);
//             if (input) {
//                 input.value = '';
//                 input.setAttribute('autocomplete', 'off');
//             }
//         }, ainputSelector);
        
//         await newPage.type(ainputSelector, accountNumber, { delay: 50 });
        
//         // Step 2: Wait for dropdown to appear
//         await new Promise(resolve => setTimeout(resolve, 1500));
//         const dropdownSelectors = [
//             "#patient-lookupUl1 li",
//             "ul.dropdown-menu li",
//             ".dropdown-menu li",
//             "li[role='option']"
//         ];
        
//         let clicked = false;
        
//         for (const sel of dropdownSelectors) {
//             const items = await newPage.$$(sel);
        
//             for (const item of items) {
//                 const text = await newPage.evaluate(el => el.textContent.trim(), item);
//                 if (text.toLowerCase().includes(accountNumber.toLowerCase())) {
//                     const box = await item.boundingBox();
//                     if (box) {
//                         await newPage.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
//                         await newPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { delay: 50 });
//                         clicked = true;
//                         break;
//                     }
//                 }
//             }
        
//             if (clicked) break;
//         }
        
//         // Step 3: Fallback if not clicked
//         if (!clicked) {
//             console.warn('Dropdown item not found with bounding box. Using keyboard fallback.');
//             await newPage.keyboard.press('ArrowDown');
//             await new Promise(resolve => setTimeout(resolve, 300));
//             await newPage.keyboard.press('Enter');
//         }
        
//         // Step 4: Trigger AngularJS model updates
//         await newPage.evaluate((selector) => {
//             const input = document.querySelector(selector);
//             if (input) {
//                 input.value = input.value.trim();
//                 const angularInput = angular.element(input);
//                 angularInput.triggerHandler('input');
//                 angularInput.triggerHandler('change');
//                 input.dispatchEvent(new Event('input', { bubbles: true }));
//                 input.dispatchEvent(new Event('change', { bubbles: true }));
//                 input.focus();
//                 setTimeout(() => input.blur(), 100);
//             }
//         }, ainputSelector);
        
        
//         // Optionally, wait a bit to let things settle
//         await new Promise(resolve => setTimeout(resolve, 500));
        
//         // Final Step - Click OK Button
//         const buttonExists = await newPage.$('#createClaimBtn');
//         if (buttonExists) {
//             console.log('Create Claim button found. Clicking...');
//             await newPage.click('#createClaimBtn');
//         } else {
//             console.error('Create Claim button not found');
//         }
//         await newPage.waitForNavigation({ waitUntil: 'load' });


//         // Check if diagnosisCodes is a string
//         console.log(`Entering Diagnosis Codes for Account: ${accountNumber}`);
//         console.log('diagnosisCodes value:', diagnosisCodes);
        
//         // Normalize diagnosis codes into an array of codes
//         let codes = [];
        
//         if (typeof diagnosisCodes === 'string') {
//             codes = diagnosisCodes.split('\n').map(code => code.trim()).filter(code => code.length > 0);
//         } else if (Array.isArray(diagnosisCodes)) {
//             codes = diagnosisCodes.map(code => code.trim()).filter(code => code.length > 0);
//         } else {
//             console.error(`Unexpected type for diagnosisCodes: ${typeof diagnosisCodes}`);
//             return;
//         }
        
//         for (let code of codes) {
//             try {
//                 await newPage.waitForSelector('#txtnewIcd1744613464218', { visible: true, timeout: 5000 });
//                 await newPage.click('#txtnewIcd1744613464218', { clickCount: 3 });
//                 await newPage.keyboard.press('Backspace');
//                 await newPage.type('#txtnewIcd1744613464218', code, { delay: 50 });
//                 await newPage.keyboard.press('Enter');
        
//                 // Check for error modal
//                 try {
//                     await newPage.waitForSelector(
//                         'body > div.bootbox.modal.fade.bluetheme.medium-width.in > div > div > div.modal-footer > button',
//                         { visible: true, timeout: 3000 }
//                     );
//                     console.log(`Error occurred while entering code: ${code}, clicking OK`);
//                     await newPage.click(
//                         'body > div.bootbox.modal.fade.bluetheme.medium-width.in > div > div > div.modal-footer > button'
//                     );
//                 } catch {
//                     console.log('No error modal appeared');
//                 }
        
//             } catch (error) {
//                 console.error('Error entering diagnosis code:', code, error);
//             }
//         }
        
//         return true; // Return success
//     } catch (error) {
//         console.error('Error processing account:', error);
//         return false; // Return failure
//     }
// }

// // Endpoint to process data and return status
// app.post("/process", upload.single("file"), async (req, res) => {
//     if (!req.file || !req.body.originalPath) {
//         return res.status(400).json({ error: "No file or original path provided" });
//     }
    
//     const privateAccessConnected = await isPrivateAccessConnected();
//     if (!privateAccessConnected) {
//         return res.status(500).json({
//             error: "Private Access is not connected. Please connect to the Private Access"
//         });
//     }
    
//     let newPage = null;
    
//     try {
//         const workbookXLSX = xlsx.read(req.file.buffer, { 
//             type: "buffer", 
//             cellDates: true,
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
        
//         // Check if "ECW #" exists in headers
//         console.log(headers);
//         if (!headers.includes("ECW #")) {
//             throw new Error("'ECW #' Column is not defined.");
//         }
        
//         // Map headers to their index positions
//         headers.forEach((header, index) => {
//             columnMap[header] = index;
//         });
        
//         // Check if required columns exist
//         const requiredColumns = ["ECW #", "Provider", "Facility Name", "Billing Date"];
//         for (const column of requiredColumns) {
//             if (!headers.includes(column)) {
//                 throw new Error(`Required column '${column}' is missing from the spreadsheet.`);
//             }
//         }
        
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
//             if (!row) continue; // Skip undefined rows
            
//             const accountNumberRaw = row[columnMap["ECW #"]] || "";
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
        
//         // Check if all remaining rows are missing Notes
//         if (newRowsToProcess > 0 && newRowsToProcess === missingCallNotes) {
//             return res.json({
//                 success: false,
//                 message: "Notes not found in Excel.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // Check if all rows are already updated
//         if (newRowsToProcess === 0) {
//             return res.json({
//                 success: true,
//                 message: "All rows were already updated.",
//                 results: { successful: [], failed: [] }
//             });
//         }
        
//         // Get the active page for automation
//         newPage = await continueWithLoggedInSession();
        
//         const results = { successful: [], failed: [] };
        
//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue;
        
//             const accountNumberRaw = row[columnMap["ECW #"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
//             if (!accountNumberRaw || status === "done") continue;
        
//             const provider = row[columnMap["Provider"]] || "";
//             const facilityName = row[columnMap["Facility Name"]] || "";
//             const billingDate = row[columnMap["Billing Date"]] || "";
//             const diagnosisText = row[columnMap["Diagnosis"]] || "";
        
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             console.log(`Processing account: ${accountNumber}`);
        
//             // Validate data before processing
//             if (!provider) console.log('Provider value is missing or undefined');
//             if (!facilityName) console.log('Facility Name is missing or undefined');
//             if (!billingDate) console.log('Billing Date is missing or undefined');
        
//             // Extract diagnosis codes from the text inside parentheses
//             const diagnosisCodes = [];
//             const regex = /\(([A-Z]\d{1,3}[A-Z]?\d{0,3})\)/gi;
//             let match;
//             while ((match = regex.exec(diagnosisText)) !== null) {
//                 diagnosisCodes.push(match[1]);
//             }
        
//             console.log(`Diagnosis Codes for ${accountNumber}:`, diagnosisCodes);
        
//             const success = await processAccount(
//                 newPage,
//                 accountNumber,
//                 provider,
//                 facilityName,
//                 billingDate,
//                 diagnosisCodes // Pass extracted codes here
//             );
        
//             if (success) {
//                 results.successful.push(accountNumber);
//                 originalData[i][columnMap["Result"]] = "Done";
//             } else {
//                 results.failed.push(accountNumber);
//                 originalData[i][columnMap["Result"]] = "Failed";
//             }
//         }
        
//         // Convert data back to an ExcelJS workbook
//         const workbook = new ExcelJS.Workbook();
//         const sheet = workbook.addWorksheet("Sheet1");

//         // Write headers with orange background
//         const headerRow = sheet.addRow(headers);
//         headerRow.height = 42;
//         headerRow.eachCell((cell) => {
//             cell.fill = {
//                 type: "pattern",
//                 pattern: "solid",
//                 fgColor: { argb: "FFA500" }
//             };
//             cell.font = { bold: true, color: { argb: "000000" }, size: 10 };
//             cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
//             cell.border = {
//                 top: { style: "thin" },
//                 left: { style: "thin" },
//                 bottom: { style: "thin" },
//                 right: { style: "thin" }
//             };
//         });

//         // In your column widths array, add width for EPS Comments if needed
//         const columnWidthsWithSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 15, 50, 20];
//         const columnWidthsWithoutSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 50, 20];
        
//         // Function to check if Specialist Copay column exists
//         const hasSpecialistCopay = headers.includes("Specialist Copay");
        
//         // Select appropriate column width array
//         const columnWidths = hasSpecialistCopay ? columnWidthsWithSpecialistCopay : columnWidthsWithoutSpecialistCopay;
        
//         // Set column widths dynamically based on actual number of columns
//         sheet.columns = headers.map((header, index) => {
//             // Use predefined width if available, otherwise use default width
//             const width = index < columnWidths.length ? columnWidths[index] : 15;
            
//             return {
//                 width,
//                 style: (header === "Notes" || header === "EPS Comments") 
//                     ? { alignment: { wrapText: true } } 
//                     : {},
//             };
//         });

//         // When formatting cells in each row, add special handling for EPS Comments
//         originalData.slice(1).forEach((row) => {
//             if (!row) return; // Skip undefined rows
            
//             const hasData = row.some(cell => cell && cell.toString().trim() !== "");
//             if (!hasData) return; // Skip empty rows
            
//             const rowData = headers.map(header => row[columnMap[header]] || "");
            
//             const newRow = sheet.addRow(rowData);
//             newRow.height = 42;

//             headers.forEach((header, colIndex) => {
//                 const cell = newRow.getCell(colIndex + 1);

//                 // Apply border to all sides
//                 cell.border = {
//                     top: { style: "thin", color: { argb: "000000" } },
//                     left: { style: "thin", color: { argb: "000000" } },
//                     bottom: { style: "thin", color: { argb: "000000" } },
//                     right: { style: "thin", color: { argb: "000000" } }
//                 };

//                 // Apply default font size
//                 cell.font = { size: 10 };

//                 if (header === "ECW #" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFF00" } // Yellow
//                     };
//                     cell.alignment = { horizontal: "center", vertical: "bottom" };
//                 }
            
//                 // Formatting for "Notes"
//                 if (header === "Notes" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFF00" } // Yellow
//                     };
//                     cell.alignment = { wrapText: true, vertical: "bottom" };
//                 }
                
//                 // Formatting for "Result"
//                 if (header === "Result" && cell.value) {
//                     const value = cell.value.toString().toLowerCase();
//                     if (value === "done") {
//                         cell.fill = {
//                             type: "pattern",
//                             pattern: "solid",
//                             fgColor: { argb: "92D050" } // Green
//                         };
//                     } else if (value === "failed") {
//                         cell.fill = {
//                             type: "pattern",
//                             pattern: "solid",
//                             fgColor: { argb: "FF0000" } // Red
//                         };
//                     }
//                     cell.alignment = { horizontal: "center", vertical: "middle" };
//                 }
//             });
//         });

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
//             message: `${results.successful.length} rows processed successfully. ${results.failed.length} rows failed.`,
//             results,
//             fileId
//         });

//     } catch (error) {
//         console.error("Error:", error);
//         res.status(500).json({ error: error.message });
//     } finally {
//         if (newPage) {
//             await new Promise(resolve => setTimeout(resolve, 1000));
//             await newPage.close();
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
import https from 'https';
import fetch from 'node-fetch';

const app = express();
const port = 5000;

app.use(cors());
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage });

// Store processed data temporarily
const processedFiles = new Map();

// Private Access Connectivity Check Function
async function isPrivateAccessConnected() {
    try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 5000);                       

        // Create an HTTPS agent that disables certificate validation
        const agent = new https.Agent({ rejectUnauthorized: false });

        const response = await fetch('https://maicllejgit6695ivkapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp', {
            method: 'HEAD',
            signal: controller.signal,
            agent
        });
        clearTimeout(timeoutId);
        return response.ok;
    } catch (error) {
        console.error("Private Access connectivity check failed:", error.message);
        return false;
    }
}

async function continueWithLoggedInSession() {
    try {
        const browser = await puppeteer.connect({
            browserURL: 'http://localhost:9222', // Connect to the existing Chrome instance
            defaultViewport: null,
            timeout: 120000, // Increased timeout
            headless: false, // Open in non-headless mode for debugging
        });

        // Get all open pages in the browser
        const pages = await browser.pages();
        
        // Find an existing page with the desired URL
        let sessionPage = pages.find(p => p.url().includes('scheduling/resourceSchedule.jsp')); 

        // If an existing session page is found, use it instead of opening a new one
        if (sessionPage) {
            console.log('Existing session found. Continuing with the current session...');
            await new Promise(resolve => setTimeout(resolve, 500)); // Wait for 500 milliseconds
            // Wait for the dropdown toggle to be visible
            console.log('Waiting for three-line menu...');
        await sessionPage.waitForSelector('#JellyBeanCountCntrl > div.fav-block > div', { visible: true, timeout: 10000 });

        console.log('Clicking three-line menu...');
        await sessionPage.click('#JellyBeanCountCntrl > div.fav-block > div');

        // Wait for the Billing option inside the menu list
        console.log('Waiting for Billing menu...');
        await sessionPage.waitForSelector('#Billing_menu', { visible: true, timeout: 10000 });

        // Scroll into view before clicking
        console.log('Scrolling to Billing menu...');
        await sessionPage.evaluate(() => {
            const element = document.querySelector('#Billing_menu');
            if (element) {
                element.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
        });

        await new Promise(resolve => setTimeout(resolve, 1000));
// Wait for smooth scrolling effect

        console.log('Clicking Billing menu...');
        await sessionPage.click('#Billing_menu');

        // Wait for Claims menu to appear
        console.log('Waiting for Claims option...');
        await sessionPage.waitForSelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)', { visible: true, timeout: 10000 });

        // Scroll to Claims menu before clicking
        console.log('Scrolling to Claims option...');
        await sessionPage.evaluate(() => {
            const element = document.querySelector('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');
            if (element) {
                element.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
        });

        await new Promise(resolve => setTimeout(resolve, 1000));


        console.log('Clicking Claims option...');
        await sessionPage.click('#Billing_menu > div > div.pad15.orange-list-lp > ul > li:nth-child(4)');

        console.log('Successfully navigated to Claims!');

            return sessionPage; // Return the existing session page
        } else {
            console.log('No existing session found. Skipping the process...');
            // Simulate an alert on the browser if no session is found
            await browser.pages().then(async (pages) => {
                const firstPage = pages[0]; // Assuming there's at least one open page
                await firstPage.evaluate(() => {
                    alert('Session not found!');
                });
            });
            return null;  // No page is returned if no session is found
        }
    } catch (error) {
        console.error("Error in continueWithLoggedInSession:", error);
        throw error;
    }
}
async function processAccount(newPage, accountNumber, provider, facilityName, billingDate, diagnosisCodes,cptCodes,admitDate) {
    try {
        // Validate required fields
        if (!provider || !facilityName || !billingDate || !accountNumber) {
            console.error('Missing required fields:', { provider, facilityName, billingDate, accountNumber });
            return false;
        }
        // Set From Date to 01/01/2023
        const fromDate = '01/01/2023';
        console.log(`Setting From Date to: ${fromDate}`);
        await newPage.waitForSelector('#fromdate', { visible: true, timeout: 5000 });
        await newPage.click('#fromdate', { clickCount: 3 });
        await newPage.keyboard.press('Backspace');
        await newPage.type('#fromdate', fromDate, { delay: 50 });

        // Set To Date to current date in MM/DD/YYYY format
        const now = new Date();
        const toDate = `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
        console.log(`Setting To Date to: ${toDate}`);
        await newPage.waitForSelector('#todate', { visible: true, timeout: 5000 });
        await newPage.click('#todate', { clickCount: 3 });
        await newPage.keyboard.press('Backspace');
        await newPage.type('#todate', toDate, { delay: 50 });

        // Select "Encounter/Claim" option from the dropdown
        console.log('Selecting Claim Status "Encounter/Claim"...');
        await newPage.waitForSelector('#claimStatusCodeId', { visible: true, timeout: 5000 });
        const optionValue = await newPage.$eval('#claimStatusCodeId > option:nth-child(2)', el => el.value);
        await newPage.select('#claimStatusCodeId', optionValue);

        // Click the Claim Lookup button
        console.log('Clicking the Claim Lookup button...');
        await newPage.waitForSelector('#claimLookupBtn10', { visible: true, timeout: 5000 });
        await newPage.click('#claimLookupBtn10');
        await new Promise(resolve => setTimeout(resolve, 3000));

        // Click Create New Claim
        console.log('Selecting Create New Claim option...');
        await newPage.waitForSelector('#claimLookupBtn11', { visible: true, timeout: 5000 });
        await newPage.click('#claimLookupBtn11');
        await new Promise(resolve => setTimeout(resolve, 3000));

        console.log(`Filling in Resource with Provider value: ${provider}`);

// Extract the second part of the provider name (last name)
const providerLastName = provider.split(' ').slice(-1).join(' ');
console.log(`Using provider's last name: ${providerLastName}`);

// Fill Provider Input - Using only the last name (second part of the provider name)
await newPage.waitForSelector("input[id^='provider-lookupIpt1']", { visible: true, timeout: 5000 });
await newPage.click("input[id^='provider-lookupIpt1']", { clickCount: 3 });
await newPage.keyboard.press('Backspace');
await newPage.type("input[id^='provider-lookupIpt1']", providerLastName, { delay: 30 });

await newPage.waitForSelector("#provider-lookupUl1 li", { visible: true, timeout: 5000 });

await newPage.evaluate((providerLastName) => {
    const options = document.querySelectorAll("#provider-lookupUl1 li");
    for (let opt of options) {
        if (opt.innerText.trim().toLowerCase() === providerLastName.trim().toLowerCase()) {
            opt.click();
            break;
        }
    }
}, providerLastName);

// Clear and Fill Resource Input
await newPage.waitForSelector("#resource-lookupIpt1", { visible: true, timeout: 8000 });

// First, clear the field completely
await newPage.click("#resource-lookupIpt1", { clickCount: 3 });
await newPage.keyboard.press('Backspace');

// Use evaluate to ensure the field is completely cleared
await newPage.evaluate(() => {
    const input = document.querySelector("#resource-lookupIpt1");
    input.value = "";
    // Trigger input event to ensure any listeners know the input changed
    input.dispatchEvent(new Event('input', { bubbles: true }));
});

// Type the provider's last name
await newPage.type("#resource-lookupIpt1", providerLastName, { delay: 50 });

// Wait for the dropdown to appear and contain options
await newPage.waitForSelector("#resource-lookupUl1 li.list-item-filter.ng-scope", { visible: true, timeout: 5000 });

// Get all dropdown items
const dropdownItems = await newPage.$$("#resource-lookupUl1 li.list-item-filter.ng-scope");

let optionFound = false;

// Loop through the dropdown items and click the matching option
for (let item of dropdownItems) {
    const itemText = await newPage.evaluate(el => el.innerText.trim(), item);
    if (itemText.toLowerCase().includes(providerLastName.toLowerCase())) {
        await item.click();
        optionFound = true;
        break;
    }
}

if (!optionFound) {
    console.error('No matching resource option found');
}

// Wait for the selection to be applied
await new Promise(resolve => setTimeout(resolve, 1000));

// Fill in Facility Name
const inputSelector = "input.createClaimFacilityCls";

// Wait for the input to be visible
await newPage.waitForSelector(inputSelector, { visible: true, timeout: 10000 });

// Clear the field first in case there's existing text
await newPage.evaluate((selector) => {
    document.querySelector(selector).value = '';
}, inputSelector);

// Type the facility name into the input field
await newPage.type(inputSelector, facilityName, { delay: 50 });

// Wait briefly to allow the application to process the input
await new Promise(resolve => setTimeout(resolve, 1000));

// Try to find the dropdown using multiple possible selectors
try {
    // First try the specific ID
    await newPage.waitForSelector("#facility-lookupUl1", { visible: true, timeout: 2000 });
} catch (e) {
    // If that fails, try a more general selector for dropdowns
    try {
        await newPage.waitForSelector("ul.dropdown-menu", { visible: true, timeout: 2000 });
    } catch (err) {
        // If no visible dropdown, press arrow down to try to trigger it
        await newPage.keyboard.press('ArrowDown');
        await new Promise(resolve => setTimeout(resolve, 500));
    }
}

// Try to select an option using multiple approaches
try {
    // First try to find and click an option with JavaScript
    const found = await newPage.evaluate((facilityName) => {
        // Try multiple selectors for dropdown options
        const selectors = [
            "#facility-lookupUl1 li", 
            "ul.dropdown-menu li",
            ".dropdown-menu li",
            "li[role='option']"
        ];
        
        let options = [];
        // Try each selector until we find options
        for (const selector of selectors) {
            options = document.querySelectorAll(selector);
            if (options.length > 0) break;
        }
        
        // If we found options, try to find a match and click it
        if (options.length > 0) {
            for (let opt of options) {
                if (opt.textContent.trim().toLowerCase().includes(facilityName.trim().toLowerCase())) {
                    opt.click();
                    return true;
                }
            }
        }
        return false;
    }, facilityName);
    
    // If JavaScript selection didn't work, use keyboard navigation
    if (!found) {
        // Press arrow down to select first option
        await newPage.keyboard.press('ArrowDown');
        await new Promise(resolve => setTimeout(resolve, 300));
        // Press Enter to confirm selection
        await newPage.keyboard.press('Enter');
    }
} catch (e) {
    // If all else fails, try keyboard navigation
    await newPage.keyboard.press('ArrowDown');
    await new Promise(resolve => setTimeout(resolve, 300));
    await newPage.keyboard.press('Enter');
}

// Wait to ensure the selection is processed
await new Promise(resolve => setTimeout(resolve, 1000));
        // Set Service Date (from Billing Date)
        console.log(`Setting Billing Date: ${billingDate}`);
        await newPage.click('#claimserviceDate', { clickCount: 3 });
        await newPage.focus('#claimserviceDate');
        await newPage.keyboard.down('Control');
        await newPage.keyboard.press('A');
        await newPage.keyboard.up('Control');
        await newPage.keyboard.press('Backspace');
        
        // Double check it's cleared
        await newPage.evaluate(() => {
          document.querySelector('#claimserviceDate').value = '';
        });
        
        await newPage.type('#claimserviceDate', billingDate, { delay: 50 });
        await newPage.keyboard.press('Tab');
        const ainputSelector = "input.patientFilterForClaim";

        // Step 1: Wait for input and type
        await newPage.waitForSelector(ainputSelector, { visible: true, timeout: 10000 });
        await newPage.evaluate(selector => {
            const input = document.querySelector(selector);
            if (input) {
                input.value = '';
                input.setAttribute('autocomplete', 'off');
            }
        }, ainputSelector);
        
        await newPage.type(ainputSelector, accountNumber, { delay: 50 });
        
        // Step 2: Wait for dropdown to appear
        await new Promise(resolve => setTimeout(resolve, 1500));
        const dropdownSelectors = [
            "#patient-lookupUl1 li",
            "ul.dropdown-menu li",
            ".dropdown-menu li",
            "li[role='option']"
        ];
        
        let clicked = false;
        
        for (const sel of dropdownSelectors) {
            const items = await newPage.$$(sel);
        
            for (const item of items) {
                const text = await newPage.evaluate(el => el.textContent.trim(), item);
                if (text.toLowerCase().includes(accountNumber.toLowerCase())) {
                    const box = await item.boundingBox();
                    if (box) {
                        await newPage.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
                        await newPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { delay: 50 });
                        clicked = true;
                        break;
                    }
                }
            }
        
            if (clicked) break;
        }
        
        // Step 3: Fallback if not clicked
        if (!clicked) {
            console.warn('Dropdown item not found with bounding box. Using keyboard fallback.');
            await newPage.keyboard.press('ArrowDown');
            await new Promise(resolve => setTimeout(resolve, 300));
            await newPage.keyboard.press('Enter');
        }
        
        // Step 4: Trigger AngularJS model updates
        await newPage.evaluate((selector) => {
            const input = document.querySelector(selector);
            if (input) {
                input.value = input.value.trim();
                const angularInput = angular.element(input);
                angularInput.triggerHandler('input');
                angularInput.triggerHandler('change');
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                input.focus();
                setTimeout(() => input.blur(), 100);
            }
        }, ainputSelector);
        
        
        // Optionally, wait a bit to let things settle
        await new Promise(resolve => setTimeout(resolve, 500));
        
        // Final Step - Click OK Button
        const buttonExists = await newPage.$('#createClaimBtn');
        if (buttonExists) {
            console.log('Create Claim button found. Clicking...');
            await newPage.click('#createClaimBtn');
        } else {
            console.error('Create Claim button not found');
        }
        await new Promise(resolve => setTimeout(resolve, 3000)); // wait for 3s (adjust as needed)

        for (let i = 0; i < diagnosisCodes.length; i++) {
            const code = diagnosisCodes[i];
            
            // Find the current active input field
            const dinputSelector = 'input[id^="txtnewIcd"]';
            await newPage.waitForSelector(dinputSelector, { visible: true, timeout: 8000 });
            
            // Find the most recently added/visible input field
            const inputs = await newPage.$$(dinputSelector);
            let activeInput = null;
            
            for (const input of inputs) {
              const isVisible = await input.evaluate(el => {
                const style = window.getComputedStyle(el);
                return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
              });
              
              if (isVisible) {
                // Check if input is empty
                const value = await input.evaluate(el => el.value);
                if (!value) {
                  activeInput = input;
                  break;
                }
              }
            }
            
            if (!activeInput) {
              console.warn(`No available input field found for code: ${code}`);
              continue;
            }
            
            // Enter the code
            await activeInput.click();
            await activeInput.type(code, { delay: 100 });
            await newPage.keyboard.press('Enter');
            
            // Wait for processing
            await new Promise(resolve => setTimeout(resolve, 1500));
            
            // Check and handle error modal if it appears
            try {
              const errorOkSelector = 'body > div.bootbox.modal.fade.bluetheme.medium-width.in > div > div > div.modal-footer > button';
              const hasErrorModal = await newPage.evaluate((selector) => {
                const element = document.querySelector(selector);
                return element !== null && element.offsetParent !== null;
              }, errorOkSelector);
              
              if (hasErrorModal) {
                console.log(`Error modal detected for code ${code}. Clicking OK...`);
                await newPage.click(errorOkSelector);
                await new Promise(resolve => setTimeout(resolve, 1000));
              }
            } catch (error) {
              console.log(`Error checking for modal: ${error.message}`);
            }
            
            // Additional wait to ensure system is ready for next input
            await new Promise(resolve => setTimeout(resolve, 1000));
          }
          for (let i = 0; i < cptCodes.length; i++) {
            const code = cptCodes[i];
            const baseCode = code.includes('-') ? code.split('-')[0].trim() : code.trim();
          
            try {
              console.log(`🔄 Processing CPT code: ${code}`);
              console.log(`🧩 Base CPT: ${baseCode}`);
          
              // Step 1: Click CPT cell
              const cptCellSelector = 'td[ng-click*="enterCPT"]';
              await newPage.waitForSelector(cptCellSelector, { visible: true, timeout: 8000 });
              await newPage.click(cptCellSelector);
              await new Promise(resolve => setTimeout(resolve, 600));
          
              // Step 2: Enter CPT code
              const cptInputSelector = '#billingClaimIpt34';
              await newPage.waitForSelector(cptInputSelector, { visible: true, timeout: 5000 });
          
              const cptInput = await newPage.$(cptInputSelector);
              await cptInput.focus();
              await newPage.evaluate(el => el.scrollIntoView(), cptInput);
              await cptInput.click({ clickCount: 3 });
              await cptInput.press('Backspace');
              await new Promise(resolve => setTimeout(resolve, 200));
              await cptInput.type(baseCode, { delay: 100 });
              await newPage.keyboard.press('Enter');
              console.log(`✅ Entered CPT code: ${baseCode}`);
          
              console.log(`✅ Done processing CPT: ${code}`);
            } catch (error) {
              console.error(`❌ Failed to process CPT code ${code}:`, error);
              await newPage.screenshot({ path: `cpt_error_${i}.png` });
            }
          
            await new Promise(resolve => setTimeout(resolve, 1000));
          }
          
          
        const claimNoSelector = '#billingClaimTbl2 > tbody > tr:nth-child(1) > td:nth-child(2) > span';

await newPage.waitForSelector(claimNoSelector, { visible: true, timeout: 5000 });

const claimNumber = await newPage.$eval(claimNoSelector, el => el.textContent.trim());

console.log(`🧾 Claim Number: ${claimNumber}`);

        // === After CPT Code Entry, Check for 'Pending' Status ===
try {
    // Find the select dropdown whose ID starts with 'claimStatusSel'
    const statusDropdownHandle = await newPage.$('select[id^="claimStatusSel"]');
    if (statusDropdownHandle) {
      const selectedText = await newPage.evaluate(el => {
        const selectedOption = el.options[el.selectedIndex];
        return selectedOption ? selectedOption.textContent.trim() : '';
      }, statusDropdownHandle);
  
      console.log(`Claim Status: ${selectedText}`);
      try {
        // Step 1: Check POS value
        const posSelector = '#billingClaimIpt24';
        await newPage.waitForSelector(posSelector, { visible: true, timeout: 8000 });
      
        const posValue = await newPage.$eval(posSelector, el => el.value.trim());
        console.log(`POS Value: ${posValue}`);
      
        if (["21", "31"].includes(posValue)) {
          console.log(`POS is ${posValue} — proceeding to click data button.`);
      
          // Step 2: Click dynamic data button (verify exact selector or fallback to partial)
          const dataBtnSelector = 'button[id^="claimDataBtn"]';
          const dataBtnHandle = await newPage.$(dataBtnSelector);
          if (dataBtnHandle) {
            const btnId = await newPage.evaluate(el => el.id, dataBtnHandle);
            console.log(`Found data button with ID: ${btnId}`);
            await dataBtnHandle.click();
            console.log(`Clicked data button`);
            await new Promise(resolve => setTimeout(resolve, 1000));
        } else {
            console.warn(`Data button with selector ${dataBtnSelector} not found`);
            return;
          }
      
          // Step 3: Extract Admit Date from the table
        // Step 4: Type date into #dtHCFAFrom16From
const admitDateInputSelector = '#dtHCFAFrom16From';
await newPage.waitForSelector(admitDateInputSelector, { visible: true, timeout: 5000 });

// ✅ Extract only the date part (e.g. from '3/19/2025 (17d)')
const admitDateOnly = admitDate.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)?.[0] || "";
console.log(`Using Admit Date: ${admitDateOnly}`);

// Clear existing input
await newPage.evaluate(selector => {
  const input = document.querySelector(selector);
  if (input) input.value = '';
}, admitDateInputSelector);

// Type the extracted date
await newPage.type(admitDateInputSelector, admitDateOnly);
console.log(`Entered admit date into input field`);

          // Step 5: Click OK button
          const okBtnSelector = '#claimDataBtn13';
          await newPage.waitForSelector(okBtnSelector, { visible: true, timeout: 5000 });
          await newPage.click(okBtnSelector);
          console.log(`Clicked OK button to finish`);
        } else {
          console.log(`POS is ${posValue} — skipping admit date flow.`);
        }
      } catch (error) {
        console.error(`❌ Error during POS/admit-date flow: ${error.message}`);
      }
      
      // If status is Pending, click the OK button
      if (selectedText.toLowerCase() === "pending") {
        const okButtonHandle = await newPage.$('button[id^="claimScreenOkBtn"]');
        if (okButtonHandle) {
          console.log(`Pending detected — clicking OK...`);
          await okButtonHandle.click();
          await new Promise(resolve => setTimeout(resolve, 1000));
        } else {
          console.warn(`OK button not found`);
        }
      }
    } else {
      console.warn(`Claim status dropdown not found`);
    }
  } catch (error) {
    console.error(`Error checking claim status: ${error.message}`);
  }

  
        return true; // Return success
    } catch (error) {
        console.error('Error processing account:', error);
        return false; // Return failure
    }
}

// Endpoint to process data and return status
app.post("/process", upload.single("file"), async (req, res) => {
    if (!req.file || !req.body.originalPath) {
        return res.status(400).json({ error: "No file or original path provided" });
    }
    
    const privateAccessConnected = await isPrivateAccessConnected();
    if (!privateAccessConnected) {
        return res.status(500).json({
            error: "Private Access is not connected. Please connect to the Private Access"
        });
    }
    
    let newPage = null;
    
    try {
        const workbookXLSX = xlsx.read(req.file.buffer, { 
            type: "buffer", 
            cellDates: true,
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
        
        // Check if "ECW #" exists in headers
        console.log(headers);
        if (!headers.includes("ECW #")) {
            throw new Error("'ECW #' Column is not defined.");
        }
        
        // Map headers to their index positions
        headers.forEach((header, index) => {
            columnMap[header] = index;
        });
        
        // Check if required columns exist
        const requiredColumns = ["ECW #", "Provider", "Facility Name", "Billing Date"];
        for (const column of requiredColumns) {
            if (!headers.includes(column)) {
                throw new Error(`Required column '${column}' is missing from the spreadsheet.`);
            }
        }
        
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
            if (!row) continue; // Skip undefined rows
            
            const accountNumberRaw = row[columnMap["ECW #"]] || "";
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
        
        // Check if all remaining rows are missing Notes
        if (newRowsToProcess > 0 && newRowsToProcess === missingCallNotes) {
            return res.json({
                success: false,
                message: "Notes not found in Excel.",
                results: { successful: [], failed: [] }
            });
        }
        
        // Check if all rows are already updated
        if (newRowsToProcess === 0) {
            return res.json({
                success: true,
                message: "All rows were already updated.",
                results: { successful: [], failed: [] }
            });
        }
        
        // Get the active page for automation
        newPage = await continueWithLoggedInSession();
        
        const results = { successful: [], failed: [] };
        
        for (let i = 1; i < originalData.length; i++) {
            const row = originalData[i];
            if (!row) continue;
        
            const accountNumberRaw = row[columnMap["ECW #"]] || "";
            const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            if (!accountNumberRaw || status === "done") continue;
        
            const provider = row[columnMap["Provider"]] || "";
            const facilityName = row[columnMap["Facility Name"]] || "";
            const billingDate = row[columnMap["Billing Date"]] || "";
            const diagnosisText = row[columnMap["Diagnoses"]] || "";
            const chargesText = row[columnMap["Charges"]] || ""; // CPT codes
            const admitDate = row[columnMap["Admit Date"]] || ""; // ✅ NEW LINE


            const accountNumber = accountNumberRaw.split("-")[0].trim();
            console.log(`Processing account: ${accountNumber}`);
        
            // Validate data before processing
            if (!provider) console.log('Provider value is missing or undefined');
            if (!facilityName) console.log('Facility Name is missing or undefined');
            if (!billingDate) console.log('Billing Date is missing or undefined');
            if (!admitDate) console.log('Admit Date is missing or undefined'); // ✅ Optional

            // Extract diagnosis codes from the text inside parentheses
            const diagnosisCodes = [];
            const regex = /\(([^)]+)\)/g;
            let match;
            
            console.log(`Diagnosis Text for ${accountNumber}:`, diagnosisText);
            
            while ((match = regex.exec(diagnosisText)) !== null) {
                const code = match[1].trim();
                
                // Only include if it matches ICD-10 format (starts with a letter, followed by numbers)
                if (/^[A-Z]\d/.test(code)) {
                    diagnosisCodes.push(code);
                }
            }
            
            console.log(`Diagnosis Codes for ${accountNumber}:`, diagnosisCodes);
            diagnosisCodes.forEach((code, index) => {
                console.log(`  Code ${index + 1}: ${code}`);
            });
            
            console.log(`Charges Text for ${accountNumber}:`, chargesText);

            const cptCodes = [];
            const cptRegex = /\b\d{5}(?:-[A-Za-z0-9]{2,4})?\b/g;
            let cptMatch;
            
            while ((cptMatch = cptRegex.exec(chargesText)) !== null) {
                const fullCode = cptMatch[0].trim();
                cptCodes.push(fullCode);
            
                const [baseCode, modifier] = fullCode.split('-');
                console.log(`🧩 Base CPT: ${baseCode}, Modifier: ${modifier || 'None'}`);
            }
            
            console.log(`Extracted CPT Codes for ${accountNumber}:`, cptCodes);
            
            
            
            // === Call processAccount (Don't send CPT length)
            const success = await processAccount(
                newPage,
                accountNumber,
                provider,
                facilityName,
                billingDate,
                diagnosisCodes,
                cptCodes, // ✅ only the array
                admitDate // ✅ Pass Admit Date here

            );
        
            if (success) {
                results.successful.push(accountNumber);
                originalData[i][columnMap["Result"]] = "Done";
            } else {
                results.failed.push(accountNumber);
                originalData[i][columnMap["Result"]] = "Failed";
            }
        }
        
        // Convert data back to an ExcelJS workbook
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("Sheet1");

        // Write headers with orange background
        const headerRow = sheet.addRow(headers);
        headerRow.height = 42;
        headerRow.eachCell((cell) => {
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFA500" }
            };
            cell.font = { bold: true, color: { argb: "000000" }, size: 10 };
            cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
            cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" }
            };
        });

        // In your column widths array, add width for EPS Comments if needed
        const columnWidthsWithSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 15, 50, 20];
        const columnWidthsWithoutSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 50, 20];
        
        // Function to check if Specialist Copay column exists
        const hasSpecialistCopay = headers.includes("Specialist Copay");
        
        // Select appropriate column width array
        const columnWidths = hasSpecialistCopay ? columnWidthsWithSpecialistCopay : columnWidthsWithoutSpecialistCopay;
        
        // Set column widths dynamically based on actual number of columns
        sheet.columns = headers.map((header, index) => {
            // Use predefined width if available, otherwise use default width
            const width = index < columnWidths.length ? columnWidths[index] : 15;
            
            return {
                width,
                style: (header === "Notes" || header === "EPS Comments") 
                    ? { alignment: { wrapText: true } } 
                    : {},
            };
        });

        // When formatting cells in each row, add special handling for EPS Comments
        originalData.slice(1).forEach((row) => {
            if (!row) return; // Skip undefined rows
            
            const hasData = row.some(cell => cell && cell.toString().trim() !== "");
            if (!hasData) return; // Skip empty rows
            
            const rowData = headers.map(header => row[columnMap[header]] || "");
            
            const newRow = sheet.addRow(rowData);
            newRow.height = 42;

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

                if (header === "ECW #" && cell.value) {
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
                
                // Formatting for "Result"
                if (header === "Result" && cell.value) {
                    const value = cell.value.toString().toLowerCase();
                    if (value === "done") {
                        cell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: { argb: "92D050" } // Green
                        };
                    } else if (value === "failed") {
                        cell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: { argb: "FF0000" } // Red
                        };
                    }
                    cell.alignment = { horizontal: "center", vertical: "middle" };
                }
            });
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
            message: `${results.successful.length} rows processed successfully. ${results.failed.length} rows failed.`,
            results,
            fileId
        });

    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({ error: error.message });
    } finally {
        if (newPage) {
            await new Promise(resolve => setTimeout(resolve, 1000));
            await newPage.close();
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