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

async function isVPNConnected() {
    try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 5000);

        // Create an HTTPS agent that disables certificate validation
        const agent = new https.Agent({ rejectUnauthorized: false });

        const response = await fetch('https://10.1.45.10/', {
            method: 'HEAD',
            signal: controller.signal,
            agent
        });
        clearTimeout(timeoutId);
        return response.ok;
    } catch (error) {
        console.error("VPN connectivity check failed:", error.message);
        return false;
    }
}

async function loginToSystem(page) {
    try {
        console.log('Navigating to login page...');
        await page.goto('https://10.1.45.10/', { waitUntil: 'networkidle0', timeout: 60000 });
    } catch (err) {
        // This error likely means the server is unreachable (e.g., VPN is off)
        throw new Error("Unable to reach the server. Please ensure you are connected to the VPN.");
    }

    console.log('Looking for login frame...');
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    const loginFrame = page.frames().find(f => f.name() === 'MsysLogIn');
    if (!loginFrame) {
        throw new Error('Login frame not found');
    }

    console.log('Filling login form...');
    await loginFrame.waitForSelector('input#UsrNm', { timeout: 5000 });
    await loginFrame.type('input#UsrNm', 'marun', { delay: 50 });
    await loginFrame.waitForSelector('input#PWCode', { timeout: 5000 });
    await loginFrame.type('input#PWCode', 'Cleveland123!', { delay: 50 });

    console.log('Submitting login form...');
    await loginFrame.click('input[type="submit"]');
    await new Promise(resolve => setTimeout(resolve, 5000));

    // Handle dropdown selection
    let targetFrame;
    for (const frame of page.frames()) {
        try {
            const hasDropdown = await frame.evaluate(() => {
                return !!document.querySelector('select[name="EntCode"]');
            });
            if (hasDropdown) {
                targetFrame = frame;
                break;
            }
        } catch (err) {
            console.log(`Error checking frame: ${err.message}`);
        }
    }

    if (!targetFrame) {
        throw new Error('Dropdown not found in any frame.');
    }

    await targetFrame.select('select[name="EntCode"]', '001001005000');
    await targetFrame.click('body > form > input[type=submit]:nth-child(2)');
    await new Promise(resolve => setTimeout(resolve, 5000));

    // Navigate to Patient Accounting
    const menuFrame = page.frames().find(f => f.name() === 'MsysMenu');
    await menuFrame.click('body > a:nth-child(23) > button');
    await new Promise(resolve => setTimeout(resolve, 5000));

    // Click Patient Account Detail
    const arMenuFrame = page.frames().find(f => f.name() === 'ARMenu');
    await arMenuFrame.click('body > a:nth-child(8) > button');
    await new Promise(resolve => setTimeout(resolve, 5000));
}

async function processAccount(page, accountNumber, notes) {
    try {
        const arPatSearchFrame = page.frames().find(f => f.name() === 'ARpatSearchBot');
        if (!arPatSearchFrame) {
            throw new Error("Frame 'ARpatSearchBot' not found");
        }

        // Enter Account Number and Search
        await arPatSearchFrame.waitForSelector('input[name="AcctNo"]', { timeout: 10000 });
        await arPatSearchFrame.evaluate((selector, accNum) => {
            const input = document.querySelector(selector);
            input.value = '';
            input.value = accNum;
            input.dispatchEvent(new Event('input', { bubbles: true }));
        }, 'input[name="AcctNo"]', accountNumber.toString());

        await arPatSearchFrame.click('input[name="ActSearch"]');
        await new Promise(resolve => setTimeout(resolve, 5000));

        // Look for the "Account Note" Button
        const buttons = await page.$$('input[type="submit"]');
        let accountNoteButtonFound = false;
        
        for (const button of buttons) {
            const text = await page.evaluate(el => el.value, button);
            if (text.includes("Account Note")) {
                await button.click();
                await new Promise(resolve => setTimeout(resolve, 5000));
                accountNoteButtonFound = true;
                break;
            }
        }
        
        // If Account Note button wasn't found, click cancel and return
        if (!accountNoteButtonFound) {
            console.log(`Account Note button not found for account ${accountNumber}. Clicking cancel.`);
            
            // Find and click the cancel button
            const cancelButtons = await page.$$('input[type="submit"], button[type="button"]');
            for (const button of cancelButtons) {
                const text = await page.evaluate(el => el.value || el.textContent, button);
                if (text.includes("Cancel") || text.includes("Close") || text.includes("Back")) {
                    await button.click();
                    await new Promise(resolve => setTimeout(resolve, 3000));
                    return false; // Return false to indicate we need to move to next account
                }
            }
            
            // If no cancel button found, just return
            return false;
        }

        // Wait for Note Frame
        let noteFrame = page.frames().find(f => f.name() === 'actnoteBot');
        if (!noteFrame) {
            throw new Error("Frame 'actnoteBot' not found");
        }

        if (notes) {
            await noteFrame.waitForSelector('body > form > p:nth-child(7) > textarea');
            await noteFrame.evaluate((selector, text) => {
                const textarea = document.querySelector(selector);
                textarea.value = text;
                textarea.dispatchEvent(new Event('input', { bubbles: true }));
            }, 'body > form > p:nth-child(7) > textarea', notes);
        }

        // Click "Add Note" Button
        await noteFrame.waitForSelector('body > form > p:nth-child(7) > input[type=submit]:nth-child(4)');
        await noteFrame.click('body > form > p:nth-child(7) > input[type=submit]:nth-child(4)');
        await new Promise(resolve => setTimeout(resolve, 5000));

        // Re-select the Note Frame (in case it reloads)
        await new Promise(resolve => setTimeout(resolve, 2000));
        noteFrame = page.frames().find(f => f.name() === 'actnoteBot');
        if (!noteFrame) {
            throw new Error("Frame 'actnoteBot' detached or not found after adding note.");
        }

        // Click "Done" Button
        await noteFrame.waitForSelector('body > form > p:nth-child(7) > input[type=submit]:nth-child(6)');
        await noteFrame.click('body > form > p:nth-child(7) > input[type=submit]:nth-child(6)');
        await new Promise(resolve => setTimeout(resolve, 5000));

        // Close the Note Window
        const cancelButtons = await page.$$('input[type="submit"]');
        if (cancelButtons.length >= 2) {
            await cancelButtons[1].click();
            await new Promise(resolve => setTimeout(resolve, 3000));
        }

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

    // First, check if VPN is connected. If not, return an error.
    const vpnConnected = await isVPNConnected();
    if (!vpnConnected) {
        return res.status(500).json({
            error: "VPN is not connected. Please connect to the VPN"
        });
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
        
        // Check if "Pt Account #" exists in headers
        if (!headers.includes("Pt Account #")) {
            throw new Error("'Pt Account #' Column is not defined."); // Display an alert if needed
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
        
        let missingCallNotes = 0;  // Track rows missing Call Notes

        for (let i = 1; i < originalData.length; i++) {
            const row = originalData[i];
            const accountNumberRaw = row[columnMap["Pt Account #"]] || "";
            const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
            if (!accountNumberRaw || status === "done") continue;  // Skip processed rows
            
            const accountNumber = accountNumberRaw.split("-")[0].trim();
            processedAccounts.add(accountNumber);
            newRowsToProcess++;
        
            // Check if Call Notes is missing
            const callNotes = row[columnMap["Call Notes"]] || "";
            if (!callNotes.trim()) {
                missingCallNotes++;
            }
        }
        
        // ✅ Check if all remaining rows are missing Call Notes
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
            const accountNumberRaw = row[columnMap["Pt Account #"]] || "";
            const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
            // Skip row if there's no account number or it's already processed.
            if (!accountNumberRaw || status === "done") continue;
            
            // Skip row if the "Call Notes" field is empty.
            const callNotes = row[columnMap["Call Notes"]] || "";
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
    style: (index === columnMap["Call Notes"] || index === columnMap["EPS Comments"]) 
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

            if (header === "Pt Account #" && cell.value) {
                cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: "FFFF00" } // Yellow
                };
                cell.alignment = { horizontal: "center", vertical: "bottom" };
            }
        
            // Formatting for "Call Notes"
            if (header === "Call Notes" && cell.value) {
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
