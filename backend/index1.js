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
            timeout: 60000, // Increased timeout
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
            console.log('Waiting for dropdown toggle to be visible...');
            await sessionPage.waitForSelector('#jellybean-panelLink66', { visible: true, timeout: 10000 });

            // Click the dropdown toggle to open it
            console.log('Clicking dropdown toggle...');
            await sessionPage.click('#jellybean-panelLink66');

            // Add a small delay using Promise-based delay
            await new Promise(resolve => setTimeout(resolve, 500)); // Wait for 500 milliseconds

            // Wait for the Patient Lookup option to be visible within the dropdown
            console.log('Waiting for Patient Lookup option...');
            await sessionPage.waitForSelector('#jellybean-panelLink67', { visible: true, timeout: 20000 });

            // Click the Patient Lookup option
            console.log('Clicking Patient Lookup option...');
            await sessionPage.evaluate(() => {
                const element = document.querySelector('#jellybean-panelLink67');
                if (element) {
                    element.click();
                } else {
                    console.error('Could not find the Patient Lookup element');
                }
            });

            console.log('Patient Lookup option selected!');
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

let isMRNSelected = false; // Flag to track if Acct No (MRN) selection has already been made

async function processAccount(newPage, accountNumber, callNotes, copayAmount) {
    try {
        if (!isMRNSelected) {
            console.log('Waiting for the dropdown to appear...');
            await newPage.waitForSelector('#patient-lookup-screen-detview > div > div:nth-child(1) > div:nth-child(1) > div > div > div > div.col-sm-5.nopadding.nameselect > div > div > span.ng-binding', { visible: true });

            console.log('Clicking the dropdown to open options...');
            await newPage.click('#patient-lookup-screen-detview > div > div:nth-child(1) > div:nth-child(1) > div > div > div > div.col-sm-5.nopadding.nameselect > div > div > span.ng-binding');

            console.log('Waiting for the dropdown options to load...');
            await new Promise(resolve => setTimeout(resolve, 1000));

            console.log('Selecting "Acct No (MRN)" from the dropdown...');
            await newPage.evaluate(() => {
                let options = document.querySelectorAll('span.ng-binding');
                for (const option of options) {
                    if (option.innerText.trim() === 'Acct No (MRN)') {
                        option.click();
                        break;
                    }
                }
            });

            isMRNSelected = true; // Mark as selected so this step is skipped in future calls
        }

        console.log('Waiting for the #searchText field...');
        await newPage.waitForSelector('#searchText', { visible: true });
        console.log(`Entering account number: ${accountNumber} in #searchText field`);
        
        // Clear the field before typing
        await newPage.evaluate(() => {
            document.querySelector('#searchText').value = '';
        });
        
        await newPage.type('#searchText', accountNumber, { delay: 50 });
        
        await new Promise(resolve => setTimeout(resolve, 3000));

        console.log('Clicking the Patient Demographics button...');
        await newPage.waitForSelector('#patientSearchBtn12', { visible: true });
        await newPage.click('#patientSearchBtn12');
        await new Promise(resolve => setTimeout(resolve, 10000));
         // Check for "Please select a patient and try again!" modal
        try {
            console.log('Checking if patient not found modal appears...');
            await newPage.waitForSelector('.bootbox.modal .modal-body', { 
                visible: true, 
                timeout: 8000 
            });
            
            // Check if the modal contains the "Please select a patient" message
            const modalText = await newPage.evaluate(() => {
                const modalBody = document.querySelector('.bootbox.modal .modal-body');
                return modalBody ? modalBody.innerText.trim() : '';
            });
            
            if (modalText.includes('Please select a patient and try again')) {
                console.log(`Patient not found for account ${accountNumber}. Clicking OK and moving to next account...`);
                
                // Click OK button on the modal
                await newPage.waitForSelector('.bootbox.modal .modal-footer button', { visible: true, timeout: 5000 });
                await newPage.click('.bootbox.modal .modal-footer button');
                
                // Wait for modal to close
                await new Promise(resolve => setTimeout(resolve, 1000));
                
                console.log(`Skipping account ${accountNumber} - patient not found`);
                return false
            }
        } catch (modalError) {
            console.log('No patient not found modal detected, proceeding normally...');
        }
        
        console.log('Waiting for the Notes field...');
        await newPage.waitForSelector('#patientdemographics-AddInformationIpt32', { visible: true });

        // Read existing text from the Notes field
        const existingNotes = await newPage.evaluate(() => {
            const notesField = document.querySelector('#patientdemographics-AddInformationIpt32');
            if (!notesField) {
                console.log('Notes field not found!');
                return ''; // Return empty string instead of throwing an error
            }

            console.log('Notes field detected:', notesField);

            // Check if the field is an input or textarea
            if (notesField.tagName.toLowerCase() === 'input' || notesField.tagName.toLowerCase() === 'textarea') {
                return notesField.value || '';
            }

            return notesField.innerText || notesField.textContent || '';
        });

        console.log(`Existing Notes: ${existingNotes}`);

        // Prepend new notes to existing notes
        const updatedNotes = existingNotes ? `${callNotes}\n${existingNotes}`.trim() : callNotes;

        console.log(`Entering updated call notes: ${updatedNotes}`);

        // Set the new value in the notes field
        await newPage.evaluate((selector, text) => {
            let notesField = document.querySelector(selector);
            if (notesField) {
                if (notesField.tagName.toLowerCase() === 'input' || notesField.tagName.toLowerCase() === 'textarea') {
                    notesField.value = text;
                } else {
                    notesField.innerText = text;
                }
                notesField.dispatchEvent(new Event('input', { bubbles: true }));
            }
        }, '#patientdemographics-AddInformationIpt32', updatedNotes);

// Process Copay BEFORE clicking Save button
if (copayAmount) {
    let trimmedCopay = copayAmount.trim().replace(/[$,-]/g, '');
    
    // Check for invalid values including N/A
    if (!trimmedCopay || trimmedCopay === '0.00' || copayAmount.trim().toUpperCase() === 'N/A') {
        console.log('Skipping copay update - invalid copay value:', copayAmount);
    } else {
        console.log(`Processing copay amount: ${trimmedCopay} for account: ${accountNumber}`);

        const insuranceExists = await newPage.evaluate(() => {
            return !!document.querySelector('#patient-demographicsTbl3 tbody tr');
        });

        if (insuranceExists) {
            try {
                console.log('Waiting for the first insurance row...');
                await newPage.waitForSelector('#patient-demographicsTbl3 > tbody > tr:nth-child(1)', {
                    visible: true,
                    timeout: 15000,
                });

                const insuranceRow = await newPage.$('#patient-demographicsTbl3 > tbody > tr:nth-child(1)');

                if (insuranceRow) {
                    console.log('First insurance row found, clicking on it...');
                    await insuranceRow.click();
                    await new Promise(resolve => setTimeout(resolve, 500));

                    await newPage.evaluate((row) => {
                        const dblClickEvent = new MouseEvent('dblclick', { bubbles: true, cancelable: true, view: window });
                        row.dispatchEvent(dblClickEvent);
                    }, insuranceRow);

                    console.log('Double-click action performed. Waiting for copay input...');
                    await newPage.waitForSelector('#addUpdateInsuranceIpt20', { visible: true, timeout: 15000 });

                    // Function to enter copay value
                    async function enterCopayValue(value) {
                        await newPage.click('#addUpdateInsuranceIpt20');
                        await new Promise(resolve => setTimeout(resolve, 200));
                        
                        // Select all text and delete
                        await newPage.keyboard.down('Control');  
                        await newPage.keyboard.press('KeyA'); 
                        await newPage.keyboard.up('Control');
                        await newPage.keyboard.press('Backspace'); 
                        await new Promise(resolve => setTimeout(resolve, 300));

                        await newPage.type('#addUpdateInsuranceIpt20', value, { delay: 50 });
                        console.log('Copay amount entered:', value);
                    }

                    // Enter copay value first time
                    await enterCopayValue(trimmedCopay);

                    console.log('Clicking copay OK button...');
                    await newPage.waitForSelector('#addUpdateInsuranceBtn24', { visible: true, timeout: 10000 });
                    await newPage.click('#addUpdateInsuranceBtn24');
                    await new Promise(resolve => setTimeout(resolve, 1000));

                    // Check for invalid copay alert with multiple possible selectors
                    let alertFound = false;
                    const alertSelectors = [
                        '.bootbox.modal .modal-footer button',
                        '.bootbox .modal-footer button',
                        '.modal-dialog .modal-footer button',
                        'div[class*="bootbox"] button',
                        '.modal .btn',
                        'button[data-bb-handler="Yes"]',
                        'button:contains("OK")',
                        '.modal-footer button'
                    ];

                    for (const selector of alertSelectors) {
                        try {
                            await newPage.waitForSelector(selector, { visible: true, timeout: 2000 });
                            console.log(`Invalid copay alert detected with selector: ${selector}. Clicking OK...`);
                            
                            // Click the OK button
                            await newPage.click(selector);
                            alertFound = true;
                            await new Promise(resolve => setTimeout(resolve, 500));
                            break;
                        } catch (e) {
                            // Continue to next selector
                            continue;
                        }
                    }

                    // If alert was found, re-enter the copay value
                    if (alertFound) {
                        console.log('Alert handled. Re-entering copay amount...');
                        
                        // Wait for the copay input to be available again
                        await newPage.waitForSelector('#addUpdateInsuranceIpt20', { visible: true, timeout: 10000 });
                        
                        // Re-enter the copay value
                        await enterCopayValue(trimmedCopay);

                        console.log('Clicking copay OK button again...');
                        await newPage.waitForSelector('#addUpdateInsuranceBtn24', { visible: true, timeout: 10000 });
                        await newPage.click('#addUpdateInsuranceBtn24');
                        await new Promise(resolve => setTimeout(resolve, 1000));
                    } else {
                        console.log('No invalid copay alert detected, proceeding normally.');
                    }

                    console.log('Copay update completed successfully.');
                } else {
                    console.log('First insurance row not found. Skipping copay update.');
                }
            } catch (copayError) {
                console.error('Error updating copay:', copayError);
                
                // Additional error handling - try to close any open dialogs
                try {
                    const closeButtons = await newPage.$$('button:contains("OK"), button:contains("Close"), .close, .btn-close');
                    for (const button of closeButtons) {
                        try {
                            await button.click();
                            await new Promise(resolve => setTimeout(resolve, 300));
                        } catch (e) {
                            // Ignore individual button click errors
                        }
                    }
                } catch (cleanupError) {
                    console.log('Cleanup attempt completed');
                }
            }
        } else {
            console.log('Insurance section not found. Skipping copay update.');
        }
    }
} else {
    console.log('No copay amount provided. Skipping copay update.');
}
        
await new Promise(resolve => setTimeout(resolve, 1000));
console.log('Waiting for the Save button...');

// Attach filtered dialog handler once (do this at top level, not inside loop)
newPage.on('dialog', async dialog => {
    const msg = dialog.message().trim();
    console.log(`🚨 Alert detected: ${msg}`);

    const shouldDismiss =
        msg.includes("This patient has both insurances and selfpay") ||
        msg.includes("SSN Number");

    if (shouldDismiss) {
        try {
            await dialog.dismiss();  // 👈 dismiss instead of accept
            console.log('❎ Dismissed alert (clicked Cancel).');
        } catch (e) {
            console.warn('⚠️ Failed to dismiss alert:', e.message);
        }
    } else {
        console.log('⛔ Unhandled alert ignored.');
    }
});

console.log('Clicking the Save button...');
await newPage.waitForSelector('#patient-demographicsBtn56', { visible: true });
await newPage.click('#patient-demographicsBtn56');

// Wait a moment for any validation to trigger
await new Promise(resolve => setTimeout(resolve, 2000));

// Check for responsible party validation error
console.log('🔍 Checking for responsible party validation error...');
const hasResponsiblePartyError = await newPage.evaluate(() => {
    // Check for the specific error message text
    const errorSpan = document.querySelector('span.formattext.ng-binding');
    if (errorSpan && errorSpan.textContent.includes('Please select Responsible party for the patient')) {
        return true;
    }
    return false;
});

if (hasResponsiblePartyError) {
    console.log('❌ RESPONSIBLE PARTY ERROR DETECTED - Clicking Cancel button...');
    
    try {
        await newPage.waitForSelector('#patient-demographicsBtn57', { visible: true, timeout: 5000 });
        await newPage.click('#patient-demographicsBtn57');
        console.log('✅ Cancel button clicked');
        
        // Wait for confirmation dialog and click "No"
        try {
            await newPage.waitForSelector('button[data-bb-handler="No"]', { 
                visible: true, 
                timeout: 8000 
            });
            
            await new Promise(resolve => setTimeout(resolve, 1000));
            await newPage.click('button[data-bb-handler="No"]');
            console.log('✅ Successfully clicked No button');
            
        } catch (noButtonError) {
            console.log('⚠️ No cancel confirmation dialog appeared');
        }
        
    } catch (cancelError) {
        console.error('❌ Failed to click Cancel button:', cancelError.message);
    }
    
    console.log(`❌ Responsible party error occurred for account ${accountNumber}. Moving to next account.`);
    return false;
}

// Handle any confirmation dialog that appears after clicking Save
console.log('Checking for confirmation dialog after Save...');
try {
    // Wait for any bootbox modal to appear
    await newPage.waitForSelector('div.bootbox.modal', { 
        visible: true, 
        timeout: 8000 
    });
    
    console.log('Modal detected after Save button click');
    
    // Add small delay to ensure modal is fully rendered
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // Check what type of modal this is
    const modalInfo = await newPage.evaluate(() => {
        const modalBody = document.querySelector('div.bootbox-body.padl5, div.bootbox-body');
        if (!modalBody) return { type: 'unknown', text: '' };
        
        const text = modalBody.textContent.trim();
        console.log('Modal text detected:', text);
        
        // Check if it's a confirmation dialog about modified values
        if (text.includes('The following values have been modified') || 
            text.includes('Are you sure you\'d like to continue') ||
            text.includes('Last Name') ||
            text.includes('First Name') ||
            text.includes('Date of Birth') ||
            text.includes('Sex')) {
            return { type: 'confirmation', text: text };
        }
        
        return { type: 'other', text: text };
    });
    
    console.log(`Modal type: ${modalInfo.type}`);
    
    if (modalInfo.type === 'confirmation') {
        console.log('✅ Confirmation dialog detected - Clicking YES to proceed with changes...');
        
        let yesClicked = false;
        
        // Try multiple methods to click the "Yes" button
        try {
            // Method 1: Click using data-bb-handler attribute
            await newPage.waitForSelector('button[data-bb-handler="Yes"]', { 
                visible: true, 
                timeout: 3000 
            });
            await newPage.click('button[data-bb-handler="Yes"]');
            yesClicked = true;
            console.log('✅ Successfully clicked Yes button (Method 1)');
        } catch (method1Error) {
            console.log('Method 1 failed, trying Method 2...');
            
            try {
                // Method 2: Find Yes button by text content
                const clickResult = await newPage.evaluate(() => {
                    const buttons = Array.from(document.querySelectorAll('button'));
                    const yesButton = buttons.find(btn => {
                        const text = btn.textContent.trim().toLowerCase();
                        return text === 'yes' || btn.getAttribute('data-bb-handler') === 'Yes';
                    });
                    
                    if (yesButton) {
                        yesButton.click();
                        return true;
                    }
                    return false;
                });
                
                if (clickResult) {
                    yesClicked = true;
                    console.log('✅ Successfully clicked Yes button (Method 2)');
                }
            } catch (method2Error) {
                console.log('Method 2 also failed');
            }
        }
        
        if (yesClicked) {
            // Wait for the modal to disappear
            try {
                await newPage.waitForSelector('div.bootbox.modal', { 
                    hidden: true, 
                    timeout: 10000 
                });
                console.log('✅ Confirmation modal closed after clicking Yes');
            } catch (closeError) {
                console.log('Warning: Modal may not have closed properly');
            }
        } else {
            console.error('❌ Failed to click Yes button');
        }
    } else {
        console.log('ℹ️ Modal is not a confirmation dialog, continuing...');
    }
    
} catch (noModalError) {
    console.log('ℹ️ No confirmation modal appeared after Save - proceeding normally');
}

// Wait for save process to complete
console.log('⏳ Waiting for save process to complete...');
await new Promise(resolve => setTimeout(resolve, 4000));

// Now check if the save was successful or if there are errors
console.log('🔍 Checking for errors after save process...');

const saveResult = await newPage.evaluate(() => {
    // Check for error message in the footer
    const errorElement = document.querySelector('#ptInfoModal > div.modal-footer.pd510.grey-bg.nomargin > span');
    const errorText = errorElement ? errorElement.innerText.trim() : '';
    
    // Check if the modal is still open (indicates potential error)
    const modalStillOpen = document.querySelector('#ptInfoModal') && 
                          getComputedStyle(document.querySelector('#ptInfoModal')).display !== 'none';
    
    // Check for any visible error indicators
    const hasVisibleError = errorText !== '' && errorText.length > 0;
    
    return {
        hasError: hasVisibleError,
        errorText: errorText,
        modalOpen: modalStillOpen
    };
});

console.log(`Save result - Has Error: ${saveResult.hasError}, Error Text: "${saveResult.errorText}", Modal Open: ${saveResult.modalOpen}`);

if (saveResult.hasError) {
    console.log('❌ ERROR DETECTED after save process. Initiating error handling...');
    console.log(`Error message: "${saveResult.errorText}"`);
    
    // Click Cancel button to exit the modal
    console.log('🔙 Clicking Cancel button...');
    try {
        await newPage.waitForSelector('#patient-demographicsBtn57', { visible: true, timeout: 5000 });
        await newPage.click('#patient-demographicsBtn57');
        console.log('✅ Cancel button clicked');
        
        // Wait for confirmation dialog and click "No"
        console.log('⏳ Waiting for cancel confirmation dialog...');
        try {
            await newPage.waitForSelector('button[data-bb-handler="No"]', { 
                visible: true, 
                timeout: 8000 
            });
            
            await new Promise(resolve => setTimeout(resolve, 1000));
            
            console.log('🚫 Clicking "No" on cancel confirmation...');
            
            let noClicked = false;
            try {
                await newPage.click('button[data-bb-handler="No"]');
                noClicked = true;
                console.log('✅ Successfully clicked No button');
            } catch (clickError) {
                // Try alternative method
                const clickResult = await newPage.evaluate(() => {
                    const noButton = document.querySelector('button[data-bb-handler="No"]') ||
                                   Array.from(document.querySelectorAll('button')).find(btn => 
                                       btn.textContent.trim().toLowerCase() === 'no'
                                   );
                    if (noButton) {
                        noButton.click();
                        return true;
                    }
                    return false;
                });
                
                if (clickResult) {
                    noClicked = true;
                    console.log('✅ Successfully clicked No button (alternative method)');
                }
            }
            
            if (!noClicked) {
                console.error('❌ Failed to click No button');
            }
            
        } catch (noButtonError) {
            console.log('⚠️ No cancel confirmation dialog appeared:', noButtonError.message);
        }
        
    } catch (cancelError) {
        console.error('❌ Failed to click Cancel button:', cancelError.message);
    }
    
    console.log(`❌ Error occurred while processing account ${accountNumber}. Moving to next account.`);
    return false;
    
} else {
    console.log('✅ SAVE COMPLETED SUCCESSFULLY - No errors detected!');
    console.log(`✅ Process completed successfully for account: ${accountNumber}`);
    return true;
}

// Error recovery section (in case of unexpected errors)
} catch (error) {
    console.error(`💥 Unexpected error processing account ${accountNumber}:`, error);
    
    // Try to recover if possible by clicking Cancel and No buttons
    try {
        console.log('🔄 Attempting error recovery...');
        
        // Check if Cancel button exists and click it
        const cancelButtonExists = await newPage.evaluate(() => {
            return !!document.querySelector('#patient-demographicsBtn57');
        });
        
        if (cancelButtonExists) {
            console.log('🔙 Clicking Cancel button during error recovery...');
            await newPage.click('#patient-demographicsBtn57');
            
            // Wait for confirmation alert and click "No"
            try {
                await newPage.waitForSelector('button[data-bb-handler="No"]', { 
                    visible: true, 
                    timeout: 8000 
                });
                
                await new Promise(resolve => setTimeout(resolve, 1000));
                
                console.log('🚫 Clicking "No" during error recovery...');
                
                try {
                    await newPage.click('button[data-bb-handler="No"]');
                    console.log('✅ Recovery: Successfully clicked No button');
                } catch (recoveryClickError) {
                    await newPage.evaluate(() => {
                        const noButton = document.querySelector('button[data-bb-handler="No"]');
                        if (noButton) noButton.click();
                    });
                    console.log('✅ Recovery: Clicked No button via evaluate');
                }
            } catch (alertError) {
                console.log('⚠️ Recovery: No confirmation alert appeared');
            }
        }
    } catch (recoveryError) {
        console.error('❌ Error during recovery attempt:', recoveryError);
    }
    
    return false;
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
    if (!row) continue; // Skip undefined rows
    
    const accountNumberRaw = row[columnMap["Acct No"]] || "";
    const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
    
if (!accountNumberRaw || status === "done" || status === "pending") continue;
    
    // ✅ CHECK NOTES BEFORE COUNTING - This is the key fix
    const callNotes = row[columnMap["Notes"]] || "";
    if (!callNotes.toString().trim()) {
        missingCallNotes++; // Count missing notes
        continue; // ❌ SKIP - Don't count this row for processing
    }
    
    // ✅ ONLY COUNT ROWS THAT HAVE NOTES AND WILL BE PROCESSED
    const accountNumber = accountNumberRaw.split("-")[0].trim();
    processedAccounts.add(accountNumber);
    newRowsToProcess++; // Now only counts rows with notes
}

// Check if all remaining rows are missing Notes
if (missingCallNotes > 0 && newRowsToProcess === 0) {
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
    if (!row) continue; // Skip undefined rows
    
    const accountNumberRaw = row[columnMap["Acct No"]] || "";
    const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
    
    // Skip row if there's no account number or it's already processed.
if (!accountNumberRaw || status === "done" || status === "pending") continue;
    
    // Skip row if the "Notes" field is empty.
    const callNotes = row[columnMap["Notes"]] || "";
    if (!callNotes.toString().trim()) continue;
    
    // Get copay amount if column exists
    let copayAmount = null;
    if (columnMap.hasOwnProperty("Specialist Copay")) {
        copayAmount = row[columnMap["Specialist Copay"]] || "";
        // Skip copay update if it's empty or already has $ symbol
        if (!copayAmount.toString().trim() || copayAmount.toString().trim().startsWith('$')) {
            console.log(`Skipping copay update for account ${accountNumberRaw}: ${copayAmount ? 'Already has $ symbol' : 'Empty copay'}`);
        } else {
            console.log(`Will update copay for account ${accountNumberRaw}: ${copayAmount}`);
        }
    }
    
    const accountNumber = accountNumberRaw.split("-")[0].trim();
    console.log(`Processing account: ${accountNumber}`);

    // Pass the newPage to processAccount
    const success = await processAccount(newPage, accountNumber, callNotes, copayAmount);

     if (success) {
                results.successful.push(accountNumber);
                originalData[i][columnMap["Result"]] = "Done";
                console.log(`✅ Account ${accountNumber} processed successfully - Status: Done`);
            } else {
                results.failed.push(accountNumber);
                originalData[i][columnMap["Result"]] = "Pending";
                console.log(`❌ Account ${accountNumber} failed to process - Status: Pending`);
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

        // Write data back to the sheet
// In your column widths array, add width for EPS Comments if needed
const columnWidthsWithSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 15, 50, 20];
const columnWidthsWithoutSpecialistCopay = [4, 10, 15, 25, 16, 10, 8, 10, 20, 14, 16, 45, 10, 15, 27, 50, 20];

// Function to check if Specialist Copay column exists
const hasSpecialistCopay = sheet.getRow(1).values.includes("Specialist Copay");

// Select appropriate column width array
const columnWidths = hasSpecialistCopay ? columnWidthsWithSpecialistCopay : columnWidthsWithoutSpecialistCopay;

// Set column widths dynamically
sheet.columns = columnWidths.map((width, index) => ({
    width,
    style: (index === columnMap["Notes"] || index === columnMap["EPS Comments"]) 
        ? { alignment: { wrapText: true } } 
        : {},
}));


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
            message: `${newRowsToProcess} rows processed successfully.`,
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
  
  // Determine content type based on file extension
  let contentType = 'application/octet-stream'; // Default
  
  if (fileData.filename.endsWith('.xlsx')) {
      contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  } else if (fileData.filename.endsWith('.xls')) {
      contentType = 'application/vnd.ms-excel';
  } else if (fileData.filename.endsWith('.csv')) {
      contentType = 'text/csv';
  }
  
  // Set the response headers
  res.setHeader('Content-Disposition', `attachment; filename="${fileData.filename}"`);
  res.setHeader('Content-Type', contentType);
  res.setHeader('Content-Length', fileData.buffer.length);
  
  // Send the file buffer
  res.send(fileData.buffer);
  
  // Optional: Remove the file from memory after some time
// Optional: Remove the file from memory after some time
setTimeout(() => {
    processedFiles.delete(fileId);
    console.log(`File ${fileId} removed from memory`);
}, 3 * 60 * 60 * 1000); // Remove after 3 hours

});

app.listen(port, () => console.log(`Server running on http://localhost:${port}`));	