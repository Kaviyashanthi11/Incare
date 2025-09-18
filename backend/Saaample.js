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
        });+
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
        const fromDate = '01/01/2023';
        console.log(`Setting From Date to: ${fromDate}`);
        await sessionPage.waitForSelector('#fromdate', { visible: true, timeout: 5000 });
        await sessionPage.click('#fromdate', { clickCount: 3 });
        await sessionPage.keyboard.press('Backspace');
        await sessionPage.type('#fromdate', fromDate, { delay: 30 });

        // Set To Date to current date in MM/DD/YYYY format
        const now = new Date();
        const toDate = `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
        console.log(`Setting To Date to: ${toDate}`);
        await sessionPage.waitForSelector('#todate', { visible: true, timeout: 5000 });
        await sessionPage.click('#todate', { clickCount: 3 });
        await sessionPage.keyboard.press('Backspace');
        await sessionPage.type('#todate', toDate, { delay: 30 });

        // Select "Encounter/Claim" option from the dropdown
        console.log('Selecting Claim Status "Encounter/Claim"...');
        await sessionPage.waitForSelector('#claimStatusCodeId', { visible: true, timeout: 5000 });
        const optionValue = await sessionPage.$eval('#claimStatusCodeId > option:nth-child(2)', el => el.value);
        await sessionPage.select('#claimStatusCodeId', optionValue);

        // Wait for the page to be fully loaded
console.log('Waiting for patient filter dropdown button to be available...');
await sessionPage.waitForSelector('#searchPtBtn', { visible: true, timeout: 10000 });

// Click the patient filter dropdown button to open the options
console.log('Clicking the Patient filter dropdown button...');
await sessionPage.click('#searchPtBtn');

// Wait for dropdown menu to appear and its options to be visible
console.log('Waiting for dropdown options to appear...');
await sessionPage.waitForSelector('#patient-lookupUl2 li', { visible: true, timeout: 5000 });

// Select the Account No option from the dropdown
console.log('Selecting Account No option...');
// Using a more robust selector with text content matching
await sessionPage.evaluate(() => {
  const options = Array.from(document.querySelectorAll('#patient-lookupUl2 li'));
  const accountNoOption = options.find(option => option.textContent.includes('Account No'));
  if (accountNoOption) {
    accountNoOption.click();
  } else {
    console.error('Account No option not found in dropdown');
  }
});

console.log('Account No option selected successfully');

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

        // Click the Claim Lookup button
        console.log('Clicking the Claim Lookup button...');
        await newPage.waitForSelector('#claimLookupBtn10', { visible: true, timeout: 5000 });
        await newPage.click('#claimLookupBtn10');
        await new Promise(resolve => setTimeout(resolve, 3000));

        // Click Create New Claim
        console.log('Selecting Create New Claim option...');
        await newPage.waitForSelector('#claimLookupBtn11', { visible: true, timeout: 8000 });
        await newPage.click('#claimLookupBtn11');
        await new Promise(resolve => setTimeout(resolve, 3000));

        console.log(`Filling in Resource with Provider value: ${provider}`);


// Extract the second part of the provider name (last name)
let providerLastName = '';
if (provider.includes(',')) {
    providerLastName = provider.split(',')[0].trim(); 
} else {
    providerLastName = provider.split(' ').slice(-1).join(' ');
}
console.log(`Using provider's last name: ${providerLastName}`);

// Fill Provider Input
await newPage.waitForSelector("input[id^='provider-lookupIpt1']", { visible: true, timeout: 5000 });
await newPage.click("input[id^='provider-lookupIpt1']", { clickCount: 3 });
await newPage.keyboard.press('Backspace');
await newPage.type("input[id^='provider-lookupIpt1']", providerLastName, { delay: 30 });

try {
    // Wait for the suggestion list to appear
    await newPage.waitForSelector("#provider-lookupUl1 li", { visible: true, timeout: 5000 });

    // Check if "No Providers Found" is displayed
    const isNoProviderFound = await newPage.evaluate(() => {
        const h4El = document.querySelector("#provider-lookupUl1 > li.ng-scope > div > h4");
        return h4El && h4El.textContent.trim() === "No Providers Found";
    });

    if (isNoProviderFound) {
        console.log("❌ No Providers Found in the list. Attempting to click Cancel...");
        try {
            await newPage.click("#createClaimBtn2");
            console.log("✅ Cancel button clicked due to no provider found.");
        } catch (err) {
            console.error("❌ Failed to click the Cancel button:", err.message);
        }
    } else {
        const providerClicked = await newPage.evaluate((providerLastName) => {
            const options = document.querySelectorAll("#provider-lookupUl1 li");
            for (let opt of options) {
                if (opt.innerText.trim().toLowerCase() === providerLastName.trim().toLowerCase()) {
                    opt.click();
                    return true;
                }
            }
            return false;
        }, providerLastName);

        if (providerClicked) {
            console.log(`✅ Provider '${providerLastName}' selected from the list.`);
        } else {
            console.log(`⚠️ Provider '${providerLastName}' not exactly matched in list. Leaving selection as-is.`);
        }
    }
} catch (timeoutErr) {
    console.error("⏳ Timeout waiting for provider suggestions. Clicking Cancel...");
    try {
        await newPage.click("#createClaimBtn2");
        console.log("✅ Cancel button clicked due to timeout in loading provider list.");
    } catch (err) {
        console.error("❌ Failed to click the Cancel button after timeout:", err.message);
    }
}

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
await newPage.type("#resource-lookupIpt1", providerLastName, { delay: 30 });

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

await new Promise(resolve => setTimeout(resolve, 1000));
const inputSelector = "input.createClaimFacilityCls";

// Wait for input field to be visible
await newPage.waitForSelector(inputSelector, { visible: true, timeout: 10000 });

// Clear existing value
await newPage.evaluate((selector) => {
    document.querySelector(selector).value = '';
}, inputSelector);

// Use provided facilityName directly
await newPage.type(inputSelector, facilityName, { delay: 30 });

// Short pause
await new Promise(resolve => setTimeout(resolve, 1000));

// Try to open the dropdown
try {
    await newPage.waitForSelector("#facility-lookupUl1", { visible: true, timeout: 2000 });
} catch {
    try {
        await newPage.waitForSelector("ul.dropdown-menu", { visible: true, timeout: 2000 });
    } catch {
        await newPage.keyboard.press('ArrowDown');
        await new Promise(resolve => setTimeout(resolve, 500));
    }
}

// Facility Selection Logic
try {
    const found = await newPage.evaluate((facilityName) => {
        const selectors = [
            "#facility-lookupUl1 li",
            "ul.dropdown-menu li",
            ".dropdown-menu li",
            "li[role='option']"
        ];

        let options = [];
        for (const selector of selectors) {
            options = document.querySelectorAll(selector);
            if (options.length > 0) break;
        }

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

    if (!found) {
        console.log(`⚠️ No Facility Found: '${facilityName}'. Attempting to click Cancel...`);
        try {
            await newPage.click("#createClaimBtn2");
            console.log("✅ Cancel button clicked due to no facility found.");
        } catch (err) {
            console.error("❌ Failed to click the Cancel button:", err.message);
        }
    }
} catch {
    console.log(`⚠️ Error while searching for facility: '${facilityName}'. Attempting to click Cancel...`);
    await newPage.click("#createClaimBtn2");
    console.log("✅ Cancel button clicked due to error.");
}


// Final pause to ensure the dropdown selection is processed
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
        
        await newPage.type('#claimserviceDate', billingDate, { delay: 30 });
        await newPage.keyboard.press('Tab');
        let patientSelected = false; // Track if a patient was successfully selected

        // Step 1: Wait for and fill the patient input field
        const ainputSelector = "input.patientFilterForClaim";
        await newPage.waitForSelector(ainputSelector, { visible: true, timeout: 10000 });
        
        // Clear the input and set autocomplete off
        await newPage.evaluate(selector => {
            const input = document.querySelector(selector);
            if (input) {
                input.value = '';
                input.setAttribute('autocomplete', 'off');
            }
        }, ainputSelector);
        
        // Type the account number
        await newPage.type(ainputSelector, accountNumber, { delay: 30 });
        
        // Step 2: Wait for dropdown to appear
        await new Promise(resolve => setTimeout(resolve, 1500));
        const dropdownSelectors = [
            "#patient-lookupUl1 li",
            "ul.dropdown-menu li",
            ".dropdown-menu li",
            "li[role='option']"
        ];
        
        let clicked = false;
        
        // Try to click the matching item in dropdown
        for (const sel of dropdownSelectors) {
            const items = await newPage.$$(sel);
            
            for (const item of items) {
                const text = await newPage.evaluate(el => el.textContent.trim(), item);
                if (text.toLowerCase().includes(accountNumber.toLowerCase())) {
                    const box = await item.boundingBox();
                    if (box) {
                        await newPage.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
                        await newPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { delay: 30 });
                        clicked = true;
                        patientSelected = true; // Mark as patient selected
                        break;
                    }
                }
            }
            
            if (clicked) break;
        }
        
        // Step 3: Fallback if not clicked
        if (!clicked && !patientSelected) {
            console.log('Dropdown item not found with bounding box. Using keyboard fallback.');
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
        
        console.log('Waiting for the patient selection dialog...');
        
        // Wait a moment for any dialog to appear
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Check for "No patient found" message
        const noPatientFound = await newPage.evaluate(() => {
            return document.body.textContent.includes('No patient found') || 
                   document.body.textContent.includes('No Patient Found');
        });
        
        if (noPatientFound && !patientSelected) {
            console.log('No patient found message detected');
            
            // Try multiple approaches to find and click the Cancel button
            try {
                const cancelButtonExists = await newPage.evaluate(() => {
                    const selectors = [
                        "button#createClaimBtn2",
                        "button.btn-lgrey:contains('Cancel')",
                        "button:contains('Cancel')",
                        ".modal-footer button[type='button']"
                    ];
                    
                    for (const selector of selectors) {
                        const elements = document.querySelectorAll(selector);
                        for (const element of elements) {
                            if (element.textContent.trim() === 'Cancel' || 
                                element.value === 'Cancel') {
                                element.click();
                                return true;
                            }
                        }
                    }
                    
                    const footerDiv = document.querySelector("#pn-createclaim > div > div > div.modal-footer.bd-mid-top.grey-bg.nomargin");
                    if (footerDiv) {
                        const buttons = footerDiv.querySelectorAll('button');
                        for (const button of buttons) {
                            if (button.textContent.trim() === 'Cancel') {
                                button.click();
                                return true;
                            }
                        }
                    }
                    
                    return false;
                });
                
                if (cancelButtonExists) {
                    console.log('Cancel button clicked via JavaScript');
                } else {
                    console.log('Trying to click on the modal footer area');
                    await newPage.waitForSelector("#pn-createclaim > div > div > div.modal-footer.bd-mid-top.grey-bg.nomargin", { 
                        visible: true, 
                        timeout: 5000 
                    });
                    
                    const footerBox = await newPage.$eval("#pn-createclaim > div > div > div.modal-footer.bd-mid-top.grey-bg.nomargin", 
                        el => {
                            const rect = el.getBoundingClientRect();
                            return { 
                                x: rect.x, 
                                y: rect.y, 
                                width: rect.width, 
                                height: rect.height 
                            };
                        }
                    );
                    
                    await newPage.mouse.click(
                        footerBox.x + footerBox.width * 0.25, 
                        footerBox.y + footerBox.height / 2,
                        { delay: 30 }
                    );
                    
                    console.log('Clicked on the left side of modal footer');
                }
                
                await new Promise(resolve => setTimeout(resolve, 1500));
                
            } catch (cancelError) {
                console.error('Error handling No Patient Found dialog:', cancelError);
            }
        } else {
            console.log('Patient found, skipping cancel');
        }
        
        // Proceed with the remaining steps
        console.log('Patient selection process completed');
        // Function to select provider from list
async function selectProvider(newPage, provider) {
    console.log(`📄 Excel Provider Value: ${provider}`);

    let providerLastName = '';
    if (provider.includes(',')) {
        providerLastName = provider.split(',')[0].trim(); 
    } else {
        providerLastName = provider.split(' ').slice(-1).join(' ');
    }
    console.log(`Using provider's last name: ${providerLastName}`);

    // Fill Provider Input
    await newPage.waitForSelector("input[id^='provider-lookupIpt1']", { visible: true, timeout: 5000 });
    await newPage.click("input[id^='provider-lookupIpt1']", { clickCount: 3 });
    await newPage.keyboard.press('Backspace');
    await newPage.type("input[id^='provider-lookupIpt1']", providerLastName, { delay: 30 });

    try {
        await newPage.waitForSelector("#provider-lookupUl1 li", { visible: true, timeout: 5000 });

        const noProviderFound = await newPage.evaluate(() => {
            const el = document.querySelector("#provider-lookupUl1 > li.ng-scope > div > h4");
            return el && el.textContent.trim() === "No Providers Found";
        });

        if (noProviderFound) {
            console.log("❌ No Providers Found. Clicking Cancel...");
            await newPage.click("#createClaimBtn2");
            return false;
        }

        // Mark the matching element with a data attribute inside page context
        const found = await newPage.$$eval("#provider-lookupUl1 li", (items, search) => {
            const match = search.toLowerCase();
            for (const item of items) {
                const text = item.innerText.trim().toLowerCase();
                if (text.includes(match) || match.includes(text)) {
                    item.setAttribute("data-to-click", "true");
                    return true;
                }
            }
            // If nothing matches exactly, mark the first item
            if (items.length > 0) {
                items[0].setAttribute("data-to-click", "true");
                return true;
            }
            return false;
        }, providerLastName);

        if (found) {
            const matchEl = await newPage.$("#provider-lookupUl1 li[data-to-click='true']");
            if (matchEl) {
                await matchEl.click();
                await newPage.evaluate(() => {
                    document.querySelectorAll("li[data-to-click]").forEach(el => el.removeAttribute("data-to-click"));
                });
                console.log("✅ Provider selected successfully.");
                await newPage.waitForTimeout(1500);
                return true;
            } else {
                console.warn("⚠️ Element was marked but not found.");
                return false;
            }
        } else {
            console.warn("⚠️ No providers matched at all.");
            return false;
        }

    } catch (err) {
        console.error("⏳ Timeout or other error:", err.message);
        try {
            await newPage.click("#createClaimBtn2");
        } catch (clickErr) {
            console.error("❌ Cancel button click failed:", clickErr.message);
        }
        return false;
    }
}
const selectedProvider = await newPage.$eval("input[id^='provider-lookupIpt1']", el => el.value.trim());
if (!selectedProvider || !selectedProvider.toLowerCase().includes(provider.split(',')[0].trim().toLowerCase())) {
    console.warn("🔁 Provider mismatch. Reselecting...");
    await selectProvider(newPage, provider);
}


const buttonExists = await newPage.$('#createClaimBtn');
if (buttonExists) {
  console.log('✅ Create Claim button found. Clicking...');
  await new Promise(resolve => setTimeout(resolve, 3000));

  await newPage.click('#createClaimBtn');

  // Wait for modal (optional)
  try {
    await newPage.waitForSelector('.modal-dialog', { timeout: 5000 });
    console.log('🟡 Modal appeared after clicking OK');
    const extractedProvider = await handleModalAndCancel(newPage); // 👉 Get extracted provider
    
    // Print detailed comparison information
    console.log(`🔍 Provider Comparison Details:`);
    console.log(`📊 Excel Provider: "${provider}" (Length: ${provider ? provider.length : 0})`);
    console.log(`🖥️ Extracted Provider: "${extractedProvider}" (Length: ${extractedProvider ? extractedProvider.length : 0})`);
    
    // Clean and normalize both providers for comparison
    const cleanExcelProvider = provider ? provider.trim().toLowerCase() : '';
    const cleanExtractedProvider = extractedProvider ? extractedProvider.trim().toLowerCase() : '';
    
    console.log(`🧹 Cleaned Excel Provider: "${cleanExcelProvider}"`);
    console.log(`🧹 Cleaned Extracted Provider: "${cleanExtractedProvider}"`);
    

    if (cleanExtractedProvider && cleanExcelProvider && cleanExtractedProvider !== cleanExcelProvider) {
        console.log(`❌ PROVIDER MISMATCH DETECTED!`);
        console.log(`📊 Excel Provider: "${provider}"`);
        console.log(`🖥️ Extracted Provider: "${extractedProvider}"`);
        console.log('🔄 Creating claim again for the same account number...');
        
        // Create claim again for the same account number - only when providers are different
        return await createClaimAgain(newPage, accountNumber, provider, facilityName, billingDate, diagnosisCodes, cptCodes, admitDate);
    } else if (!cleanExtractedProvider) {
        console.log('⚠️ EXTRACTED PROVIDER IS EMPTY - No action needed, continuing with normal process');
    } else if (!cleanExcelProvider) {
        console.log('⚠️ EXCEL PROVIDER IS EMPTY - No action needed, continuing with normal process');
    } else {
        console.log('✅ PROVIDERS MATCH - No action needed, continuing with normal process');
    }
    
    console.log('➡️ Continuing with normal claim processing...');
    
  } catch (error) {
    console.log('🟢 No modal appeared. Continuing with claim process...');
    console.log('➡️ Proceeding with normal claim creation...');
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
const billedFees = [];
const processedCPTs = new Map(); // Track processed CPT codes and their units from asterisk

// Extract CPT codes and their units (from asterisk values)
for (const code of cptCodes) {
  let baseCode, units = 1;
  
  if (code.includes('*')) {
    // Split by asterisk to get code and units
    const parts = code.split('*');
    baseCode = parts[0].split('-')[0].trim(); // Get base code (before any modifier)
    units = parseInt(parts[1].trim()) || 1;
  } else {
    baseCode = code.split('-')[0].trim();
    units = 1;
  }
  
  // Store the CPT code with its units (don't count occurrences, use asterisk value)
  if (!processedCPTs.has(baseCode)) {
    processedCPTs.set(baseCode, units);
  }
}

console.log('CPT codes with units from asterisk:', Object.fromEntries(processedCPTs));

// Process each unique CPT code
const uniqueCPTCodes = [...processedCPTs.keys()];

for (let i = 0; i < uniqueCPTCodes.length; i++) {
  const baseCode = uniqueCPTCodes[i];
  let modifier = "";
  
  // Get the original code with modifier if present
  const originalCode = cptCodes.find(code => {
    const codeBase = code.includes('*') ? code.split('*')[0] : code;
    return codeBase.split('-')[0].trim() === baseCode;
  });
  
  if (originalCode) {
    const codeWithoutUnits = originalCode.includes('*') ? originalCode.split('*')[0] : originalCode;
    if (codeWithoutUnits.includes('-')) {
      modifier = codeWithoutUnits.split('-')[1].trim();
    }
  }
  
  // Get units from the asterisk value for this CPT code
  const units = processedCPTs.get(baseCode);
  
  try {
    console.log(`🔄 Processing CPT code ${i + 1} of ${uniqueCPTCodes.length}: ${baseCode} (units: ${units})`);
    
    // Click on the CPT cell to begin entry
    const cptCellSelector = 'td[ng-click*="enterCPT"]';
    await newPage.waitForSelector(cptCellSelector, { visible: true, timeout: 8000 });
    await newPage.click(cptCellSelector);
    await new Promise(resolve => setTimeout(resolve, 600));
    
    // Enter the CPT code
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
    
      try {
        console.log('Waiting for CPT confirmation alert...');
        const confirmYesSelector = 'body > div.bootbox.modal.fade.bluetheme.medium-width.in > div > div > div.modal-footer > button.btn.btn.btn-lblue.btn-lgrey.btn-xs.btn-default.btn-yes';
        await newPage.waitForSelector(confirmYesSelector, { visible: true, timeout: 8000 });
        await new Promise(resolve => setTimeout(resolve, 1000));
        console.log('Clicking "Yes" on CPT confirmation...');
        await newPage.click(confirmYesSelector);
      } catch (modalErr) {
        console.log('No CPT confirmation modal appeared or failed to click:', modalErr.message);
      }
      
      // Check and set POS (Place of Service) and apply modifiers if needed
      const posSelector = '#billingClaimIpt24';
      await newPage.waitForSelector(posSelector, { visible: true, timeout: 8000 });
      const posValue = await newPage.$eval(posSelector, el => el.value.trim());
      console.log(`POS Value: ${posValue}`);
      
      if (posValue !== "11") {
        const has99291 = cptCodes.some(c => c.includes("99291"));
        const has99292 = cptCodes.some(c => c.includes("99292"));
        
        if ((baseCode === "99291" || baseCode === "99292") && has99291 && has99292) {
          modifier = "25";
          console.log(`Applying rule: Both 99291 & 99292 exist — modifier 25`);
        } else if (baseCode === "99291" && uniqueCPTCodes.length > i + 1 && uniqueCPTCodes[i + 1] !== "99292") {
          modifier = "25";
          console.log(`Applying rule: 99291 followed by non-99292 — modifier 25`);
        } else if (baseCode === "99291") {
          modifier = "";
          console.log(`Applying rule: Solo 99291 — no modifier`);
        }
        
        if (modifier) {
          const modifierSelector = `#billingClaimIpt8ngR${i}`;
          await newPage.waitForSelector(modifierSelector, { visible: true, timeout: 5000 });
          
          console.log(`Using modifier field selector for row ${i}: ${modifierSelector}`);
          
          const modifierInput = await newPage.$(modifierSelector);
          await modifierInput.focus();
          await newPage.evaluate(el => el.scrollIntoView(), modifierInput);
          await modifierInput.click({ clickCount: 3 });
          await modifierInput.press('Backspace');
          await new Promise(resolve => setTimeout(resolve, 200));
          await modifierInput.type(modifier, { delay: 100 });
          console.log(`✅ Entered modifier: ${modifier} in row ${i}`);
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      } else {
        console.log(`POS is ${posValue} — skipping modifiers`);
      }
    
    // Set units from asterisk value if more than 1
    if (units > 1) {
      try {
        // Find the units input field for the current row
        const unitsSelector = `#billingClaimIpt16ngR${i}`;
        await newPage.waitForSelector(unitsSelector, { visible: true, timeout: 5000 });
        
        const unitsInput = await newPage.$(unitsSelector);
        await unitsInput.focus();
        await newPage.evaluate(el => el.scrollIntoView(), unitsInput);
        await unitsInput.click({ clickCount: 3 });
        await unitsInput.press('Backspace');
        await new Promise(resolve => setTimeout(resolve, 200));
        await unitsInput.type(units.toString(), { delay: 100 });
        console.log(`✅ Updated units to ${units} for CPT ${baseCode} (from asterisk value)`);
        
        // Click the OK button to confirm the update
        const okButtonSelector = '#CPTUpdateOkBtn';
        await newPage.waitForSelector(okButtonSelector, { visible: true, timeout: 5000 });
        await newPage.click(okButtonSelector);
        console.log('Clicked OK button to confirm units update');
      } catch (unitsErr) {
        console.error(`⚠️ Failed to update units for CPT ${baseCode}:`, unitsErr.message);
      }
    }
    
      // Check billed fee after entering CPT and modifier
      try {
        const billedFeeSelector = `#billingClaimIpt17ngR${i}`;
        await newPage.waitForSelector(billedFeeSelector, { visible: true, timeout: 8000 });
        
        // Get the value from the input field directly
        const billedFeeValue = await newPage.$eval(billedFeeSelector, el => {
          // If it's an input element, get its value
          if (el.tagName === 'INPUT') {
            return el.value.trim();
          } 
          // Otherwise get the text content
          return el.textContent.trim();
        });
        
        // Extract numeric value, handling currency symbols and formatting
        const numericValue = billedFeeValue.replace(/[^0-9.]/g, '');
        const billedFee = parseFloat(numericValue);
        
        // Use isNaN to check if we got a valid number
        const validBilledFee = !isNaN(billedFee) ? billedFee : 0;
        
        // For multiple units, calculate the per-unit fee
        const perUnitFee = units > 1 ? validBilledFee / units : validBilledFee;
        
        console.log(`💲 Billed Fee for CPT ${baseCode} (Row ${i}): ${validBilledFee} (${perUnitFee} per unit)`);
        
        billedFees.push({ 
          code: baseCode, 
          units: units,
          fee: validBilledFee,
          perUnitFee: perUnitFee,
          originalText: billedFeeValue 
        });
      } catch (feeErr) {
        console.error(`⚠️ Failed to check billed fee for CPT ${baseCode}:`, feeErr.message);
      }
      
      console.log(`✅ Done processing CPT: ${baseCode}`);
      
    } catch (error) {
      console.error(`❌ Failed to process CPT code ${baseCode}:`, error);
      await newPage.screenshot({ path: `cpt_error_${baseCode}.png` });
    }
  
    await new Promise(resolve => setTimeout(resolve, 1000));
  }
  
  // After processing all CPT codes, check if we need to handle dates
  // We'll check POS value again and if needed, enter the dates once
  try {
    // Check POS value from the page (get it after all CPTs are processed)
    const posSelector = '#billingClaimIpt24';
    await newPage.waitForSelector(posSelector, { visible: true, timeout: 8000 });
    const posValue = await newPage.$eval(posSelector, el => el.value.trim());
    
    console.log(`Final POS check: ${posValue}`);
    
    // Handle date entry for POS values 21 or 31 after all CPTs are processed
    if (["21", "31"].includes(posValue)) {
      console.log(`POS is ${posValue} — proceeding to date input.`);
      
      // Find and click the data button to open the date modal
      const dataBtnSelector = 'button[id^="claimDataBtn"]';
      const dataBtnHandle = await newPage.$(dataBtnSelector);
      if (dataBtnHandle) {
        const btnId = await newPage.evaluate(el => el.id, dataBtnHandle);
        console.log(`Found data button with ID: ${btnId}`);
        await dataBtnHandle.click();
        console.log(`Clicked data button`);
        await new Promise(resolve => setTimeout(resolve, 1500)); // Longer wait for modal to open fully
      } else {
        console.warn(`Data button with selector ${dataBtnSelector} not found`);
        throw new Error("Data button not found");
      }
      
      // Wait for admit date input to be visible
      const admitDateInputSelector = '#dtHCFAFrom16From';
      const dischargeInputSelector = '#dtHCFAFrom16To';
      await newPage.waitForSelector(admitDateInputSelector, { visible: true, timeout: 5000 });
      
      // Helper function to format date with leading zeros (MM/DD/YYYY)
      const formatDateWithLeadingZeros = (dateStr) => {
        const date = new Date(dateStr);
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        const year = date.getFullYear();
        return `${month}/${day}/${year}`;     
      };
      
      const formattedBillingDate = formatDateWithLeadingZeros(billingDate);
      
// After calculation of admit date, add this validation check before entering dates:

// Determine admit date based on CPT codes present
let admitDateOnly;
let dischargeDateToUse = formattedBillingDate;
      
// Check if we have any special CPT codes that affect date handling
const has99222or99223 = cptCodes.some(code => {
  const baseCode = code.split('-')[0].trim();
  return ["99222", "99223"].includes(baseCode);
});
      
// Check if we have discharge CPT codes
const has99238or99239 = cptCodes.some(code => {
  const baseCode = code.split('-')[0].trim();
  return ["99238", "99239"].includes(baseCode);
});
      
// Special case: if we have 99222 or 99223, use billing date as admit date
if (has99222or99223) {
  console.log(`🔍 Special CPT codes 99222/99223 detected - using billing date as admit date`);
  admitDateOnly = formattedBillingDate;
} else {
  // For other codes, process normally
  // Clean admit date if present in expected format
  admitDateOnly = admitDate.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)?.[0] || "";
  
  // Format with leading zeros if date is present
  if (admitDateOnly) {
    admitDateOnly = formatDateWithLeadingZeros(admitDateOnly);
  }
  
  // If admit date is NOT provided, calculate it from discharge date
  if (!admitDateOnly) {
    console.log(`❗ Admit Date not provided. Calculating from Discharge Date: ${formattedBillingDate}`);
    
    // Calculate admit date as 1 day before discharge date
    const admit = new Date(billingDate);
    admit.setDate(admit.getDate() - 1);
    
    // Format calculated admit date with leading zeros
    const month = String(admit.getMonth() + 1).padStart(2, '0');
    const day = String(admit.getDate()).padStart(2, '0');
    const year = admit.getFullYear();
    admitDateOnly = `${month}/${day}/${year}`;
    
    console.log(`Calculated Admit Date: ${admitDateOnly}`);
  } else {
    console.log(`✅ Using provided Admit Date: ${admitDateOnly}`);
  } 
}
// Only check for same dates when dealing with discharge codes
if (has99238or99239 && admitDateOnly === dischargeDateToUse) {
  console.log(`⚠️ Admit date and discharge date are the same for discharge codes: ${admitDateOnly}`);
  
  // Calculate new admit date as 1 day before discharge date
  const admit = new Date(dischargeDateToUse);
  admit.setDate(admit.getDate() - 1);
  
  // Format calculated admit date with leading zeros
  const month = String(admit.getMonth() + 1).padStart(2, '0');
  const day = String(admit.getDate()).padStart(2, '0');
  const year = admit.getFullYear();
  admitDateOnly = `${month}/${day}/${year}`;
  
  console.log(`Updated Admit Date to be one day before discharge: ${admitDateOnly}`);
}

  
// ADD THIS NEW VALIDATION: Check if admit date is greater than billing date
const admitDateObj = new Date(admitDateOnly);
const billingDateObj = new Date(billingDate);

if (admitDateObj > billingDateObj) {
  console.log(`⚠️ Admit date (${admitDateOnly}) is greater than billing date (${formattedBillingDate})`);
  
  // Set admit date to one day before billing date
  const newAdmitDate = new Date(billingDate);
  newAdmitDate.setDate(newAdmitDate.getDate() - 1);
  
  // Format the new admit date
  const month = String(newAdmitDate.getMonth() + 1).padStart(2, '0');
  const day = String(newAdmitDate.getDate()).padStart(2, '0');
  const year = newAdmitDate.getFullYear();
  admitDateOnly = `${month}/${day}/${year}`;
  
  console.log(`Corrected admit date to one day before billing date: ${admitDateOnly}`);
}

// Now enter the validated admit date
await newPage.type(admitDateInputSelector, admitDateOnly);
console.log(`Entered admit date: ${admitDateOnly}`);
await new Promise(resolve => setTimeout(resolve, 500));

// Enter discharge date ONLY if we have discharge CPT codes
if (has99238or99239) {
  await newPage.type(dischargeInputSelector, dischargeDateToUse);
  await newPage.keyboard.press('Tab'); // Move focus away 
  console.log(`Entered discharge date: ${dischargeDateToUse}`);
  await new Promise(resolve => setTimeout(resolve, 500));
} else {
  console.log(`No discharge codes (99238/99239) present - skipping discharge date entry`);
}

// Add extra wait time before attempting to click OK
await new Promise(resolve => setTimeout(resolve, 1500));

// Try multiple click strategies for the OK button
const okBtnSelector = '#claimDataBtn13';
const bootboxOkSelector = 'body > div.bootbox.modal.fade.bluetheme.medium-width.bootbox-alert.in > div > div > div.modal-footer > button';

try {
  let modalStillOpen = true;
  let attempts = 0;

  while (modalStillOpen && attempts < 3) {
    attempts++;

    // Check if OK button is visible
    await newPage.waitForSelector(okBtnSelector, { visible: true, timeout: 5000 });
    console.log(`Attempt ${attempts}: OK button is visible, trying click...`);

    // Try direct click
    await newPage.click(okBtnSelector);
    await new Promise(resolve => setTimeout(resolve, 1000));

    // Check for bootbox alert
    const bootboxExists = await newPage.$(bootboxOkSelector);
    if (bootboxExists) {
      console.log("Bootbox alert detected. Clicking OK...");
      await newPage.click(bootboxOkSelector);
      await new Promise(resolve => setTimeout(resolve, 1000));

      // Re-enter admit date
      await newPage.click(admitDateInputSelector, { clickCount: 3 });
      await newPage.type(admitDateInputSelector, admitDateOnly);
      console.log(`Re-entered admit date: ${admitDateOnly}`);
      await new Promise(resolve => setTimeout(resolve, 500));

      // Re-enter discharge date if needed
      if (has99238or99239) {
        await newPage.click(dischargeInputSelector, { clickCount: 3 });
        await newPage.type(dischargeInputSelector, dischargeDateToUse);
        await newPage.keyboard.press('Tab');
        console.log(`Re-entered discharge date: ${dischargeDateToUse}`);
        await new Promise(resolve => setTimeout(resolve, 500));
      }

      continue; // Retry clicking OK
    }

    // Check if modal is still open
    modalStillOpen = await newPage.evaluate(() => {
      return !!document.querySelector('.modal-dialog:not([style*="display: none"])');
    });

    if (!modalStillOpen) {
      console.log(`✅ Modal closed successfully after ${attempts} attempt(s)`);
      break;
    } else {
      console.log("Modal still open, retrying...");
    }
  }
} catch (error) {
  console.error(`Failed to click OK button or handle modal:`, error);
  await newPage.screenshot({ path: `ok_button_error_date.png` });
}

    } else {
      console.log(`POS is ${posValue} — no date entry needed.`);
    }
    
    console.log(`✅ Completed processing all CPT codes and dates.`);
  } catch (error) {
    console.error(`❌ Failed during date processing:`, error);
    await newPage.screenshot({ path: `date_processing_error.png` });
  }
const claimNoSelector = '#billingClaimTbl2 > tbody > tr:nth-child(1) > td:nth-child(2) > span';

        await newPage.waitForSelector(claimNoSelector, { visible: true, timeout: 5000 });
        
        const claimNumber = await newPage.$eval(claimNoSelector, el => el.textContent.trim());
        
        console.log(`🧾 Claim Number: ${claimNumber}`);
        try {
            // Step 1: Find the claim status dropdown
            const statusSelector = 'select[id^="claimStatusSel"]';
            await newPage.waitForSelector(statusSelector, { visible: true, timeout: 5000 });
            const statusDropdownHandle = await newPage.$(statusSelector);
          
            if (statusDropdownHandle) {
              const selectedText = await newPage.evaluate(el => {
                const selectedOption = el.options[el.selectedIndex];
                return selectedOption ? selectedOption.textContent.trim() : '';
              }, statusDropdownHandle);  
          
              console.log(`Current Claim Status: ${selectedText}`);
          
              // Step 2: If not pending, set to Pending
              if (selectedText.toLowerCase() !== "pending") {
                await newPage.select(statusSelector, "PEN");
                console.log("Changed Claim Status to Pending");
                await new Promise(resolve => setTimeout(resolve, 500)); // wait for change event
              }
              await new Promise(resolve => setTimeout(resolve, 1000));

// Step 3: Click the first OK button
try {
  const okSelector = 'button[id^="claimScreenOkBtn"]';
  await newPage.waitForSelector(okSelector, { visible: true, timeout: 5000 });
  
  const okButtonHandle = await newPage.$(okSelector);
  if (okButtonHandle) {
    await okButtonHandle.evaluate(btn => btn.scrollIntoView({ behavior: "smooth", block: "center" }));
    await new Promise(resolve => setTimeout(resolve, 300));
    
    // Use bounding box + mouse click for higher reliability
    const box = await okButtonHandle.boundingBox();
    if (box) {
      await newPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
      console.log("Clicked first OK button.");
    } else {
      console.warn("First OK button not clickable (no bounding box).");
    }
    
    // Wait for error popup to appear after clicking first OK
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    // Step 4: Now handle the error popup that appears - click the second OK button
    try {
      // Error popup OK button selector - from modal-footer
      const errorOkSelector = 'div.bootbox.modal.fade.bluetheme.medium-width.in button.btn.btn-blue.btn-sm.btn-xs, button[data-bb-handler="Yes"]';
      await newPage.waitForSelector(errorOkSelector, { visible: true, timeout: 5000 });
      
      const errorOkButton = await newPage.$(errorOkSelector);
      if (errorOkButton) {
        await errorOkButton.evaluate(btn => btn.scrollIntoView({ behavior: "smooth", block: "center" }));
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // Click the error popup's OK button
        const errorOkBox = await errorOkButton.boundingBox();
        if (errorOkBox) {
          await newPage.mouse.click(errorOkBox.x + errorOkBox.width / 2, errorOkBox.y + errorOkBox.height / 2);
          console.log("Clicked error popup OK button.");
        } else {
          console.warn("Error popup OK button not clickable (no bounding box).");
          await errorOkButton.click();
        }
        
        // Wait for UI to update after clicking error OK button
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Step 5: Finally click the Cancel button
        const cancelSelector = '#billingClaimBtn36';
        await newPage.waitForSelector(cancelSelector, { visible: true, timeout: 5000 });
        
        const cancelButton = await newPage.$(cancelSelector);
        if (cancelButton) {
          await cancelButton.evaluate(btn => btn.scrollIntoView({ behavior: "smooth", block: "center" }));
          await new Promise(resolve => setTimeout(resolve, 300));
          
          // Click the Cancel button
          const cancelBox = await cancelButton.boundingBox();
          if (cancelBox) {
            await newPage.mouse.click(cancelBox.x + cancelBox.width / 2, cancelBox.y + cancelBox.height / 2);
            console.log("Clicked Cancel button after handling error.");
          } else {
            console.warn("Cancel button not clickable (no bounding box).");
            await cancelButton.click();
          }
          
          await new Promise(resolve => setTimeout(resolve, 1000));
        } else {
          console.warn("Cancel button not found after error handling.");
        }
      } else {
        console.warn("Error popup OK button not found.");
      }
    } catch (errorPopupError) {
      console.error(`Error handling popup: ${errorPopupError.message}`);
      
      // Fallback: try to find any visible OK button in a modal
      try {
        const anyOkSelector = 'div.modal.in button:contains("OK"), div.modal.in button.btn-blue';
        const anyOkButton = await newPage.$(anyOkSelector);
        if (anyOkButton) {
          await anyOkButton.click();
          console.log("Clicked fallback OK button in error popup.");
          
          // Still try to click Cancel after fallback
          await new Promise(resolve => setTimeout(resolve, 1000));
          await newPage.click('#billingClaimBtn36').catch(e => console.warn("Cancel button click failed after fallback."));
        }
      } catch (fallbackError) {
        console.error(`Fallback error handling failed: ${fallbackError.message}`);
      }
    }
  } else {
    console.warn("First OK button not found.");
  }
} catch (error) {
  console.error(`Error in claim status handling sequence: ${error.message}`);
}
            }
          } catch (error) {
            console.error(`Error handling claim status: ${error.message}`);
          }
          
  
  return { success: true, claimNumber,billedFees };
  } else {
   return { success: false, error: 'Create Claim button not found' };

  }
async function createClaimAgain(newPage, accountNumber, provider, facilityName, billingDate, diagnosisCodes,cptCodes,admitDate) {
    try {
        // Validate required fields
        if (!provider || !facilityName || !billingDate || !accountNumber) {
            console.error('Missing required fields:', { provider, facilityName, billingDate, accountNumber });
            return false;
        }

        // Click the Claim Lookup button
        console.log('Clicking the Claim Lookup button...');
        await newPage.waitForSelector('#claimLookupBtn10', { visible: true, timeout: 5000 });
        await newPage.click('#claimLookupBtn10');
        await new Promise(resolve => setTimeout(resolve, 3000));

        // Click Create New Claim
        console.log('Selecting Create New Claim option...');
        await newPage.waitForSelector('#claimLookupBtn11', { visible: true, timeout: 8000 });
        await newPage.click('#claimLookupBtn11');
        await new Promise(resolve => setTimeout(resolve, 3000));

        console.log(`Filling in Resource with Provider value: ${provider}`);


// Extract the second part of the provider name (last name)
let providerLastName = '';
if (provider.includes(',')) {
    providerLastName = provider.split(',')[0].trim(); 
} else {
    providerLastName = provider.split(' ').slice(-1).join(' ');
}
console.log(`Using provider's last name: ${providerLastName}`);

// Fill Provider Input
await newPage.waitForSelector("input[id^='provider-lookupIpt1']", { visible: true, timeout: 5000 });
await newPage.click("input[id^='provider-lookupIpt1']", { clickCount: 3 });
await newPage.keyboard.press('Backspace');
await newPage.type("input[id^='provider-lookupIpt1']", providerLastName, { delay: 30 });

try {
    // Wait for the suggestion list to appear
    await newPage.waitForSelector("#provider-lookupUl1 li", { visible: true, timeout: 5000 });

    // Check if "No Providers Found" is displayed
    const isNoProviderFound = await newPage.evaluate(() => {
        const h4El = document.querySelector("#provider-lookupUl1 > li.ng-scope > div > h4");
        return h4El && h4El.textContent.trim() === "No Providers Found";
    });

    if (isNoProviderFound) {
        console.log("❌ No Providers Found in the list. Attempting to click Cancel...");
        try {
            await newPage.click("#createClaimBtn2");
            console.log("✅ Cancel button clicked due to no provider found.");
        } catch (err) {
            console.error("❌ Failed to click the Cancel button:", err.message);
        }
    } else {
        const providerClicked = await newPage.evaluate((providerLastName) => {
            const options = document.querySelectorAll("#provider-lookupUl1 li");
            for (let opt of options) {
                if (opt.innerText.trim().toLowerCase() === providerLastName.trim().toLowerCase()) {
                    opt.click();
                    return true;
                }
            }
            return false;
        }, providerLastName);

        if (providerClicked) {
            console.log(`✅ Provider '${providerLastName}' selected from the list.`);
        } else {
            console.log(`⚠️ Provider '${providerLastName}' not exactly matched in list. Leaving selection as-is.`);
        }
    }
} catch (timeoutErr) {
    console.error("⏳ Timeout waiting for provider suggestions. Clicking Cancel...");
    try {
        await newPage.click("#createClaimBtn2");
        console.log("✅ Cancel button clicked due to timeout in loading provider list.");
    } catch (err) {
        console.error("❌ Failed to click the Cancel button after timeout:", err.message);
    }
}

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
await newPage.type("#resource-lookupIpt1", providerLastName, { delay: 30 });

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

await new Promise(resolve => setTimeout(resolve, 1000));
const inputSelector = "input.createClaimFacilityCls";

// Wait for input field to be visible
await newPage.waitForSelector(inputSelector, { visible: true, timeout: 10000 });

// Clear existing value
await newPage.evaluate((selector) => {
    document.querySelector(selector).value = '';
}, inputSelector);

// Use provided facilityName directly
await newPage.type(inputSelector, facilityName, { delay: 30 });

// Short pause
await new Promise(resolve => setTimeout(resolve, 1000));

// Try to open the dropdown
try {
    await newPage.waitForSelector("#facility-lookupUl1", { visible: true, timeout: 2000 });
} catch {
    try {
        await newPage.waitForSelector("ul.dropdown-menu", { visible: true, timeout: 2000 });
    } catch {
        await newPage.keyboard.press('ArrowDown');
        await new Promise(resolve => setTimeout(resolve, 500));
    }
}

// Facility Selection Logic
try {
    const found = await newPage.evaluate((facilityName) => {
        const selectors = [
            "#facility-lookupUl1 li",
            "ul.dropdown-menu li",
            ".dropdown-menu li",
            "li[role='option']"
        ];

        let options = [];
        for (const selector of selectors) {
            options = document.querySelectorAll(selector);
            if (options.length > 0) break;
        }

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

    if (!found) {
        console.log(`⚠️ No Facility Found: '${facilityName}'. Attempting to click Cancel...`);
        try {
            await newPage.click("#createClaimBtn2");
            console.log("✅ Cancel button clicked due to no facility found.");
        } catch (err) {
            console.error("❌ Failed to click the Cancel button:", err.message);
        }
    }
} catch {
    console.log(`⚠️ Error while searching for facility: '${facilityName}'. Attempting to click Cancel...`);
    await newPage.click("#createClaimBtn2");
    console.log("✅ Cancel button clicked due to error.");
}


// Final pause to ensure the dropdown selection is processed
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
        
        await newPage.type('#claimserviceDate', billingDate, { delay: 30 });
        await newPage.keyboard.press('Tab');
        let patientSelected = false; // Track if a patient was successfully selected

        // Step 1: Wait for and fill the patient input field
        const ainputSelector = "input.patientFilterForClaim";
        await newPage.waitForSelector(ainputSelector, { visible: true, timeout: 10000 });
        
        // Clear the input and set autocomplete off
        await newPage.evaluate(selector => {
            const input = document.querySelector(selector);
            if (input) {
                input.value = '';
                input.setAttribute('autocomplete', 'off');
            }
        }, ainputSelector);
        
        // Type the account number
        await newPage.type(ainputSelector, accountNumber, { delay: 30 });
        
        // Step 2: Wait for dropdown to appear
        await new Promise(resolve => setTimeout(resolve, 1500));
        const dropdownSelectors = [
            "#patient-lookupUl1 li",
            "ul.dropdown-menu li",
            ".dropdown-menu li",
            "li[role='option']"
        ];
        
        let clicked = false;
        
        // Try to click the matching item in dropdown
        for (const sel of dropdownSelectors) {
            const items = await newPage.$$(sel);
            
            for (const item of items) {
                const text = await newPage.evaluate(el => el.textContent.trim(), item);
                if (text.toLowerCase().includes(accountNumber.toLowerCase())) {
                    const box = await item.boundingBox();
                    if (box) {
                        await newPage.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
                        await newPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { delay: 30 });
                        clicked = true;
                        patientSelected = true; // Mark as patient selected
                        break;
                    }
                }
            }
            
            if (clicked) break;
        }
        
        // Step 3: Fallback if not clicked
        if (!clicked && !patientSelected) {
            console.log('Dropdown item not found with bounding box. Using keyboard fallback.');
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
        
        console.log('Waiting for the patient selection dialog...');
        
        // Wait a moment for any dialog to appear
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Check for "No patient found" message
        const noPatientFound = await newPage.evaluate(() => {
            return document.body.textContent.includes('No patient found') || 
                   document.body.textContent.includes('No Patient Found');
        });
        
        if (noPatientFound && !patientSelected) {
            console.log('No patient found message detected');
            
            // Try multiple approaches to find and click the Cancel button
            try {
                const cancelButtonExists = await newPage.evaluate(() => {
                    const selectors = [
                        "button#createClaimBtn2",
                        "button.btn-lgrey:contains('Cancel')",
                        "button:contains('Cancel')",
                        ".modal-footer button[type='button']"
                    ];
                    
                    for (const selector of selectors) {
                        const elements = document.querySelectorAll(selector);
                        for (const element of elements) {
                            if (element.textContent.trim() === 'Cancel' || 
                                element.value === 'Cancel') {
                                element.click();
                                return true;
                            }
                        }
                    }
                    
                    const footerDiv = document.querySelector("#pn-createclaim > div > div > div.modal-footer.bd-mid-top.grey-bg.nomargin");
                    if (footerDiv) {
                        const buttons = footerDiv.querySelectorAll('button');
                        for (const button of buttons) {
                            if (button.textContent.trim() === 'Cancel') {
                                button.click();
                                return true;
                            }
                        }
                    }
                    
                    return false;
                });
                
                if (cancelButtonExists) {
                    console.log('Cancel button clicked via JavaScript');
                } else {
                    console.log('Trying to click on the modal footer area');
                    await newPage.waitForSelector("#pn-createclaim > div > div > div.modal-footer.bd-mid-top.grey-bg.nomargin", { 
                        visible: true, 
                        timeout: 5000 
                    });
                    
                    const footerBox = await newPage.$eval("#pn-createclaim > div > div > div.modal-footer.bd-mid-top.grey-bg.nomargin", 
                        el => {
                            const rect = el.getBoundingClientRect();
                            return { 
                                x: rect.x, 
                                y: rect.y, 
                                width: rect.width, 
                                height: rect.height 
                            };
                        }
                    );
                    
                    await newPage.mouse.click(
                        footerBox.x + footerBox.width * 0.25, 
                        footerBox.y + footerBox.height / 2,
                        { delay: 30 }
                    );
                    
                    console.log('Clicked on the left side of modal footer');
                }
                
                await new Promise(resolve => setTimeout(resolve, 1500));
                
            } catch (cancelError) {
                console.error('Error handling No Patient Found dialog:', cancelError);
            }
        } else {
            console.log('Patient found, skipping cancel');
        }
        
        // Proceed with the remaining steps
        console.log('Patient selection process completed');
         
   // Final Step - Click OK Button
        const buttonExists = await newPage.$('#createClaimBtn');
if (buttonExists) {
    console.log('Create Claim button found. Clicking...');
    await newPage.click('#createClaimBtn');
    
    // Wait for the modal to appear
    await newPage.waitForSelector('.modal-dialog', { timeout: 5000 });
    
    // Wait a bit for the dialog to fully render
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    try {
        // Try using evaluate to click the button via JavaScript
        await newPage.evaluate(() => {
            const yesButton = document.querySelector('#createClaimConfirmationBtn');
            if (yesButton) {
                yesButton.click();
                return true;
            }
            // Fallback to trying by value
            const buttonByValue = document.querySelector('input[value="Yes"]');
            if (buttonByValue) {
                buttonByValue.click();
                return true;
            }
            return false;
        });
        console.log('Clicked Yes button via JavaScript');
    } catch (error) {
        console.error('Error clicking Yes button:', error);
        
        // Alternative approach - try to use the ng-click function directly
        try {
            await newPage.evaluate(() => {
                if (typeof onclick_yes === 'function') {
                    onclick_yes();
                    return true;
                }
                return false;
            });
            console.log('Triggered onclick_yes() function');
        } catch (clickError) {
            console.error('Error triggering onclick function:', clickError);
        }
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
const billedFees = [];
const processedCPTs = new Map(); // Track processed CPT codes and their units from asterisk

// Extract CPT codes and their units (from asterisk values)
for (const code of cptCodes) {
  let baseCode, units = 1;
  
  if (code.includes('*')) {
    // Split by asterisk to get code and units
    const parts = code.split('*');
    baseCode = parts[0].split('-')[0].trim(); // Get base code (before any modifier)
    units = parseInt(parts[1].trim()) || 1;
  } else {
    baseCode = code.split('-')[0].trim();
    units = 1;
  }
  
  // Store the CPT code with its units (don't count occurrences, use asterisk value)
  if (!processedCPTs.has(baseCode)) {
    processedCPTs.set(baseCode, units);
  }
}

console.log('CPT codes with units from asterisk:', Object.fromEntries(processedCPTs));

// Process each unique CPT code
const uniqueCPTCodes = [...processedCPTs.keys()];

for (let i = 0; i < uniqueCPTCodes.length; i++) {
  const baseCode = uniqueCPTCodes[i];
  let modifier = "";
  
  // Get the original code with modifier if present
  const originalCode = cptCodes.find(code => {
    const codeBase = code.includes('*') ? code.split('*')[0] : code;
    return codeBase.split('-')[0].trim() === baseCode;
  });
  
  if (originalCode) {
    const codeWithoutUnits = originalCode.includes('*') ? originalCode.split('*')[0] : originalCode;
    if (codeWithoutUnits.includes('-')) {
      modifier = codeWithoutUnits.split('-')[1].trim();
    }
  }
  
  // Get units from the asterisk value for this CPT code
  const units = processedCPTs.get(baseCode);
  
  try {
    console.log(`🔄 Processing CPT code ${i + 1} of ${uniqueCPTCodes.length}: ${baseCode} (units: ${units})`);
    
    // Click on the CPT cell to begin entry
    const cptCellSelector = 'td[ng-click*="enterCPT"]';
    await newPage.waitForSelector(cptCellSelector, { visible: true, timeout: 8000 });
    await newPage.click(cptCellSelector);
    await new Promise(resolve => setTimeout(resolve, 600));
    
    // Enter the CPT code
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
    
      try {
        console.log('Waiting for CPT confirmation alert...');
        const confirmYesSelector = 'body > div.bootbox.modal.fade.bluetheme.medium-width.in > div > div > div.modal-footer > button.btn.btn.btn-lblue.btn-lgrey.btn-xs.btn-default.btn-yes';
        await newPage.waitForSelector(confirmYesSelector, { visible: true, timeout: 8000 });
        await new Promise(resolve => setTimeout(resolve, 1000));
        console.log('Clicking "Yes" on CPT confirmation...');
        await newPage.click(confirmYesSelector);
      } catch (modalErr) {
        console.log('No CPT confirmation modal appeared or failed to click:', modalErr.message);
      }
      
      // Check and set POS (Place of Service) and apply modifiers if needed
      const posSelector = '#billingClaimIpt24';
      await newPage.waitForSelector(posSelector, { visible: true, timeout: 8000 });
      const posValue = await newPage.$eval(posSelector, el => el.value.trim());
      console.log(`POS Value: ${posValue}`);
      
      if (posValue !== "11") {
        const has99291 = cptCodes.some(c => c.includes("99291"));
        const has99292 = cptCodes.some(c => c.includes("99292"));
        
        if ((baseCode === "99291" || baseCode === "99292") && has99291 && has99292) {
          modifier = "25";
          console.log(`Applying rule: Both 99291 & 99292 exist — modifier 25`);
        } else if (baseCode === "99291" && uniqueCPTCodes.length > i + 1 && uniqueCPTCodes[i + 1] !== "99292") {
          modifier = "25";
          console.log(`Applying rule: 99291 followed by non-99292 — modifier 25`);
        } else if (baseCode === "99291") {
          modifier = "";
          console.log(`Applying rule: Solo 99291 — no modifier`);
        }
        
        if (modifier) {
          const modifierSelector = `#billingClaimIpt8ngR${i}`;
          await newPage.waitForSelector(modifierSelector, { visible: true, timeout: 5000 });
          
          console.log(`Using modifier field selector for row ${i}: ${modifierSelector}`);
          
          const modifierInput = await newPage.$(modifierSelector);
          await modifierInput.focus();
          await newPage.evaluate(el => el.scrollIntoView(), modifierInput);
          await modifierInput.click({ clickCount: 3 });
          await modifierInput.press('Backspace');
          await new Promise(resolve => setTimeout(resolve, 200));
          await modifierInput.type(modifier, { delay: 100 });
          console.log(`✅ Entered modifier: ${modifier} in row ${i}`);
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      } else {
        console.log(`POS is ${posValue} — skipping modifiers`);
      }
    
    // Set units from asterisk value if more than 1
    if (units > 1) {
      try {
        // Find the units input field for the current row
        const unitsSelector = `#billingClaimIpt16ngR${i}`;
        await newPage.waitForSelector(unitsSelector, { visible: true, timeout: 5000 });
        
        const unitsInput = await newPage.$(unitsSelector);
        await unitsInput.focus();
        await newPage.evaluate(el => el.scrollIntoView(), unitsInput);
        await unitsInput.click({ clickCount: 3 });
        await unitsInput.press('Backspace');
        await new Promise(resolve => setTimeout(resolve, 200));
        await unitsInput.type(units.toString(), { delay: 100 });
        console.log(`✅ Updated units to ${units} for CPT ${baseCode} (from asterisk value)`);
        
        // Click the OK button to confirm the update
        const okButtonSelector = '#CPTUpdateOkBtn';
        await newPage.waitForSelector(okButtonSelector, { visible: true, timeout: 5000 });
        await newPage.click(okButtonSelector);
        console.log('Clicked OK button to confirm units update');
      } catch (unitsErr) {
        console.error(`⚠️ Failed to update units for CPT ${baseCode}:`, unitsErr.message);
      }
    }
    
      // Check billed fee after entering CPT and modifier
      try {
        const billedFeeSelector = `#billingClaimIpt17ngR${i}`;
        await newPage.waitForSelector(billedFeeSelector, { visible: true, timeout: 8000 });
        
        // Get the value from the input field directly
        const billedFeeValue = await newPage.$eval(billedFeeSelector, el => {
          // If it's an input element, get its value
          if (el.tagName === 'INPUT') {
            return el.value.trim();
          } 
          // Otherwise get the text content
          return el.textContent.trim();
        });
        
        // Extract numeric value, handling currency symbols and formatting
        const numericValue = billedFeeValue.replace(/[^0-9.]/g, '');
        const billedFee = parseFloat(numericValue);
        
        // Use isNaN to check if we got a valid number
        const validBilledFee = !isNaN(billedFee) ? billedFee : 0;
        
        // For multiple units, calculate the per-unit fee
        const perUnitFee = units > 1 ? validBilledFee / units : validBilledFee;
        
        console.log(`💲 Billed Fee for CPT ${baseCode} (Row ${i}): ${validBilledFee} (${perUnitFee} per unit)`);
        
        billedFees.push({ 
          code: baseCode, 
          units: units,
          fee: validBilledFee,
          perUnitFee: perUnitFee,
          originalText: billedFeeValue 
        });
      } catch (feeErr) {
        console.error(`⚠️ Failed to check billed fee for CPT ${baseCode}:`, feeErr.message);
      }
      
      console.log(`✅ Done processing CPT: ${baseCode}`);
      
    } catch (error) {
      console.error(`❌ Failed to process CPT code ${baseCode}:`, error);
      await newPage.screenshot({ path: `cpt_error_${baseCode}.png` });
    }
  
    await new Promise(resolve => setTimeout(resolve, 1000));
  }
  
  // After processing all CPT codes, check if we need to handle dates
  // We'll check POS value again and if needed, enter the dates once
  try {
    // Check POS value from the page (get it after all CPTs are processed)
    const posSelector = '#billingClaimIpt24';
    await newPage.waitForSelector(posSelector, { visible: true, timeout: 8000 });
    const posValue = await newPage.$eval(posSelector, el => el.value.trim());
    
    console.log(`Final POS check: ${posValue}`);
    
    // Handle date entry for POS values 21 or 31 after all CPTs are processed
    if (["21", "31"].includes(posValue)) {
      console.log(`POS is ${posValue} — proceeding to date input.`);
      
      // Find and click the data button to open the date modal
      const dataBtnSelector = 'button[id^="claimDataBtn"]';
      const dataBtnHandle = await newPage.$(dataBtnSelector);
      if (dataBtnHandle) {
        const btnId = await newPage.evaluate(el => el.id, dataBtnHandle);
        console.log(`Found data button with ID: ${btnId}`);
        await dataBtnHandle.click();
        console.log(`Clicked data button`);
        await new Promise(resolve => setTimeout(resolve, 1500)); // Longer wait for modal to open fully
      } else {
        console.warn(`Data button with selector ${dataBtnSelector} not found`);
        throw new Error("Data button not found");
      }
      
      // Wait for admit date input to be visible
      const admitDateInputSelector = '#dtHCFAFrom16From';
      const dischargeInputSelector = '#dtHCFAFrom16To';
      await newPage.waitForSelector(admitDateInputSelector, { visible: true, timeout: 5000 });
      
      // Helper function to format date with leading zeros (MM/DD/YYYY)
      const formatDateWithLeadingZeros = (dateStr) => {
        const date = new Date(dateStr);
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        const year = date.getFullYear();
        return `${month}/${day}/${year}`;     
      };
      
      const formattedBillingDate = formatDateWithLeadingZeros(billingDate);
      
// After calculation of admit date, add this validation check before entering dates:

// Determine admit date based on CPT codes present
let admitDateOnly;
let dischargeDateToUse = formattedBillingDate;
      
// Check if we have any special CPT codes that affect date handling
const has99222or99223 = cptCodes.some(code => {
  const baseCode = code.split('-')[0].trim();
  return ["99222", "99223"].includes(baseCode);
});
      
// Check if we have discharge CPT codes
const has99238or99239 = cptCodes.some(code => {
  const baseCode = code.split('-')[0].trim();
  return ["99238", "99239"].includes(baseCode);
});
      
// Special case: if we have 99222 or 99223, use billing date as admit date
if (has99222or99223) {
  console.log(`🔍 Special CPT codes 99222/99223 detected - using billing date as admit date`);
  admitDateOnly = formattedBillingDate;
} else {
  // For other codes, process normally
  // Clean admit date if present in expected format
  admitDateOnly = admitDate.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)?.[0] || "";
  
  // Format with leading zeros if date is present
  if (admitDateOnly) {
    admitDateOnly = formatDateWithLeadingZeros(admitDateOnly);
  }
  
  // If admit date is NOT provided, calculate it from discharge date
  if (!admitDateOnly) {
    console.log(`❗ Admit Date not provided. Calculating from Discharge Date: ${formattedBillingDate}`);
    
    // Calculate admit date as 1 day before discharge date
    const admit = new Date(billingDate);
    admit.setDate(admit.getDate() - 1);
    
    // Format calculated admit date with leading zeros
    const month = String(admit.getMonth() + 1).padStart(2, '0');
    const day = String(admit.getDate()).padStart(2, '0');
    const year = admit.getFullYear();
    admitDateOnly = `${month}/${day}/${year}`;
    
    console.log(`Calculated Admit Date: ${admitDateOnly}`);
  } else {
    console.log(`✅ Using provided Admit Date: ${admitDateOnly}`);
  } 
}
// Only check for same dates when dealing with discharge codes
if (has99238or99239 && admitDateOnly === dischargeDateToUse) {
  console.log(`⚠️ Admit date and discharge date are the same for discharge codes: ${admitDateOnly}`);
  
  // Calculate new admit date as 1 day before discharge date
  const admit = new Date(dischargeDateToUse);
  admit.setDate(admit.getDate() - 1);
  
  // Format calculated admit date with leading zeros
  const month = String(admit.getMonth() + 1).padStart(2, '0');
  const day = String(admit.getDate()).padStart(2, '0');
  const year = admit.getFullYear();
  admitDateOnly = `${month}/${day}/${year}`;
  
  console.log(`Updated Admit Date to be one day before discharge: ${admitDateOnly}`);
}

  
// ADD THIS NEW VALIDATION: Check if admit date is greater than billing date
const admitDateObj = new Date(admitDateOnly);
const billingDateObj = new Date(billingDate);

if (admitDateObj > billingDateObj) {
  console.log(`⚠️ Admit date (${admitDateOnly}) is greater than billing date (${formattedBillingDate})`);
  
  // Set admit date to one day before billing date
  const newAdmitDate = new Date(billingDate);
  newAdmitDate.setDate(newAdmitDate.getDate() - 1);
  
  // Format the new admit date
  const month = String(newAdmitDate.getMonth() + 1).padStart(2, '0');
  const day = String(newAdmitDate.getDate()).padStart(2, '0');
  const year = newAdmitDate.getFullYear();
  admitDateOnly = `${month}/${day}/${year}`;
  
  console.log(`Corrected admit date to one day before billing date: ${admitDateOnly}`);
}

// Now enter the validated admit date
await newPage.type(admitDateInputSelector, admitDateOnly);
console.log(`Entered admit date: ${admitDateOnly}`);
await new Promise(resolve => setTimeout(resolve, 500));

// Enter discharge date ONLY if we have discharge CPT codes
if (has99238or99239) {
  await newPage.type(dischargeInputSelector, dischargeDateToUse);
  await newPage.keyboard.press('Tab'); // Move focus away 
  console.log(`Entered discharge date: ${dischargeDateToUse}`);
  await new Promise(resolve => setTimeout(resolve, 500));
} else {
  console.log(`No discharge codes (99238/99239) present - skipping discharge date entry`);
}

// Add extra wait time before attempting to click OK
await new Promise(resolve => setTimeout(resolve, 1500));

// Try multiple click strategies for the OK button
const okBtnSelector = '#claimDataBtn13';
const bootboxOkSelector = 'body > div.bootbox.modal.fade.bluetheme.medium-width.bootbox-alert.in > div > div > div.modal-footer > button';

try {
  let modalStillOpen = true;
  let attempts = 0;

  while (modalStillOpen && attempts < 3) {
    attempts++;

    // Check if OK button is visible
    await newPage.waitForSelector(okBtnSelector, { visible: true, timeout: 5000 });
    console.log(`Attempt ${attempts}: OK button is visible, trying click...`);

    // Try direct click
    await newPage.click(okBtnSelector);
    await new Promise(resolve => setTimeout(resolve, 1000));

    // Check for bootbox alert
    const bootboxExists = await newPage.$(bootboxOkSelector);
    if (bootboxExists) {
      console.log("Bootbox alert detected. Clicking OK...");
      await newPage.click(bootboxOkSelector);
      await new Promise(resolve => setTimeout(resolve, 1000));

      // Re-enter admit date
      await newPage.click(admitDateInputSelector, { clickCount: 3 });
      await newPage.type(admitDateInputSelector, admitDateOnly);
      console.log(`Re-entered admit date: ${admitDateOnly}`);
      await new Promise(resolve => setTimeout(resolve, 500));

      // Re-enter discharge date if needed
      if (has99238or99239) {
        await newPage.click(dischargeInputSelector, { clickCount: 3 });
        await newPage.type(dischargeInputSelector, dischargeDateToUse);
        await newPage.keyboard.press('Tab');
        console.log(`Re-entered discharge date: ${dischargeDateToUse}`);
        await new Promise(resolve => setTimeout(resolve, 500));
      }

      continue; // Retry clicking OK
    }

    // Check if modal is still open
    modalStillOpen = await newPage.evaluate(() => {
      return !!document.querySelector('.modal-dialog:not([style*="display: none"])');
    });

    if (!modalStillOpen) {
      console.log(`✅ Modal closed successfully after ${attempts} attempt(s)`);
      break;
    } else {
      console.log("Modal still open, retrying...");
    }
  }
} catch (error) {
  console.error(`Failed to click OK button or handle modal:`, error);
  await newPage.screenshot({ path: `ok_button_error_date.png` });
}

    } else {
      console.log(`POS is ${posValue} — no date entry needed.`);
    }
    
    console.log(`✅ Completed processing all CPT codes and dates.`);
  } catch (error) {
    console.error(`❌ Failed during date processing:`, error);
    await newPage.screenshot({ path: `date_processing_error.png` });
  }
const claimNoSelector = '#billingClaimTbl2 > tbody > tr:nth-child(1) > td:nth-child(2) > span';

        await newPage.waitForSelector(claimNoSelector, { visible: true, timeout: 5000 });
        
        const claimNumber = await newPage.$eval(claimNoSelector, el => el.textContent.trim());
        
        console.log(`🧾 Claim Number: ${claimNumber}`);
        try {
            // Step 1: Find the claim status dropdown
            const statusSelector = 'select[id^="claimStatusSel"]';
            await newPage.waitForSelector(statusSelector, { visible: true, timeout: 5000 });
            const statusDropdownHandle = await newPage.$(statusSelector);
          
            if (statusDropdownHandle) {
              const selectedText = await newPage.evaluate(el => {
                const selectedOption = el.options[el.selectedIndex];
                return selectedOption ? selectedOption.textContent.trim() : '';
              }, statusDropdownHandle);  
          
              console.log(`Current Claim Status: ${selectedText}`);
          
              // Step 2: If not pending, set to Pending
              if (selectedText.toLowerCase() !== "pending") {
                await newPage.select(statusSelector, "PEN");
                console.log("Changed Claim Status to Pending");
                await new Promise(resolve => setTimeout(resolve, 500)); // wait for change event
              }
              await new Promise(resolve => setTimeout(resolve, 1000));

// Step 3: Click the first OK button
try {
  const okSelector = 'button[id^="claimScreenOkBtn"]';
  await newPage.waitForSelector(okSelector, { visible: true, timeout: 5000 });
  
  const okButtonHandle = await newPage.$(okSelector);
  if (okButtonHandle) {
    await okButtonHandle.evaluate(btn => btn.scrollIntoView({ behavior: "smooth", block: "center" }));
    await new Promise(resolve => setTimeout(resolve, 300));
    
    // Use bounding box + mouse click for higher reliability
    const box = await okButtonHandle.boundingBox();
    if (box) {
      await newPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
      console.log("Clicked first OK button.");
    } else {
      console.warn("First OK button not clickable (no bounding box).");
    }
    
    // Wait for error popup to appear after clicking first OK
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    // Step 4: Now handle the error popup that appears - click the second OK button
    try {
      // Error popup OK button selector - from modal-footer
      const errorOkSelector = 'div.bootbox.modal.fade.bluetheme.medium-width.in button.btn.btn-blue.btn-sm.btn-xs, button[data-bb-handler="Yes"]';
      await newPage.waitForSelector(errorOkSelector, { visible: true, timeout: 5000 });
      
      const errorOkButton = await newPage.$(errorOkSelector);
      if (errorOkButton) {
        await errorOkButton.evaluate(btn => btn.scrollIntoView({ behavior: "smooth", block: "center" }));
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // Click the error popup's OK button
        const errorOkBox = await errorOkButton.boundingBox();
        if (errorOkBox) {
          await newPage.mouse.click(errorOkBox.x + errorOkBox.width / 2, errorOkBox.y + errorOkBox.height / 2);
          console.log("Clicked error popup OK button.");
        } else {
          console.warn("Error popup OK button not clickable (no bounding box).");
          await errorOkButton.click();
        }
        
        // Wait for UI to update after clicking error OK button
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Step 5: Finally click the Cancel button
        const cancelSelector = '#billingClaimBtn36';
        await newPage.waitForSelector(cancelSelector, { visible: true, timeout: 5000 });
        
        const cancelButton = await newPage.$(cancelSelector);
        if (cancelButton) {
          await cancelButton.evaluate(btn => btn.scrollIntoView({ behavior: "smooth", block: "center" }));
          await new Promise(resolve => setTimeout(resolve, 300));
          
          // Click the Cancel button
          const cancelBox = await cancelButton.boundingBox();
          if (cancelBox) {
            await newPage.mouse.click(cancelBox.x + cancelBox.width / 2, cancelBox.y + cancelBox.height / 2);
            console.log("Clicked Cancel button after handling error.");
          } else {
            console.warn("Cancel button not clickable (no bounding box).");
            await cancelButton.click();
          }
          
          await new Promise(resolve => setTimeout(resolve, 1000));
        } else {
          console.warn("Cancel button not found after error handling.");
        }
      } else {
        console.warn("Error popup OK button not found.");
      }
    } catch (errorPopupError) {
      console.error(`Error handling popup: ${errorPopupError.message}`);
      
      // Fallback: try to find any visible OK button in a modal
      try {
        const anyOkSelector = 'div.modal.in button:contains("OK"), div.modal.in button.btn-blue';
        const anyOkButton = await newPage.$(anyOkSelector);
        if (anyOkButton) {
          await anyOkButton.click();
          console.log("Clicked fallback OK button in error popup.");
          
          // Still try to click Cancel after fallback
          await new Promise(resolve => setTimeout(resolve, 1000));
          await newPage.click('#billingClaimBtn36').catch(e => console.warn("Cancel button click failed after fallback."));
        }
      } catch (fallbackError) {
        console.error(`Fallback error handling failed: ${fallbackError.message}`);
      }
    }
  } else {
    console.warn("First OK button not found.");
  }
} catch (error) {
  console.error(`Error in claim status handling sequence: ${error.message}`);
}
            }
          } catch (error) {
            console.error(`Error handling claim status: ${error.message}`);
          }
          
     
            return { success: true, claimNumber, billedFees: [] };
        } else {
            return { success: false, error: 'Create Claim button not found' };
        }
    } catch (error) {
        console.error(`Error creating claim again: ${error.message}`);
        return { success: false, error: error.message };
    }
}
async function handleModalAndProceed(page) {
   try {
        // Try using evaluate to click the button via JavaScript
        await newPage.evaluate(() => {
            const yesButton = document.querySelector('#createClaimConfirmationBtn');
            if (yesButton) {
                yesButton.click();
                return true;
            }
            // Fallback to trying by value
            const buttonByValue = document.querySelector('input[value="Yes"]');
            if (buttonByValue) {
                buttonByValue.click();
                return true;
            }
            return false;
        });
        console.log('Clicked Yes button via JavaScript');
    } catch (error) {
        console.error('Error clicking Yes button:', error);
        
        // Alternative approach - try to use the ng-click function directly
        try {
            await newPage.evaluate(() => {
                if (typeof onclick_yes === 'function') {
                    onclick_yes();
                    return true;
                }
                return false;
            });
            console.log('Triggered onclick_yes() function');
        } catch (clickError) {
            console.error('Error triggering onclick function:', clickError);
        }
    }
}

async function handleModalAndCancel(page) {
  let extractedProvider = null;
  
  try {
    // 1) WAIT FOR "No" BUTTON IN THE MODAL, THEN CLICK IT
    const noButtonSelector = '#createClaimIpt6';
    await page.waitForSelector(noButtonSelector, { visible: true, timeout: 5000 });
    await page.click(noButtonSelector);
    console.log('🟡 Clicked "No" in modal');

  } catch (err) {
    console.warn('⚠️ Could not find or click "No" button:', err.message);
    return extractedProvider;
  }

  // 2) WAIT FOR "SERVICING PROVIDER" FIELD (⚠️ CORRECT SELECTOR "billingClaimIpt25")
  try {
    // 2) Extract provider from the field before cancel
    const servicingProviderSelector = '#billingClaimIpt25';
    await page.waitForSelector(servicingProviderSelector, { visible: true, timeout: 10000 });
    extractedProvider = await page.$eval(servicingProviderSelector, el => el.value);
    console.log(`🔎 Extracted provider before cancel: "${extractedProvider}"`);
  } catch (err) {
    console.warn('⚠️ Could not extract provider:', err.message);
  }

  // 3) WAIT FOR "CANCEL" BUTTON AND CLICK IT
  try {
    const cancelSelector = '#billingClaimBtn36';
    await page.waitForSelector(cancelSelector, { visible: true, timeout: 5000 });

    const cancelButton = await page.$(cancelSelector);
    if (cancelButton) {
      // Scroll into view
      await cancelButton.evaluate(el =>
        el.scrollIntoView({ behavior: 'smooth', block: 'center' })
      );

      const box = await cancelButton.boundingBox();
      if (box) {
        await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
        console.log('🟢 Clicked Cancel button');
      } else {
        console.warn('⚠️ Cancel button bounding box not found – using fallback click()');
        await cancelButton.click();
      }
    } else {
      console.warn('❗ Cancel button (#billingClaimBtn36) not found');
    }
  } catch (err) {
    console.warn('⚠️ Error while clicking "Cancel": ', err.message);
  }
  
  return extractedProvider;
}
} catch (error) {
  console.error('❌ Error processing account:', error.message);
  return { success: false, claimNumber: '', billedFees: [] };
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
        
        const requiredColumns = ["ECW #", "Provider", "Hospital Location", "Billing Date"];
        for (const column of requiredColumns) {
            if (!headers.includes(column)) {
                throw new Error(`Required column '${column}' is missing from the spreadsheet.`);
            }
        }
    
        // Add missing headers if needed
        const extraHeaders = ["Claim Number", "Billed Fee", "Result"];
        for (const header of extraHeaders) {
            if (!headers.includes(header)) {
                columnMap[header] = headers.length;
                headers.push(header);
                for (let i = 1; i < originalData.length; i++) {
                    originalData[i] = originalData[i] || [];
                    originalData[i].push("");
                }
            }
        }
    
        headers.forEach((header, idx) => columnMap[header] = idx);
    
        let newRowsToProcess = 0;
        const processedAccounts = new Set();
        
        let missingCallDiagnoses = 0;  // Track rows missing Diagnoses

        for (let i = 1; i < originalData.length; i++) {
            const row = originalData[i];
            if (!row) continue; // Skip undefined rows
            
            const accountNumberRaw = row[columnMap["ECW #"]] || "";
            const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
            if (!accountNumberRaw || status === "done" || status === "failed") continue;  // Skip processed or failed rows

            
            const accountNumber = accountNumberRaw.split("-")[0].trim();
            processedAccounts.add(accountNumber);
            newRowsToProcess++;
        
            // Check if Diagnoses is missing
            const callDiagnoses = row[columnMap["Diagnoses"]] || "";
            if (!callDiagnoses.trim()) {
                missingCallDiagnoses++;
            }
        }
        
        // Check if all remaining rows are missing Diagnoses
        if (newRowsToProcess > 0 && newRowsToProcess === missingCallDiagnoses) {
            return res.json({
                success: false,
                message: "Diagnoses not found in Excel.",
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
            const facilityName = row[columnMap["Hospital Location"]] || "";
            const billingDate = row[columnMap["Billing Date"]] || "";
            const diagnosisText = row[columnMap["Diagnoses"]] || "";
            const chargesText = row[columnMap["Charges"]] || ""; // CPT codes
            const admitDate = row[columnMap["Admit Date"]] || ""; // ✅ NEW LINE


            const accountNumber = accountNumberRaw.split("-")[0].trim();
            console.log(`Processing account: ${accountNumber}`);
        
            // Validate data before processing
            if (!provider) console.log('Provider value is missing or undefined');
            if (!facilityName) console.log('Hospital Location is missing or undefined');
            if (!billingDate) console.log('Billing Date is missing or undefined');
            if (!admitDate) console.log('Admit Date is missing or undefined'); // ✅ Optional

            // Extract diagnosis codes from the text inside parentheses
            const diagnosisCodes = [];

            // Regex for ICD-10 codes and procedure codes:
            // - ICD-10: Starts with a letter, 2–4 digits, optional dot, 1–4 alphanums
            // - ICD-10-PCS (procedure): 7 alphanumeric characters
            const dxCodeRegex = /\b(?:[A-Z][0-9]{2,4}(?:\.[0-9A-Z]{1,4})?|[A-Z0-9]{7})\b/g;
          
            // Match all valid codes
            const matches = diagnosisText.match(dxCodeRegex);
          
            if (matches) {
              // Remove duplicates and trim spaces
              const uniqueCodes = [...new Set(matches.map(code => code.trim()))];
              diagnosisCodes.push(...uniqueCodes);
            }
          
            // Filter: Only alphanumeric codes (must contain at least one letter & one digit)
            const alphanumericCodes = diagnosisCodes.filter(code => /[A-Z]/i.test(code) && /\d/.test(code));
          
            // Output
            console.log("Diagnosis Codes:", diagnosisCodes);
            console.log("Total Codes:", diagnosisCodes.length);
            console.log("Alphanumeric Codes:", alphanumericCodes.length);
            alphanumericCodes.forEach((code, index) => {
              console.log(`  Code ${index + 1}: ${code}`);
            });
          
          
          
            console.log(`Charges Text for ${accountNumber}:`, chargesText);

const cptCodes = [];
const cptRegex = /\b[A-Za-z]?\d{4,5}(?:-[A-Za-z0-9]{2,4})?(?:\*\d+)?\b/g;
let cptMatch;

while ((cptMatch = cptRegex.exec(chargesText)) !== null) {
    const fullCode = cptMatch[0].trim();
    cptCodes.push(fullCode);

    // Parse the full code to extract base, modifier, and units
    let baseCode, modifier, units;
    
    if (fullCode.includes('*')) {
        // Has asterisk with units
        const [codeWithModifier, unitsStr] = fullCode.split('*');
        units = unitsStr;
        
        if (codeWithModifier.includes('-')) {
            [baseCode, modifier] = codeWithModifier.split('-');
        } else {
            baseCode = codeWithModifier;
            modifier = 'None';
        }
    } else {
        // No asterisk
        units = '1'; // default
        
        if (fullCode.includes('-')) {
            [baseCode, modifier] = fullCode.split('-');
        } else {
            baseCode = fullCode;
            modifier = 'None';
        }
    }
    
    console.log(`🧩 Base CPT: ${baseCode}, Modifier: ${modifier}, Units: ${units}`);
}

console.log('Extracted CPT Codes with asterisk units:', cptCodes);

            // === Call processAccount (Don't send CPT length)
            const { success, claimNumber ,billedFees} = await processAccount(
                newPage,
                accountNumber,
                provider,
                facilityName,
                billingDate,
                diagnosisCodes,
                cptCodes, // ✅ only the array
                admitDate // ✅ Pass Admit Date here

            );
        
         // Set Result
            if (success) {
                results.successful.push(accountNumber);
                row[columnMap["Result"]] = "Done";
            } else {
                results.failed.push(accountNumber);
                row[columnMap["Result"]] = "Failed";
            }

            // Always write Claim Number if available
            row[columnMap["Claim Number"]] = claimNumber ? claimNumber.toString() : "";

            // Write Billed Fee: only "Billed Fee is zero" if any fee is zero, else leave blank
            if (Array.isArray(billedFees)) {
                const anyFeeZero = billedFees.some(f => typeof f.fee === 'number' && f.fee === 0);
                row[columnMap["Billed Fee"]] = anyFeeZero ? "Billed Fee is zero" : "";
            } else {
                row[columnMap["Billed Fee"]] = ""; // Write blank if no valid array
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
                fgColor: { argb: "A9A9A9" }
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

        const columnWidths = [6, 10, 10, 10, 15, 16, 16, 8, 40,15, 11, 16, 50, 30,13];

        // When setting up columns, ensure EPS Comments has wrap text enabled
        sheet.columns = columnWidths.map((width, index) => ({
            width,
            style: (index === columnMap["Diagnoses"] || index === columnMap["Charges"])                                
                ? { alignment: { wrapText: true } } 
                : {},
        }));
        
        // When formatting cells in each row, add special handling for EPS Comments
        originalData.slice(1).forEach((row) => {
            if (!row) return; // Skip undefined rows
            
            const hasData = row.some(cell => cell && cell.toString().trim() !== "");
            if (!hasData) return; // Skip empty rows
            
            const rowData = headers.map(header => {
                if (header === "Claim Number") {
                    return row[columnMap["Claim Number"]] || "";  // Use the value of Claim Number
                }
                return row[columnMap[header]] || "";  // Use existing data for other columns
            });
            
            const newRow = sheet.addRow(rowData);
            newRow.height = 100;
            newRow.getCell(columnMap["Claim Number"] + 1).alignment = { horizontal: "center", vertical: "middle" };

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
                        fgColor: { argb: "FFFFFF" } 
                    };
                    cell.alignment = { horizontal: "center", vertical: "bottom" };
                }
                if (header === "Claim Number" && cell.value) {
                    cell.fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: "FFFFFF" } 
                    };
                    cell.alignment = { horizontal: "center", vertical: "bottom" };
                }
                if (header === "Billed Fee" && cell.value) {
                    // Set the column width to 20
                    sheet.getColumn(columnMap["Billed Fee"] + 1).width = 20;
                
                    // Set text color to red
                    cell.font = {
                        color: { argb: "FF0000" }, // Red text
                    };
                
                    // Align the text
                    cell.alignment = { horizontal: "center", vertical: "bottom" };
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
            message: `${results.successful.length} Charges Entered successfully. ${results.failed.length} rows failed.`,
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