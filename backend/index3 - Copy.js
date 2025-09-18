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



// async function continueWithLoggedInSession() {
//     try {
//         const browser = await puppeteer.connect({
//             browserURL: 'http://localhost:9222', // Connect to existing Chrome
//             defaultViewport: null,
//             headless: false,
//             timeout: 120000,
//         });

//         const pages = await browser.pages();
//         const sessionPage = pages.find(p =>
//             p.url().includes('Appointments/ViewAppointments.aspx')
//         );

//         if (!sessionPage) {
//             console.log('❌ No matching tab found for the appointment page.');
//             return null;
//         }

//         console.log('✅ Found existing page. Waiting for Patient Visit tab...');
//         await sessionPage.waitForSelector('#patient-visits_tab > span', {
//             visible: true,
//             timeout: 10000,
//         });

//         console.log('🖱️ Clicking Patient Visit tab...');
//         await sessionPage.evaluate(() => {
//             const tab = document.querySelector('#patient-visits_tab > span');
//             if (tab) {
//                 tab.scrollIntoView({ behavior: 'smooth', block: 'center' });
//                 tab.click();
//             }
//         });

//         console.log('✅ Patient Visit tab clicked successfully.');
//         return sessionPage;
//     } catch (error) {
//         console.error('🚨 Error in continueWithLoggedInSession:', error.message);
//         return null;
//     }
// }
// async function processAccount(page,browser,accountNumber, provider, facilityName, billingDate, diagnosisCodes,cptCodes,admitDate,pos,modifier) {
//    try {
//       // Wait and click Add New Visit
// // Wait and click Add New Visit
// await new Promise(resolve => setTimeout(resolve, 3000));
// console.log('⏳ Waiting for Add New Visit button to appear...');
// await page.waitForSelector('#addNewVisit', { visible: true, timeout: 15000 });

// console.log('🖱️ Setting up popup listener...');

// // Set up popup listener BEFORE clicking the button
// const popupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Popup detected!');
//         browser.off('targetcreated', onPopup); // Fixed: Use .off() instead of .removeListener()
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// console.log('🖱️ Clicking Add New Visit button...');
// await page.evaluate(() => {
//     const btn = document.querySelector('#addNewVisit');
//     if (btn) {
//         btn.scrollIntoView({ behavior: 'smooth', block: 'center' });  
//         btn.click();
//     } else {
//         console.error('❌ Add New Visit button not found!');
//     }
// });
//  // ⏳ Wait for popup to load
//         console.log('⏳ Waiting for Search button to appear...');
//         await page.waitForSelector('#ctl00_phFolderContent_Button1', { visible: true, timeout: 6000 });

//         // 🔘 Click initial popup search button
//         console.log('🖱️ Clicking popup search open button...');
//         await page.click('#ctl00_phFolderContent_Button1');
// // Wait for popup to appear
// console.log('⏳ Waiting for popup to load...');
// let newPage;
// try {
//     // Wait for popup with timeout
//     newPage = await Promise.race([
//         popupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Popup timeout')), 6000))
//     ]);
//     console.log('✅ Popup loaded successfully');
// } catch (error) {
//     console.log('⚠️ No popup detected, checking for new tabs...');
    
//     // Fallback: Check for new tabs
//     await new Promise(resolve => setTimeout(resolve, 2000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//         newPage = allPages[allPages.length - 1]; // Get the newest page
//         console.log('✅ New tab detected, using it...');
//     } else {
//         newPage = page; // Use current page if no new tab
//         console.log('📝 No new tab, using current page...');
//     }
// }

// // Wait for the new page to load completely
// await new Promise(resolve => setTimeout(resolve, 3000));

// // ⏳ Wait for Search button to appear in the new page
// console.log('⏳ Waiting for Search button to appear in new page...');
// try {
//     await newPage.waitForSelector('#ctl00_phFolderContent_Button1', { visible: true, timeout: 6000 });
// } catch (error) {
//     // Try alternative selectors
//     console.log('🔍 Trying alternative search button selectors...');
//     const searchButtonSelectors = [
//         'input[value="Search"]',
//         'button[onclick*="Search"]',
//         '.button:contains("Search")',
//         '#Button1'
//     ];
    
//     let found = false;
//     for (const selector of searchButtonSelectors) {
//         try {
//             await newPage.waitForSelector(selector, { visible: true, timeout: 5000 });
//             console.log(`✅ Found search button with selector: ${selector}`);
//             found = true;
//             break;
//         } catch (e) {
//             continue;
//         }
//     }
    
//     if (!found) {
//         throw new Error('Search button not found with any selector');
//     }
// }

// // 🔘 Click initial popup search button
// console.log('🖱️ Clicking popup search open button...');
// try {
//     await newPage.click('#ctl00_phFolderContent_Button1');
// } catch (error) {
//     // Try clicking any search button
//     await newPage.evaluate(() => {
//         const searchButtons = document.querySelectorAll('input[value*="Search"], button');
//         for (const btn of searchButtons) {
//             if (btn.value?.includes('Search') || btn.textContent?.includes('Search')) {
//                 btn.click();
//                 break;
//             }
//         }
//     });
// }

// // Wait for the dropdown to load and be visible
// console.log('⏳ Waiting for dropdown to appear...');
// try {
//     await newPage.waitForSelector('#ctl04_popupBase_ddlSearch', { visible: true, timeout: 6000 });
// } catch (error) {
//     console.log('🔍 Dropdown not found, trying alternative selectors...');
//     // Try to find any dropdown
//     const dropdownSelectors = [
//         'select[name*="ddlSearch"]',
//         'select[id*="Search"]',
//         'select:first-of-type',
//         '#ctl04_popupBase_ddlSearch'
//     ];
    
//     for (const selector of dropdownSelectors) {
//         try {
//             await newPage.waitForSelector(selector, { visible: true, timeout: 5000 });
//             console.log(`✅ Found dropdown with selector: ${selector}`);
//             break;
//         } catch (e) {
//             continue;
//         }
//     }
// }

// await new Promise(resolve => setTimeout(resolve, 1000)); // Extra wait for dropdown to fully load

// // 🔽 Select "Patient ID" (using index 2 for 3rd option, 0-based indexing)
// console.log('🔽 Selecting "Patient ID" from dropdown...');
// await newPage.evaluate(() => {
//     const dropdown = document.querySelector('#ctl04_popupBase_ddlSearch') || 
//                     document.querySelector('select[name*="ddlSearch"]') ||
//                     document.querySelector('select[id*="Search"]') ||
//                     document.querySelector('select:first-of-type');
    
//     if (dropdown) {
//         console.log('Available options:');
//         for (let i = 0; i < dropdown.options.length; i++) {
//             console.log(`${i}: ${dropdown.options[i].text} (value: ${dropdown.options[i].value})`);
//         }
        
//         // Try to find Patient ID option
//         let patientIdIndex = -1;
//         for (let i = 0; i < dropdown.options.length; i++) {
//             if (dropdown.options[i].text.includes('Patient ID') || 
//                 dropdown.options[i].value.includes('PatientID')) {
//                 patientIdIndex = i;
//                 break;
//             }
//         }
        
//         if (patientIdIndex !== -1) {
//             dropdown.selectedIndex = patientIdIndex;
//             console.log(`Selected Patient ID option at index ${patientIdIndex}`);
//         } else {
//             // Fallback to index 2 (3rd option)
//             dropdown.selectedIndex = 2;
//             console.log('Fallback: Selected 3rd option (index 2)');
//         }
        
//         // Trigger change event
//         dropdown.dispatchEvent(new Event('change', { bubbles: true }));
//         console.log('Selected option:', dropdown.options[dropdown.selectedIndex].text);
//     } else {
//         console.error('❌ Dropdown not found!');
//     }
// });

// // Wait a moment for the selection to register
// await new Promise(resolve => setTimeout(resolve, 500));

// // Clear the text field first, then enter the patient ID
// console.log(`⌨️ Entering Patient ID: ${accountNumber}`);
// const textFieldSelector = '#ctl04_popupBase_txtSearch';
// try {
//     await newPage.waitForSelector(textFieldSelector, { visible: true, timeout: 5000 });
// } catch (error) {
//     console.log('🔍 Text field not found, trying alternatives...');
//     const textSelectors = [
//         'input[name*="txtSearch"]',
//         'input[type="text"]:visible',
//         'input[maxlength="50"]'
//     ];
//     // Will use first available selector
// }

// await newPage.evaluate(() => {
//     const textField = document.querySelector('#ctl04_popupBase_txtSearch') ||
//                      document.querySelector('input[name*="txtSearch"]') ||
//                      document.querySelector('input[type="text"]:not([style*="display: none"])');
//     if (textField) {
//         textField.value = '';
//         textField.focus();
//     }
// });

// await newPage.type('#ctl04_popupBase_txtSearch', accountNumber, { delay: 100 });

// // Verify the text was entered
// const enteredText = await newPage.$eval('#ctl04_popupBase_txtSearch', el => el.value).catch(() => 'Could not verify');
// console.log('✅ Text entered:', enteredText);

// // 🔍 Click the actual search button
// console.log('🖱️ Clicking search button...');
// try {
//     await newPage.click('#ctl04_popupBase_btnSearch');
// } catch (error) {
//     // Try alternative search button
//     await newPage.evaluate(() => {
//         const searchBtn = document.querySelector('#ctl04_popupBase_btnSearch') ||
//                          document.querySelector('input[name*="btnSearch"]') ||
//                          document.querySelector('input[value="Search"]');
//         if (searchBtn) {
//             searchBtn.click();
//         }
//     });
// }

// // Wait for search results
// await new Promise(resolve => setTimeout(resolve, 2000));
// console.log('✅ Search completed.');
// // ✅ Wait for search results to appear (if not already handled)
// console.log('⏳ Waiting for search result row selector to appear...');
// try {
//     await newPage.waitForSelector('#ctl04_popupBase_grvPopup_ctl02_lnkSelect', {
//         visible: true,
//         timeout: 10000
//     });

//     // 🖱️ Click the first row's "Select" link
//     console.log('🖱️ Clicking first patient select link...');
//     await newPage.evaluate(() => {
//         const selectBtn = document.querySelector('#ctl04_popupBase_grvPopup_ctl02_lnkSelect');
//         if (selectBtn) {
//             selectBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             selectBtn.click();
//         } else {
//             console.error('❌ Select button not found!');
//         }
//     });

//     console.log('✅ Patient selected successfully.');

// await new Promise(resolve => setTimeout(resolve, 1000));

//     // Wait for the Visit Date fields to load
// console.log('⏳ Waiting for Visit Date input fields...');
// await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Month', { visible: true, timeout: 8000 });
// await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Day', { visible: true, timeout: 8000 });
// await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Year', { visible: true, timeout: 8000 });

// // Split billingDate into MM/DD/YYYY
// const [month, day, year] = billingDate.split('/');
// console.log(`⌨️ Entering Visit Date: ${month}/${day}/${year}`);

// // Clear and type values into fields
// await page.evaluate(({ month, day, year }) => {
//     const setFieldValue = (selector, value) => {
//         const field = document.querySelector(selector);
//         if (field) {
//             field.value = '';
//             field.focus();
//             field.value = value;
//         }
//     };

//     setFieldValue('#ctl00_phFolderContent_DateVisited_Month', month);
//     setFieldValue('#ctl00_phFolderContent_DateVisited_Day', day);
//     setFieldValue('#ctl00_phFolderContent_DateVisited_Year', year);
// }, { month, day, year });

// console.log('✅ Visit Date fields populated successfully.');
// // ⏳ Wait for popup to load
//       // ⏳ Wait for popup to load
// console.log('⏳ Waiting for Provider Search button to appear...');
// await page.waitForSelector('#ctl00_phFolderContent_Button2', { visible: true, timeout: 6000 });

// console.log('🖱️ Setting up provider popup listener...');

// // Set up popup listener BEFORE clicking the button
// const providerPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Provider popup detected!');
//         browser.off('targetcreated', onPopup);
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// // 🔘 Click provider search button
// console.log('🖱️ Clicking provider search button...');
// await page.click('#ctl00_phFolderContent_Button2');

// // Wait for popup to appear
// console.log('⏳ Waiting for provider popup to load...');
// let providerPage;
// try {
//     // Wait for popup with timeout
//     providerPage = await Promise.race([
//         providerPopupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Provider popup timeout')), 8000))
//     ]);
//     console.log('✅ Provider popup loaded successfully');
// } catch (error) {
//     console.log('⚠️ No provider popup detected, checking for new tabs...');
    
//     // Fallback: Check for new tabs
//     await new Promise(resolve => setTimeout(resolve, 3000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//         providerPage = allPages[allPages.length - 1]; // Get the newest page
//         console.log('✅ New tab detected for provider search...');
//     } else {
//         throw new Error('No provider popup or new tab found');
//     }
// }

// // Wait for the provider page to load completely
// await new Promise(resolve => setTimeout(resolve, 3000));

// // Extract last name from provider variable (e.g., "Michael Lin, MD" -> "Lin")
// const providerSearchTerm = provider.split(' ')[1] || provider; // Get second word or fallback to full name
// console.log(`⌨️ Will search for Provider: ${providerSearchTerm}`);

// // Wait for the provider search input field to appear in the NEW TAB
// console.log('⏳ Waiting for Provider search field in new tab...');
// try {
//     await providerPage.waitForSelector('input[type="text"]', { visible: true, timeout: 8000 });
//     console.log('✅ Provider search field found in new tab');
// } catch (error) {
//     console.log('🔍 Provider search field not found, trying to find any text input...');
    
//     // Try to find any text input in the new tab
//     const textInputs = await providerPage.$$('input[type="text"]');
//     if (textInputs.length === 0) {
//         throw new Error('No text input found in provider popup');
//     }
//     console.log(`Found ${textInputs.length} text input(s) in provider popup`);
// }

// // Clear and enter provider search term in the NEW TAB
// console.log(`⌨️ Entering Provider search term in new tab: ${providerSearchTerm}`);

// // Find and clear the search field
// await providerPage.evaluate(() => {
//     const searchField = document.querySelector('input[type="text"]');
//     if (searchField) {
//         searchField.value = '';
//         searchField.focus();
//         console.log('Search field cleared and focused');
//     }
// });

// // Type the provider search term
// await providerPage.type('input[type="text"]', providerSearchTerm, { delay: 100 });

// // Verify the text was entered
// const enteredText = await providerPage.$eval('input[type="text"]', el => el.value).catch(() => 'Could not verify');
// console.log('✅ Provider search text entered in new tab:', enteredText);

// // Find and click the Search button in the NEW TAB
// console.log('🔍 Looking for Search button in new tab...');
// try {
//     // Wait for search button to be available
//     await providerPage.waitForSelector('input[value="Search"]', { visible: true, timeout: 5000 });
//     console.log('🖱️ Clicking Search button in new tab...');
//     await providerPage.click('input[value="Search"]');
// } catch (error) {
//     console.log('🔍 Search button not found, trying alternatives...');
    
//     // Try clicking any button that might be the search button
//     await providerPage.evaluate(() => {
//         const searchButtons = document.querySelectorAll('input[type="button"], input[type="submit"], button');
//         for (const btn of searchButtons) {
//             if (btn.value?.toLowerCase().includes('search') || 
//                 btn.textContent?.toLowerCase().includes('search') ||
//                 btn.onclick?.toString().includes('search')) {
//                 console.log('Found search button, clicking...');
//                 btn.click();
//                 return;
//             }
//         }
//         console.log('No search button found, trying first button...');
//         if (searchButtons.length > 0) {
//             searchButtons[0].click();
//         }
//     });
// }

// // Wait for search results to load
// console.log('⏳ Waiting for search results...');
// await new Promise(resolve => setTimeout(resolve, 3000));

// // Wait for search results and select first provider
// console.log('⏳ Waiting for provider search results to appear...');
// try {
//     // Wait for any select links to appear
//     await providerPage.waitForSelector('a[href*="Select"], a[onclick*="Select"]', {
//         visible: true,
//         timeout: 10000
//     });

//     console.log('🖱️ Selecting first provider from search results...');
    
//     // Click the first Select link
//     await providerPage.evaluate(() => {
//         const selectLinks = document.querySelectorAll('a[href*="Select"], a[onclick*="Select"]');
//         if (selectLinks.length > 0) {
//             console.log(`Found ${selectLinks.length} select link(s), clicking first one...`);
//             selectLinks[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
//             selectLinks[0].click();
//         } else {
//             console.error('❌ No select links found!');
//         }
//     });

//     console.log('✅ Provider selected successfully from new tab.');

// } catch (error) {
//     console.error('🚨 Failed to find or select provider:', error.message);
    
//     // Try alternative selection - look for any clickable row
//     console.log('🔄 Trying alternative provider selection...');
//     try {
//         await providerPage.evaluate(() => {
//             // Look for table rows that might be clickable
//             const rows = document.querySelectorAll('tr, td');
//             for (const row of rows) {
//                 if (row.onclick || row.querySelector('a')) {
//                     console.log('Found clickable row, attempting click...');
//                     row.click();
//                     return;
//                 }
//             }
//         });
//     } catch (altError) {
//         console.error('🚨 Alternative provider selection failed:', altError.message);
//     }
// }

// // Wait for the selection to process and popup to close
// await new Promise(resolve => setTimeout(resolve, 2000));

// console.log('🎉 Provider selection process completed.');
// // ⏳ Wait for the Billing Info tab to be visible
// console.log('⏳ Waiting for Billing Info tab...');
// await page.waitForSelector('#VisitTabs > ul > li:nth-child(2)', { visible: true, timeout: 6000 });

// // 🖱️ Click the Billing Info tab
// console.log('🖱️ Clicking Billing Info tab...');
// await page.click('#VisitTabs > ul > li:nth-child(2)');

// console.log('✅ Billing Info tab clicked successfully.');


// } catch (error) {
//     console.error('🚨 Failed to find or click select button:', error.message);
// }

// // Optional: Wait for search results to load
// try {
//     await newPage.waitForSelector('table[style*="margin:0;padding:0;"]', { timeout: 10000 });
//     console.log('✅ Search results table loaded.');
// } catch (error) {
//     console.log('⚠️ Search results may still be loading...');
// }
//         await new Promise(resolve => setTimeout(resolve, 3000)); // wait for 3s (adjust as needed)

//       for (let i = 0; i < diagnosisCodes.length; i++) {
//   const code = diagnosisCodes[i];

//   // Construct the field selector dynamically based on index (starts from 1)
//   const inputSelector = `#ctl00_phFolderContent_ucDiagnosisCodes_dc_10_${i + 1}`;
//   const suggestionSelector = `#divICD_10 > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(${i + 1}) > td:nth-child(2) > span`;

//   try {
//     console.log(`⏳ Waiting for diagnosis input ${i + 1}...`);
//     await page.waitForSelector(inputSelector, { visible: true, timeout: 8000 });

//     const input = await page.$(inputSelector);
//     if (!input) {
//       console.warn(`⚠️ Input field ${inputSelector} not found`);
//       continue;
//     }

//     console.log(`⌨️ Typing diagnosis code: ${code}`);
//     await input.click({ clickCount: 3 });
//     await input.type(code, { delay: 100 });

//     // Wait for the suggestion dropdown and click the visible option
//     console.log(`⏳ Waiting for suggestion dropdown ${i + 1}...`);
//     await page.waitForSelector(suggestionSelector, { visible: true, timeout: 5000 });

//     const suggestionSpan = await page.$(suggestionSelector);
//     if (suggestionSpan) {
//       console.log(`✅ Selecting suggestion for ${code}`);
//       await suggestionSpan.click();
//     } else {
//       console.warn(`⚠️ Suggestion not found for code: ${code}`);
//     }

//     // Extra wait for system to process
//     await new Promise(resolve => setTimeout(resolve, 1000));

//   } catch (error) {
//     console.error(`❌ Error entering diagnosis code ${code}:`, error.message);
//   }
// }
// // Function to extract and clean CPT codes
// function extractCPTCodes(cptInput) {
//   if (!cptInput) return [];
  
//   // Remove any non-digit characters and split by spaces, commas, or other separators
//   const cptArray = cptInput.toString()
//     .replace(/[^\d\s,]/g, '') // Remove non-digits except spaces and commas
//     .split(/[\s,]+/) // Split by spaces or commas
//     .filter(code => code.length === 5 && /^\d{5}$/.test(code)); // Only keep 5-digit codes
  
//   return cptArray;
// }

// // Extract CPT codes from the input
// const cptCodeArray = extractCPTCodes(cptCodes);
// console.log(`📋 Found ${cptCodeArray.length} valid CPT codes:`, cptCodeArray);

// if (cptCodeArray.length === 0) {
//   console.error('❌ No valid CPT codes found!');
//   return;
// }

// // Process each CPT code in sequence
// for (let i = 0; i < cptCodeArray.length; i++) {
//   const currentCPT = cptCodeArray[i];
//   console.log(`\n🔄 Processing CPT ${i + 1}/${cptCodeArray.length}: ${currentCPT}`);
  
//   // Step 1: Enter FROM date and TO date
//   console.log(`⌨️ Step 1: Entering billing dates for CPT ${currentCPT}`);
  
//   // Wait for DOS From field
//   const dosFromSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_DOS${i}`;
//   await page.waitForSelector(dosFromSelector, { visible: true, timeout: 8000 });
//   await page.focus(dosFromSelector);
//   await page.click(dosFromSelector, { clickCount: 3 });
//   await page.keyboard.type(billingDate, { delay: 50 });
//   console.log(`✅ FROM date entered: ${billingDate}`);

//   // Wait for DOS To field
//   const dosToSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_ToDOS${i}`;
//   await page.waitForSelector(dosToSelector, { visible: true, timeout: 8000 });
//   await page.focus(dosToSelector);
//   await page.click(dosToSelector, { clickCount: 3 });
//   await page.keyboard.type(billingDate, { delay: 50 });
//   console.log(`✅ TO date entered: ${billingDate}`);

//   // Step 2: Search and select CPT code
//   console.log(`🔍 Step 2: Searching for CPT code: ${currentCPT}`);
  
//   // Wait for CPT search button to appear
//   const searchButtonSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_btnUserCPT${i}`;
//   await page.waitForSelector(searchButtonSelector, { visible: true, timeout: 6000 });

//   // Set up popup listener BEFORE clicking the CPT button
//   const cptPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//       console.log('✅ CPT popup detected!');
//       browser.off('targetcreated', onPopup);
//       const popup = await target.page();
//       resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
//   });

//   // Click CPT search button
//   console.log('🖱️ Clicking CPT search button...');
//   await page.click(searchButtonSelector);

//   // Wait for popup to appear
//   console.log('⏳ Waiting for CPT popup to load...');
//   let cptPage;
//   try {
//     // Wait for popup with timeout
//     cptPage = await Promise.race([
//       cptPopupPromise,
//       new Promise((_, reject) => setTimeout(() => reject(new Error('CPT popup timeout')), 8000))
//     ]);
//     console.log('✅ CPT popup loaded successfully');
//   } catch (error) {
//     console.log('⚠️ No CPT popup detected, checking for new tabs...');
    
//     // Fallback: Check for new tabs
//     await new Promise(resolve => setTimeout(resolve, 3000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//       cptPage = allPages[allPages.length - 1]; // Get the newest page
//       console.log('✅ New tab detected for CPT search...');
//     } else {
//       throw new Error('No CPT popup or new tab found');
//     }
//   }

//   // Wait for the CPT page to load completely
//   await new Promise(resolve => setTimeout(resolve, 3000));

//   // Wait for the CPT search input field to appear in the popup
//   console.log('⏳ Waiting for CPT search field in popup...');
//   try {
//     await cptPage.waitForSelector('input[name="ctl04$popupBase$txtSearch"]', { visible: true, timeout: 8000 });
//     console.log('✅ CPT search field found in popup');
//   } catch (error) {
//     console.log('🔍 Specific CPT search field not found, trying generic text input...');
    
//     const textInputs = await cptPage.$$('input[type="text"]');
//     if (textInputs.length === 0) {
//       throw new Error('No text input found in CPT popup');
//     }
//     console.log(`Found ${textInputs.length} text input(s) in CPT popup`);
//   }

//   // Clear and enter CPT search term in the popup
//   console.log(`⌨️ Entering CPT search term in popup: ${currentCPT}`);

//   // Find and clear the search field
//   await cptPage.evaluate(() => {
//     const searchField = document.querySelector('input[name="ctl04$popupBase$txtSearch"]') || 
//                         document.querySelector('input[type="text"]');
//     if (searchField) {
//       searchField.value = '';
//       searchField.focus();
//     }
//   });

//   // Type the CPT search term
//   const searchSelector = 'input[name="ctl04$popupBase$txtSearch"]';
//   try {
//     await cptPage.type(searchSelector, currentCPT, { delay: 100 });
//   } catch (error) {
//     await cptPage.type('input[type="text"]', currentCPT, { delay: 100 });
//   }

//   // Verify the text was entered
//   const enteredText = await cptPage.evaluate(() => {
//     const field = document.querySelector('input[name="ctl04$popupBase$txtSearch"]') || 
//                     document.querySelector('input[type="text"]');
//     return field ? field.value : 'Could not verify';
//   });
//   console.log('✅ CPT search text entered in popup:', enteredText);

//   // Find and click the Search button in the popup
//   console.log('🔍 Looking for Search button in CPT popup...');
//   try {
//     await cptPage.waitForSelector('input[name="ctl04$popupBase$btnSearch"]', { visible: true, timeout: 5000 });
//     console.log('🖱️ Clicking Search button in CPT popup...');
//     await cptPage.click('input[name="ctl04$popupBase$btnSearch"]');
//   } catch (error) {
//     console.log('🔍 Specific search button not found, trying alternatives...');
    
//     await cptPage.evaluate(() => {
//       const searchButtons = document.querySelectorAll('input[value="Search"], input[type="button"], input[type="submit"], button');
//       for (const btn of searchButtons) {
//         if (btn.value?.toLowerCase().includes('search') || 
//             btn.textContent?.toLowerCase().includes('search') ||
//             btn.onclick?.toString().includes('search')) {
//           btn.click();
//           return;
//         }
//       }
//       if (searchButtons.length > 0) {
//         searchButtons[0].click();
//       }
//     });
//   }

//   // Wait for search results to load
//   console.log('⏳ Waiting for CPT search results...');
//   await new Promise(resolve => setTimeout(resolve, 3000));

//   // Select the CPT code from search results
//   console.log('⏳ Waiting for CPT search results to appear...');
//   try {
//     await cptPage.waitForSelector('#ctl04_popupBase_grvPopup_ctl02_lnkSelect', {
//       visible: true,
//       timeout: 10000
//     });

//     console.log('🖱️ Selecting CPT code from search results...');
//     await cptPage.click('#ctl04_popupBase_grvPopup_ctl02_lnkSelect');
//     console.log('✅ CPT code selected successfully');

//   } catch (error) {
//     console.error('🚨 Failed to find specific select link, trying alternatives...');
    
//     const selectLinkFound = await cptPage.evaluate(() => {
//       const selectLinks = document.querySelectorAll('a[id*="lnkSelect"]');
//       if (selectLinks.length > 0) {
//         selectLinks[0].click();
//         return true;
//       }
      
//       const allLinks = document.querySelectorAll('a');
//       for (const link of allLinks) {
//         if (link.textContent.trim().toLowerCase() === 'select') {
//           link.click();
//           return true;
//         }
//       }
      
//       const resultRows = document.querySelectorAll('table[id*="grvPopup"] tbody tr, #ctl04_popupBase_grvPopup tbody tr');
//       if (resultRows.length > 0) {
//         const firstLink = resultRows[0].querySelector('a');
//         if (firstLink) {
//           firstLink.click();
//         } else {
//           resultRows[0].click();
//         }
//         return true;
//       }
//       return false;
//     });

//     if (selectLinkFound) {
//       console.log('✅ Selected using alternative method');
//     } else {
//       console.error('🚨 Failed to select CPT code');
//     }
//   }

//   // Wait for the selection to process and popup to close
//   await new Promise(resolve => setTimeout(resolve, 3000));
//   console.log(`✅ CPT code ${currentCPT} selected and popup closed`);

//   // Step 3: Enter POS, Modifier, and ICD-10 pointer
//   console.log(`📝 Step 3: Entering additional fields for CPT ${currentCPT}`);

//   // Enter the POS value
//   if (typeof pos === 'string' && pos.trim() !== '') {
//     console.log(`🏥 Entering POS: ${pos}`);
//     const posSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_PlaceOfService${i}`;
//     try {
//       await page.waitForSelector(posSelector, { visible: true, timeout: 5000 });
//       await page.evaluate((value, selector) => {
//         const input = document.querySelector(selector);
//         if (input) input.value = value;
//       }, pos, posSelector);
//       console.log('✅ POS entered successfully');
//     } catch (error) {
//       console.error('❌ Error entering POS:', error.message);
//     }
//   }

//   // Enter the Modifier value
//   if (
//     modifier != null &&
//     typeof modifier === 'string' &&
//     modifier.trim() !== '' &&
//     modifier.trim() !== '-'
//   ) {
//     console.log(`🔧 Entering Modifier: ${modifier}`);
//     const modifierSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_ModifierA${i}`;
//     try {
//       await page.waitForSelector(modifierSelector, { visible: true, timeout: 5000 });
//       await page.evaluate((value, selector) => {
//         const input = document.querySelector(selector);
//         if (input) input.value = value;
//       }, modifier.trim(), modifierSelector);
//       console.log('✅ Modifier entered successfully');
//     } catch (error) {
//       console.error('❌ Error entering Modifier:', error.message);
//     }
//   }

//   // Set ICD-10 Pointer
//   try {
//     const pointerSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_DiagnosisCode${i}`;
//     let icdPointer = '';

//     if (diagnosisCodes && diagnosisCodes.length > 0) {
//       if (diagnosisCodes.length === 1) {
//         icdPointer = '1';
//       } else if (diagnosisCodes.length === 2) {
//         icdPointer = '12';
//       } else if (diagnosisCodes.length === 3) {
//         icdPointer = '123';
//       } else if (diagnosisCodes.length >= 4) {
//         icdPointer = '1234';
//       }

//       console.log(`🩺 Setting ICD-10 Pointer to: ${icdPointer}`);
//       await page.waitForSelector(pointerSelector, { visible: true, timeout: 5000 });
//       await page.evaluate((value, selector) => {
//         const input = document.querySelector(selector);
//         if (input) input.value = value;
//       }, icdPointer, pointerSelector);
//       console.log('✅ ICD-10 pointer set successfully');
//     }

//   } catch (error) {
//     console.error('❌ Error setting ICD-10 pointer:', error.message);
//   }

//   console.log(`🎉 CPT ${currentCPT} processing completed for row ${i}`);

//   // Add delay before processing next CPT (if any)
//   if (i < cptCodeArray.length - 1) {
//     console.log('⏳ Preparing for next CPT code...');
//     await new Promise(resolve => setTimeout(resolve, 2000));
//   }
// }

// console.log(`🎊 All ${cptCodeArray.length} CPT codes processed successfully!
// Summary:
// - CPT Codes: ${cptCodeArray.join(', ')}
// - Billing Date: ${billingDate}
// - POS: ${pos || 'Not provided'}
// - Modifier: ${modifier || 'Not provided'}
// - ICD-10 Codes: ${diagnosisCodes?.length || 0} codes`);

// // Optional: Add a final verification step
// console.log('🔍 Processing complete. Please verify all entries on the form.');
//   // After processing all CPT codes, check if we need to handle dates
//   // We'll check POS value again and if needed, enter the dates once
//   try {
//     // Check POS value from the page (get it after all CPTs are processed)
//     const posSelector = '#billingClaimIpt24';
//     await newPage.waitForSelector(posSelector, { visible: true, timeout: 8000 });
//     const posValue = await newPage.$eval(posSelector, el => el.value.trim());
    
//     console.log(`Final POS check: ${posValue}`);
    
//     // Handle date entry for POS values 21 or 31 after all CPTs are processed
//     if (["21", "31"].includes(posValue)) {
//       console.log(`POS is ${posValue} — proceeding to date input.`);
      
//       // Find and click the data button to open the date modal
//       const dataBtnSelector = 'button[id^="claimDataBtn"]';
//       const dataBtnHandle = await newPage.$(dataBtnSelector);
//       if (dataBtnHandle) {
//         const btnId = await newPage.evaluate(el => el.id, dataBtnHandle);
//         console.log(`Found data button with ID: ${btnId}`);
//         await dataBtnHandle.click();
//         console.log(`Clicked data button`);
//         await new Promise(resolve => setTimeout(resolve, 1500)); // Longer wait for modal to open fully
//       } else {
//         console.warn(`Data button with selector ${dataBtnSelector} not found`);
//         throw new Error("Data button not found");
//       }
      
//       // Wait for admit date input to be visible
//       const admitDateInputSelector = '#dtHCFAFrom16From';
//       const dischargeInputSelector = '#dtHCFAFrom16To';
//       await newPage.waitForSelector(admitDateInputSelector, { visible: true, timeout: 5000 });
      
//       // Helper function to format date with leading zeros (MM/DD/YYYY)
//       const formatDateWithLeadingZeros = (dateStr) => {
//         const date = new Date(dateStr);
//         const month = String(date.getMonth() + 1).padStart(2, '0');
//         const day = String(date.getDate()).padStart(2, '0');
//         const year = date.getFullYear();
//         return `${month}/${day}/${year}`;     
//       };
      
//       const formattedBillingDate = formatDateWithLeadingZeros(billingDate);
      
// // After calculation of admit date, add this validation check before entering dates:

// // Determine admit date based on CPT codes present
// let admitDateOnly;
// let dischargeDateToUse = formattedBillingDate;
      
// // Check if we have any special CPT codes that affect date handling
// const has99222or99223 = cptCodes.some(code => {
//   const baseCode = code.split('-')[0].trim();
//   return ["99222", "99223"].includes(baseCode);
// });
      
// // Check if we have discharge CPT codes
// const has99238or99239 = cptCodes.some(code => {
//   const baseCode = code.split('-')[0].trim();
//   return ["99238", "99239"].includes(baseCode);
// });
      
// // Special case: if we have 99222 or 99223, use DOS as admit date
// if (has99222or99223) {
//   console.log(`🔍 Special CPT codes 99222/99223 detected - using DOS as admit date`);
//   admitDateOnly = formattedBillingDate;
// } else {
//   // For other codes, process normally
//   // Clean admit date if present in expected format
//   admitDateOnly = admitDate.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)?.[0] || "";
  
//   // Format with leading zeros if date is present
//   if (admitDateOnly) {
//     admitDateOnly = formatDateWithLeadingZeros(admitDateOnly);
//   }
  
//   // If admit date is NOT provided, calculate it from discharge date
//   if (!admitDateOnly) {
//     console.log(`❗ Admit Date not provided. Calculating from Discharge Date: ${formattedBillingDate}`);
    
//     // Calculate admit date as 1 day before discharge date
//     const admit = new Date(billingDate);
//     admit.setDate(admit.getDate() - 1);
    
//     // Format calculated admit date with leading zeros
//     const month = String(admit.getMonth() + 1).padStart(2, '0');
//     const day = String(admit.getDate()).padStart(2, '0');
//     const year = admit.getFullYear();
//     admitDateOnly = `${month}/${day}/${year}`;
    
//     console.log(`Calculated Admit Date: ${admitDateOnly}`);
//   } else {
//     console.log(`✅ Using provided Admit Date: ${admitDateOnly}`);
//   } 
// }
// // Only check for same dates when dealing with discharge codes
// if (has99238or99239 && admitDateOnly === dischargeDateToUse) {
//   console.log(`⚠️ Admit date and discharge date are the same for discharge codes: ${admitDateOnly}`);
  
//   // Calculate new admit date as 1 day before discharge date
//   const admit = new Date(dischargeDateToUse);
//   admit.setDate(admit.getDate() - 1);
  
//   // Format calculated admit date with leading zeros
//   const month = String(admit.getMonth() + 1).padStart(2, '0');
//   const day = String(admit.getDate()).padStart(2, '0');
//   const year = admit.getFullYear();
//   admitDateOnly = `${month}/${day}/${year}`;
  
//   console.log(`Updated Admit Date to be one day before discharge: ${admitDateOnly}`);
// }

  
// // ADD THIS NEW VALIDATION: Check if admit date is greater than DOS
// const admitDateObj = new Date(admitDateOnly);
// const billingDateObj = new Date(billingDate);

// if (admitDateObj > billingDateObj) {
//   console.log(`⚠️ Admit date (${admitDateOnly}) is greater than DOS (${formattedBillingDate})`);
  
//   // Set admit date to one day before DOS
//   const newAdmitDate = new Date(billingDate);
//   newAdmitDate.setDate(newAdmitDate.getDate() - 1);
  
//   // Format the new admit date
//   const month = String(newAdmitDate.getMonth() + 1).padStart(2, '0');
//   const day = String(newAdmitDate.getDate()).padStart(2, '0');
//   const year = newAdmitDate.getFullYear();
//   admitDateOnly = `${month}/${day}/${year}`;
  
//   console.log(`Corrected admit date to one day before DOS: ${admitDateOnly}`);
// }

// // Now enter the validated admit date
// await newPage.type(admitDateInputSelector, admitDateOnly);
// console.log(`Entered admit date: ${admitDateOnly}`);
// await new Promise(resolve => setTimeout(resolve, 500));

// // Enter discharge date ONLY if we have discharge CPT codes
// if (has99238or99239) {
//   await newPage.type(dischargeInputSelector, dischargeDateToUse);
//   await newPage.keyboard.press('Tab'); // Move focus away 
//   console.log(`Entered discharge date: ${dischargeDateToUse}`);
//   await new Promise(resolve => setTimeout(resolve, 500));
// } else {
//   console.log(`No discharge codes (99238/99239) present - skipping discharge date entry`);
// }

// // Add extra wait time before attempting to click OK
// await new Promise(resolve => setTimeout(resolve, 1500));

// // Try multiple click strategies for the OK button
// const okBtnSelector = '#claimDataBtn13';
// const bootboxOkSelector = 'body > div.bootbox.modal.fade.bluetheme.medium-width.bootbox-alert.in > div > div > div.modal-footer > button';

// try {
//   let modalStillOpen = true;
//   let attempts = 0;

//   while (modalStillOpen && attempts < 3) {
//     attempts++;

//     // Check if OK button is visible
//     await newPage.waitForSelector(okBtnSelector, { visible: true, timeout: 5000 });
//     console.log(`Attempt ${attempts}: OK button is visible, trying click...`);

//     // Try direct click
//     await newPage.click(okBtnSelector);
//     await new Promise(resolve => setTimeout(resolve, 1000));

//     // Check for bootbox alert
//     const bootboxExists = await newPage.$(bootboxOkSelector);
//     if (bootboxExists) {
//       console.log("Bootbox alert detected. Clicking OK...");
//       await newPage.click(bootboxOkSelector);
//       await new Promise(resolve => setTimeout(resolve, 1000));

//       // Re-enter admit date
//       await newPage.click(admitDateInputSelector, { clickCount: 3 });
//       await newPage.type(admitDateInputSelector, admitDateOnly);
//       console.log(`Re-entered admit date: ${admitDateOnly}`);
//       await new Promise(resolve => setTimeout(resolve, 500));

//       // Re-enter discharge date if needed
//       if (has99238or99239) {
//         await newPage.click(dischargeInputSelector, { clickCount: 3 });
//         await newPage.type(dischargeInputSelector, dischargeDateToUse);
//         await newPage.keyboard.press('Tab');
//         console.log(`Re-entered discharge date: ${dischargeDateToUse}`);
//         await new Promise(resolve => setTimeout(resolve, 500));
//       }

//       continue; // Retry clicking OK
//     }

//     // Check if modal is still open
//     modalStillOpen = await newPage.evaluate(() => {
//       return !!document.querySelector('.modal-dialog:not([style*="display: none"])');
//     });

//     if (!modalStillOpen) {
//       console.log(`✅ Modal closed successfully after ${attempts} attempt(s)`);
//       break;
//     } else {
//       console.log("Modal still open, retrying...");
//     }
//   }
// } catch (error) {
//   console.error(`Failed to click OK button or handle modal:`, error);
//   await newPage.screenshot({ path: `ok_button_error_date.png` });
// }

//     } else {
//       console.log(`POS is ${posValue} — no date entry needed.`);
//     }
    
//     console.log(`✅ Completed processing all CPT codes and dates.`);
//   } catch (error) {
//     console.error(`❌ Failed during date processing:`, error);
//     await newPage.screenshot({ path: `date_processing_error.png` });
//   }
// const claimNoSelector = '#billingClaimTbl2 > tbody > tr:nth-child(1) > td:nth-child(2) > span';

//         await newPage.waitForSelector(claimNoSelector, { visible: true, timeout: 5000 });
        
//         const claimNumber = await newPage.$eval(claimNoSelector, el => el.textContent.trim());
        
//         console.log(`🧾 Claim Number: ${claimNumber}`);
//         try {
//             // Step 1: Find the claim status dropdown
//             const statusSelector = 'select[id^="claimStatusSel"]';
//             await newPage.waitForSelector(statusSelector, { visible: true, timeout: 5000 });
//             const statusDropdownHandle = await newPage.$(statusSelector);
          
//             if (statusDropdownHandle) {
//               const selectedText = await newPage.evaluate(el => {
//                 const selectedOption = el.options[el.selectedIndex];
//                 return selectedOption ? selectedOption.textContent.trim() : '';
//               }, statusDropdownHandle);  
          
//               console.log(`Current Claim Status: ${selectedText}`);
          
//               // Step 2: If not pending, set to Pending
//               if (selectedText.toLowerCase() !== "pending") {
//                 await newPage.select(statusSelector, "PEN");
//                 console.log("Changed Claim Status to Pending");
//                 await new Promise(resolve => setTimeout(resolve, 500)); // wait for change event
//               }
//               await new Promise(resolve => setTimeout(resolve, 1000));

// // Step 3: Click the first OK button
// try {
//   const okSelector = 'button[id^="claimScreenOkBtn"]';
//   await newPage.waitForSelector(okSelector, { visible: true, timeout: 5000 });
  
//   const okButtonHandle = await newPage.$(okSelector);
//   if (okButtonHandle) {
//     await okButtonHandle.evaluate(btn => btn.scrollIntoView({ behavior: "smooth", block: "center" }));
//     await new Promise(resolve => setTimeout(resolve, 300));
    
//     // Use bounding box + mouse click for higher reliability
//     const box = await okButtonHandle.boundingBox();
//     if (box) {
//       await newPage.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
//       console.log("Clicked first OK button.");
//     } else {
//       console.warn("First OK button not clickable (no bounding box).");
//     }
    
//     // Wait for error popup to appear after clicking first OK
//     await new Promise(resolve => setTimeout(resolve, 1500));
    
//     // Step 4: Now handle the error popup that appears - click the second OK button
//     try {
//       // Error popup OK button selector - from modal-footer
//       const errorOkSelector = 'div.bootbox.modal.fade.bluetheme.medium-width.in button.btn.btn-blue.btn-sm.btn-xs, button[data-bb-handler="Yes"]';
//       await newPage.waitForSelector(errorOkSelector, { visible: true, timeout: 5000 });
      
//       const errorOkButton = await newPage.$(errorOkSelector);
//       if (errorOkButton) {
//         await errorOkButton.evaluate(btn => btn.scrollIntoView({ behavior: "smooth", block: "center" }));
//         await new Promise(resolve => setTimeout(resolve, 300));
        
//         // Click the error popup's OK button
//         const errorOkBox = await errorOkButton.boundingBox();
//         if (errorOkBox) {
//           await newPage.mouse.click(errorOkBox.x + errorOkBox.width / 2, errorOkBox.y + errorOkBox.height / 2);
//           console.log("Clicked error popup OK button.");
//         } else {
//           console.warn("Error popup OK button not clickable (no bounding box).");
//           await errorOkButton.click();
//         }
        
//         // Wait for UI to update after clicking error OK button
//         await new Promise(resolve => setTimeout(resolve, 1000));
        
//         // Step 5: Finally click the Cancel button
//         const cancelSelector = '#billingClaimBtn36';
//         await newPage.waitForSelector(cancelSelector, { visible: true, timeout: 5000 });
        
//         const cancelButton = await newPage.$(cancelSelector);
//         if (cancelButton) {
//           await cancelButton.evaluate(btn => btn.scrollIntoView({ behavior: "smooth", block: "center" }));
//           await new Promise(resolve => setTimeout(resolve, 300));
          
//           // Click the Cancel button
//           const cancelBox = await cancelButton.boundingBox();
//           if (cancelBox) {
//             await newPage.mouse.click(cancelBox.x + cancelBox.width / 2, cancelBox.y + cancelBox.height / 2);
//             console.log("Clicked Cancel button after handling error.");
//           } else {
//             console.warn("Cancel button not clickable (no bounding box).");
//             await cancelButton.click();
//           }
          
//           await new Promise(resolve => setTimeout(resolve, 1000));
//         } else {
//           console.warn("Cancel button not found after error handling.");
//         }
//       } else {
//         console.warn("Error popup OK button not found.");
//       }
//     } catch (errorPopupError) {
//       console.error(`Error handling popup: ${errorPopupError.message}`);
      
//       // Fallback: try to find any visible OK button in a modal
//       try {
//         const anyOkSelector = 'div.modal.in button:contains("OK"), div.modal.in button.btn-blue';
//         const anyOkButton = await newPage.$(anyOkSelector);
//         if (anyOkButton) {
//           await anyOkButton.click();
//           console.log("Clicked fallback OK button in error popup.");
          
//           // Still try to click Cancel after fallback
//           await new Promise(resolve => setTimeout(resolve, 1000));
//           await newPage.click('#billingClaimBtn36').catch(e => console.warn("Cancel button click failed after fallback."));
//         }
//       } catch (fallbackError) {
//         console.error(`Fallback error handling failed: ${fallbackError.message}`);
//       }
//     }
//   } else {
//     console.warn("First OK button not found.");
//   }
// } catch (error) {
//   console.error(`Error in claim status handling sequence: ${error.message}`);
// }
//             }
//           } catch (error) {
//             console.error(`Error handling claim status: ${error.message}`);
//           }
          
  
//   return { success: true, claimNumber,billedFees };
// } catch (error) {
//         console.error('Error processing account:', error);
//         return false; // Return failure
//     }
// }

// // Endpoint to process data and return status
// app.post("/process", upload.single("file"), async (req, res) => {
//     if (!req.file || !req.body.originalPath) {
//         return res.status(400).json({ error: "No file or original path provided" });
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
        
//         // Check if "Patient ID" exists in headers
//         console.log(headers);
//         if (!headers.includes("Patient ID")) {
//             throw new Error("'Patient ID' Column is not defined.");
//         }
        
//         // Map headers to their index positions
//         headers.forEach((header, index) => {
//             columnMap[header] = index;
//         });
        
//         const requiredColumns = ["Patient ID", "Provider", "Facility", "DOS"];
//         for (const column of requiredColumns) {
//             if (!headers.includes(column)) {
//                 throw new Error(`Required column '${column}' is missing from the spreadsheet.`);
//             }
//         }
    
//         // Add missing headers if needed
//         const extraHeaders = ["Claim Number", "Billed Fee", "Result"];
//         for (const header of extraHeaders) {
//             if (!headers.includes(header)) {
//                 columnMap[header] = headers.length;
//                 headers.push(header);
//                 for (let i = 1; i < originalData.length; i++) {
//                     originalData[i] = originalData[i] || [];
//                     originalData[i].push("");
//                 }
//             }
//         }
    
//         headers.forEach((header, idx) => columnMap[header] = idx);
    
//         let newRowsToProcess = 0;
//         const processedAccounts = new Set();
        
//         let missingCallDiagnoses = 0;  // Track rows missing Diagnoses

//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue; // Skip undefined rows
            
//             const accountNumberRaw = row[columnMap["Patient ID"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
//             if (!accountNumberRaw || status === "done" || status === "failed") continue;  // Skip processed or failed rows

            
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             processedAccounts.add(accountNumber);
//             newRowsToProcess++;
        
//             // Check if Diagnoses is missing
//             const callDiagnoses = row[columnMap["Diagnoses"]] || "";
//             if (!callDiagnoses.trim()) {
//                 missingCallDiagnoses++;
//             }
//         }
        
//         // Check if all remaining rows are missing Diagnoses
//         if (newRowsToProcess > 0 && newRowsToProcess === missingCallDiagnoses) {
//             return res.json({
//                 success: false,
//                 message: "Diagnoses not found in Excel.",
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
        
//         const newPage = await continueWithLoggedInSession();
// if (!newPage) {
//     console.error("❌ Failed to get active page.");
//     return;
// }

        
//         const results = { successful: [], failed: [] };
        
//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue;
        
//             const accountNumberRaw = row[columnMap["Patient ID"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
//             if (!accountNumberRaw || status === "done") continue;
        
//             const provider = row[columnMap["Provider"]] || "";
//             const facilityName = row[columnMap["Facility"]] || "";
//             const billingDate = row[columnMap["DOS"]] || "";
//             const diagnosisText = row[columnMap["Diagnoses"]] || "";
//             const chargesText = row[columnMap["CPT"]] || ""; // CPT codes
//             const admitDate = row[columnMap["Admit Date"]] || ""; // ✅ NEW LINE
//             const pos = row[columnMap["POS"]] || ""; // ✅ NEW LINE
//             const modifier = row[columnMap["Modifier"]] || ""; // ✅ NEW LINE

//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             console.log(`Processing account: ${accountNumber}`);
        
//             // Validate data before processing
//             if (!provider) console.log('Provider value is missing or undefined');
//             if (!facilityName) console.log('Facility is missing or undefined');
//             if (!billingDate) console.log('DOS is missing or undefined');
//             if (!admitDate) console.log('Admit Date is missing or undefined'); // ✅ Optional

//             // Extract diagnosis codes from the text inside parentheses
//             const diagnosisCodes = [];

//             // Regex for ICD-10 codes and procedure codes:
//             // - ICD-10: Starts with a letter, 2–4 digits, optional dot, 1–4 alphanums
//             // - ICD-10-PCS (procedure): 7 alphanumeric characters
//             const dxCodeRegex = /\b(?:[A-Z][0-9]{2,4}(?:\.[0-9A-Z]{1,4})?|[A-Z0-9]{7})\b/g;
          
//             // Match all valid codes
//             const matches = diagnosisText.match(dxCodeRegex);
          
//             if (matches) {
//               // Remove duplicates and trim spaces
//               const uniqueCodes = [...new Set(matches.map(code => code.trim()))];
//               diagnosisCodes.push(...uniqueCodes);
//             }
          
//             // Filter: Only alphanumeric codes (must contain at least one letter & one digit)
//             const alphanumericCodes = diagnosisCodes.filter(code => /[A-Z]/i.test(code) && /\d/.test(code));
          
//             // Output
//             console.log("Diagnosis Codes:", diagnosisCodes);
//             console.log("Total Codes:", diagnosisCodes.length);
//             console.log("Alphanumeric Codes:", alphanumericCodes.length);
//             alphanumericCodes.forEach((code, index) => {
//               console.log(`  Code ${index + 1}: ${code}`);
//             });
          
          
          
//          const cptCodes = [];
// const validCptRegex = /\b\d{5}\b/g;
// let cptMatch;

// while ((cptMatch = validCptRegex.exec(chargesText)) !== null) {
//   const baseCode = cptMatch[0];
//   cptCodes.push(baseCode);
//   console.log(`🧩 CPT Code: ${baseCode}`);
// }

// console.log('Extracted CPT Codes with asterisk units:', cptCodes);

//             // === Call processAccount (Don't send CPT length)
//             const { success, claimNumber ,billedFees} = await processAccount(
//                 newPage,
//                 newPage.browser(), // Add browser parameter
//                 accountNumber,
//                 provider,
//                 facilityName,
//                 billingDate,
//                 diagnosisCodes,
//                 cptCodes, // ✅ only the array
//                 admitDate,
//                 pos,
//                 modifier // ✅ Pass Admit Date here

//             );
        
//          // Set Result
//             if (success) {
//                 results.successful.push(accountNumber);
//                 row[columnMap["Result"]] = "Done";
//             } else {
//                 results.failed.push(accountNumber);
//                 row[columnMap["Result"]] = "Failed";
//             }

//             // Always write Claim Number if available
//             row[columnMap["Claim Number"]] = claimNumber ? claimNumber.toString() : "";

//             // Write Billed Fee: only "Billed Fee is zero" if any fee is zero, else leave blank
//             if (Array.isArray(billedFees)) {
//                 const anyFeeZero = billedFees.some(f => typeof f.fee === 'number' && f.fee === 0);
//                 row[columnMap["Billed Fee"]] = anyFeeZero ? "Billed Fee is zero" : "";
//             } else {
//                 row[columnMap["Billed Fee"]] = ""; // Write blank if no valid array
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
//                 fgColor: { argb: "A9A9A9" }
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

//         const columnWidths = [6, 10, 10, 10, 15, 16, 16, 8, 40,15, 11, 16, 50, 30,13];

//         // When setting up columns, ensure EPS Comments has wrap text enabled
//         sheet.columns = columnWidths.map((width, index) => ({
//             width,
//             style: (index === columnMap["Diagnoses"] || index === columnMap["Charges"])                                
//                 ? { alignment: { wrapText: true } } 
//                 : {},
//         }));
        
//         // When formatting cells in each row, add special handling for EPS Comments
//         originalData.slice(1).forEach((row) => {
//             if (!row) return; // Skip undefined rows
            
//             const hasData = row.some(cell => cell && cell.toString().trim() !== "");
//             if (!hasData) return; // Skip empty rows
            
//             const rowData = headers.map(header => {
//                 if (header === "Claim Number") {
//                     return row[columnMap["Claim Number"]] || "";  // Use the value of Claim Number
//                 }
//                 return row[columnMap[header]] || "";  // Use existing data for other columns
//             });
            
//             const newRow = sheet.addRow(rowData);
//             newRow.height = 100;
//             newRow.getCell(columnMap["Claim Number"] + 1).alignment = { horizontal: "center", vertical: "middle" };

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

//                 if (header === "Patient ID" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFFFF" } 
//                     };
//                     cell.alignment = { horizontal: "center", vertical: "bottom" };
//                 }
//                 if (header === "Claim Number" && cell.value) {
//                     cell.fill = {
//                         type: "pattern",
//                         pattern: "solid",
//                         fgColor: { argb: "FFFFFF" } 
//                     };
//                     cell.alignment = { horizontal: "center", vertical: "bottom" };
//                 }
//                 if (header === "Billed Fee" && cell.value) {
//                     // Set the column width to 20
//                     sheet.getColumn(columnMap["Billed Fee"] + 1).width = 20;
                
//                     // Set text color to red
//                     cell.font = {
//                         color: { argb: "FF0000" }, // Red text
//                     };
                
//                     // Align the text
//                     cell.alignment = { horizontal: "center", vertical: "bottom" };
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
//             message: `${results.successful.length} Charges Entered successfully. ${results.failed.length} rows failed.`,
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
//   const fileId = req.params.fileId;
//   const fileData = processedFiles.get(fileId);
  
//   if (!fileData) {
//       return res.status(404).json({ error: "File not found" });
//   }
  
//   // Determine content type based on file extension
//   let contentType = 'application/octet-stream'; // Default
  
//   if (fileData.filename.endsWith('.xlsx')) {
//       contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
//   } else if (fileData.filename.endsWith('.xls')) {
//       contentType = 'application/vnd.ms-excel';
//   } else if (fileData.filename.endsWith('.csv')) {
//       contentType = 'text/csv';
//   }
  
//   // Set the response headers
//   res.setHeader('Content-Disposition', `attachment; filename="${fileData.filename}"`);
//   res.setHeader('Content-Type', contentType);
//   res.setHeader('Content-Length', fileData.buffer.length);
  
//   // Send the file buffer
//   res.send(fileData.buffer);
  
//   // Optional: Remove the file from memory after some time
// // Optional: Remove the file from memory after some time
// setTimeout(() => {
//     processedFiles.delete(fileId);
//     console.log(`File ${fileId} removed from memory`);
// }, 3 * 60 * 60 * 1000); // Remove after 3 hours

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



// async function continueWithLoggedInSession() {
//     try {
//         const browser = await puppeteer.connect({
//             browserURL: 'http://localhost:9222', // Connect to existing Chrome
//             defaultViewport: null,
//             headless: false,
//             timeout: 120000,
//         });

//         const pages = await browser.pages();
//         const sessionPage = pages.find(p =>
//             p.url().includes('Appointments/ViewAppointments.aspx')
//         );

//         if (!sessionPage) {
//             console.log('❌ No matching tab found for the appointment page.');

//             if (pages.length > 0) {
//                 // Show alert in the first available page
//                 await pages[0].bringToFront();
//                 await pages[0].evaluate(() => {
//                     alert('⚠️ Session not found');
//                 });
//             } else {
//                 console.log('⚠️ No pages available to show alert.');
//             }

//             return null;
//         }

//         console.log('✅ Found existing page. Waiting for Patient Visit tab...');
//         await sessionPage.waitForSelector('#patient-visits_tab > span', {
//             visible: true,
//             timeout: 10000,
//         });

//         console.log('🖱️ Clicking Patient Visit tab...');
//         await sessionPage.evaluate(() => {
//             const tab = document.querySelector('#patient-visits_tab > span');
//             if (tab) {
//                 tab.scrollIntoView({ behavior: 'smooth', block: 'center' });
//                 tab.click();
//             }
//         });

//         console.log('✅ Patient Visit tab clicked successfully.');
//         return sessionPage;
//     } catch (error) {
//         console.error('🚨 Error in continueWithLoggedInSession:', error.message);
//         return null;
//     }
// }

// async function processAccount(page,browser,accountNumber, provider, facilityName, billingDate, diagnosisCodes,cptCodes,admitDate,pos,modifier,priornumber) {
//    try {
//       // Wait and click Add New Visit
// // Wait and click Add New Visit
// await new Promise(resolve => setTimeout(resolve, 3000));
// console.log('⏳ Waiting for Add New Visit button to appear...');
// await page.waitForSelector('#addNewVisit', { visible: true, timeout: 15000 });

// console.log('🖱️ Setting up popup listener...');

// // Set up popup listener BEFORE clicking the button
// const popupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Popup detected!');
//         browser.off('targetcreated', onPopup); // Fixed: Use .off() instead of .removeListener()
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// console.log('🖱️ Clicking Add New Visit button...');
// await page.evaluate(() => {
//     const btn = document.querySelector('#addNewVisit');
//     if (btn) {
//         btn.scrollIntoView({ behavior: 'smooth', block: 'center' });  
//         btn.click();
//     } else {
//         console.error('❌ Add New Visit button not found!');
//     }
// });
//  // ⏳ Wait for popup to load
//         console.log('⏳ Waiting for Search button to appear...');
//         await page.waitForSelector('#ctl00_phFolderContent_Button1', { visible: true, timeout: 9000 });

//         // 🔘 Click initial popup search button
//         console.log('🖱️ Clicking popup search open button...');
//         await page.click('#ctl00_phFolderContent_Button1');
// // Wait for popup to appear with reduced timeout
// console.log('⏳ Waiting for popup to load...');
// let newPage;
// try {
//     // Reduced timeout from 6000ms to 3000ms
//     newPage = await Promise.race([
//         popupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Popup timeout')), 3000))
//     ]);
//     console.log('✅ Popup loaded successfully');
// } catch (error) {
//     console.log('⚠️ No popup detected, checking for new tabs...');
    
//     // Reduced wait time from 2000ms to 1000ms
//     await new Promise(resolve => setTimeout(resolve, 1000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//         newPage = allPages[allPages.length - 1];
//         console.log('✅ New tab detected, using it...');
//     } else {
//         newPage = page;
//         console.log('📝 No new tab, using current page...');
//     }
// }

// // Reduced wait time from 3000ms to 1500ms
// await new Promise(resolve => setTimeout(resolve, 1500));

// // ⏳ Wait for Search button with reduced timeout and parallel checking
// console.log('⏳ Waiting for Search button to appear in new page...');
// const searchButtonSelectors = [
//     '#ctl00_phFolderContent_Button1',
//     'input[value="Search"]',
//     'button[onclick*="Search"]',
//     '.button:contains("Search")',
//     '#Button1'
// ];

// let searchButtonFound = false;
// let workingSelector = null;

// // Check all selectors in parallel instead of sequentially
// try {
//     const promises = searchButtonSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 3000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const results = await Promise.allSettled(promises);
//     workingSelector = results.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (workingSelector) {
//         console.log(`✅ Found search button with selector: ${workingSelector}`);
//         searchButtonFound = true;
//     }
// } catch (error) {
//     console.log('🔍 No search button found with standard selectors');
// }

// if (!searchButtonFound) {
//     throw new Error('Search button not found with any selector');
// }

// // 🔘 Click initial popup search button
// console.log('🖱️ Clicking popup search open button...');
// try {
//     await newPage.click(workingSelector);
// } catch (error) {
//     // Fallback click method
//     await newPage.evaluate(() => {
//         const searchButtons = document.querySelectorAll('input[value*="Search"], button');
//         for (const btn of searchButtons) {
//             if (btn.value?.includes('Search') || btn.textContent?.includes('Search')) {
//                 btn.click();
//                 break;
//             }
//         }
//     });
// }

// // Wait for dropdown with parallel selector checking
// console.log('⏳ Waiting for dropdown to appear...');
// const dropdownSelectors = [
//     '#ctl04_popupBase_ddlSearch',
//     'select[name*="ddlSearch"]',
//     'select[id*="Search"]',
//     'select:first-of-type'
// ];

// let dropdownSelector = null;

// try {
//     const dropdownPromises = dropdownSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 2000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const dropdownResults = await Promise.allSettled(dropdownPromises);
//     dropdownSelector = dropdownResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (dropdownSelector) {
//         console.log(`✅ Found dropdown with selector: ${dropdownSelector}`);
//     } else {
//         throw new Error('Dropdown not found');
//     }
// } catch (error) {
//     console.log('❌ Dropdown not found with any selector');
//     throw error;
// }

// // Reduced wait time from 1000ms to 500ms
// await new Promise(resolve => setTimeout(resolve, 500));

// // 🔽 Select "Patient ID" with improved logic
// console.log('🔽 Selecting "Patient ID" from dropdown...');
// await newPage.evaluate((selector) => {
//     const dropdown = document.querySelector(selector);
    
//     if (dropdown) {
//         // Find Patient ID option more efficiently
//         let patientIdIndex = -1;
//         const options = Array.from(dropdown.options);
        
//         // Search for Patient ID option
//         patientIdIndex = options.findIndex(option => 
//             option.text.toLowerCase().includes('patient id') || 
//             option.value.toLowerCase().includes('patientid') ||
//             option.text.toLowerCase().includes('patient')
//         );
        
//         if (patientIdIndex !== -1) {
//             dropdown.selectedIndex = patientIdIndex;
//             console.log(`Selected Patient ID option at index ${patientIdIndex}: ${options[patientIdIndex].text}`);
//         } else {
//             // Fallback to index 2 (3rd option) if Patient ID not found
//             dropdown.selectedIndex = Math.min(2, options.length - 1);
//             console.log(`Fallback: Selected option at index ${dropdown.selectedIndex}: ${options[dropdown.selectedIndex].text}`);
//         }
        
//         // Trigger change event
//         dropdown.dispatchEvent(new Event('change', { bubbles: true }));
//         return true;
//     }
//     return false;
// }, dropdownSelector);

// // Reduced wait time from 500ms to 200ms
// await new Promise(resolve => setTimeout(resolve, 200));

// // Enhanced text field handling with parallel selector checking
// console.log(`⌨️ Entering Patient ID: ${accountNumber}`);
// const textFieldSelectors = [
//     '#ctl04_popupBase_txtSearch',
//     'input[name*="txtSearch"]',
//     'input[type="text"]:not([style*="display: none"])',
//     'input[maxlength="50"]'
// ];

// let textFieldSelector = null;

// try {
//     const textPromises = textFieldSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 2000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const textResults = await Promise.allSettled(textPromises);
//     textFieldSelector = textResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (!textFieldSelector) {
//         throw new Error('Text field not found');
//     }
// } catch (error) {
//     console.log('❌ Text field not found with any selector');
//     throw error;
// }

// // Clear and enter text more efficiently
// await newPage.evaluate((selector) => {
//     const textField = document.querySelector(selector);
//     if (textField) {
//         textField.value = '';
//         textField.focus();
//     }
// }, textFieldSelector);

// await newPage.type(textFieldSelector, accountNumber, { delay: 50 }); // Reduced delay from 100ms to 50ms

// // 🔍 Click the search button with parallel selector checking
// console.log('🖱️ Clicking search button...');
// const searchBtnSelectors = [
//     '#ctl04_popupBase_btnSearch',
//     'input[name*="btnSearch"]',
//     'input[value="Search"]',
//     'button[onclick*="Search"]'
// ];

// let searchBtnClicked = false;

// try {
//     const searchBtnPromises = searchBtnSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 1000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const searchBtnResults = await Promise.allSettled(searchBtnPromises);
//     const searchBtnSelector = searchBtnResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (searchBtnSelector) {
//         await newPage.click(searchBtnSelector);
//         searchBtnClicked = true;
//     }
// } catch (error) {
//     console.log('⚠️ Standard search button click failed, trying fallback...');
// }

// if (!searchBtnClicked) {
//     // Fallback click method
//     await newPage.evaluate(() => {
//         const searchBtn = document.querySelector('#ctl04_popupBase_btnSearch') ||
//                          document.querySelector('input[name*="btnSearch"]') ||
//                          document.querySelector('input[value="Search"]');
//         if (searchBtn) {
//             searchBtn.click();
//             return true;
//         }
//         return false;
//     });
// }

// // Reduced wait time from 2000ms to 1000ms for search results
// await new Promise(resolve => setTimeout(resolve, 1000));
// console.log('✅ Search completed.');

// // ✅ Wait for search results with reduced timeout
// console.log('⏳ Waiting for search result row selector to appear...');
// const resultSelectors = [
//     '#ctl04_popupBase_grvPopup_ctl02_lnkSelect',
//     'a[id*="lnkSelect"]',
//     'a[href*="Select"]',
//     'input[value="Select"]'
// ];

// let resultSelector = null;

// try {
//     const resultPromises = resultSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 5000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const resultResults = await Promise.allSettled(resultPromises);
//     resultSelector = resultResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (resultSelector) {
//         console.log(`✅ Found result selector: ${resultSelector}`);
        
//         // 🖱️ Click the first row's "Select" link
//         console.log('🖱️ Clicking first patient select link...');
//         await newPage.evaluate((selector) => {
//             const selectBtn = document.querySelector(selector);
//             if (selectBtn) {
//                 selectBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
//                 selectBtn.click();
//                 return true;
//             }
//             return false;
//         }, resultSelector);
        
//         console.log('✅ Patient selected successfully.');
//     } else {
//         console.log('❌ No search results found');
//     }
//     await new Promise(resolve => setTimeout(resolve, 1000)); // Reduced from 2500 to 600

//     // Wait for the Visit Date fields to load
// console.log('⏳ Waiting for Visit Date input fields...');
// await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Month', { visible: true, timeout: 8000 });
// await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Day', { visible: true, timeout: 8000 });
// await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Year', { visible: true, timeout: 8000 });

// // Split billingDate into MM/DD/YYYY
// const [month, day, year] = billingDate.split('/');
// console.log(`⌨️ Entering Visit Date: ${month}/${day}/${year}`);

// // Clear and type values into fields
// await page.evaluate(({ month, day, year }) => {
//     const setFieldValue = (selector, value) => {
//         const field = document.querySelector(selector);
//         if (field) {
//             field.value = '';
//             field.focus();
//             field.value = value;
//         }
//     };

//     setFieldValue('#ctl00_phFolderContent_DateVisited_Month', month);
//     setFieldValue('#ctl00_phFolderContent_DateVisited_Day', day);
//     setFieldValue('#ctl00_phFolderContent_DateVisited_Year', year);
// }, { month, day, year });

// console.log('✅ Visit Date fields populated successfully.');
// await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close

// // ⏳ Wait for popup to load
//       // ⏳ Wait for popup to load
// console.log('⏳ Waiting for Provider Search button to appear...');
// await page.waitForSelector('#ctl00_phFolderContent_Button2', { visible: true, timeout: 6000 });

// console.log('🖱️ Setting up provider popup listener...');

// // Set up popup listener BEFORE clicking the button
// const providerPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Provider popup detected!');
//         browser.off('targetcreated', onPopup);
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// // 🔘 Click provider search button
// console.log('🖱️ Clicking provider search button...');
// await page.click('#ctl00_phFolderContent_Button2');

// // Wait for popup to appear
// console.log('⏳ Waiting for provider popup to load...');
// let providerPage;
// try {
//     // Wait for popup with timeout
//     providerPage = await Promise.race([
//         providerPopupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Provider popup timeout')), 8000))
//     ]);
//     console.log('✅ Provider popup loaded successfully');
// } catch (error) {
//     console.log('⚠️ No provider popup detected, checking for new tabs...');
    
//     // Fallback: Check for new tabs
//     await new Promise(resolve => setTimeout(resolve, 3000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//         providerPage = allPages[allPages.length - 1]; // Get the newest page
//         console.log('✅ New tab detected for provider search...');
//     } else {
//         throw new Error('No provider popup or new tab found');
//     }
// }

// // Wait for the provider page to load completely
// await new Promise(resolve => setTimeout(resolve, 3000));

// const providerSearchTerm = provider.split(',')[0].trim().split(' ').slice(-1)[0] || provider;
// console.log(`⌨️ Will search for Provider: ${providerSearchTerm}`);


// // Wait for the provider search input field to appear in the NEW TAB
// console.log('⏳ Waiting for Provider search field in new tab...');
// try {
//     await providerPage.waitForSelector('input[type="text"]', { visible: true, timeout: 8000 });
//     console.log('✅ Provider search field found in new tab');
// } catch (error) {
//     console.log('🔍 Provider search field not found, trying to find any text input...');
    
//     // Try to find any text input in the new tab
//     const textInputs = await providerPage.$$('input[type="text"]');
//     if (textInputs.length === 0) {
//         throw new Error('No text input found in provider popup');
//     }
//     console.log(`Found ${textInputs.length} text input(s) in provider popup`);
// }

// // Clear and enter provider search term in the NEW TAB
// console.log(`⌨️ Entering Provider search term in new tab: ${providerSearchTerm}`);

// // Find and clear the search field
// await providerPage.evaluate(() => {
//     const searchField = document.querySelector('input[type="text"]');
//     if (searchField) {
//         searchField.value = '';
//         searchField.focus();
//         console.log('Search field cleared and focused');
//     }
// });

// // Type the provider search term
// await providerPage.type('input[type="text"]', providerSearchTerm, { delay: 100 });

// // Verify the text was entered
// const enteredText = await providerPage.$eval('input[type="text"]', el => el.value).catch(() => 'Could not verify');
// console.log('✅ Provider search text entered in new tab:', enteredText);

// // Find and click the Search button in the NEW TAB
// console.log('🔍 Looking for Search button in new tab...');
// try {
//     // Wait for search button to be available
//     await providerPage.waitForSelector('input[value="Search"]', { visible: true, timeout: 5000 });
//     console.log('🖱️ Clicking Search button in new tab...');
//     await providerPage.click('input[value="Search"]');
// } catch (error) {
//     console.log('🔍 Search button not found, trying alternatives...');
    
//     // Try clicking any button that might be the search button
//     await providerPage.evaluate(() => {
//         const searchButtons = document.querySelectorAll('input[type="button"], input[type="submit"], button');
//         for (const btn of searchButtons) {
//             if (btn.value?.toLowerCase().includes('search') || 
//                 btn.textContent?.toLowerCase().includes('search') ||
//                 btn.onclick?.toString().includes('search')) {
//                 console.log('Found search button, clicking...');
//                 btn.click();
//                 return;
//             }
//         }
//         console.log('No search button found, trying first button...');
//         if (searchButtons.length > 0) {
//             searchButtons[0].click();
//         }
//     });
// }

// // Wait for search results to load
// console.log('⏳ Waiting for search results...');
// await new Promise(resolve => setTimeout(resolve, 3000));

// // Wait for search results and select first provider
// console.log('⏳ Waiting for provider search results to appear...');
// try {
//     // Wait for any select links to appear
//     await providerPage.waitForSelector('a[href*="Select"], a[onclick*="Select"]', {
//         visible: true,
//         timeout: 10000
//     });

//     console.log('🖱️ Selecting first provider from search results...');
    
//     // Click the first Select link
//     await providerPage.evaluate(() => {
//         const selectLinks = document.querySelectorAll('a[href*="Select"], a[onclick*="Select"]');
//         if (selectLinks.length > 0) {
//             console.log(`Found ${selectLinks.length} select link(s), clicking first one...`);
//             selectLinks[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
//             selectLinks[0].click();
//         } else {
//             console.error('❌ No select links found!');
//         }
//     });

//     console.log('✅ Provider selected successfully from new tab.');

// } catch (error) {
//     console.error('🚨 Failed to find or select provider:', error.message);
    
//     // Try alternative selection - look for any clickable row
//     console.log('🔄 Trying alternative provider selection...');
//     try {
//         await providerPage.evaluate(() => {
//             // Look for table rows that might be clickable
//             const rows = document.querySelectorAll('tr, td');
//             for (const row of rows) {
//                 if (row.onclick || row.querySelector('a')) {
//                     console.log('Found clickable row, attempting click...');
//                     row.click();
//                     return;
//                 }
//             }
//         });
//     } catch (altError) {
//         console.error('🚨 Alternative provider selection failed:', altError.message);
//     }
// }

// // Wait for the selection to process and popup to close
// await new Promise(resolve => setTimeout(resolve, 3000));

// console.log('🎉 Provider selection process completed.');
// // ⏳ Wait for the Billing Info tab to be visible
// console.log('⏳ Waiting for Billing Info tab...');
// await page.waitForSelector('#VisitTabs > ul > li:nth-child(2)', { visible: true, timeout: 6000 });

// // 🖱️ Click the Billing Info tab
// console.log('🖱️ Clicking Billing Info tab...');
// await page.click('#VisitTabs > ul > li:nth-child(2)');

// console.log('✅ Billing Info tab clicked successfully.');


// } catch (error) {
//     console.error('🚨 Failed to find or click select button:', error.message);
// }

// // Optional: Wait for search results to load
// try {
//     await newPage.waitForSelector('table[style*="margin:0;padding:0;"]', { timeout: 10000 });
//     console.log('✅ Search results table loaded.');
// } catch (error) {
//     console.log('⚠️ Search results may still be loading...');
// }
// await new Promise(resolve => setTimeout(resolve, 2000)); // wait for 2s
// for (let i = 0; i < diagnosisCodes.length; i++) {
//   const code = diagnosisCodes[i];
//   const inputSelector = `#ctl00_phFolderContent_ucDiagnosisCodes_dc_10_${i + 1}`;

//   try {
//     console.log(`⏳ Processing diagnosis ${i + 1}: ${code}`);
    
//     // Wait for input field
//     await page.waitForSelector(inputSelector, { visible: true, timeout: 5000 });
//     const input = await page.$(inputSelector);
    
//     if (!input) {
//       console.warn(`⚠️ Input field not found: ${inputSelector}`);
//       continue;
//     }

//     // Clear and type
//     await input.click({ clickCount: 3 });
//     await input.type(code, { delay: 50 });

//     // Wait for autocomplete to appear (not just delay)
//     try {
//       await page.waitForSelector('.ui-autocomplete li', { visible: true, timeout: 3000 });
//     } catch (suggestionTimeout) {
//       console.warn(`⚠️ Autocomplete not shown in time for ${code}`);
//     }

//     let selected = false;

//     // Try keyboard selection first
//     try {
//       await input.focus();
//       await page.keyboard.press('ArrowDown');
//       await new Promise(resolve => setTimeout(resolve, 300));
//       await page.keyboard.press('Enter');
//       selected = true;
//       console.log(`✅ Selected via keyboard: ${code}`);
//     } catch (keyboardError) {
//       console.warn(`⚠️ Keyboard selection failed for ${code}`);

//       // Fallback: try clicking
//       try {
//         const suggestion = await page.$('.ui-autocomplete li:first-child');
//         if (suggestion) {
//           await suggestion.click();
//           selected = true;
//           console.log(`✅ Selected via click: ${code}`);
//         }
//       } catch (clickError) {
//         console.warn(`⚠️ Click selection failed for ${code}`);

//         // Fallback: Tab key
//         try {
//           await input.focus();
//           await page.keyboard.press('Tab');
//           selected = true;
//           console.log(`✅ Selected via Tab: ${code}`);
//         } catch (tabError) {
//           console.warn(`❌ All selection methods failed for: ${code}`);
//         }
//       }
//     }

//     await new Promise(resolve => setTimeout(resolve, 600));

//   } catch (error) {
//     console.error(`❌ Error processing ${code}:`, error.message);
//     await page.keyboard.press('Escape').catch(() => {});
//     await new Promise(resolve => setTimeout(resolve, 300));
//   }
// }

// console.log('🏁 Completed processing all diagnosis codes');


// // Function to extract and clean CPT codes
// function extractCPTCodes(cptInput) {
//   if (!cptInput) return [];
  
//   // Remove any non-digit characters and split by spaces, commas, or other separators
//   const cptArray = cptInput.toString()
//     .replace(/[^\d\s,]/g, '') // Remove non-digits except spaces and commas
//     .split(/[\s,]+/) // Split by spaces or commas
//     .filter(code => code.length === 5 && /^\d{5}$/.test(code)); // Only keep 5-digit codes
  
//   return cptArray;
// }

// // Extract CPT codes from the input
// const cptCodeArray = extractCPTCodes(cptCodes);
// console.log(`📋 Found ${cptCodeArray.length} valid CPT codes:`, cptCodeArray);

// if (cptCodeArray.length === 0) {
//   console.error('❌ No valid CPT codes found!');
//   return;
// }

// // Process each CPT code in sequence
// for (let i = 0; i < cptCodeArray.length; i++) {
//   const currentCPT = cptCodeArray[i];
//   console.log(`\n🔄 Processing CPT ${i + 1}/${cptCodeArray.length}: ${currentCPT}`);
  
//   // Step 1: Enter FROM date and TO date
//   console.log(`⌨️ Step 1: Entering billing dates for CPT ${currentCPT}`);
  
//   // Wait for DOS From field
//   const dosFromSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_DOS${i}`;
//   await page.waitForSelector(dosFromSelector, { visible: true, timeout: 8000 });
//   await page.focus(dosFromSelector);
//   await page.click(dosFromSelector, { clickCount: 3 });
//   await page.keyboard.type(billingDate, { delay: 50 });
//   console.log(`✅ FROM date entered: ${billingDate}`);
//   await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close


//   // Wait for DOS To field
//   const dosToSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_ToDOS${i}`;
//   await page.waitForSelector(dosToSelector, { visible: true, timeout: 8000 });
//   await page.focus(dosToSelector);
//   await page.click(dosToSelector, { clickCount: 3 });
//   await page.keyboard.type(billingDate, { delay: 50 });
//   console.log(`✅ TO date entered: ${billingDate}`);
//   await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close


//   // Step 2: Search and select CPT code
//   console.log(`🔍 Step 2: Searching for CPT code: ${currentCPT}`);
  
//   // Wait for CPT search button to appear
//   const searchButtonSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_btnUserCPT${i}`;
//   await page.waitForSelector(searchButtonSelector, { visible: true, timeout: 6000 });

//   // Set up popup listener BEFORE clicking the CPT button
//   const cptPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//       console.log('✅ CPT popup detected!');
//       browser.off('targetcreated', onPopup);
//       const popup = await target.page();
//       resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
//   });

//   // Click CPT search button
//   console.log('🖱️ Clicking CPT search button...');
//   await page.click(searchButtonSelector);

//   // Wait for popup to appear
//   console.log('⏳ Waiting for CPT popup to load...');
//   let cptPage;
//   try {
//     // Wait for popup with timeout
//     cptPage = await Promise.race([
//       cptPopupPromise,
//       new Promise((_, reject) => setTimeout(() => reject(new Error('CPT popup timeout')), 8000))
//     ]);
//     console.log('✅ CPT popup loaded successfully');
//   } catch (error) {
//     console.log('⚠️ No CPT popup detected, checking for new tabs...');
    
//     // Fallback: Check for new tabs
//     await new Promise(resolve => setTimeout(resolve, 3000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//       cptPage = allPages[allPages.length - 1]; // Get the newest page
//       console.log('✅ New tab detected for CPT search...');
//     } else {
//       throw new Error('No CPT popup or new tab found');
//     }
//   }

//   // Wait for the CPT page to load completely
//   await new Promise(resolve => setTimeout(resolve, 3000));

//   // Wait for the CPT search input field to appear in the popup
//   console.log('⏳ Waiting for CPT search field in popup...');
//   try {
//     await cptPage.waitForSelector('input[name="ctl04$popupBase$txtSearch"]', { visible: true, timeout: 8000 });
//     console.log('✅ CPT search field found in popup');
//   } catch (error) {
//     console.log('🔍 Specific CPT search field not found, trying generic text input...');
    
//     const textInputs = await cptPage.$$('input[type="text"]');
//     if (textInputs.length === 0) {
//       throw new Error('No text input found in CPT popup');
//     }
//     console.log(`Found ${textInputs.length} text input(s) in CPT popup`);
//   }

//   // Clear and enter CPT search term in the popup
//   console.log(`⌨️ Entering CPT search term in popup: ${currentCPT}`);

//   // Find and clear the search field
//   await cptPage.evaluate(() => {
//     const searchField = document.querySelector('input[name="ctl04$popupBase$txtSearch"]') || 
//                         document.querySelector('input[type="text"]');
//     if (searchField) {
//       searchField.value = '';
//       searchField.focus();
//     }
//   });

//   // Type the CPT search term
//   const searchSelector = 'input[name="ctl04$popupBase$txtSearch"]';
//   try {
//     await cptPage.type(searchSelector, currentCPT, { delay: 100 });
//   } catch (error) {
//     await cptPage.type('input[type="text"]', currentCPT, { delay: 100 });
//   }

//   // Verify the text was entered
//   const enteredText = await cptPage.evaluate(() => {
//     const field = document.querySelector('input[name="ctl04$popupBase$txtSearch"]') || 
//                     document.querySelector('input[type="text"]');
//     return field ? field.value : 'Could not verify';
//   });
//   console.log('✅ CPT search text entered in popup:', enteredText);

//   // Find and click the Search button in the popup
//   console.log('🔍 Looking for Search button in CPT popup...');
//   try {
//     await cptPage.waitForSelector('input[name="ctl04$popupBase$btnSearch"]', { visible: true, timeout: 5000 });
//     console.log('🖱️ Clicking Search button in CPT popup...');
//     await cptPage.click('input[name="ctl04$popupBase$btnSearch"]');
//   } catch (error) {
//     console.log('🔍 Specific search button not found, trying alternatives...');
    
//     await cptPage.evaluate(() => {
//       const searchButtons = document.querySelectorAll('input[value="Search"], input[type="button"], input[type="submit"], button');
//       for (const btn of searchButtons) {
//         if (btn.value?.toLowerCase().includes('search') || 
//             btn.textContent?.toLowerCase().includes('search') ||
//             btn.onclick?.toString().includes('search')) {
//           btn.click();
//           return;
//         }
//       }
//       if (searchButtons.length > 0) {
//         searchButtons[0].click();
//       }
//     });
//   }

//   // Wait for search results to load
//   console.log('⏳ Waiting for CPT search results...');
//   await new Promise(resolve => setTimeout(resolve, 3000));

//   // Select the CPT code from search results
//   console.log('⏳ Waiting for CPT search results to appear...');
//   try {
//     await cptPage.waitForSelector('#ctl04_popupBase_grvPopup_ctl02_lnkSelect', {
//       visible: true,
//       timeout: 10000
//     });

//     console.log('🖱️ Selecting CPT code from search results...');
//     await cptPage.click('#ctl04_popupBase_grvPopup_ctl02_lnkSelect');
//     console.log('✅ CPT code selected successfully');

//   } catch (error) {
//     console.error('🚨 Failed to find specific select link, trying alternatives...');
    
//     const selectLinkFound = await cptPage.evaluate(() => {
//       const selectLinks = document.querySelectorAll('a[id*="lnkSelect"]');
//       if (selectLinks.length > 0) {
//         selectLinks[0].click();
//         return true;
//       }
      
//       const allLinks = document.querySelectorAll('a');
//       for (const link of allLinks) {
//         if (link.textContent.trim().toLowerCase() === 'select') {
//           link.click();
//           return true;
//         }
//       }
      
//       const resultRows = document.querySelectorAll('table[id*="grvPopup"] tbody tr, #ctl04_popupBase_grvPopup tbody tr');
//       if (resultRows.length > 0) {
//         const firstLink = resultRows[0].querySelector('a');
//         if (firstLink) {
//           firstLink.click();
//         } else {
//           resultRows[0].click();
//         }
//         return true;
//       }
//       return false;
//     });

//     if (selectLinkFound) {
//       console.log('✅ Selected using alternative method');
//     } else {
//       console.error('🚨 Failed to select CPT code');
//     }
//   }

//   // Wait for the selection to process and popup to close
//   await new Promise(resolve => setTimeout(resolve, 3000));
//   console.log(`✅ CPT code ${currentCPT} selected and popup closed`);

//   // Step 3: Enter POS, Modifier, and ICD-10 pointer
//   console.log(`📝 Step 3: Entering additional fields for CPT ${currentCPT}`);
// await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close

//   // Enter the POS value
//   if (typeof pos === 'string' && pos.trim() !== '') {
//     console.log(`🏥 Entering POS: ${pos}`);
//     const posSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_PlaceOfService${i}`;
//     try {
//       await page.waitForSelector(posSelector, { visible: true, timeout: 5000 });
//       await page.evaluate((value, selector) => {
//         const input = document.querySelector(selector);
//         if (input) input.value = value;
//       }, pos, posSelector);
//       console.log('✅ POS entered successfully');
//     } catch (error) {
//       console.error('❌ Error entering POS:', error.message);
//     }
//   }
// await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close

//   // Enter the Modifier value
//   if (
//     modifier != null &&
//     typeof modifier === 'string' &&
//     modifier.trim() !== '' &&
//     modifier.trim() !== '-'
//   ) {
//     console.log(`🔧 Entering Modifier: ${modifier}`);
//     const modifierSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_ModifierA${i}`;
//     try {
//       await page.waitForSelector(modifierSelector, { visible: true, timeout: 5000 });
//       await page.evaluate((value, selector) => {
//         const input = document.querySelector(selector);
//         if (input) input.value = value;
//       }, modifier.trim(), modifierSelector);
//       console.log('✅ Modifier entered successfully');
//     } catch (error) {
//       console.error('❌ Error entering Modifier:', error.message);
//     }
//   }
// await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close

//   // Set ICD-10 Pointer
//   try {
//     const pointerSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_DiagnosisCode${i}`;
//     let icdPointer = '';

//     if (diagnosisCodes && diagnosisCodes.length > 0) {
//       if (diagnosisCodes.length === 1) {
//         icdPointer = '1';
//       } else if (diagnosisCodes.length === 2) {
//         icdPointer = '12';
//       } else if (diagnosisCodes.length === 3) {
//         icdPointer = '123';
//       } else if (diagnosisCodes.length >= 4) {
//         icdPointer = '1234';
//       }

//       console.log(`🩺 Setting ICD-10 Pointer to: ${icdPointer}`);
//       await page.waitForSelector(pointerSelector, { visible: true, timeout: 5000 });
//       await page.evaluate((value, selector) => {
//         const input = document.querySelector(selector);
//         if (input) input.value = value;
//       }, icdPointer, pointerSelector);
//       console.log('✅ ICD-10 pointer set successfully');
//     }

//   } catch (error) {
//     console.error('❌ Error setting ICD-10 pointer:', error.message);
//   }
// await new Promise(resolve => setTimeout(resolve, 1000));
//   console.log(`🎉 CPT ${currentCPT} processing completed for row ${i}`);

//   // Add delay before processing next CPT (if any)
//   if (i < cptCodeArray.length - 1) {
//     console.log('⏳ Preparing for next CPT code...');
//     await new Promise(resolve => setTimeout(resolve, 2000));
//   }
// }
// // ⏳ Wait for the Billing Info tab to be visible
// console.log('⏳ Waiting for Billing Options tab...');
// await page.waitForSelector('#VisitTabs > ul > li:nth-child(3)', { visible: true, timeout: 6000 });

// // 🖱️ Click the Billing Info tab
// console.log('🖱️ Clicking Billing Info tab...');
// await page.click('#VisitTabs > ul > li:nth-child(3)');

// console.log('✅ Billing Options tab clicked successfully.');
// await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close

// // ⏳ Wait for popup to load
// console.log('⏳ Waiting for Provider Search button to appear...');
// await page.waitForSelector('#div1 > table > tbody > tr:nth-child(2) > td:nth-child(2) > input.button', { visible: true, timeout: 6000 });

// console.log('🖱️ Setting up provider popup listener...');

// // Set up popup listener BEFORE clicking the button
// const providerPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Provider popup detected!');
//         browser.off('targetcreated', onPopup);
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// // 🔘 Click provider search button
// console.log('🖱️ Clicking provider search button...');
// await page.click('#div1 > table > tbody > tr:nth-child(2) > td:nth-child(2) > input.button');

// // Wait for popup to appear
// console.log('⏳ Waiting for provider popup to load...');
// let providerPage;
// try {
//     // Wait for popup with timeout
//     providerPage = await Promise.race([
//         providerPopupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Provider popup timeout')), 8000))
//     ]);
//     console.log('✅ Provider popup loaded successfully');
// } catch (error) {
//     console.log('⚠️ No provider popup detected, checking for new tabs...');
    
//     // Fallback: Check for new tabs
//     await new Promise(resolve => setTimeout(resolve, 3000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//         providerPage = allPages[allPages.length - 1]; // Get the newest page
//         console.log('✅ New tab detected for provider search...');
//     } else {
//         throw new Error('No provider popup or new tab found');
//     }
// }

// // Wait for the provider page to load completely
// await new Promise(resolve => setTimeout(resolve, 3000));

// const providerSearchTerm = provider.split(',')[0].trim().split(' ').slice(-1)[0] || provider;
// console.log(`⌨️ Will search for Provider: ${providerSearchTerm}`);


// // Wait for the specific provider search input field to appear
// console.log('⏳ Waiting for Provider search field (#ctl04_popupBase_txtSearch)...');
// try {
//     await providerPage.waitForSelector('#ctl04_popupBase_txtSearch', { visible: true, timeout: 8000 });
//     console.log('✅ Provider search field found');
// } catch (error) {
//     console.log('🔍 Specific search field not found, trying fallback selectors...');
    
//     // Try alternative selectors
//     const fallbackSelectors = [
//         'input[name*="txtSearch"]',
//         'input[id*="txtSearch"]',
//         'input[type="text"]'
//     ];
    
//     let fieldFound = false;
//     for (const selector of fallbackSelectors) {
//         try {
//             await providerPage.waitForSelector(selector, { visible: true, timeout: 3000 });
//             console.log(`✅ Found search field with selector: ${selector}`);
//             fieldFound = true;
//             break;
//         } catch (e) {
//             continue;
//         }
//     }
    
//     if (!fieldFound) {
//         throw new Error('No provider search field found');
//     }
// }

// // Clear and enter provider search term
// console.log(`⌨️ Entering Provider search term: ${providerSearchTerm}`);

// try {
//     // First try the specific selector
//     await providerPage.evaluate(() => {
//         const searchField = document.querySelector('#ctl04_popupBase_txtSearch');
//         if (searchField) {
//             searchField.value = '';
//             searchField.focus();
//             console.log('Search field cleared and focused');
//             return true;
//         }
//         return false;
//     });
    
//     await providerPage.type('#ctl04_popupBase_txtSearch', providerSearchTerm, { delay: 100 });
// } catch (error) {
//     // Fallback to any text input
//     console.log('🔄 Using fallback method to enter text...');
//     await providerPage.evaluate(() => {
//         const searchField = document.querySelector('input[type="text"]');
//         if (searchField) {
//             searchField.value = '';
//             searchField.focus();
//         }
//     });
    
//     await providerPage.type('input[type="text"]', providerSearchTerm, { delay: 100 });
// }

// // Verify the text was entered
// const enteredText1 = await providerPage.evaluate(() => {
//     const field = document.querySelector('#ctl04_popupBase_txtSearch') || 
//                   document.querySelector('input[type="text"]');
//     return field ? field.value : 'Could not verify';
// });
// console.log('✅ Provider search text entered:', enteredText1);

// // Find and click the Search button
// console.log('🔍 Looking for Search button...');
// try {
//     // Wait for search button to be available - try multiple selectors
//     const searchButtonSelectors = [
//         'input[value="Search"]',
//         'button[onclick*="Search"]',
//         '#ctl04_popupBase_btnSearch',
//         'input[id*="Search"]',
//         'input[name*="Search"]'
//     ];
    
//     let searchButtonFound = false;
//     for (const selector of searchButtonSelectors) {
//         try {
//             await providerPage.waitForSelector(selector, { visible: true, timeout: 2000 });
//             console.log(`🖱️ Clicking Search button with selector: ${selector}`);
//             await providerPage.click(selector);
//             searchButtonFound = true;
//             break;
//         } catch (e) {
//             continue;
//         }
//     }
    
//     if (!searchButtonFound) {
//         // Try clicking any button that might be the search button
//         await providerPage.evaluate(() => {
//             const searchButtons = document.querySelectorAll('input[type="button"], input[type="submit"], button');
//             for (const btn of searchButtons) {
//                 if (btn.value?.toLowerCase().includes('search') || 
//                     btn.textContent?.toLowerCase().includes('search') ||
//                     btn.onclick?.toString().includes('search')) {
//                     console.log('Found search button by content, clicking...');
//                     btn.click();
//                     return true;
//                 }
//             }
//             return false;
//         });
//     }
// } catch (error) {
//     console.log('🚨 Search button click failed:', error.message);
// }

// // Wait for search results to load
// console.log('⏳ Waiting for search results...');
// await new Promise(resolve => setTimeout(resolve, 4000));

// // Wait for search results and select first provider
// console.log('⏳ Waiting for provider search results to appear...');
// try {
//     // Wait for the results table or select links to appear
//     const resultSelectors = [
//         'a[onclick*="Select"]',
//         'a[href*="Select"]',
//         'input[value="Select"]',
//         'button[onclick*="Select"]',
//         '.grid tbody tr', // Table rows that might contain select options
//         'table tr td a' // Links within table cells
//     ];
    
//     let resultsFound = false;
//     for (const selector of resultSelectors) {
//         try {
//             await providerPage.waitForSelector(selector, { visible: true, timeout: 3000 });
//             console.log(`✅ Found search results with selector: ${selector}`);
//             resultsFound = true;
//             break;
//         } catch (e) {
//             continue;
//         }
//     }
    
//     if (resultsFound) {
//         console.log('🖱️ Selecting first provider from search results...');
        
//         // Try to click the first Select link/button
//         const selectionSuccess = await providerPage.evaluate(() => {
//             // Look for Select links first
//             const selectLinks = document.querySelectorAll('a[onclick*="Select"], a[href*="Select"]');
//             if (selectLinks.length > 0) {
//                 console.log(`Found ${selectLinks.length} select link(s), clicking first one...`);
//                 selectLinks[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
//                 selectLinks[0].click();
//                 return true;
//             }
            
//             // Look for Select buttons
//             const selectButtons = document.querySelectorAll('input[value="Select"], button[onclick*="Select"]');
//             if (selectButtons.length > 0) {
//                 console.log(`Found ${selectButtons.length} select button(s), clicking first one...`);
//                 selectButtons[0].click();
//                 return true;
//             }
            
//             // Look for any clickable rows in results table
//             const rows = document.querySelectorAll('table tr');
//             for (let i = 1; i < rows.length; i++) { // Skip header row
//                 const row = rows[i];
//                 const selectElement = row.querySelector('a, button, input[type="button"]');
//                 if (selectElement) {
//                     console.log('Found clickable element in row, clicking...');
//                     selectElement.click();
//                     return true;
//                 }
//             }
            
//             return false;
//         });
        
//         if (selectionSuccess) {
//             console.log('✅ Provider selected successfully.');
//         } else {
//             console.log('⚠️ Could not find selectable provider in results');
//         }
//     } else {
//         console.log('⚠️ No search results found');
//     }

// } catch (error) {
//     console.error('🚨 Failed to find or select provider:', error.message);
    
//     // Final fallback - try to click any link in the page
//     console.log('🔄 Trying final fallback selection...');
//     try {
//         await providerPage.evaluate(() => {
//             const allLinks = document.querySelectorAll('a, button, input[type="button"]');
//             for (const link of allLinks) {
//                 if (link.textContent?.toLowerCase().includes('select') || 
//                     link.value?.toLowerCase().includes('select') ||
//                     link.onclick?.toString().includes('select')) {
//                     console.log('Found potential select element, clicking...');
//                     link.click();
//                     return;
//                 }
//             }
//         });
//     } catch (altError) {
//         console.error('🚨 Final fallback selection failed:', altError.message);
//     }
// }

// // Wait for the selection to process and popup to close
// await new Promise(resolve => setTimeout(resolve, 5000));

// console.log('🎉 Provider selection process completed.');

// // ⏳ Wait for popup trigger button
// console.log('⏳ Waiting for Facility Search button...');
// await page.waitForSelector('#ctl00_phFolderContent_Button35', { visible: true, timeout: 6000 });

// console.log('🖱️ Setting up facility popup listener...');

// // Set up popup listener BEFORE clicking the button
// const facilityPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Facility popup detected!');
//         browser.off('targetcreated', onPopup);
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// // 🔘 Click facility search button
// console.log('🖱️ Clicking facility search button...');
// await page.click('#ctl00_phFolderContent_Button35');

// // Wait for popup
// console.log('⏳ Waiting for facility popup to load...');
// let facilityPage;
// try {
//     facilityPage = await Promise.race([
//         facilityPopupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Facility popup timeout')), 8000))
//     ]);
//     console.log('✅ Facility popup loaded');
// } catch (error) {
//     const allPages = await browser.pages();
//     if (allPages.length > 1) {
//         facilityPage = allPages[allPages.length - 1];
//         console.log('✅ Detected new tab as facility popup');
//     } else {
//         throw new Error('❌ No popup or new tab for facility found');
//     }
// }

// // ⏳ Wait for search field
// console.log('⏳ Waiting for facility search input...');
// await facilityPage.waitForSelector('#ctl04_popupBase_txtSearch', { visible: true, timeout: 6000 });

// // 💬 Enter facility name
// console.log(`⌨️ Entering Facility Name: ${facilityName}`);
// await facilityPage.evaluate((name) => {
//     const input = document.querySelector('#ctl04_popupBase_txtSearch');
//     if (input) {
//         input.value = '';
//         input.focus();
//     }
// }, facilityName);

// await facilityPage.type('#ctl04_popupBase_txtSearch', facilityName, { delay: 100 });

// // 🔍 Click Search button
// console.log('🖱️ Clicking facility search button...');
// await facilityPage.waitForSelector('#ctl04_popupBase_btnSearch', { visible: true, timeout: 4000 });
// await facilityPage.click('#ctl04_popupBase_btnSearch');

// // ⏳ Wait for results
// console.log('⏳ Waiting for facility search results...');
// await facilityPage.waitForSelector('a[id^="ctl04_popupBase_grvPopup_ctl"][id$="_lnkSelect"]', { visible: true, timeout: 6000 });
// await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close

// // 🖱️ Click the first "Select" link
// console.log('🖱️ Selecting the first facility...');
// await facilityPage.evaluate(() => {
//     const selectLink = document.querySelector('a[id^="ctl04_popupBase_grvPopup_ctl"][id$="_lnkSelect"]');
//     if (selectLink) {
//         selectLink.scrollIntoView({ behavior: 'smooth', block: 'center' });
//         selectLink.click();
//     }
// });

// // ✅ Done
// console.log('🎉 Facility selection completed!');
// await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close

// console.log('🔍 Processing complete. Please verify all entries on the form.');
// if (['21', '22', '31'].includes(pos)) {
//   console.log(`🩺 POS is ${pos}, entering Hospital From Date: ${admitDate}`);

//   try {
//     const [month, day, year] = admitDate.split('/');

//     // Month field
//     const monthSelector = '#ctl00_phFolderContent_HospitalFromDate_Month';
//     await page.waitForSelector(monthSelector, { visible: true, timeout: 10000 });
//     await page.focus(monthSelector);
//     await page.type(monthSelector, month, { delay: 50 });
//     await page.keyboard.press('Tab');

//     // Day field
//     const daySelector = '#ctl00_phFolderContent_HospitalFromDate_Day';
//     await page.waitForSelector(daySelector, { visible: true, timeout: 3000 });
//     await page.focus(daySelector);
//     await page.type(daySelector, day, { delay: 50 });
//     await page.keyboard.press('Tab');

//     // Year field
//     const yearSelector = '#ctl00_phFolderContent_HospitalFromDate_Year';
//     await page.waitForSelector(yearSelector, { visible: true, timeout: 3000 });
//     await page.focus(yearSelector);
//     await page.type(yearSelector, year, { delay: 50 });

//     // ✅ Do NOT press Tab or click again — just let it be
//     console.log('✅ Hospital From Date filled successfully.');
//   } catch (error) {
//     console.log(`❌ Error filling Hospital From Date: ${error.message}`);
//   }
// } else {
//   console.log(`ℹ️ POS is ${pos}. Skipping Hospital From Date entry.`);
// }

// await new Promise(resolve => setTimeout(resolve, 1500)); // Shorter wait


// // Handle Prior Authorization Number
// if (priornumber && priornumber.trim() !== '') {
//     console.log(`📝 Entering Prior Authorization Number: ${priornumber}`);
    
//     try {
//         // Wait for the prior auth field with longer timeout
//         await page.waitForSelector('#ctl00_phFolderContent_PriorAuthorizationNumber', { visible: true, timeout: 10000 });

//         // Scroll to the field to ensure it's visible
//         await page.evaluate(() => {
//             const priorField = document.querySelector('#ctl00_phFolderContent_PriorAuthorizationNumber');
//             if (priorField) {
//                 priorField.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         // Wait for scroll to complete
//         await new Promise(resolve => setTimeout(resolve, 1000));

//         // Click on the field to focus it
//         await page.click('#ctl00_phFolderContent_PriorAuthorizationNumber');

//         // Clear the field using keyboard shortcuts
//         await page.keyboard.down('Control');
//         await page.keyboard.press('KeyA');
//         await page.keyboard.up('Control');
//         await page.keyboard.press('Delete');

//         // Type the prior authorization number
//         await page.type('#ctl00_phFolderContent_PriorAuthorizationNumber', priornumber);

//         // Verify the value was entered
//         const enteredValue = await page.$eval('#ctl00_phFolderContent_PriorAuthorizationNumber', el => el.value);
//         console.log(`✅ Prior Authorization Number entered. Value: ${enteredValue}`);

//         // Trigger change event to ensure the form recognizes the input
//         await page.evaluate(() => {
//             const priorField = document.querySelector('#ctl00_phFolderContent_PriorAuthorizationNumber');
//             if (priorField) {
//                 priorField.dispatchEvent(new Event('change', { bubbles: true }));
//                 priorField.dispatchEvent(new Event('input', { bubbles: true }));
//             }
//         });

//     } catch (error) {
//         console.log(`❌ Error entering Prior Authorization Number: ${error.message}`);
        
//         // Try alternative approach if the main selector fails
//         try {
//             console.log('🔄 Trying alternative approach for Prior Auth field...');
            
//             // Try to find the field by partial name match
//             const alternativeSelector = 'input[name*="PriorAuthorizationNumber"]';
//             await page.waitForSelector(alternativeSelector, { visible: true, timeout: 5000 });
            
//             await page.click(alternativeSelector);
//             await page.keyboard.down('Control');
//             await page.keyboard.press('KeyA');
//             await page.keyboard.up('Control');
//             await page.keyboard.press('Delete');
//             await page.type(alternativeSelector, priornumber);
            
//             console.log('✅ Prior Authorization Number entered using alternative selector.');
            
//         } catch (altError) {
//             console.log(`❌ Alternative approach also failed: ${altError.message}`);
//         }
//     }
// } else {
//     console.log('ℹ️ Prior Authorization Number not provided. Skipping field.');
// }

// // Final wait before proceeding
// await new Promise(resolve => setTimeout(resolve, 1000));
// // ✅ Click the update button after processing
// // Wait for the update button to appear
// await page.waitForSelector('#ctl00_phFolderContent_btnUpdate', { visible: true });

// // Click the button using Puppeteer
// await page.click('#ctl00_phFolderContent_btnUpdate');
// console.log("🚀 Update button clicked.");
// await new Promise(resolve => setTimeout(resolve, 1000));

// return { success: true };

// } catch (error) {
//         console.error('Error processing account:', error);
//         return false; // Return failure
//     }
// }


// // Endpoint to process data and return status
// app.post("/process", upload.single("file"), async (req, res) => {
//     if (!req.file || !req.body.originalPath) {
//         return res.status(400).json({ error: "No file or original path provided" });
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
        
//         // Check if "Patient ID" exists in headers
//         console.log(headers);
//         if (!headers.includes("Patient ID")) {
//             throw new Error("'Patient ID' Column is not defined.");
//         }
        
//         // Map headers to their index positions
//         headers.forEach((header, index) => {
//             columnMap[header] = index;
//         });
        
//         const requiredColumns = ["Patient ID", "Provider", "Facility", "DOS"];
//         for (const column of requiredColumns) {
//             if (!headers.includes(column)) {
//                 throw new Error(`Required column '${column}' is missing from the spreadsheet.`);
//             }
//         }
    
//   // Ensure only "Result" column is added if missing
//     const resultHeader = "Result";
//     if (!headers.includes(resultHeader)) {
//         columnMap[resultHeader] = headers.length;
//         headers.push(resultHeader);
//         for (let i = 1; i < originalData.length; i++) {
//             originalData[i] = originalData[i] || [];
//             originalData[i].push("");
//         }
//     }

//     headers.forEach((header, idx) => columnMap[header] = idx);
    
//         let newRowsToProcess = 0;
//         const processedAccounts = new Set();
        
//         let missingCallDiagnoses = 0;  // Track rows missing Diagnoses

//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue; // Skip undefined rows
            
//             const accountNumberRaw = row[columnMap["Patient ID"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
//             if (!accountNumberRaw || status === "done" || status === "failed") continue;  // Skip processed or failed rows

            
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             processedAccounts.add(accountNumber);
//             newRowsToProcess++;
        
//             // Check if Diagnoses is missing
//             const callDiagnoses = row[columnMap["Diagnoses"]] || "";
//             if (!callDiagnoses.trim()) {
//                 missingCallDiagnoses++;
//             }
//         }
        
//         // Check if all remaining rows are missing Diagnoses
//         if (newRowsToProcess > 0 && newRowsToProcess === missingCallDiagnoses) {
//             return res.json({
//                 success: false,
//                 message: "Diagnoses not found in Excel.",
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
        
//         newPage = await continueWithLoggedInSession();

        
//         const results = { successful: [], failed: [] };
        
//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue;
        
//             const accountNumberRaw = row[columnMap["Patient ID"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
//             if (!accountNumberRaw || status === "done") continue;
        
//             const provider = row[columnMap["Provider"]] || "";
//             const facilityName = row[columnMap["Facility"]] || "";
//             const billingDate = row[columnMap["DOS"]] || "";
//             const diagnosisText = row[columnMap["Diagnoses"]] || "";
//             const chargesText = row[columnMap["CPT"]] || ""; // CPT codes
//             const admitDate = row[columnMap["Hospital Date"]] || ""; // ✅ NEW LINE
//             const pos = row[columnMap["POS"]] || ""; // ✅ NEW LINE
//             const modifier = row[columnMap["Modifier"]] || ""; // ✅ NEW LINE
//             const priornumber = row[columnMap["Prior Authorization Number"]] || ""; // ✅ NEW LINE
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             console.log(`Processing account: ${accountNumber}`);
        
//             // Validate data before processing
//             if (!provider) console.log('Provider value is missing or undefined');
//             if (!facilityName) console.log('Facility is missing or undefined');
//             if (!billingDate) console.log('DOS is missing or undefined');
//             if (!admitDate) console.log('Hospital Date is missing or undefined'); // ✅ Optional

//           // Array to hold valid diagnosis codes
// const diagnosisCodes = [];

// // Regex:
// // - ICD-10-CM: Starts with a letter, 2–4 digits, optional dot, followed by 1–4 alphanumerics
// // - ICD-10-PCS: Exactly 7 alphanumeric characters (A-Z, 0-9)
// const dxCodeRegex = /\b(?:[A-Z][0-9]{2,4}(?:\.[0-9A-Z]{1,4})?|[A-Z0-9]{7})\b/g;

// // Match codes
// const matches = diagnosisText.match(dxCodeRegex);

// if (matches) {
//   // Normalize: trim, uppercase, remove duplicates
//   const uniqueCodes = [...new Set(matches.map(code => code.trim().toUpperCase()))];

//   // Filter: Must contain at least one letter and one digit
//   const validCodes = uniqueCodes.filter(code => /[A-Z]/i.test(code) && /\d/.test(code));

//   diagnosisCodes.push(...validCodes);

//   // Output
//   console.log("✅ Final Unique, Valid Diagnosis Codes:");
//   diagnosisCodes.forEach((code, index) => {
//     console.log(`  ${index + 1}. ${code}`);
//   });
//   console.log("🔢 Total:", diagnosisCodes.length);
// } else {
//   console.log("⚠️ No valid diagnosis codes found.");
// }
          
          
// // Normalize input: remove non-breaking spaces, tabs, newlines, etc.
// const cleanedText = chargesText.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();

// const cptCodes = [];
// const validCptRegex = /\b\d{5}\b/g;
// let cptMatch;

// // Extract all 5-digit CPT codes
// while ((cptMatch = validCptRegex.exec(cleanedText)) !== null) {
//   const baseCode = cptMatch[0].replace(/\D/g, '').trim();  // extra clean
//   cptCodes.push(baseCode);
//   console.log(`🧩 CPT Code: ${baseCode}`);
// }

// // Remove duplicates using Set (case-insensitive, clean)
// const uniqueCptCodes = [...new Set(cptCodes.map(code => code.trim()))];

// // Final output
// console.log('✅ Extracted Unique CPT Codes:', uniqueCptCodes);
// uniqueCptCodes.forEach((code, index) => {
//   console.log(`  ${index + 1}. ${code}`);
// });


//           // === Call processAccount (Don't send CPT length)
// const success = await processAccount(
//     newPage,
//     newPage.browser(), // Add browser parameter
//     accountNumber,
//     provider,
//     facilityName,
//     billingDate,
//     diagnosisCodes,
//     cptCodes, // ✅ only the array
//     admitDate,
//     pos,
//     modifier,
//     priornumber
// );

// // Set Result
// if (success) {
//     results.successful.push(accountNumber);
//     row[columnMap["Result"]] = "Done";
// } else {
//     results.failed.push(accountNumber);
//     row[columnMap["Result"]] = "Failed";
// }

//         }            
//         // Convert data back to an ExcelJS workbook
//        const workbook = new ExcelJS.Workbook();
// const sheet = workbook.addWorksheet("Sheet1");

// // Write headers with gray background
// const headerRow = sheet.addRow(headers);
// headerRow.height = 42;
// headerRow.eachCell((cell) => {
//     cell.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: "A9A9A9" }
//     };
//     cell.font = { bold: true, color: { argb: "000000" }, size: 10 };
//     cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
//     cell.border = {
//         top: { style: "thin" },
//         left: { style: "thin" },
//         bottom: { style: "thin" },
//         right: { style: "thin" }
//     };
// });

// // Define column widths (adjusted to remove Billed Fee and Claim Number)
// const columnWidths = [6, 11, 30, 15, 15, 23, 54, 10,10, 11, 50, 20, 30]; 
// // You may want to adjust the size or remove two widths if those columns were exclusively for Claim Number and Billed Fee

// sheet.columns = columnWidths.map((width, index) => ({
//     width,
//     style: (index === columnMap["Diagnoses"] || index === columnMap["CPT"])
//         ? { alignment: { wrapText: true } }
//         : {},
// }));

// originalData.slice(1).forEach((row) => {
//     if (!row) return;

//     const hasData = row.some(cell => cell && cell.toString().trim() !== "");
//     if (!hasData) return;

//     const rowData = headers.map(header => row[columnMap[header]] || "");
//     const newRow = sheet.addRow(rowData);
//     newRow.height = 100;

//     headers.forEach((header, colIndex) => {
//         const cell = newRow.getCell(colIndex + 1);

//         // Apply border
//         cell.border = {
//             top: { style: "thin", color: { argb: "000000" } },
//             left: { style: "thin", color: { argb: "000000" } },
//             bottom: { style: "thin", color: { argb: "000000" } },
//             right: { style: "thin", color: { argb: "000000" } }
//         };

//         cell.font = { size: 10 };

//         // Format Patient ID
//         if (header === "Patient ID" && cell.value) {
//             cell.fill = {
//                 type: "pattern",
//                 pattern: "solid",
//                 fgColor: { argb: "FFFFFF" }
//             };
//             cell.alignment = { horizontal: "center", vertical: "bottom" };
//         }

//         // Format Result
//         if (header === "Result" && cell.value) {
//             const value = cell.value.toString().toLowerCase();
//             if (value === "done") {
//                 cell.fill = {
//                     type: "pattern",
//                     pattern: "solid",
//                     fgColor: { argb: "92D050" } // Green
//                 };
//             } else if (value === "failed") {
//                 cell.fill = {
//                     type: "pattern",
//                     pattern: "solid",
//                     fgColor: { argb: "FF0000" } // Red
//                 };
//             }
//             cell.alignment = { horizontal: "center", vertical: "middle" };
//         }
//     });
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
//             message: `${results.successful.length} Charges Entered successfully. ${results.failed.length} rows failed.`,
//             results,
//             fileId
//         });

//     } catch (error) {
//         console.error("Error:", error);
//         res.status(500).json({ error: error.message });
//     } finally {
//         if (newPage) {
//             await new Promise(resolve => setTimeout(resolve, 2000));
//             await newPage.close();
//             console.log("Browser closed");
//         }
//     }
// });

// // Separate endpoint to download the processed file
// app.get("/download/:fileId", (req, res) => {
//   const fileId = req.params.fileId;
//   const fileData = processedFiles.get(fileId);
  
//   if (!fileData) {
//       return res.status(404).json({ error: "File not found" });
//   }
  
//   // Determine content type based on file extension
//   let contentType = 'application/octet-stream'; // Default
  
//   if (fileData.filename.endsWith('.xlsx')) {
//       contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
//   } else if (fileData.filename.endsWith('.xls')) {
//       contentType = 'application/vnd.ms-excel';
//   } else if (fileData.filename.endsWith('.csv')) {
//       contentType = 'text/csv';
//   }
  
//   // Set the response headers
//   res.setHeader('Content-Disposition', `attachment; filename="${fileData.filename}"`);
//   res.setHeader('Content-Type', contentType);
//   res.setHeader('Content-Length', fileData.buffer.length);
  
//   // Send the file buffer
//   res.send(fileData.buffer);
  
//   // Optional: Remove the file from memory after some time
// // Optional: Remove the file from memory after some time
// setTimeout(() => {
//     processedFiles.delete(fileId);
//     console.log(`File ${fileId} removed from memory`);
// }, 3 * 60 * 60 * 1000); // Remove after 3 hours

// });

// app.listen(port, () => console.log(`Server running on http://localhost:${port}`));	// 07/07/2025






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



// async function continueWithLoggedInSession() {
//     try {
//         const browser = await puppeteer.connect({
//             browserURL: 'http://localhost:9222', // Connect to existing Chrome
//             defaultViewport: null,
//             headless: false,
//             timeout: 120000,
//         });

//         const pages = await browser.pages();
//         const sessionPage = pages.find(p =>
//             p.url().includes('Appointments/ViewAppointments.aspx')
//         );

//         if (!sessionPage) {
//             console.log('❌ No matching tab found for the appointment page.');

//             if (pages.length > 0) {
//                 // Show alert in the first available page
//                 await pages[0].bringToFront();
//                 await pages[0].evaluate(() => {
//                     alert('⚠️ Session not found');
//                 });
//             } else {
//                 console.log('⚠️ No pages available to show alert.');
//             }

//             return null;
//         }

//         console.log('✅ Found existing page. Waiting for Patient Visit tab...');
//         await sessionPage.waitForSelector('#patient-visits_tab > span', {
//             visible: true,
//             timeout: 10000,
//         });

//         console.log('🖱️ Clicking Patient Visit tab...');
//         await sessionPage.evaluate(() => {
//             const tab = document.querySelector('#patient-visits_tab > span');
//             if (tab) {
//                 tab.scrollIntoView({ behavior: 'smooth', block: 'center' });
//                 tab.click();
//             }
//         });

//         console.log('✅ Patient Visit tab clicked successfully.');
//         return sessionPage;
//     } catch (error) {
//         console.error('🚨 Error in continueWithLoggedInSession:', error.message);
//         return null;
//     }
// }

// async function processAccount(page,browser,accountNumber, provider, facilityName, billingDate, diagnosisCodes,cptCodes,admitDate,pos,modifier,priornumber,cptUnits) {
//    try {
//       // Wait and click Add New Visit
// // Wait and click Add New Visit
// await new Promise(resolve => setTimeout(resolve, 3000));
// console.log('⏳ Waiting for Add New Visit button to appear...');
// await page.waitForSelector('#addNewVisit', { visible: true, timeout: 15000 });

// console.log('🖱️ Setting up popup listener...');

// // Set up popup listener BEFORE clicking the button
// const popupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Popup detected!');
//         browser.off('targetcreated', onPopup); // Fixed: Use .off() instead of .removeListener()
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// console.log('🖱️ Clicking Add New Visit button...');
// await page.evaluate(() => {
//     const btn = document.querySelector('#addNewVisit');
//     if (btn) {
//         btn.scrollIntoView({ behavior: 'smooth', block: 'center' });  
//         btn.click();
//     } else {
//         console.error('❌ Add New Visit button not found!');
//     }
// });
//  // ⏳ Wait for popup to load
//         console.log('⏳ Waiting for Search button to appear...');
//         await page.waitForSelector('#ctl00_phFolderContent_Button1', { visible: true, timeout: 9000 });

//         // 🔘 Click initial popup search button
//         console.log('🖱️ Clicking popup search open button...');
//         await page.click('#ctl00_phFolderContent_Button1');
// // Wait for popup to appear with reduced timeout
// console.log('⏳ Waiting for popup to load...');
// let newPage;
// try {
//     // Reduced timeout from 6000ms to 3000ms
//     newPage = await Promise.race([
//         popupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Popup timeout')), 3000))
//     ]);
//     console.log('✅ Popup loaded successfully');
// } catch (error) {
//     console.log('⚠️ No popup detected, checking for new tabs...');
    
//     // Reduced wait time from 2000ms to 1000ms
//     await new Promise(resolve => setTimeout(resolve, 1000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//         newPage = allPages[allPages.length - 1];
//         console.log('✅ New tab detected, using it...');
//     } else {
//         newPage = page;
//         console.log('📝 No new tab, using current page...');
//     }
// }

// // Reduced wait time from 3000ms to 1500ms
// await new Promise(resolve => setTimeout(resolve, 1500));

// // ⏳ Wait for Search button with reduced timeout and parallel checking
// console.log('⏳ Waiting for Search button to appear in new page...');
// const searchButtonSelectors = [
//     '#ctl00_phFolderContent_Button1',
//     'input[value="Search"]',
//     'button[onclick*="Search"]',
//     '.button:contains("Search")',
//     '#Button1'
// ];

// let searchButtonFound = false;
// let workingSelector = null;

// // Check all selectors in parallel instead of sequentially
// try {
//     const promises = searchButtonSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 3000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const results = await Promise.allSettled(promises);
//     workingSelector = results.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (workingSelector) {
//         console.log(`✅ Found search button with selector: ${workingSelector}`);
//         searchButtonFound = true;
//     }
// } catch (error) {
//     console.log('🔍 No search button found with standard selectors');
// }

// if (!searchButtonFound) {
//     throw new Error('Search button not found with any selector');
// }

// // 🔘 Click initial popup search button
// console.log('🖱️ Clicking popup search open button...');
// try {
//     await newPage.click(workingSelector);
// } catch (error) {
//     // Fallback click method
//     await newPage.evaluate(() => {
//         const searchButtons = document.querySelectorAll('input[value*="Search"], button');
//         for (const btn of searchButtons) {
//             if (btn.value?.includes('Search') || btn.textContent?.includes('Search')) {
//                 btn.click();
//                 break;
//             }
//         }
//     });
// }

// // Wait for dropdown with parallel selector checking
// console.log('⏳ Waiting for dropdown to appear...');
// const dropdownSelectors = [
//     '#ctl04_popupBase_ddlSearch',
//     'select[name*="ddlSearch"]',
//     'select[id*="Search"]',
//     'select:first-of-type'
// ];

// let dropdownSelector = null;

// try {
//     const dropdownPromises = dropdownSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 2000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const dropdownResults = await Promise.allSettled(dropdownPromises);
//     dropdownSelector = dropdownResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (dropdownSelector) {
//         console.log(`✅ Found dropdown with selector: ${dropdownSelector}`);
//     } else {
//         throw new Error('Dropdown not found');
//     }
// } catch (error) {
//     console.log('❌ Dropdown not found with any selector');
//     throw error;
// }

// // Reduced wait time from 1000ms to 500ms
// await new Promise(resolve => setTimeout(resolve, 500));

// // 🔽 Select "Patient ID" with improved logic
// console.log('🔽 Selecting "Patient ID" from dropdown...');
// await newPage.evaluate((selector) => {
//     const dropdown = document.querySelector(selector);
    
//     if (dropdown) {
//         // Find Patient ID option more efficiently
//         let patientIdIndex = -1;
//         const options = Array.from(dropdown.options);
        
//         // Search for Patient ID option
//         patientIdIndex = options.findIndex(option => 
//             option.text.toLowerCase().includes('patient id') || 
//             option.value.toLowerCase().includes('patientid') ||
//             option.text.toLowerCase().includes('patient')
//         );
        
//         if (patientIdIndex !== -1) {
//             dropdown.selectedIndex = patientIdIndex;
//             console.log(`Selected Patient ID option at index ${patientIdIndex}: ${options[patientIdIndex].text}`);
//         } else {
//             // Fallback to index 2 (3rd option) if Patient ID not found
//             dropdown.selectedIndex = Math.min(2, options.length - 1);
//             console.log(`Fallback: Selected option at index ${dropdown.selectedIndex}: ${options[dropdown.selectedIndex].text}`);
//         }
        
//         // Trigger change event
//         dropdown.dispatchEvent(new Event('change', { bubbles: true }));
//         return true;
//     }
//     return false;
// }, dropdownSelector);

// // Reduced wait time from 500ms to 200ms
// await new Promise(resolve => setTimeout(resolve, 200));

// // Enhanced text field handling with parallel selector checking
// console.log(`⌨️ Entering Patient ID: ${accountNumber}`);
// const textFieldSelectors = [
//     '#ctl04_popupBase_txtSearch',
//     'input[name*="txtSearch"]',
//     'input[type="text"]:not([style*="display: none"])',
//     'input[maxlength="50"]'
// ];

// let textFieldSelector = null;

// try {
//     const textPromises = textFieldSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 2000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const textResults = await Promise.allSettled(textPromises);
//     textFieldSelector = textResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (!textFieldSelector) {
//         throw new Error('Text field not found');
//     }
// } catch (error) {
//     console.log('❌ Text field not found with any selector');
//     throw error;
// }

// // Clear and enter text more efficiently
// await newPage.evaluate((selector) => {
//     const textField = document.querySelector(selector);
//     if (textField) {
//         textField.value = '';
//         textField.focus();
//     }
// }, textFieldSelector);

// await newPage.type(textFieldSelector, accountNumber, { delay: 50 }); // Reduced delay from 100ms to 50ms

// // 🔍 Click the search button with parallel selector checking
// console.log('🖱️ Clicking search button...');
// const searchBtnSelectors = [
//     '#ctl04_popupBase_btnSearch',
//     'input[name*="btnSearch"]',
//     'input[value="Search"]',
//     'button[onclick*="Search"]'
// ];

// let searchBtnClicked = false;

// try {
//     const searchBtnPromises = searchBtnSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 1000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const searchBtnResults = await Promise.allSettled(searchBtnPromises);
//     const searchBtnSelector = searchBtnResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (searchBtnSelector) {
//         await newPage.click(searchBtnSelector);
//         searchBtnClicked = true;
//     }
// } catch (error) {
//     console.log('⚠️ Standard search button click failed, trying fallback...');
// }

// if (!searchBtnClicked) {
//     // Fallback click method
//     await newPage.evaluate(() => {
//         const searchBtn = document.querySelector('#ctl04_popupBase_btnSearch') ||
//                          document.querySelector('input[name*="btnSearch"]') ||
//                          document.querySelector('input[value="Search"]');
//         if (searchBtn) {
//             searchBtn.click();
//             return true;
//         }
//         return false;
//     });
// }

// // Reduced wait time from 2000ms to 1000ms for search results
// await new Promise(resolve => setTimeout(resolve, 1000));
// console.log('✅ Search completed.');

// // ✅ Wait for search results with reduced timeout
// console.log('⏳ Waiting for search result row selector to appear...');
// const resultSelectors = [
//     '#ctl04_popupBase_grvPopup_ctl02_lnkSelect',
//     'a[id*="lnkSelect"]',
//     'a[href*="Select"]',
//     'input[value="Select"]'
// ];

// let resultSelector = null;

// try {
//     const resultPromises = resultSelectors.map(selector => 
//         newPage.waitForSelector(selector, { visible: true, timeout: 5000 })
//             .then(() => selector)
//             .catch(() => null)
//     );
    
//     const resultResults = await Promise.allSettled(resultPromises);
//     resultSelector = resultResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
//     if (resultSelector) {
//         console.log(`✅ Found result selector: ${resultSelector}`);
        
//         // 🖱️ Click the first row's "Select" link
//         console.log('🖱️ Clicking first patient select link...');
//         await newPage.evaluate((selector) => {
//             const selectBtn = document.querySelector(selector);
//             if (selectBtn) {
//                 selectBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
//                 selectBtn.click();
//                 return true;
//             }
//             return false;
//         }, resultSelector);
        
//         console.log('✅ Patient selected successfully.');
//     } else {
//         console.log('❌ No search results found');
//     }
//     await new Promise(resolve => setTimeout(resolve, 1000)); // Reduced from 2500 to 600

//     // Wait for the Visit Date fields to load
// console.log('⏳ Waiting for Visit Date input fields...');
// await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Month', { visible: true, timeout: 8000 });
// await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Day', { visible: true, timeout: 8000 });
// await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Year', { visible: true, timeout: 8000 });

// // Split billingDate into MM/DD/YYYY
// const [month, day, year] = billingDate.split('/');
// console.log(`⌨️ Entering Visit Date: ${month}/${day}/${year}`);

// // Clear and type values into fields
// await page.evaluate(({ month, day, year }) => {
//     const setFieldValue = (selector, value) => {
//         const field = document.querySelector(selector);
//         if (field) {
//             field.value = '';
//             field.focus();
//             field.value = value;
//         }
//     };

//     setFieldValue('#ctl00_phFolderContent_DateVisited_Month', month);
//     setFieldValue('#ctl00_phFolderContent_DateVisited_Day', day);
//     setFieldValue('#ctl00_phFolderContent_DateVisited_Year', year);
// }, { month, day, year });

// console.log('✅ Visit Date fields populated successfully.');
// await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close

// // ⏳ Wait for popup to load
//       // ⏳ Wait for popup to load
// console.log('⏳ Waiting for Provider Search button to appear...');
// await page.waitForSelector('#ctl00_phFolderContent_Button2', { visible: true, timeout: 6000 });

// console.log('🖱️ Setting up provider popup listener...');

// // Set up popup listener BEFORE clicking the button
// const providerPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Provider popup detected!');
//         browser.off('targetcreated', onPopup);
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// // 🔘 Click provider search button
// console.log('🖱️ Clicking provider search button...');
// await page.click('#ctl00_phFolderContent_Button2');

// // Wait for popup to appear
// console.log('⏳ Waiting for provider popup to load...');
// let providerPage;
// try {
//     // Wait for popup with timeout
//     providerPage = await Promise.race([
//         providerPopupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Provider popup timeout')), 8000))
//     ]);
//     console.log('✅ Provider popup loaded successfully');
// } catch (error) {
//     console.log('⚠️ No provider popup detected, checking for new tabs...');
    
//     // Fallback: Check for new tabs
//     await new Promise(resolve => setTimeout(resolve, 3000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//         providerPage = allPages[allPages.length - 1]; // Get the newest page
//         console.log('✅ New tab detected for provider search...');
//     } else {
//         throw new Error('No provider popup or new tab found');
//     }
// }

// // Wait for the provider page to load completely
// await new Promise(resolve => setTimeout(resolve, 3000));

// const providerSearchTerm = provider.split(',')[0].trim().split(' ').slice(-1)[0] || provider;
// console.log(`⌨️ Will search for Provider: ${providerSearchTerm}`);


// // Wait for the provider search input field to appear in the NEW TAB
// console.log('⏳ Waiting for Provider search field in new tab...');
// try {
//     await providerPage.waitForSelector('input[type="text"]', { visible: true, timeout: 8000 });
//     console.log('✅ Provider search field found in new tab');
// } catch (error) {
//     console.log('🔍 Provider search field not found, trying to find any text input...');
    
//     // Try to find any text input in the new tab
//     const textInputs = await providerPage.$$('input[type="text"]');
//     if (textInputs.length === 0) {
//         throw new Error('No text input found in provider popup');
//     }
//     console.log(`Found ${textInputs.length} text input(s) in provider popup`);
// }

// // Clear and enter provider search term in the NEW TAB
// console.log(`⌨️ Entering Provider search term in new tab: ${providerSearchTerm}`);

// // Find and clear the search field
// await providerPage.evaluate(() => {
//     const searchField = document.querySelector('input[type="text"]');
//     if (searchField) {
//         searchField.value = '';
//         searchField.focus();
//         console.log('Search field cleared and focused');
//     }
// });

// // Type the provider search term
// await providerPage.type('input[type="text"]', providerSearchTerm, { delay: 100 });

// // Verify the text was entered
// const enteredText = await providerPage.$eval('input[type="text"]', el => el.value).catch(() => 'Could not verify');
// console.log('✅ Provider search text entered in new tab:', enteredText);

// // Find and click the Search button in the NEW TAB
// console.log('🔍 Looking for Search button in new tab...');
// try {
//     // Wait for search button to be available
//     await providerPage.waitForSelector('input[value="Search"]', { visible: true, timeout: 5000 });
//     console.log('🖱️ Clicking Search button in new tab...');
//     await providerPage.click('input[value="Search"]');
// } catch (error) {
//     console.log('🔍 Search button not found, trying alternatives...');
    
//     // Try clicking any button that might be the search button
//     await providerPage.evaluate(() => {
//         const searchButtons = document.querySelectorAll('input[type="button"], input[type="submit"], button');
//         for (const btn of searchButtons) {
//             if (btn.value?.toLowerCase().includes('search') || 
//                 btn.textContent?.toLowerCase().includes('search') ||
//                 btn.onclick?.toString().includes('search')) {
//                 console.log('Found search button, clicking...');
//                 btn.click();
//                 return;
//             }
//         }
//         console.log('No search button found, trying first button...');
//         if (searchButtons.length > 0) {
//             searchButtons[0].click();
//         }
//     });
// }

// // Wait for search results to load
// console.log('⏳ Waiting for search results...');
// await new Promise(resolve => setTimeout(resolve, 3000));

// // Wait for search results and select first provider
// console.log('⏳ Waiting for provider search results to appear...');
// try {
//     // Wait for any select links to appear
//     await providerPage.waitForSelector('a[href*="Select"], a[onclick*="Select"]', {
//         visible: true,
//         timeout: 10000
//     });

//     console.log('🖱️ Selecting first provider from search results...');
    
//     // Click the first Select link
//     await providerPage.evaluate(() => {
//         const selectLinks = document.querySelectorAll('a[href*="Select"], a[onclick*="Select"]');
//         if (selectLinks.length > 0) {
//             console.log(`Found ${selectLinks.length} select link(s), clicking first one...`);
//             selectLinks[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
//             selectLinks[0].click();
//         } else {
//             console.error('❌ No select links found!');
//         }
//     });

//     console.log('✅ Provider selected successfully from new tab.');

// } catch (error) {
//     console.error('🚨 Failed to find or select provider:', error.message);
    
//     // Try alternative selection - look for any clickable row
//     console.log('🔄 Trying alternative provider selection...');
//     try {
//         await providerPage.evaluate(() => {
//             // Look for table rows that might be clickable
//             const rows = document.querySelectorAll('tr, td');
//             for (const row of rows) {
//                 if (row.onclick || row.querySelector('a')) {
//                     console.log('Found clickable row, attempting click...');
//                     row.click();
//                     return;
//                 }
//             }
//         });
//     } catch (altError) {
//         console.error('🚨 Alternative provider selection failed:', altError.message);
//     }
// }

// // Wait for the selection to process and popup to close
// await new Promise(resolve => setTimeout(resolve, 3000));

// console.log('🎉 Provider selection process completed.');
// // ⏳ Wait for the Billing Info tab to be visible
// console.log('⏳ Waiting for Billing Info tab...');
// await page.waitForSelector('#VisitTabs > ul > li:nth-child(2)', { visible: true, timeout: 6000 });

// // 🖱️ Click the Billing Info tab
// console.log('🖱️ Clicking Billing Info tab...');
// await page.click('#VisitTabs > ul > li:nth-child(2)');

// console.log('✅ Billing Info tab clicked successfully.');


// } catch (error) {
//     console.error('🚨 Failed to find or click select button:', error.message);
// }

// // Optional: Wait for search results to load
// try {
//     await newPage.waitForSelector('table[style*="margin:0;padding:0;"]', { timeout: 10000 });
//     console.log('✅ Search results table loaded.');
// } catch (error) {
//     console.log('⚠️ Search results may still be loading...');
// }
// await new Promise(resolve => setTimeout(resolve, 2000)); // wait for 2s
// for (let i = 0; i < diagnosisCodes.length; i++) {
//   const code = diagnosisCodes[i];
//   const inputSelector = `#ctl00_phFolderContent_ucDiagnosisCodes_dc_10_${i + 1}`;

//   try {
//     console.log(`⏳ Processing diagnosis ${i + 1}: ${code}`);
    
//     // Wait for input field
//     await page.waitForSelector(inputSelector, { visible: true, timeout: 5000 });
//     const input = await page.$(inputSelector);
    
//     if (!input) {
//       console.warn(`⚠️ Input field not found: ${inputSelector}`);
//       continue;
//     }

//     // Clear and type
//     await input.click({ clickCount: 3 });
//     await input.type(code, { delay: 50 });

//     // Wait for autocomplete to appear (not just delay)
//     try {
//       await page.waitForSelector('.ui-autocomplete li', { visible: true, timeout: 3000 });
//     } catch (suggestionTimeout) {
//       console.warn(`⚠️ Autocomplete not shown in time for ${code}`);
//     }

//     let selected = false;

//     // Try keyboard selection first
//     try {
//       await input.focus();
//       await page.keyboard.press('ArrowDown');
//       await new Promise(resolve => setTimeout(resolve, 300));
//       await page.keyboard.press('Enter');
//       selected = true;
//       console.log(`✅ Selected via keyboard: ${code}`);
//     } catch (keyboardError) {
//       console.warn(`⚠️ Keyboard selection failed for ${code}`);

//       // Fallback: try clicking
//       try {
//         const suggestion = await page.$('.ui-autocomplete li:first-child');
//         if (suggestion) {
//           await suggestion.click();
//           selected = true;
//           console.log(`✅ Selected via click: ${code}`);
//         }
//       } catch (clickError) {
//         console.warn(`⚠️ Click selection failed for ${code}`);

//         // Fallback: Tab key
//         try {
//           await input.focus();
//           await page.keyboard.press('Tab');
//           selected = true;
//           console.log(`✅ Selected via Tab: ${code}`);
//         } catch (tabError) {
//           console.warn(`❌ All selection methods failed for: ${code}`);
//         }
//       }
//     }

//     await new Promise(resolve => setTimeout(resolve, 600));

//   } catch (error) {
//     console.error(`❌ Error processing ${code}:`, error.message);
//     await page.keyboard.press('Escape').catch(() => {});
//     await new Promise(resolve => setTimeout(resolve, 300));
//   }
// }

// console.log('🏁 Completed processing all diagnosis codes');


// // Function to extract and clean CPT codes
// function extractCPTCodes(cptInput) {
//   if (!cptInput) return [];
  
//   // Remove any non-digit characters and split by spaces, commas, or other separators
//   const cptArray = cptInput.toString()
//     .replace(/[^\d\s,]/g, '') // Remove non-digits except spaces and commas
//     .split(/[\s,]+/) // Split by spaces or commas
//     .filter(code => code.length === 5 && /^\d{5}$/.test(code)); // Only keep 5-digit codes
  
//   return cptArray;
// }

// // Extract CPT codes from the input
// const cptCodeArray = extractCPTCodes(cptCodes);
// console.log(`📋 Found ${cptCodeArray.length} valid CPT codes:`, cptCodeArray);

// if (cptCodeArray.length === 0) {
//   console.error('❌ No valid CPT codes found!');
//   return;
// }

// // Process each CPT code in sequence
// for (let i = 0; i < cptCodeArray.length; i++) {
//   const currentCPT = cptCodeArray[i];
//   console.log(`\n🔄 Processing CPT ${i + 1}/${cptCodeArray.length}: ${currentCPT}`);
  
//   // Step 1: Enter FROM date and TO date
//   console.log(`⌨️ Step 1: Entering billing dates for CPT ${currentCPT}`);
  
//   // Wait for DOS From field
//   const dosFromSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_DOS${i}`;
//   await page.waitForSelector(dosFromSelector, { visible: true, timeout: 8000 });
//   await page.focus(dosFromSelector);
//   await page.click(dosFromSelector, { clickCount: 3 });
//   await page.keyboard.type(billingDate, { delay: 50 });
//   console.log(`✅ FROM date entered: ${billingDate}`);
//   await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close


//   // Wait for DOS To field
//   const dosToSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_ToDOS${i}`;
//   await page.waitForSelector(dosToSelector, { visible: true, timeout: 8000 });
//   await page.focus(dosToSelector);
//   await page.click(dosToSelector, { clickCount: 3 });
//   await page.keyboard.type(billingDate, { delay: 50 });
//   console.log(`✅ TO date entered: ${billingDate}`);
//   await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close


//   // Step 2: Search and select CPT code
//   console.log(`🔍 Step 2: Searching for CPT code: ${currentCPT}`);
  
//   // Wait for CPT search button to appear
//   const searchButtonSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_btnUserCPT${i}`;
//   await page.waitForSelector(searchButtonSelector, { visible: true, timeout: 6000 });

//   // Set up popup listener BEFORE clicking the CPT button
//   const cptPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//       console.log('✅ CPT popup detected!');
//       browser.off('targetcreated', onPopup);
//       const popup = await target.page();
//       resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
//   });

//   // Click CPT search button
//   console.log('🖱️ Clicking CPT search button...');
//   await page.click(searchButtonSelector);

//   // Wait for popup to appear
//   console.log('⏳ Waiting for CPT popup to load...');
//   let cptPage;
//   try {
//     // Wait for popup with timeout
//     cptPage = await Promise.race([
//       cptPopupPromise,
//       new Promise((_, reject) => setTimeout(() => reject(new Error('CPT popup timeout')), 8000))
//     ]);
//     console.log('✅ CPT popup loaded successfully');
//   } catch (error) {
//     console.log('⚠️ No CPT popup detected, checking for new tabs...');
    
//     // Fallback: Check for new tabs
//     await new Promise(resolve => setTimeout(resolve, 3000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//       cptPage = allPages[allPages.length - 1]; // Get the newest page
//       console.log('✅ New tab detected for CPT search...');
//     } else {
//       throw new Error('No CPT popup or new tab found');
//     }
//   }

//   // Wait for the CPT page to load completely
//   await new Promise(resolve => setTimeout(resolve, 3000));

//   // Wait for the CPT search input field to appear in the popup
//   console.log('⏳ Waiting for CPT search field in popup...');
//   try {
//     await cptPage.waitForSelector('input[name="ctl04$popupBase$txtSearch"]', { visible: true, timeout: 8000 });
//     console.log('✅ CPT search field found in popup');
//   } catch (error) {
//     console.log('🔍 Specific CPT search field not found, trying generic text input...');
    
//     const textInputs = await cptPage.$$('input[type="text"]');
//     if (textInputs.length === 0) {
//       throw new Error('No text input found in CPT popup');
//     }
//     console.log(`Found ${textInputs.length} text input(s) in CPT popup`);
//   }

//   // Clear and enter CPT search term in the popup
//   console.log(`⌨️ Entering CPT search term in popup: ${currentCPT}`);

//   // Find and clear the search field
//   await cptPage.evaluate(() => {
//     const searchField = document.querySelector('input[name="ctl04$popupBase$txtSearch"]') || 
//                         document.querySelector('input[type="text"]');
//     if (searchField) {
//       searchField.value = '';
//       searchField.focus();
//     }
//   });

//   // Type the CPT search term
//   const searchSelector = 'input[name="ctl04$popupBase$txtSearch"]';
//   try {
//     await cptPage.type(searchSelector, currentCPT, { delay: 100 });
//   } catch (error) {
//     await cptPage.type('input[type="text"]', currentCPT, { delay: 100 });
//   }

//   // Verify the text was entered
//   const enteredText = await cptPage.evaluate(() => {
//     const field = document.querySelector('input[name="ctl04$popupBase$txtSearch"]') || 
//                     document.querySelector('input[type="text"]');
//     return field ? field.value : 'Could not verify';
//   });
//   console.log('✅ CPT search text entered in popup:', enteredText);

//   // Find and click the Search button in the popup
//   console.log('🔍 Looking for Search button in CPT popup...');
//   try {
//     await cptPage.waitForSelector('input[name="ctl04$popupBase$btnSearch"]', { visible: true, timeout: 5000 });
//     console.log('🖱️ Clicking Search button in CPT popup...');
//     await cptPage.click('input[name="ctl04$popupBase$btnSearch"]');
//   } catch (error) {
//     console.log('🔍 Specific search button not found, trying alternatives...');
    
//     await cptPage.evaluate(() => {
//       const searchButtons = document.querySelectorAll('input[value="Search"], input[type="button"], input[type="submit"], button');
//       for (const btn of searchButtons) {
//         if (btn.value?.toLowerCase().includes('search') || 
//             btn.textContent?.toLowerCase().includes('search') ||
//             btn.onclick?.toString().includes('search')) {
//           btn.click();
//           return;
//         }
//       }
//       if (searchButtons.length > 0) {
//         searchButtons[0].click();
//       }
//     });
//   }

//   // Wait for search results to load
//   console.log('⏳ Waiting for CPT search results...');
//   await new Promise(resolve => setTimeout(resolve, 3000));

//   // Select the CPT code from search results
//   console.log('⏳ Waiting for CPT search results to appear...');
//   try {
//     await cptPage.waitForSelector('#ctl04_popupBase_grvPopup_ctl02_lnkSelect', {
//       visible: true,
//       timeout: 10000
//     });

//     console.log('🖱️ Selecting CPT code from search results...');
//     await cptPage.click('#ctl04_popupBase_grvPopup_ctl02_lnkSelect');
//     console.log('✅ CPT code selected successfully');

//   } catch (error) {
//     console.error('🚨 Failed to find specific select link, trying alternatives...');
    
//     const selectLinkFound = await cptPage.evaluate(() => {
//       const selectLinks = document.querySelectorAll('a[id*="lnkSelect"]');
//       if (selectLinks.length > 0) {
//         selectLinks[0].click();
//         return true;
//       }
      
//       const allLinks = document.querySelectorAll('a');
//       for (const link of allLinks) {
//         if (link.textContent.trim().toLowerCase() === 'select') {
//           link.click();
//           return true;
//         }
//       }
      
//       const resultRows = document.querySelectorAll('table[id*="grvPopup"] tbody tr, #ctl04_popupBase_grvPopup tbody tr');
//       if (resultRows.length > 0) {
//         const firstLink = resultRows[0].querySelector('a');
//         if (firstLink) {
//           firstLink.click();
//         } else {
//           resultRows[0].click();
//         }
//         return true;
//       }
//       return false;
//     });

//     if (selectLinkFound) {
//       console.log('✅ Selected using alternative method');
//     } else {
//       console.error('🚨 Failed to select CPT code');
//     }
//   }

//   // Wait for the selection to process and popup to close
//   await new Promise(resolve => setTimeout(resolve, 3000));
//   console.log(`✅ CPT code ${currentCPT} selected and popup closed`);

//   // Step 3: Enter POS, Modifier, and ICD-10 pointer
//   console.log(`📝 Step 3: Entering additional fields for CPT ${currentCPT}`);
// await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close

//   // Enter the POS value
//   if (typeof pos === 'string' && pos.trim() !== '') {
//     console.log(`🏥 Entering POS: ${pos}`);
//     const posSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_PlaceOfService${i}`;
//     try {
//       await page.waitForSelector(posSelector, { visible: true, timeout: 5000 });
//       await page.evaluate((value, selector) => {
//         const input = document.querySelector(selector);
//         if (input) input.value = value;
//       }, pos, posSelector);
//       console.log('✅ POS entered successfully');
//     } catch (error) {
//       console.error('❌ Error entering POS:', error.message);
//     }
//   }
// await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close

//   // Enter the Modifier value
//   if (
//     modifier != null &&
//     typeof modifier === 'string' &&
//     modifier.trim() !== '' &&
//     modifier.trim() !== '-'
//   ) {
//     console.log(`🔧 Entering Modifier: ${modifier}`);
//     const modifierSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_ModifierA${i}`;
//     try {
//       await page.waitForSelector(modifierSelector, { visible: true, timeout: 5000 });
//       await page.evaluate((value, selector) => {
//         const input = document.querySelector(selector);
//         if (input) input.value = value;
//       }, modifier.trim(), modifierSelector);
//       console.log('✅ Modifier entered successfully');
//     } catch (error) {
//       console.error('❌ Error entering Modifier:', error.message);
//     }
//   }
// await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close

//   // Set ICD-10 Pointer
//   try {
//     const pointerSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_DiagnosisCode${i}`;
//     let icdPointer = '';

//     if (diagnosisCodes && diagnosisCodes.length > 0) {
//       if (diagnosisCodes.length === 1) {
//         icdPointer = '1';
//       } else if (diagnosisCodes.length === 2) {
//         icdPointer = '12';
//       } else if (diagnosisCodes.length === 3) {
//         icdPointer = '123';
//       } else if (diagnosisCodes.length >= 4) {
//         icdPointer = '1234';
//       }

//       console.log(`🩺 Setting ICD-10 Pointer to: ${icdPointer}`);
//       await page.waitForSelector(pointerSelector, { visible: true, timeout: 5000 });
//       await page.evaluate((value, selector) => {
//         const input = document.querySelector(selector);
//         if (input) input.value = value;
//       }, icdPointer, pointerSelector);
//       console.log('✅ ICD-10 pointer set successfully');
//     }

//   } catch (error) {
//     console.error('❌ Error setting ICD-10 pointer:', error.message);
//   }
//   // ✅ Enter Days or Units field
// const unitSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_Quantity${i}`;
// const unitValue = cptUnits[i] || '';  // default to 1 if not found

// console.log(`📦 Entering Units for CPT ${currentCPT}: ${unitValue}`);
// await page.waitForSelector(unitSelector, { visible: true, timeout: 5000 });
// await page.evaluate((value, selector) => {
//   const input = document.querySelector(selector);
//   if (input) {
//     input.focus();
//     input.value = value;
//   }
// }, unitValue, unitSelector);
// console.log('✅ Units field set successfully');

// await new Promise(resolve => setTimeout(resolve, 1000));
//   console.log(`🎉 CPT ${currentCPT} processing completed for row ${i}`);

//   // Add delay before processing next CPT (if any)
//   if (i < cptCodeArray.length - 1) {
//     console.log('⏳ Preparing for next CPT code...');
//     await new Promise(resolve => setTimeout(resolve, 2000));
//   }
// }
// // ⏳ Wait for the Billing Info tab to be visible
// console.log('⏳ Waiting for Billing Options tab...');
// await page.waitForSelector('#VisitTabs > ul > li:nth-child(3)', { visible: true, timeout: 6000 });

// // 🖱️ Click the Billing Info tab
// console.log('🖱️ Clicking Billing Info tab...');
// await page.click('#VisitTabs > ul > li:nth-child(3)');

// console.log('✅ Billing Options tab clicked successfully.');
// await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close
// // ⏳ Wait for popup to load
// console.log('⏳ Waiting for Provider Search button to appear...');
// await page.waitForSelector('#div1 > table > tbody > tr:nth-child(2) > td:nth-child(2) > input.button', { visible: true, timeout: 6000 });

// console.log('🖱️ Setting up provider popup listener...');

// // Set up popup listener BEFORE clicking the button
// const providerPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Provider popup detected!');
//         browser.off('targetcreated', onPopup);
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// // 🔘 Click provider search button
// console.log('🖱️ Clicking provider search button...');
// await page.click('#div1 > table > tbody > tr:nth-child(2) > td:nth-child(2) > input.button');

// // Wait for popup to appear
// console.log('⏳ Waiting for provider popup to load...');
// let providerPage;
// try {
//     // Wait for popup with timeout
//     providerPage = await Promise.race([
//         providerPopupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Provider popup timeout')), 8000))
//     ]);
//     console.log('✅ Provider popup loaded successfully');
// } catch (error) {
//     console.log('⚠️ No provider popup detected, checking for new tabs...');
    
//     // Fallback: Check for new tabs
//     await new Promise(resolve => setTimeout(resolve, 3000));
//     const allPages = await browser.pages();
    
//     if (allPages.length > 1) {
//         providerPage = allPages[allPages.length - 1]; // Get the newest page
//         console.log('✅ New tab detected for provider search...');
//     } else {
//         throw new Error('No provider popup or new tab found');
//     }
// }

// // Wait for the provider page to load completely
// await new Promise(resolve => setTimeout(resolve, 3000));

// // Extract FIRST NAMES from provider list
// let providerFirstNames = [];
// if (typeof provider !== 'undefined' && provider) {
//     // Split by newlines or commas to get individual providers
//     const providers = provider.split(/\n|,(?=\s*[A-Z])/); // Split by newlines or commas followed by capital letter
    
//     providers.forEach(providerEntry => {
//         const trimmedProvider = providerEntry.trim();
//         if (trimmedProvider) {
//             // Format: "FirstName LastName, MD" - e.g., "Michael Lin, MD"
//             if (trimmedProvider.includes(',')) {
//                 // Split by comma and take the part before comma
//                 const beforeComma = trimmedProvider.split(',')[0].trim();
//                 // Get first word (first name)
//                 const firstName = beforeComma.split(' ')[0];
//                 providerFirstNames.push(firstName);
//             } else {
//                 // If no comma, just take first word
//                 const firstName = trimmedProvider.split(' ')[0];
//                 providerFirstNames.push(firstName);
//             }
//         }
//     });
    
//     console.log(`📝 Extracted FIRST NAMES from providers: ${providerFirstNames.join(', ')}`);
// }

// // If you want just the first provider's first name:
// let providerFirstName = providerFirstNames.length > 0 ? providerFirstNames[0] : '';

// // NEW: Select "First Name" from dropdown
// console.log('🔍 Looking for search type dropdown to select First Name...');
// try {
//     // Wait for the dropdown to be available
//     await new Promise(resolve => setTimeout(resolve, 2000));
    
//     const searchTypeDropdownSelectors = [
//         'select[name*="LastName"]', // This might be the dropdown name even if it has First Name option
//         'select:first-of-type',
//         'select'
//     ];
    
//     let searchTypeDropdown = null;
//     for (const selector of searchTypeDropdownSelectors) {
//         try {
//             await providerPage.waitForSelector(selector, { visible: true, timeout: 2000 });
//             searchTypeDropdown = selector;
//             console.log(`✅ Found search type dropdown with selector: ${selector}`);
//             break;
//         } catch (e) {
//             continue;
//         }
//     }
    
//     if (searchTypeDropdown) {
//         // Get available options from the dropdown
//         const availableOptions = await providerPage.evaluate((selector) => {
//             const dropdown = document.querySelector(selector);
//             if (!dropdown) return [];
            
//             const options = [];
//             for (let i = 0; i < dropdown.options.length; i++) {
//                 const option = dropdown.options[i];
//                 options.push({
//                     value: option.value,
//                     text: option.text
//                 });
//             }
//             return options;
//         }, searchTypeDropdown);
        
//         console.log('📋 Available search type options:', availableOptions);
        
//         // Select "First Name" option
//         const firstNameOption = availableOptions.find(opt => 
//             opt.text.toLowerCase().includes('first name') || 
//             opt.text.toLowerCase().includes('firstname') ||
//             opt.value.toLowerCase().includes('firstname') ||
//             opt.value.toLowerCase().includes('first')
//         );
        
//         if (firstNameOption) {
//             console.log(`🎯 Selecting search type: ${firstNameOption.text} (value: ${firstNameOption.value})`);
//             await providerPage.select(searchTypeDropdown, firstNameOption.value);
//             console.log('✅ First Name selected from dropdown');
//         } else {
//             console.log('⚠️ First Name option not found in dropdown, using default');
//         }
//     }
    
//     // Wait a moment for the selection to process
//     await new Promise(resolve => setTimeout(resolve, 1000));
    
// } catch (error) {
//     console.log('🚨 Error selecting First Name from dropdown:', error.message);
// }

// // Wait for the specific provider search input field to appear
// console.log('⏳ Waiting for Provider search field...');
// try {
//     const searchFieldSelectors = [
//         'input[type="text"]',
//         'input[name*="txtSearch"]',
//         'input[id*="txtSearch"]',
//         'input[name*="Search"]'
//     ];
    
//     let searchField = null;
//     for (const selector of searchFieldSelectors) {
//         try {
//             await providerPage.waitForSelector(selector, { visible: true, timeout: 3000 });
//             searchField = selector;
//             console.log(`✅ Found search field with selector: ${selector}`);
//             break;
//         } catch (e) {
//             continue;
//         }
//     }
    
//     if (!searchField) {
//         throw new Error('No provider search field found');
//     }
    
//     // Clear and enter provider FIRST NAME
//     console.log(`⌨️ Entering Provider FIRST NAME: ${providerFirstName}`);
    
//     // Clear the field first
//     await providerPage.evaluate((selector) => {
//         const field = document.querySelector(selector);
//         if (field) {
//             field.value = '';
//             field.focus();
//         }
//     }, searchField);
    
//     // Type the first name
//     await providerPage.type(searchField, providerFirstName, { delay: 100 });
    
//     // Verify the text was entered
//     const enteredText = await providerPage.evaluate((selector) => {
//         const field = document.querySelector(selector);
//         return field ? field.value : 'Could not verify';
//     }, searchField);
//     console.log('✅ Provider FIRST NAME entered:', enteredText);
    
// } catch (error) {
//     console.log('🚨 Error with search field:', error.message);
//     throw error;
// }

// // Find and click the Search button
// console.log('🔍 Looking for Search button...');
// try {
//     const searchButtonSelectors = [
//         'input[value="Search"]',
//         'button[onclick*="Search"]',
//         'input[type="button"][value*="Search"]',
//         'input[id*="Search"]',
//         'input[name*="Search"]'
//     ];
    
//     let searchButtonFound = false;
//     for (const selector of searchButtonSelectors) {
//         try {
//             await providerPage.waitForSelector(selector, { visible: true, timeout: 2000 });
//             console.log(`🖱️ Clicking Search button with selector: ${selector}`);
//             await providerPage.click(selector);
//             searchButtonFound = true;
//             break;
//         } catch (e) {
//             continue;
//         }
//     }
    
//     if (!searchButtonFound) {
//         // Try clicking any button that might be the search button
//         await providerPage.evaluate(() => {
//             const searchButtons = document.querySelectorAll('input[type="button"], input[type="submit"], button');
//             for (const btn of searchButtons) {
//                 if (btn.value?.toLowerCase().includes('search') || 
//                     btn.textContent?.toLowerCase().includes('search') ||
//                     btn.onclick?.toString().includes('search')) {
//                     console.log('Found search button by content, clicking...');
//                     btn.click();
//                     return true;
//                 }
//             }
//             return false;
//         });
//     }
// } catch (error) {
//     console.log('🚨 Search button click failed:', error.message);
// }

// // Wait for search results to load
// console.log('⏳ Waiting for search results...');
// await new Promise(resolve => setTimeout(resolve, 4000));

// // Wait for search results and select first provider
// console.log('⏳ Waiting for provider search results to appear...');
// try {
//     // Wait for the results table or select links to appear
//     const resultSelectors = [
//         'a[onclick*="Select"]',
//         'a[href*="Select"]',
//         'input[value="Select"]',
//         'button[onclick*="Select"]',
//         '.grid tbody tr', // Table rows that might contain select options
//         'table tr td a' // Links within table cells
//     ];
    
//     let resultsFound = false;
//     for (const selector of resultSelectors) {
//         try {
//             await providerPage.waitForSelector(selector, { visible: true, timeout: 3000 });
//             console.log(`✅ Found search results with selector: ${selector}`);
//             resultsFound = true;
//             break;
//         } catch (e) {
//             continue;
//         }
//     }
    
//     if (resultsFound) {
//         console.log('🖱️ Selecting first provider from search results...');
        
//         // Try to click the first Select link/button
//         const selectionSuccess = await providerPage.evaluate(() => {
//             // Look for Select links first
//             const selectLinks = document.querySelectorAll('a[onclick*="Select"], a[href*="Select"]');
//             if (selectLinks.length > 0) {
//                 console.log(`Found ${selectLinks.length} select link(s), clicking first one...`);
//                 selectLinks[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
//                 selectLinks[0].click();
//                 return true;
//             }
            
//             // Look for Select buttons
//             const selectButtons = document.querySelectorAll('input[value="Select"], button[onclick*="Select"]');
//             if (selectButtons.length > 0) {
//                 console.log(`Found ${selectButtons.length} select button(s), clicking first one...`);
//                 selectButtons[0].click();
//                 return true;
//             }
            
//             // Look for any clickable rows in results table
//             const rows = document.querySelectorAll('table tr');
//             for (let i = 1; i < rows.length; i++) { // Skip header row
//                 const row = rows[i];
//                 const selectElement = row.querySelector('a, button, input[type="button"]');
//                 if (selectElement) {
//                     console.log('Found clickable element in row, clicking...');
//                     selectElement.click();
//                     return true;
//                 }
//             }
            
//             return false;
//         });
        
//         if (selectionSuccess) {
//             console.log('✅ Provider selected successfully.');
//         } else {
//             console.log('⚠️ Could not find selectable provider in results');
//         }
//     } else {
//         console.log('⚠️ No search results found');
//     }

// } catch (error) {
//     console.error('🚨 Failed to find or select provider:', error.message);
    
//     // Final fallback - try to click any link in the page
//     console.log('🔄 Trying final fallback selection...');
//     try {
//         await providerPage.evaluate(() => {
//             const allLinks = document.querySelectorAll('a, button, input[type="button"]');
//             for (const link of allLinks) {
//                 if (link.textContent?.toLowerCase().includes('select') || 
//                     link.value?.toLowerCase().includes('select') ||
//                     link.onclick?.toString().includes('select')) {
//                     console.log('Found potential select element, clicking...');
//                     link.click();
//                     return;
//                 }
//             }
//         });
//     } catch (altError) {
//         console.error('🚨 Final fallback selection failed:', altError.message);
//     }
// }

// // Wait for the selection to process and popup to close
// await new Promise(resolve => setTimeout(resolve, 5000));

// console.log('🎉 Provider selection process completed.');

// // ⏳ Wait for popup trigger button
// console.log('⏳ Waiting for Facility Search button...');
// await page.waitForSelector('#ctl00_phFolderContent_Button35', { visible: true, timeout: 6000 });

// console.log('🖱️ Setting up facility popup listener...');

// // Set up popup listener BEFORE clicking the button
// const facilityPopupPromise = new Promise((resolve) => {
//     const onPopup = async (target) => {
//         console.log('✅ Facility popup detected!');
//         browser.off('targetcreated', onPopup);
//         const popup = await target.page();
//         resolve(popup);
//     };
//     browser.on('targetcreated', onPopup);
// });

// // 🔘 Click facility search button
// console.log('🖱️ Clicking facility search button...');
// await page.click('#ctl00_phFolderContent_Button35');

// // Wait for popup
// console.log('⏳ Waiting for facility popup to load...');
// let facilityPage;
// try {
//     facilityPage = await Promise.race([
//         facilityPopupPromise,
//         new Promise((_, reject) => setTimeout(() => reject(new Error('Facility popup timeout')), 8000))
//     ]);
//     console.log('✅ Facility popup loaded');
// } catch (error) {
//     const allPages = await browser.pages();
//     if (allPages.length > 1) {
//         facilityPage = allPages[allPages.length - 1];
//         console.log('✅ Detected new tab as facility popup');
//     } else {
//         throw new Error('❌ No popup or new tab for facility found');
//     }
// }

// // ⏳ Wait for search field
// console.log('⏳ Waiting for facility search input...');
// await facilityPage.waitForSelector('#ctl04_popupBase_txtSearch', { visible: true, timeout: 6000 });

// // 💬 Enter facility name
// console.log(`⌨️ Entering Facility Name: ${facilityName}`);
// await facilityPage.evaluate((name) => {
//     const input = document.querySelector('#ctl04_popupBase_txtSearch');
//     if (input) {
//         input.value = '';
//         input.focus();
//     }
// }, facilityName);

// await facilityPage.type('#ctl04_popupBase_txtSearch', facilityName, { delay: 100 });

// // 🔍 Click Search button
// console.log('🖱️ Clicking facility search button...');
// await facilityPage.waitForSelector('#ctl04_popupBase_btnSearch', { visible: true, timeout: 4000 });
// await facilityPage.click('#ctl04_popupBase_btnSearch');

// // ⏳ Wait for results
// console.log('⏳ Waiting for facility search results...');
// await facilityPage.waitForSelector('a[id^="ctl04_popupBase_grvPopup_ctl"][id$="_lnkSelect"]', { visible: true, timeout: 6000 });
// await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close

// // 🖱️ Click the first "Select" link
// console.log('🖱️ Selecting the first facility...');
// await facilityPage.evaluate(() => {
//     const selectLink = document.querySelector('a[id^="ctl04_popupBase_grvPopup_ctl"][id$="_lnkSelect"]');
//     if (selectLink) {
//         selectLink.scrollIntoView({ behavior: 'smooth', block: 'center' });
//         selectLink.click();
//     }
// });

// // ✅ Done
// console.log('🎉 Facility selection completed!');
// await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close

// console.log('🔍 Processing complete. Please verify all entries on the form.');
// if (['21', '61', '31'].includes(pos)) {
//   console.log(`🩺 POS is ${pos}, entering Hospital From Date: ${admitDate}`);

//   try {
//     const [month, day, year] = admitDate.split('/');

//     // Month field
//     const monthSelector = '#ctl00_phFolderContent_HospitalFromDate_Month';
//     await page.waitForSelector(monthSelector, { visible: true, timeout: 10000 });
//     await page.focus(monthSelector);
//     await page.type(monthSelector, month, { delay: 50 });
//     await page.keyboard.press('Tab');

//     // Day field
//     const daySelector = '#ctl00_phFolderContent_HospitalFromDate_Day';
//     await page.waitForSelector(daySelector, { visible: true, timeout: 3000 });
//     await page.focus(daySelector);
//     await page.type(daySelector, day, { delay: 50 });
//     await page.keyboard.press('Tab');

//     // Year field
//     const yearSelector = '#ctl00_phFolderContent_HospitalFromDate_Year';
//     await page.waitForSelector(yearSelector, { visible: true, timeout: 3000 });
//     await page.focus(yearSelector);
//     await page.type(yearSelector, year, { delay: 50 });

//     // ✅ Do NOT press Tab or click again — just let it be
//     console.log('✅ Hospital From Date filled successfully.');
//   } catch (error) {
//     console.log(`❌ Error filling Hospital From Date: ${error.message}`);
//   }
// } else {
//   console.log(`ℹ️ POS is ${pos}. Skipping Hospital From Date entry.`);
// }

// await new Promise(resolve => setTimeout(resolve, 1500)); // Shorter wait


// // Handle Prior Authorization Number
// if (priornumber && priornumber.trim() !== '') {
//     console.log(`📝 Entering Prior Authorization Number: ${priornumber}`);
    
//     try {
//         // Wait for the prior auth field with longer timeout
//         await page.waitForSelector('#ctl00_phFolderContent_PriorAuthorizationNumber', { visible: true, timeout: 10000 });

//         // Scroll to the field to ensure it's visible
//         await page.evaluate(() => {
//             const priorField = document.querySelector('#ctl00_phFolderContent_PriorAuthorizationNumber');
//             if (priorField) {
//                 priorField.scrollIntoView({ behavior: 'smooth', block: 'center' });
//             }
//         });

//         // Wait for scroll to complete
//         await new Promise(resolve => setTimeout(resolve, 1000));

//         // Click on the field to focus it
//         await page.click('#ctl00_phFolderContent_PriorAuthorizationNumber');

//         // Clear the field using keyboard shortcuts
//         await page.keyboard.down('Control');
//         await page.keyboard.press('KeyA');
//         await page.keyboard.up('Control');
//         await page.keyboard.press('Delete');

//         // Type the prior authorization number
//         await page.type('#ctl00_phFolderContent_PriorAuthorizationNumber', priornumber);

//         // Verify the value was entered
//         const enteredValue = await page.$eval('#ctl00_phFolderContent_PriorAuthorizationNumber', el => el.value);
//         console.log(`✅ Prior Authorization Number entered. Value: ${enteredValue}`);

//         // Trigger change event to ensure the form recognizes the input
//         await page.evaluate(() => {
//             const priorField = document.querySelector('#ctl00_phFolderContent_PriorAuthorizationNumber');
//             if (priorField) {
//                 priorField.dispatchEvent(new Event('change', { bubbles: true }));
//                 priorField.dispatchEvent(new Event('input', { bubbles: true }));
//             }
//         });

//     } catch (error) {
//         console.log(`❌ Error entering Prior Authorization Number: ${error.message}`);
        
//         // Try alternative approach if the main selector fails
//         try {
//             console.log('🔄 Trying alternative approach for Prior Auth field...');
            
//             // Try to find the field by partial name match
//             const alternativeSelector = 'input[name*="PriorAuthorizationNumber"]';
//             await page.waitForSelector(alternativeSelector, { visible: true, timeout: 5000 });
            
//             await page.click(alternativeSelector);
//             await page.keyboard.down('Control');
//             await page.keyboard.press('KeyA');
//             await page.keyboard.up('Control');
//             await page.keyboard.press('Delete');
//             await page.type(alternativeSelector, priornumber);
            
//             console.log('✅ Prior Authorization Number entered using alternative selector.');
            
//         } catch (altError) {
//             console.log(`❌ Alternative approach also failed: ${altError.message}`);
//         }
//     }
// } else {
//     console.log('ℹ️ Prior Authorization Number not provided. Skipping field.');
// }

// // Final wait before proceeding
// await new Promise(resolve => setTimeout(resolve, 1000));
// // ✅ Click the update button after processing
// // Wait for the update button to appear
// await page.waitForSelector('#ctl00_phFolderContent_btnUpdate', { visible: true });

// // Click the button using Puppeteer
// await page.click('#ctl00_phFolderContent_btnUpdate');
// console.log("🚀 Update button clicked.");
// await new Promise(resolve => setTimeout(resolve, 1000));

// // Wait for the message to appear
// await page.waitForSelector('.message-header span', { visible: true });

// // Get the message text
// const messageText = await page.$eval('.message-header span', el => el.textContent.trim());

// // Extract just the Visit ID
// const visitIdMatch = messageText.match(/Visit\s+(\d+)/);

// let visitId = null; // ✅ Declare outside the block

// if (visitIdMatch) {
//   visitId = visitIdMatch[1];
//   console.log(visitId); // 🔥 Only print the ID
// } else {
//   console.log("❌ Visit ID not found.");
// }

// return { success: true, visitId }; // ✅ Safe to use


// } catch (error) {
//         console.error('Error processing account:', error);
//         return false; // Return failure
//     }
// }


// // Endpoint to process data and return status
// app.post("/process", upload.single("file"), async (req, res) => {
//     if (!req.file || !req.body.originalPath) {
//         return res.status(400).json({ error: "No file or original path provided" });
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
        
//         // Check if "Patient ID" exists in headers
//         console.log(headers);
//         if (!headers.includes("Patient ID")) {
//             throw new Error("'Patient ID' Column is not defined.");
//         }
        
//         // Map headers to their index positions
//         headers.forEach((header, index) => {
//             columnMap[header] = index;
//         });
        
//         const requiredColumns = ["Patient ID", "Provider", "Facility", "DOS"];
//         for (const column of requiredColumns) {
//             if (!headers.includes(column)) {
//                 throw new Error(`Required column '${column}' is missing from the spreadsheet.`);
//             }
//         }
    
//         // Ensure "Visit ID" column is added if missing (before Result)
//         const visitIdHeader = "Visit ID";
//         if (!headers.includes(visitIdHeader)) {
//             // Find the position of Result column or add at the end
//             const resultIndex = headers.indexOf("Result");
//             const insertIndex = resultIndex !== -1 ? resultIndex : headers.length;
            
//             // Insert Visit ID column
//             headers.splice(insertIndex, 0, visitIdHeader);
            
//             // Add empty Visit ID cells to all existing rows
//             for (let i = 1; i < originalData.length; i++) {
//                 originalData[i] = originalData[i] || [];
//                 originalData[i].splice(insertIndex, 0, "");
//             }
//         }

//         // Ensure only "Result" column is added if missing
//         const resultHeader = "Result";
//         if (!headers.includes(resultHeader)) {
//             columnMap[resultHeader] = headers.length;
//             headers.push(resultHeader);
//             for (let i = 1; i < originalData.length; i++) {
//                 originalData[i] = originalData[i] || [];
//                 originalData[i].push("");
//             }
//         }

//         // Update column mapping after adding new columns
//         headers.forEach((header, idx) => columnMap[header] = idx);
        
//         let newRowsToProcess = 0;
//         const processedAccounts = new Set();
        
//         let missingCallDiagnoses = 0;  // Track rows missing Diagnoses

//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue; // Skip undefined rows
            
//             const accountNumberRaw = row[columnMap["Patient ID"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            
//             if (!accountNumberRaw || status === "done" || status === "failed") continue;  // Skip processed or failed rows

            
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             processedAccounts.add(accountNumber);
//             newRowsToProcess++;
        
//             // Check if Diagnoses is missing
//             const callDiagnoses = row[columnMap["Diagnoses"]] || "";
//             if (!callDiagnoses.trim()) {
//                 missingCallDiagnoses++;
//             }
//         }
        
//         // Check if all remaining rows are missing Diagnoses
//         if (newRowsToProcess > 0 && newRowsToProcess === missingCallDiagnoses) {
//             return res.json({
//                 success: false,
//                 message: "Diagnoses not found in Excel.",
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
        
//         newPage = await continueWithLoggedInSession();

        
//         const results = { successful: [], failed: [] };
        
//         for (let i = 1; i < originalData.length; i++) {
//             const row = originalData[i];
//             if (!row) continue;
        
//             const accountNumberRaw = row[columnMap["Patient ID"]] || "";
//             const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
//             if (!accountNumberRaw || status === "done") continue;
        
//             const provider = row[columnMap["Provider"]] || "";
//             const facilityName = row[columnMap["Facility"]] || "";
//             const billingDate = row[columnMap["DOS"]] || "";
//             const diagnosisText = row[columnMap["Diagnoses"]] || "";
//             const chargesText = row[columnMap["CPT"]] || ""; // CPT codes
//             const admitDate = row[columnMap["Hospital Date"]] || ""; // ✅ NEW LINE
//             const pos = row[columnMap["POS"]] || ""; // ✅ NEW LINE
//             const modifier = row[columnMap["Modifier"]] || ""; // ✅ NEW LINE
//             const priornumber = row[columnMap["Prior Authorization Number"]] || ""; // ✅ NEW LINE
//             const accountNumber = accountNumberRaw.split("-")[0].trim();
//             console.log(`Processing account: ${accountNumber}`);
        
//             // Validate data before processing
//             if (!provider) console.log('Provider value is missing or undefined');
//             if (!facilityName) console.log('Facility is missing or undefined');
//             if (!billingDate) console.log('DOS is missing or undefined');
//             if (!admitDate) console.log('Hospital Date is missing or undefined'); // ✅ Optional

//           // Array to hold valid diagnosis codes
// const diagnosisCodes = [];

// // Regex:
// // - ICD-10-CM: Starts with a letter, 2–4 digits, optional dot, followed by 1–4 alphanumerics
// // - ICD-10-PCS: Exactly 7 alphanumeric characters (A-Z, 0-9)
// const dxCodeRegex = /\b(?:[A-Z][0-9]{2,4}(?:\.[0-9A-Z]{1,4})?|[A-Z0-9]{7})\b/g;

// // Match codes
// const matches = diagnosisText.match(dxCodeRegex);

// if (matches) {
//   // Normalize: trim, uppercase, remove duplicates
//   const uniqueCodes = [...new Set(matches.map(code => code.trim().toUpperCase()))];

//   // Filter: Must contain at least one letter and one digit
//   const validCodes = uniqueCodes.filter(code => /[A-Z]/i.test(code) && /\d/.test(code));

//   diagnosisCodes.push(...validCodes);

//   // Output
//   console.log("✅ Final Unique, Valid Diagnosis Codes:");
//   diagnosisCodes.forEach((code, index) => {
//     console.log(`  ${index + 1}. ${code}`);
//   });
//   console.log("🔢 Total:", diagnosisCodes.length);
// } else {
//   console.log("⚠️ No valid diagnosis codes found.");
// }
          
          
// // ✅ Normalize input: remove non-breaking spaces, tabs, newlines, etc.
// const cleanedText = chargesText.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();

// const cptCodes = [];
// const cptUnits = []; // Store corresponding units

// const cptRegex = /\b\d{5}(?:\*\d+)?\b/g;
// let cptMatch;

// // ✅ Extract all 5-digit CPT codes, optionally with *units
// while ((cptMatch = cptRegex.exec(cleanedText)) !== null) {
//   const fullCode = cptMatch[0].trim();
//   let baseCode = fullCode;
//   let units = '1'; // default if no units provided

//   if (fullCode.includes('*')) {
//     [baseCode, units] = fullCode.split('*');
//   }

//   baseCode = baseCode.replace(/\D/g, '').trim(); // extra cleaning
//   cptCodes.push(baseCode);
//   cptUnits.push(units);

//   console.log(`🧩 CPT Code: ${baseCode}, Units: ${units}`);
// }

// // ✅ Remove duplicates in CPT codes (if needed)
// const uniqueCptCodes = [...new Set(cptCodes)];
// console.log('✅ Extracted Unique CPT Codes:', uniqueCptCodes);
// uniqueCptCodes.forEach((code, index) => {
//   console.log(`  ${index + 1}. ${code}`);
// });


//           // === Call processAccount (Don't send CPT length)
//  const { success, visitId} = await processAccount(
//     newPage,
//     newPage.browser(), // Add browser parameter
//     accountNumber,
//     provider,
//     facilityName,
//     billingDate,
//     diagnosisCodes,
//     cptCodes, // ✅ only the array
//     admitDate,
//     pos,
//     modifier,
//     priornumber,
//     cptUnits
// );

// // Set Result and Visit ID
// if (success) {
//     results.successful.push(accountNumber);
//     row[columnMap["Result"]] = "Done";
//     row[columnMap["Visit ID"]] = visitId ? visitId.toString() : "";
// } else {
//     results.failed.push(accountNumber);
//     row[columnMap["Result"]] = "Failed";
//     row[columnMap["Visit ID"]] = visitId ? visitId.toString() : "";
// }
//         }            
//         // Convert data back to an ExcelJS workbook
//        const workbook = new ExcelJS.Workbook();
// const sheet = workbook.addWorksheet("Sheet1");

// // Write headers with gray background
// const headerRow = sheet.addRow(headers);
// headerRow.height = 42;
// headerRow.eachCell((cell) => {
//     cell.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: "A9A9A9" }
//     };
//     cell.font = { bold: true, color: { argb: "000000" }, size: 10 };
//     cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
//     cell.border = {
//         top: { style: "thin" },
//         left: { style: "thin" },
//         bottom: { style: "thin" },
//         right: { style: "thin" }
//     };
// });

// // Define column widths (adjusted to include Visit ID column)
// const columnWidths = [6, 11, 30, 15, 15, 23, 54, 10, 10, 11, 50, 20, 15, 30 ,15]; 
// // Added one more width for Visit ID column (15)

// sheet.columns = columnWidths.map((width, index) => ({
//     width,
//     style: (index === columnMap["Diagnoses"] || index === columnMap["CPT"])
//         ? { alignment: { wrapText: true } }
//         : {},
// }));

// originalData.slice(1).forEach((row) => {
//     if (!row) return;

//     const hasData = row.some(cell => cell && cell.toString().trim() !== "");
//     if (!hasData) return;

//     const rowData = headers.map(header => row[columnMap[header]] || "");
//     const newRow = sheet.addRow(rowData);
//     newRow.height = 100;

//     headers.forEach((header, colIndex) => {
//         const cell = newRow.getCell(colIndex + 1);

//         // Apply border
//         cell.border = {
//             top: { style: "thin", color: { argb: "000000" } },
//             left: { style: "thin", color: { argb: "000000" } },
//             bottom: { style: "thin", color: { argb: "000000" } },
//             right: { style: "thin", color: { argb: "000000" } }
//         };

//         cell.font = { size: 10 };

//         // Format Patient ID
//         if (header === "Patient ID" && cell.value) {
//             cell.fill = {
//                 type: "pattern",
//                 pattern: "solid",
//                 fgColor: { argb: "FFFFFF" }
//             };
//             cell.alignment = { horizontal: "center", vertical: "bottom" };
//         }

//         // Format Visit ID
//         if (header === "Visit ID" && cell.value) {
//             cell.fill = {
//                 type: "pattern",
//                 pattern: "solid",
//                 fgColor: { argb: "E6F3FF" } // Light blue background
//             };
//             cell.alignment = { horizontal: "center", vertical: "middle" };
//         }

//         // Format Result
//         if (header === "Result" && cell.value) {
//             const value = cell.value.toString().toLowerCase();
//             if (value === "done") {
//                 cell.fill = {
//                     type: "pattern",
//                     pattern: "solid",
//                     fgColor: { argb: "92D050" } // Green
//                 };
//             } else if (value === "failed") {
//                 cell.fill = {
//                     type: "pattern",
//                     pattern: "solid",
//                     fgColor: { argb: "FF0000" } // Red
//                 };
//             }
//             cell.alignment = { horizontal: "center", vertical: "middle" };
//         }
//     });
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
//             message: `${results.successful.length} Charges Entered successfully. ${results.failed.length} rows failed.`,
//             results,
//             fileId
//         });

//     } catch (error) {
//         console.error("Error:", error);
//         res.status(500).json({ error: error.message });
//     } finally {
//         if (newPage) {
//             await new Promise(resolve => setTimeout(resolve, 2000));
//             await newPage.close();
//             console.log("Browser closed");
//         }
//     }
// });
// // Separate endpoint to download the processed file
// app.get("/download/:fileId", (req, res) => {
//   const fileId = req.params.fileId;
//   const fileData = processedFiles.get(fileId);
  
//   if (!fileData) {
//       return res.status(404).json({ error: "File not found" });
//   }
  
//   // Determine content type based on file extension
//   let contentType = 'application/octet-stream'; // Default
  
//   if (fileData.filename.endsWith('.xlsx')) {
//       contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
//   } else if (fileData.filename.endsWith('.xls')) {
//       contentType = 'application/vnd.ms-excel';
//   } else if (fileData.filename.endsWith('.csv')) {
//       contentType = 'text/csv';
//   }
  
//   // Set the response headers
//   res.setHeader('Content-Disposition', `attachment; filename="${fileData.filename}"`);
//   res.setHeader('Content-Type', contentType);
//   res.setHeader('Content-Length', fileData.buffer.length);
  
//   // Send the file buffer
//   res.send(fileData.buffer);
  
//   // Optional: Remove the file from memory after some time
// // Optional: Remove the file from memory after some time
// setTimeout(() => {
//     processedFiles.delete(fileId);
//     console.log(`File ${fileId} removed from memory`);
// }, 3 * 60 * 60 * 1000); // Remove after 3 hours

// });

// app.listen(port, () => console.log(`Server running on http://localhost:${port}`));	//07/09/2025 - After Fetch Visit ID except Modifier 




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



async function continueWithLoggedInSession() {
    try {
        const browser = await puppeteer.connect({
            browserURL: 'http://localhost:9222', // Connect to existing Chrome
            defaultViewport: null,
            headless: false,
            timeout: 120000,
        });

        const pages = await browser.pages();
        const sessionPage = pages.find(p =>
            p.url().includes('Appointments/ViewAppointments.aspx')
        );

        if (!sessionPage) {
            console.log('❌ No matching tab found for the appointment page.');

            if (pages.length > 0) {
                // Show alert in the first available page
                await pages[0].bringToFront();
                await pages[0].evaluate(() => {
                    alert('⚠️ Session not found');
                });
            } else {
                console.log('⚠️ No pages available to show alert.');
            }

            return null;
        }

        console.log('✅ Found existing page. Waiting for Patient Visit tab...');
        await sessionPage.waitForSelector('#patient-visits_tab > span', {
            visible: true,
            timeout: 10000,
        });

        console.log('🖱️ Clicking Patient Visit tab...');
        await sessionPage.evaluate(() => {
            const tab = document.querySelector('#patient-visits_tab > span');
            if (tab) {
                tab.scrollIntoView({ behavior: 'smooth', block: 'center' });
                tab.click();
            }
        });

        console.log('✅ Patient Visit tab clicked successfully.');
        return sessionPage;
    } catch (error) {
        console.error('🚨 Error in continueWithLoggedInSession:', error.message);
        return null;
    }
}

async function processAccount(page,browser,accountNumber, provider, facilityName, billingDate, diagnosisCodes,uniqueCptCodes,admitDate,pos,modifier,priornumber,cptUnits) {
   try {
      // Wait and click Add New Visit
// Wait and click Add New Visit
await new Promise(resolve => setTimeout(resolve, 3000));
console.log('⏳ Waiting for Add New Visit button to appear...');
await page.waitForSelector('#addNewVisit', { visible: true, timeout: 15000 });

console.log('🖱️ Setting up popup listener...');

// Set up popup listener BEFORE clicking the button
const popupPromise = new Promise((resolve) => {
    const onPopup = async (target) => {
        console.log('✅ Popup detected!');
        browser.off('targetcreated', onPopup); // Fixed: Use .off() instead of .removeListener()
        const popup = await target.page();
        resolve(popup);
    };
    browser.on('targetcreated', onPopup);
});

console.log('🖱️ Clicking Add New Visit button...');
await page.evaluate(() => {
    const btn = document.querySelector('#addNewVisit');
    if (btn) {
        btn.scrollIntoView({ behavior: 'smooth', block: 'center' });  
        btn.click();
    } else {
        console.error('❌ Add New Visit button not found!');
    }
});
 // ⏳ Wait for popup to load
        console.log('⏳ Waiting for Search button to appear...');
        await page.waitForSelector('#ctl00_phFolderContent_Button1', { visible: true, timeout: 10000 });

        // 🔘 Click initial popup search button
        console.log('🖱️ Clicking popup search open button...');
        await page.click('#ctl00_phFolderContent_Button1');
// Wait for popup to appear with reduced timeout
console.log('⏳ Waiting for popup to load...');
let newPage;
try {
    // Reduced timeout from 6000ms to 3000ms
    newPage = await Promise.race([
        popupPromise,
        new Promise((_, reject) => setTimeout(() => reject(new Error('Popup timeout')), 7000))
    ]);
    console.log('✅ Popup loaded successfully');
} catch (error) {
    console.log('⚠️ No popup detected, checking for new tabs...');
    
    // Reduced wait time from 2000ms to 1000ms
    await new Promise(resolve => setTimeout(resolve, 1000));
    const allPages = await browser.pages();
    
    if (allPages.length > 1) {
        newPage = allPages[allPages.length - 1];
        console.log('✅ New tab detected, using it...');
    } else {
        newPage = page;
        console.log('📝 No new tab, using current page...');
    }
}

// Reduced wait time from 3000ms to 1500ms
await new Promise(resolve => setTimeout(resolve, 1500));

// ⏳ Wait for Search button with reduced timeout and parallel checking
console.log('⏳ Waiting for Search button to appear in new page...');
const searchButtonSelectors = [
    '#ctl00_phFolderContent_Button1',
    'input[value="Search"]',
    'button[onclick*="Search"]',
    '.button:contains("Search")',
    '#Button1'
];

let searchButtonFound = false;
let workingSelector = null;

// Check all selectors in parallel instead of sequentially
try {
    const promises = searchButtonSelectors.map(selector => 
        newPage.waitForSelector(selector, { visible: true, timeout: 3000 })
            .then(() => selector)
            .catch(() => null)
    );
    
    const results = await Promise.allSettled(promises);
    workingSelector = results.find(result => result.status === 'fulfilled' && result.value)?.value;
    
    if (workingSelector) {
        console.log(`✅ Found search button with selector: ${workingSelector}`);
        searchButtonFound = true;
    }
} catch (error) {
    console.log('🔍 No search button found with standard selectors');
}

if (!searchButtonFound) {
    throw new Error('Search button not found with any selector');
}

// 🔘 Click initial popup search button
console.log('🖱️ Clicking popup search open button...');
try {
    await newPage.click(workingSelector);
} catch (error) {
    // Fallback click method
    await newPage.evaluate(() => {
        const searchButtons = document.querySelectorAll('input[value*="Search"], button');
        for (const btn of searchButtons) {
            if (btn.value?.includes('Search') || btn.textContent?.includes('Search')) {
                btn.click();
                break;
            }
        }
    });
}

// Wait for dropdown with parallel selector checking
console.log('⏳ Waiting for dropdown to appear...');
const dropdownSelectors = [
    '#ctl04_popupBase_ddlSearch',
    'select[name*="ddlSearch"]',
    'select[id*="Search"]',
    'select:first-of-type'
];

let dropdownSelector = null;

try {
    const dropdownPromises = dropdownSelectors.map(selector => 
        newPage.waitForSelector(selector, { visible: true, timeout: 2000 })
            .then(() => selector)
            .catch(() => null)
    );
    
    const dropdownResults = await Promise.allSettled(dropdownPromises);
    dropdownSelector = dropdownResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
    if (dropdownSelector) {
        console.log(`✅ Found dropdown with selector: ${dropdownSelector}`);
    } else {
        throw new Error('Dropdown not found');
    }
} catch (error) {
    console.log('❌ Dropdown not found with any selector');
    throw error;
}

// Reduced wait time from 1000ms to 500ms
await new Promise(resolve => setTimeout(resolve, 500));

// 🔽 Select "Patient ID" with improved logic
console.log('🔽 Selecting "Patient ID" from dropdown...');
await newPage.evaluate((selector) => {
    const dropdown = document.querySelector(selector);
    
    if (dropdown) {
        // Find Patient ID option more efficiently
        let patientIdIndex = -1;
        const options = Array.from(dropdown.options);
        
        // Search for Patient ID option
        patientIdIndex = options.findIndex(option => 
            option.text.toLowerCase().includes('patient id') || 
            option.value.toLowerCase().includes('patientid') ||
            option.text.toLowerCase().includes('patient')
        );
        
        if (patientIdIndex !== -1) {
            dropdown.selectedIndex = patientIdIndex;
            console.log(`Selected Patient ID option at index ${patientIdIndex}: ${options[patientIdIndex].text}`);
        } else {
            // Fallback to index 2 (3rd option) if Patient ID not found
            dropdown.selectedIndex = Math.min(2, options.length - 1);
            console.log(`Fallback: Selected option at index ${dropdown.selectedIndex}: ${options[dropdown.selectedIndex].text}`);
        }
        
        // Trigger change event
        dropdown.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
    }
    return false;
}, dropdownSelector);

// Reduced wait time from 500ms to 200ms
await new Promise(resolve => setTimeout(resolve, 200));

// Enhanced text field handling with parallel selector checking
console.log(`⌨️ Entering Patient ID: ${accountNumber}`);
const textFieldSelectors = [
    '#ctl04_popupBase_txtSearch',
    'input[name*="txtSearch"]',
    'input[type="text"]:not([style*="display: none"])',
    'input[maxlength="50"]'
];

let textFieldSelector = null;

try {
    const textPromises = textFieldSelectors.map(selector => 
        newPage.waitForSelector(selector, { visible: true, timeout: 2000 })
            .then(() => selector)
            .catch(() => null)
    );
    
    const textResults = await Promise.allSettled(textPromises);
    textFieldSelector = textResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
    if (!textFieldSelector) {
        throw new Error('Text field not found');
    }
} catch (error) {
    console.log('❌ Text field not found with any selector');
    throw error;
}

// Clear and enter text more efficiently
await newPage.evaluate((selector) => {
    const textField = document.querySelector(selector);
    if (textField) {
        textField.value = '';
        textField.focus();
    }
}, textFieldSelector);

await newPage.type(textFieldSelector, accountNumber, { delay: 50 }); // Reduced delay from 100ms to 50ms

// 🔍 Click the search button with parallel selector checking
console.log('🖱️ Clicking search button...');
const searchBtnSelectors = [
    '#ctl04_popupBase_btnSearch',
    'input[name*="btnSearch"]',
    'input[value="Search"]',
    'button[onclick*="Search"]'
];

let searchBtnClicked = false;

try {
    const searchBtnPromises = searchBtnSelectors.map(selector => 
        newPage.waitForSelector(selector, { visible: true, timeout: 1000 })
            .then(() => selector)
            .catch(() => null)
    );
    
    const searchBtnResults = await Promise.allSettled(searchBtnPromises);
    const searchBtnSelector = searchBtnResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
    if (searchBtnSelector) {
        await newPage.click(searchBtnSelector);
        searchBtnClicked = true;
    }
} catch (error) {
    console.log('⚠️ Standard search button click failed, trying fallback...');
}

if (!searchBtnClicked) {
    // Fallback click method
    await newPage.evaluate(() => {
        const searchBtn = document.querySelector('#ctl04_popupBase_btnSearch') ||
                         document.querySelector('input[name*="btnSearch"]') ||
                         document.querySelector('input[value="Search"]');
        if (searchBtn) {
            searchBtn.click();
            return true;
        }
        return false;
    });
}

// Reduced wait time from 2000ms to 1000ms for search results
await new Promise(resolve => setTimeout(resolve, 1000));
console.log('✅ Search completed.');

// ✅ Wait for search results with reduced timeout
console.log('⏳ Waiting for search result row selector to appear...');
const resultSelectors = [
    '#ctl04_popupBase_grvPopup_ctl02_lnkSelect',
    'a[id*="lnkSelect"]',
    'a[href*="Select"]',
    'input[value="Select"]'
];

let resultSelector = null;

try {
    const resultPromises = resultSelectors.map(selector => 
        newPage.waitForSelector(selector, { visible: true, timeout: 5000 })
            .then(() => selector)
            .catch(() => null)
    );
    
    const resultResults = await Promise.allSettled(resultPromises);
    resultSelector = resultResults.find(result => result.status === 'fulfilled' && result.value)?.value;
    
    if (resultSelector) {
        console.log(`✅ Found result selector: ${resultSelector}`);
        
        // 🖱️ Click the first row's "Select" link
        console.log('🖱️ Clicking first patient select link...');
        await newPage.evaluate((selector) => {
            const selectBtn = document.querySelector(selector);
            if (selectBtn) {
                selectBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
                selectBtn.click();
                return true;
            }
            return false;
        }, resultSelector);
        
        console.log('✅ Patient selected successfully.');
    } else {
        console.log('❌ No search results found');
    }
    await new Promise(resolve => setTimeout(resolve, 1000)); // Reduced from 2500 to 600

    // Wait for the Visit Date fields to load
console.log('⏳ Waiting for Visit Date input fields...');
await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Month', { visible: true, timeout: 8000 });
await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Day', { visible: true, timeout: 8000 });
await page.waitForSelector('#ctl00_phFolderContent_DateVisited_Year', { visible: true, timeout: 8000 });

// Split billingDate into MM/DD/YYYY
const [month, day, year] = billingDate.split('/');
console.log(`⌨️ Entering Visit Date: ${month}/${day}/${year}`);

// Clear and type values into fields
await page.evaluate(({ month, day, year }) => {
    const setFieldValue = (selector, value) => {
        const field = document.querySelector(selector);
        if (field) {
            field.value = '';
            field.focus();
            field.value = value;
        }
    };

    setFieldValue('#ctl00_phFolderContent_DateVisited_Month', month);
    setFieldValue('#ctl00_phFolderContent_DateVisited_Day', day);
    setFieldValue('#ctl00_phFolderContent_DateVisited_Year', year);
}, { month, day, year });

console.log('✅ Visit Date fields populated successfully.');
await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close

// ⏳ Wait for popup to load
      // ⏳ Wait for popup to load
console.log('⏳ Waiting for Provider Search button to appear...');
await page.waitForSelector('#ctl00_phFolderContent_Button2', { visible: true, timeout: 6000 });

console.log('🖱️ Setting up provider popup listener...');

// Set up popup listener BEFORE clicking the button
const providerPopupPromise = new Promise((resolve) => {
    const onPopup = async (target) => {
        console.log('✅ Provider popup detected!');
        browser.off('targetcreated', onPopup);
        const popup = await target.page();
        resolve(popup);
    };
    browser.on('targetcreated', onPopup);
});

// 🔘 Click provider search button
console.log('🖱️ Clicking provider search button...');
await page.click('#ctl00_phFolderContent_Button2');

// Wait for popup to appear
console.log('⏳ Waiting for provider popup to load...');
let providerPage;
try {
    // Wait for popup with timeout
    providerPage = await Promise.race([
        providerPopupPromise,
        new Promise((_, reject) => setTimeout(() => reject(new Error('Provider popup timeout')), 8000))
    ]);
    console.log('✅ Provider popup loaded successfully');
} catch (error) {
    console.log('⚠️ No provider popup detected, checking for new tabs...');
    
    // Fallback: Check for new tabs
    await new Promise(resolve => setTimeout(resolve, 3000));
    const allPages = await browser.pages();
    
    if (allPages.length > 1) {
        providerPage = allPages[allPages.length - 1]; // Get the newest page
        console.log('✅ New tab detected for provider search...');
    } else {
        throw new Error('No provider popup or new tab found');
    }
}

// Wait for the provider page to load completely
await new Promise(resolve => setTimeout(resolve, 3000));

const providerSearchTerm = provider.split(',')[0].trim().split(' ').slice(-1)[0] || provider;
console.log(`⌨️ Will search for Provider: ${providerSearchTerm}`);


// Wait for the provider search input field to appear in the NEW TAB
console.log('⏳ Waiting for Provider search field in new tab...');
try {
    await providerPage.waitForSelector('input[type="text"]', { visible: true, timeout: 8000 });
    console.log('✅ Provider search field found in new tab');
} catch (error) {
    console.log('🔍 Provider search field not found, trying to find any text input...');
    
    // Try to find any text input in the new tab
    const textInputs = await providerPage.$$('input[type="text"]');
    if (textInputs.length === 0) {
        throw new Error('No text input found in provider popup');
    }
    console.log(`Found ${textInputs.length} text input(s) in provider popup`);
}

// Clear and enter provider search term in the NEW TAB
console.log(`⌨️ Entering Provider search term in new tab: ${providerSearchTerm}`);

// Find and clear the search field
await providerPage.evaluate(() => {
    const searchField = document.querySelector('input[type="text"]');
    if (searchField) {
        searchField.value = '';
        searchField.focus();
        console.log('Search field cleared and focused');
    }
});

// Type the provider search term
await providerPage.type('input[type="text"]', providerSearchTerm, { delay: 100 });

// Verify the text was entered
const enteredText = await providerPage.$eval('input[type="text"]', el => el.value).catch(() => 'Could not verify');
console.log('✅ Provider search text entered in new tab:', enteredText);

// Find and click the Search button in the NEW TAB
console.log('🔍 Looking for Search button in new tab...');
try {
    // Wait for search button to be available
    await providerPage.waitForSelector('input[value="Search"]', { visible: true, timeout: 5000 });
    console.log('🖱️ Clicking Search button in new tab...');
    await providerPage.click('input[value="Search"]');
} catch (error) {
    console.log('🔍 Search button not found, trying alternatives...');
    
    // Try clicking any button that might be the search button
    await providerPage.evaluate(() => {
        const searchButtons = document.querySelectorAll('input[type="button"], input[type="submit"], button');
        for (const btn of searchButtons) {
            if (btn.value?.toLowerCase().includes('search') || 
                btn.textContent?.toLowerCase().includes('search') ||
                btn.onclick?.toString().includes('search')) {
                console.log('Found search button, clicking...');
                btn.click();
                return;
            }
        }
        console.log('No search button found, trying first button...');
        if (searchButtons.length > 0) {
            searchButtons[0].click();
        }
    });
}

// Wait for search results to load
console.log('⏳ Waiting for search results...');
await new Promise(resolve => setTimeout(resolve, 3000));

// ⏳ Checking if provider search results are empty
console.log('⏳ Checking provider search results...');
let noResults = false;

try {
    noResults = await providerPage.evaluate(() => {
        const tdElements = Array.from(document.querySelectorAll('td'));
        return tdElements.some(td => td.textContent.trim().toLowerCase() === 'no data results to display.');
    });

    if (noResults) {
        console.warn('❌ No provider search results found.');
    }
} catch (checkError) {
    console.warn('⚠️ Error while checking for no results:', checkError.message);
}

if (noResults) {
    // Close the provider tab if nothing is found
    try {
        await providerPage.close();
        console.log('🛑 Provider popup closed due to no results.');
    } catch (closeError) {
        console.warn('⚠️ Error closing provider popup:', closeError.message);
    }
} else {
    // Continue with provider selection
    try {
        await providerPage.waitForSelector('a[href*="Select"], a[onclick*="Select"]', {
            visible: true,
            timeout: 10000
        });

        console.log('🖱️ Selecting first provider from search results...');
        await providerPage.evaluate(() => {
            const selectLinks = document.querySelectorAll('a[href*="Select"], a[onclick*="Select"]');
            if (selectLinks.length > 0) {
                selectLinks[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
                selectLinks[0].click();
            } else {
                console.error('❌ No select links found!');
            }
        });

        console.log('✅ Provider selected successfully from new tab.');
    } catch (error) {
        console.error('🚨 Failed to select provider:', error.message);
    }

    // Wait for the selection to process and popup to close
    await new Promise(resolve => setTimeout(resolve, 3000));
}

console.log('🎉 Provider selection process completed.');

// ⏳ Wait for the Billing Info tab to be visible
console.log('⏳ Waiting for Billing Info tab...');
await page.waitForSelector('#VisitTabs > ul > li:nth-child(2)', { visible: true, timeout: 6000 });

// 🖱️ Click the Billing Info tab
console.log('🖱️ Clicking Billing Info tab...');
await page.click('#VisitTabs > ul > li:nth-child(2)');

console.log('✅ Billing Info tab clicked successfully.');


} catch (error) {
    console.error('🚨 Failed to find or click select button:', error.message);
}

// Optional: Wait for search results to load
try {
    await newPage.waitForSelector('table[style*="margin:0;padding:0;"]', { timeout: 10000 });
    console.log('✅ Search results table loaded.');
} catch (error) {
    console.log('⚠️ Search results may still be loading...');
}
await new Promise(resolve => setTimeout(resolve, 2000)); // wait for 2s
let diagnosisNotFoundList = [];
let validDiagnosisIndexes = []; // ← stores valid fieldIndexes (1-based)

let i = 0;
let fieldIndex = 1;

while (i < diagnosisCodes.length) {
  const code = diagnosisCodes[i];
  const inputSelector = `#ctl00_phFolderContent_ucDiagnosisCodes_dc_10_${fieldIndex}`;

  try {
    console.log(`⏳ Processing diagnosis ${fieldIndex}: ${code}`);

    await page.waitForSelector(inputSelector, { visible: true, timeout: 5000 });
    const input = await page.$(inputSelector);
    if (!input) {
      console.warn(`⚠️ Input field not found: ${inputSelector}`);
      diagnosisNotFoundList.push(code);
      i++;
      continue;
    }

    await input.click({ clickCount: 3 });
    await page.keyboard.press('Delete');
    await new Promise(resolve => setTimeout(resolve, 200));
    await input.type(code, { delay: 50 });
    await new Promise(resolve => setTimeout(resolve, 1500));

    const noResults = await page.evaluate(() => {
      const noResultsElement = document.querySelector('#autocomplete-noresultsICD10');
      if (noResultsElement && noResultsElement.style.display !== 'none') return true;

      const items = document.querySelectorAll('.ui-autocomplete li');
      for (const item of items) {
        const text = item.textContent.trim().toLowerCase();
        if (text.includes('no results found')) return true;
      }
      return false;
    });

    if (noResults) {
      console.warn(`❌ No results for ${code}, trying next in same field`);
      diagnosisNotFoundList.push(code);
      await input.click({ clickCount: 3 });
      await page.keyboard.press('Delete');
      await page.keyboard.press('Escape');
      await new Promise(resolve => setTimeout(resolve, 500));
      i++; // next code, same field
      continue;
    }

    let selected = false;
    try {
      await input.focus();
      await new Promise(resolve => setTimeout(resolve, 200));
      await page.keyboard.press('ArrowDown');
      await new Promise(resolve => setTimeout(resolve, 300));
      await page.keyboard.press('Enter');
      selected = true;
      console.log(`✅ Selected ${code} via keyboard`);
    } catch {
      try {
        const firstSuggestion = await page.$('.ui-autocomplete li:first-child');
        if (firstSuggestion) {
          await firstSuggestion.click();
          selected = true;
          console.log(`✅ Selected ${code} via click`);
        }
      } catch {
        try {
          await input.focus();
          await page.keyboard.press('Tab');
          selected = true;
          console.log(`✅ Selected ${code} via Tab fallback`);
        } catch {
          console.warn(`⚠️ Could not select ${code}`);
        }
      }
    }

    if (selected) {
      validDiagnosisIndexes.push(fieldIndex); // Track valid field index
      i++;
      fieldIndex++; // Next field
    } else {
      // clear and stay in same field
      await input.click({ clickCount: 3 });
      await page.keyboard.press('Delete');
    }

    await new Promise(resolve => setTimeout(resolve, 800));

  } catch (err) {
    console.error(`❌ Error with ${code}:`, err.message);
    try {
      const input = await page.$(inputSelector);
      if (input) {
        await input.click({ clickCount: 3 });
        await page.keyboard.press('Delete');
      }
    } catch {}
    await page.keyboard.press('Escape').catch(() => {});
    await new Promise(resolve => setTimeout(resolve, 500));
    i++; // Move to next code even on error
  }
}

console.log('🏁 Completed diagnosis processing.');
if (diagnosisNotFoundList.length > 0) {
  console.log('📋 Diagnosis codes not found:', diagnosisNotFoundList);
} else {
  console.log('✅ All diagnosis codes processed successfully.');
}

const zeroBalanceCPTs = [];
// At the beginning of your main function
const cptNotFoundList = [];

// Example usage:
const cptCodeArray = uniqueCptCodes;
console.log(`📋 Found ${cptCodeArray.length} valid CPT codes:`, cptCodeArray);

if (cptCodeArray.length === 0) {
  console.error('❌ No valid CPT codes found!');
  return;
}

for (let i = 0; i < cptCodeArray.length; i++) {
  const currentCPT = cptCodeArray[i];
  console.log(`🔄 Processing CPT ${i + 1}/${cptCodeArray.length}: ${currentCPT}`);
  
  try {
    // Step 1: Enter FROM date and TO date
    console.log(`⌨️ Step 1: Entering billing dates for CPT ${currentCPT}`);
    
    // Wait for DOS From field
    const dosFromSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_DOS${i}`;
    await page.waitForSelector(dosFromSelector, { visible: true, timeout: 8000 });
    await page.focus(dosFromSelector);
    await page.click(dosFromSelector, { clickCount: 3 });
    await page.keyboard.type(billingDate, { delay: 50 });
    console.log(`✅ FROM date entered: ${billingDate}`);
    await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close


    // Wait for DOS To field
    const dosToSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_ToDOS${i}`;
    await page.waitForSelector(dosToSelector, { visible: true, timeout: 8000 });
    await page.focus(dosToSelector);
    await page.click(dosToSelector, { clickCount: 3 });
    await page.keyboard.type(billingDate, { delay: 50 });
    console.log(`✅ TO date entered: ${billingDate}`);
    await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close


    // Step 2: Search and select CPT code
    console.log(`🔍 Step 2: Searching for CPT code: ${currentCPT}`);
    
    // Wait for CPT search button to appear
    const searchButtonSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_btnUserCPT${i}`;
    await page.waitForSelector(searchButtonSelector, { visible: true, timeout: 6000 });

    // Set up popup listener BEFORE clicking the CPT button
    const cptPopupPromise = new Promise((resolve) => {
      const onPopup = async (target) => {
        console.log('✅ CPT popup detected!');
        browser.off('targetcreated', onPopup);
        const popup = await target.page();
        resolve(popup);
      };
      browser.on('targetcreated', onPopup);
    });

    // Click CPT search button
    console.log('🖱️ Clicking CPT search button...');
    await page.click(searchButtonSelector);

    // Wait for popup to appear
    console.log('⏳ Waiting for CPT popup to load...');
    let cptPage;
    try {
      // Wait for popup with timeout
      cptPage = await Promise.race([
        cptPopupPromise,
        new Promise((_, reject) => setTimeout(() => reject(new Error('CPT popup timeout')), 8000))
      ]);
      console.log('✅ CPT popup loaded successfully');
    } catch (error) {
      console.log('⚠️ No CPT popup detected, checking for new tabs...');
      
      // Fallback: Check for new tabs
      await new Promise(resolve => setTimeout(resolve, 3000));
      const allPages = await browser.pages();
      
      if (allPages.length > 1) {
        cptPage = allPages[allPages.length - 1]; // Get the newest page
        console.log('✅ New tab detected for CPT search...');
      } else {
        console.error(`🚨 No CPT popup or new tab found for CPT: ${currentCPT}`);
        cptNotFoundList.push(currentCPT);
        continue; // Skip to next CPT
      }
    }

    // Wait for the CPT page to load completely
    await new Promise(resolve => setTimeout(resolve, 3000));

    // Wait for the CPT search input field to appear in the popup
    console.log('⏳ Waiting for CPT search field in popup...');
    try {
      await cptPage.waitForSelector('input[name="ctl04$popupBase$txtSearch"]', { visible: true, timeout: 8000 });
      console.log('✅ CPT search field found in popup');
    } catch (error) {
      console.log('🔍 Specific CPT search field not found, trying generic text input...');
      
      const textInputs = await cptPage.$$('input[type="text"]');
      if (textInputs.length === 0) {
        console.error(`🚨 No text input found in CPT popup for CPT: ${currentCPT}`);
        cptNotFoundList.push(currentCPT);
        await cptPage.close();
        continue; // Skip to next CPT
      }
      console.log(`Found ${textInputs.length} text input(s) in CPT popup`);
    }

    // Clear and enter CPT search term in the popup
    console.log(`⌨️ Entering CPT search term in popup: ${currentCPT}`);

    // Find and clear the search field
    await cptPage.evaluate(() => {
      const searchField = document.querySelector('input[name="ctl04$popupBase$txtSearch"]') || 
                          document.querySelector('input[type="text"]');
      if (searchField) {
        searchField.value = '';
        searchField.focus();
      }
    });

    // Type the CPT search term
    const searchSelector = 'input[name="ctl04$popupBase$txtSearch"]';
    try {
      await cptPage.type(searchSelector, currentCPT, { delay: 100 });
    } catch (error) {
      try {
        await cptPage.type('input[type="text"]', currentCPT, { delay: 100 });
      } catch (typeError) {
        console.error(`🚨 Failed to type CPT code ${currentCPT}:`, typeError.message);
        cptNotFoundList.push(currentCPT);
        await cptPage.close();
        continue; // Skip to next CPT
      }
    }

    // Verify the text was entered
    const enteredText = await cptPage.evaluate(() => {
      const field = document.querySelector('input[name="ctl04$popupBase$txtSearch"]') || 
                      document.querySelector('input[type="text"]');
      return field ? field.value : 'Could not verify';
    });
    console.log('✅ CPT search text entered in popup:', enteredText);

    // Find and click the Search button in the popup
    console.log('🔍 Looking for Search button in CPT popup...');
    try {
      await cptPage.waitForSelector('input[name="ctl04$popupBase$btnSearch"]', { visible: true, timeout: 5000 });
      console.log('🖱️ Clicking Search button in CPT popup...');
      await cptPage.click('input[name="ctl04$popupBase$btnSearch"]');
    } catch (error) {
      console.log('🔍 Specific search button not found, trying alternatives...');
      
      const searchButtonClicked = await cptPage.evaluate(() => {
        const searchButtons = document.querySelectorAll('input[value="Search"], input[type="button"], input[type="submit"], button');
        for (const btn of searchButtons) {
          if (btn.value?.toLowerCase().includes('search') || 
              btn.textContent?.toLowerCase().includes('search') ||
              btn.onclick?.toString().includes('search')) {
            btn.click();
            return true;
          }
        }
        if (searchButtons.length > 0) {
          searchButtons[0].click();
          return true;
        }
        return false;
      });
      
      if (!searchButtonClicked) {
        console.error(`🚨 No search button found for CPT: ${currentCPT}`);
        cptNotFoundList.push(currentCPT);
        await cptPage.close();
        continue; // Skip to next CPT
      }
    }
// Wait for search results to load
console.log('⏳ Waiting for CPT search results...');
await new Promise(resolve => setTimeout(resolve, 3000));

// ⏳ Wait for CPT search results to load
console.log('⏳ Waiting for CPT search results...');
await new Promise(resolve => setTimeout(resolve, 3000));

// ✅ ONLY set this if "No data results to display." is found
let cptNoResults = false;

try {
  cptNoResults = await cptPage.evaluate(() => {
    const tdElements = Array.from(document.querySelectorAll('td'));
    return tdElements.some(td => td.textContent.trim().toLowerCase() === 'no data results to display.');
  });

  if (cptNoResults) {
    console.warn(`❌ CPT not found: ${currentCPT}`);
    cptNotFoundList.push(currentCPT);  // ✅ Only here CPT is logged
    await cptPage.close();
    continue; // ✅ Skip this CPT and move to the next
  }
} catch (err) {
  // 🔇 Do NOT log CPT not found here — this is just a generic failure
  console.warn(`⚠️ Could not evaluate CPT results for ${currentCPT}:`, err.message);
  await cptPage.close();
  continue; // Gracefully skip without logging CPT as not found
}

// If reached here, CPT was found in popup
console.log(`✅ CPT found: ${currentCPT}`);


try {
  console.log('⏳ Waiting for CPT search results to appear...');
  await cptPage.waitForSelector('#ctl04_popupBase_grvPopup_ctl02_lnkSelect', {
    visible: true,
    timeout: 10000
  });

  console.log('🖱️ Selecting CPT code from search results...');
  await cptPage.click('#ctl04_popupBase_grvPopup_ctl02_lnkSelect');
  console.log('✅ CPT code selected successfully');

} catch (error) {
  console.error(`🚨 Failed to find specific select link for CPT ${currentCPT}, trying alternatives...`);

  const selectLinkFound = await cptPage.evaluate(() => {
    const selectLinks = document.querySelectorAll('a[id*="lnkSelect"]');
    if (selectLinks.length > 0) {
      selectLinks[0].click();
      return true;
    }

    const allLinks = document.querySelectorAll('a');
    for (const link of allLinks) {
      if (link.textContent.trim().toLowerCase() === 'select') {
        link.click();
        return true;
      }
    }

    const resultRows = document.querySelectorAll('table[id*="grvPopup"] tbody tr, #ctl04_popupBase_grvPopup tbody tr');
    if (resultRows.length > 0) {
      const firstLink = resultRows[0].querySelector('a');
      if (firstLink) {
        firstLink.click();
      } else {
        resultRows[0].click();
      }
      return true;
    }
    return false;
  });

  if (selectLinkFound) {
    console.log(`✅ Selected CPT ${currentCPT} using alternative method`);
  } else {
    console.error(`🚨 Failed to select CPT code: ${currentCPT}`);
    cptNotFoundList.push(currentCPT);

    if (!cptPage.isClosed()) {
      await cptPage.close(); // ✅ Only closes popup, not parent
      console.log(`❌ Closed CPT popup for ${currentCPT}`);
    }

    continue; // 🔒 Skip to next CPT
  }
}

// ✅ Wait until popup disappears or force close it if needed
try {
  await cptPage.waitForFunction(() => window.closed || document.hidden, { timeout: 5000 });
  console.log(`✅ CPT popup closed after selecting: ${currentCPT}`);
} catch (error) {
  if (!cptPage.isClosed()) {
    await cptPage.close(); // ✅ Safe close
    console.log(`🔄 Force closed CPT popup for: ${currentCPT}`);
  }
}

} catch (error) {
  console.error(`🚨 Error processing CPT ${currentCPT}:`, error.message);

  try {
    if (!cptPage.isClosed()) {
      await cptPage.close(); // ✅ Close only popup
      console.log(`❌ Closed CPT popup for ${currentCPT} due to error`);
    }
  } catch (cleanupError) {
    console.warn('⚠️ Error during CPT popup cleanup:', cleanupError.message);
  }

  continue; // Skip to next CPT
}


// Log final results
console.log('\n📊 Final CPT Processing Results:');
console.log('✅ Zero Balance CPTs:', zeroBalanceCPTs);
console.log('❌ CPT Not Found List:', cptNotFoundList);
console.log(`📈 Total CPTs processed: ${cptCodeArray.length}`);
console.log(`❌ Total CPTs not found: ${cptNotFoundList.length}`);


  // Step 3: Enter POS, Modifier, and ICD-10 pointer
  console.log(`📝 Step 3: Entering additional fields for CPT ${currentCPT}`);
await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close

  // Enter the POS value
  if (typeof pos === 'string' && pos.trim() !== '') {
    console.log(`🏥 Entering POS: ${pos}`);
    const posSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_PlaceOfService${i}`;
    try {
      await page.waitForSelector(posSelector, { visible: true, timeout: 5000 });
      await page.evaluate((value, selector) => {
        const input = document.querySelector(selector);
        if (input) input.value = value;
      }, pos, posSelector);
      console.log('✅ POS entered successfully');
    } catch (error) {
      console.error('❌ Error entering POS:', error.message);
    }
  }
await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close

// Enter the Modifier value — only for the first CPT code (i === 0)
if (
  i === 0 &&
  modifier != null &&
  typeof modifier === 'string' &&
  modifier.trim() !== '' &&
  modifier.trim() !== '-'
) {
  console.log(`🔧 Entering Modifier (only in first CPT row): ${modifier}`);
  const modifierSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_ModifierA${i}`;
  try {
    await page.waitForSelector(modifierSelector, { visible: true, timeout: 5000 });
    await page.evaluate((value, selector) => {
      const input = document.querySelector(selector);
      if (input) input.value = value;
    }, modifier.trim(), modifierSelector);
    console.log('✅ Modifier entered successfully');
  } catch (error) {
    console.error('❌ Error entering Modifier:', error.message);
  }
}

await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for popup to close

  // Set ICD-10 Pointer
  try {
  const pointerSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_DiagnosisCode${i}`;
  const icdPointer = validDiagnosisIndexes.map(index => index.toString()).join('').slice(0, 4);

  if (icdPointer.length > 0) {
    console.log(`🩺 Setting ICD-10 Pointer to: ${icdPointer}`);
    await page.waitForSelector(pointerSelector, { visible: true, timeout: 5000 });
    await page.evaluate((value, selector) => {
      const input = document.querySelector(selector);
      if (input) input.value = value;
    }, icdPointer, pointerSelector);
    console.log('✅ ICD-10 pointer set successfully');
  } else {
    console.warn('⚠️ No valid diagnosis codes selected. Pointer skipped.');
  }
} catch (error) {
  console.error('❌ Error setting ICD-10 pointer:', error.message);
}

  await new Promise(resolve => setTimeout(resolve, 1000));

const unitSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_Quantity${i}`;
const unitValue = cptUnits[i];

if (unitValue) {
  console.log(`📦 Entering Units for CPT ${currentCPT}: ${unitValue}`);

  await page.waitForSelector(unitSelector, { visible: true, timeout: 7000 });
  const [popupPromise] = await Promise.all([
    new Promise(resolve => browser.once('targetcreated', target => resolve(target.page()))),
    page.click(unitSelector, { clickCount: 2 }),  // Double-click opens popup
  ]);

  const popup = await popupPromise;
  await popup.bringToFront();

  await popup.waitForSelector('#ucEditCharge_Quantity', { visible: true, timeout: 5000 });

  // Clear and type new unit value
  await popup.evaluate((value) => {
    const input = document.querySelector('#ucEditCharge_Quantity');
    if (input) {
      input.focus();
      input.value = '';
    }
  }, unitValue);
  await new Promise(resolve => setTimeout(resolve, 1000));
  await popup.type('#ucEditCharge_Quantity', unitValue.toString(), { delay: 100 });

  console.log('📝 Unit value entered in popup window');
await new Promise(resolve => setTimeout(resolve, 100));

  // Click update
  await popup.click('#btnUpdate');
  console.log('✅ Clicked update in popup');
await new Promise(resolve => setTimeout(resolve, 1000));

  // Gracefully handle window close
  try {
    await popup.waitForNavigation({ timeout: 3000 }).catch(() => {}); // optional wait
    if (!popup.isClosed()) {
      await popup.close();
      console.log('🧹 Popup manually closed');
    } else {
      console.log('🧹 Popup was auto-closed after update');
    }
  } catch (err) {
    console.log('⚠️ Error during popup close (probably already closed):', err.message);
  }

} else {
  console.log(`🚫 Skipping Units for CPT ${currentCPT} as no units specified.`);
}



// 🔎 Check balance after CPT entry
try {
    const balanceSelector = `#ctl00_phFolderContent_ucVisitLineItem_ucBillingCPT_Balance${i}`;
    await page.waitForSelector(balanceSelector, { visible: true, timeout: 5000 });

    const balanceText = await page.$eval(balanceSelector, el => {
        if (el.tagName === 'INPUT') return el.value.trim();
        return el.textContent.trim();
    });

    const cleanedBalance = balanceText.replace(/[^0-9.-]/g, '');
    const balanceValue = parseFloat(cleanedBalance);

    if (cleanedBalance === "0" || cleanedBalance === "0.00") {
    console.log(`💰 CPT ${currentCPT} (Row ${i + 1}) - ZERO BALANCE DETECTED: $${cleanedBalance}`);
    console.log(`🔍 ${currentCPT} Billed Fee is Zero`);

        zeroBalanceCPTs.push({
            CPTCode: currentCPT,
            Row: i + 1,
            Balance: cleanedBalance,
            Modifier: i === 0 ? modifier || '' : '',
            Units: cptUnits[i] || '1',
            POS: pos || '',
            Date: billingDate,
            Status: 'Zero Balance'
        });
    } else if (!isNaN(balanceValue) && balanceValue > 0.01) {
        console.log(`💲 CPT ${currentCPT} (Row ${i + 1}) - NON-ZERO BALANCE: $${cleanedBalance}`);
        console.log(`✅ ${currentCPT} billed fee is $${cleanedBalance}`);
    } else {
        console.log(`⚠️ CPT ${currentCPT} (Row ${i + 1}) - UNABLE TO PARSE BALANCE: "${balanceText}"`);
    }

} catch (balErr) {
    console.error(`⚠️ Failed to fetch balance for CPT ${currentCPT} (Row ${i + 1}):`, balErr.message);
}


// ✅ Now continue with the rest
await new Promise(resolve => setTimeout(resolve, 1000));
console.log(`🎉 CPT ${currentCPT} processing completed for row ${i}`);


  // Add delay before processing next CPT (if any)
  if (i < cptCodeArray.length - 1) {
    console.log('⏳ Preparing for next CPT code...');
    await new Promise(resolve => setTimeout(resolve, 2000));
  }
}
  await new Promise(resolve => setTimeout(resolve, 1000));

// ⏳ Wait for the Billing Info tab to be visible
console.log('⏳ Waiting for Billing Options tab...');
await page.waitForSelector('#VisitTabs > ul > li:nth-child(3)', { visible: true, timeout: 6000 });

// 🖱️ Click the Billing Info tab
console.log('🖱️ Clicking Billing Info tab...');
await page.click('#VisitTabs > ul > li:nth-child(3)');

console.log('✅ Billing Options tab clicked successfully.');
await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close
// ⏳ Wait for popup to load
console.log('⏳ Waiting for Provider Search button to appear...');
await page.waitForSelector('#div1 > table > tbody > tr:nth-child(2) > td:nth-child(2) > input.button', { visible: true, timeout: 6000 });

console.log('🖱️ Setting up provider popup listener...');

// Set up popup listener BEFORE clicking the button
const providerPopupPromise = new Promise((resolve) => {
    const onPopup = async (target) => {
        console.log('✅ Provider popup detected!');
        browser.off('targetcreated', onPopup);
        const popup = await target.page();
        resolve(popup);
    };
    browser.on('targetcreated', onPopup);
});

// 🔘 Click provider search button
console.log('🖱️ Clicking provider search button...');
await page.click('#div1 > table > tbody > tr:nth-child(2) > td:nth-child(2) > input.button');

// Wait for popup to appear
console.log('⏳ Waiting for provider popup to load...');
let providerPage;
try {
    // Wait for popup with timeout
    providerPage = await Promise.race([
        providerPopupPromise,
        new Promise((_, reject) => setTimeout(() => reject(new Error('Provider popup timeout')), 8000))
    ]);
    console.log('✅ Provider popup loaded successfully');
} catch (error) {
    console.log('⚠️ No provider popup detected, checking for new tabs...');
    
    // Fallback: Check for new tabs
    await new Promise(resolve => setTimeout(resolve, 3000));
    const allPages = await browser.pages();
    
    if (allPages.length > 1) {
        providerPage = allPages[allPages.length - 1]; // Get the newest page
        console.log('✅ New tab detected for provider search...');
    } else {
        throw new Error('No provider popup or new tab found');
    }
}

// Wait for the provider page to load completely
await new Promise(resolve => setTimeout(resolve, 3000));

// Extract FIRST NAMES from provider list
let providerFirstNames = [];
if (typeof provider !== 'undefined' && provider) {
    // Split by newlines or commas to get individual providers
    const providers = provider.split(/\n|,(?=\s*[A-Z])/); // Split by newlines or commas followed by capital letter
    
    providers.forEach(providerEntry => {
        const trimmedProvider = providerEntry.trim();
        if (trimmedProvider) {
            // Format: "FirstName LastName, MD" - e.g., "Michael Lin, MD"
            if (trimmedProvider.includes(',')) {
                // Split by comma and take the part before comma
                const beforeComma = trimmedProvider.split(',')[0].trim();
                // Get first word (first name)
                const firstName = beforeComma.split(' ')[0];
                providerFirstNames.push(firstName);
            } else {
                // If no comma, just take first word
                const firstName = trimmedProvider.split(' ')[0];
                providerFirstNames.push(firstName);
            }
        }
    });
    
    console.log(`📝 Extracted FIRST NAMES from providers: ${providerFirstNames.join(', ')}`);
}

// If you want just the first provider's first name:
let providerFirstName = providerFirstNames.length > 0 ? providerFirstNames[0] : '';

// NEW: Select "First Name" from dropdown
console.log('🔍 Looking for search type dropdown to select First Name...');
try {
    // Wait for the dropdown to be available
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    const searchTypeDropdownSelectors = [
        'select[name*="LastName"]', // This might be the dropdown name even if it has First Name option
        'select:first-of-type',
        'select'
    ];
    
    let searchTypeDropdown = null;
    for (const selector of searchTypeDropdownSelectors) {
        try {
            await providerPage.waitForSelector(selector, { visible: true, timeout: 2000 });
            searchTypeDropdown = selector;
            console.log(`✅ Found search type dropdown with selector: ${selector}`);
            break;
        } catch (e) {
            continue;
        }
    }
    
    if (searchTypeDropdown) {
        // Get available options from the dropdown
        const availableOptions = await providerPage.evaluate((selector) => {
            const dropdown = document.querySelector(selector);
            if (!dropdown) return [];
            
            const options = [];
            for (let i = 0; i < dropdown.options.length; i++) {
                const option = dropdown.options[i];
                options.push({
                    value: option.value,
                    text: option.text
                });
            }
            return options;
        }, searchTypeDropdown);
        
        console.log('📋 Available search type options:', availableOptions);
        
        // Select "First Name" option
        const firstNameOption = availableOptions.find(opt => 
            opt.text.toLowerCase().includes('first name') || 
            opt.text.toLowerCase().includes('firstname') ||
            opt.value.toLowerCase().includes('firstname') ||
            opt.value.toLowerCase().includes('first')
        );
        
        if (firstNameOption) {
            console.log(`🎯 Selecting search type: ${firstNameOption.text} (value: ${firstNameOption.value})`);
            await providerPage.select(searchTypeDropdown, firstNameOption.value);
            console.log('✅ First Name selected from dropdown');
        } else {
            console.log('⚠️ First Name option not found in dropdown, using default');
        }
    }
    
    // Wait a moment for the selection to process
    await new Promise(resolve => setTimeout(resolve, 1000));
    
} catch (error) {
    console.log('🚨 Error selecting First Name from dropdown:', error.message);
}

// Wait for the specific provider search input field to appear
console.log('⏳ Waiting for Provider search field...');
try {
    const searchFieldSelectors = [
        'input[type="text"]',
        'input[name*="txtSearch"]',
        'input[id*="txtSearch"]',
        'input[name*="Search"]'
    ];
    
    let searchField = null;
    for (const selector of searchFieldSelectors) {
        try {
            await providerPage.waitForSelector(selector, { visible: true, timeout: 3000 });
            searchField = selector;
            console.log(`✅ Found search field with selector: ${selector}`);
            break;
        } catch (e) {
            continue;
        }
    }
    
    if (!searchField) {
        throw new Error('No provider search field found');
    }
    
    // Clear and enter provider FIRST NAME
    console.log(`⌨️ Entering Provider FIRST NAME: ${providerFirstName}`);
    
    // Clear the field first
    await providerPage.evaluate((selector) => {
        const field = document.querySelector(selector);
        if (field) {
            field.value = '';
            field.focus();
        }
    }, searchField);
    
    // Type the first name
    await providerPage.type(searchField, providerFirstName, { delay: 100 });
    
    // Verify the text was entered
    const enteredText = await providerPage.evaluate((selector) => {
        const field = document.querySelector(selector);
        return field ? field.value : 'Could not verify';
    }, searchField);
    console.log('✅ Provider FIRST NAME entered:', enteredText);
    
} catch (error) {
    console.log('🚨 Error with search field:', error.message);
    throw error;
}

// Find and click the Search button
console.log('🔍 Looking for Search button...');
try {
    const searchButtonSelectors = [
        'input[value="Search"]',
        'button[onclick*="Search"]',
        'input[type="button"][value*="Search"]',
        'input[id*="Search"]',
        'input[name*="Search"]'
    ];
    
    let searchButtonFound = false;
    for (const selector of searchButtonSelectors) {
        try {
            await providerPage.waitForSelector(selector, { visible: true, timeout: 2000 });
            console.log(`🖱️ Clicking Search button with selector: ${selector}`);
            await providerPage.click(selector);
            searchButtonFound = true;
            break;
        } catch (e) {
            continue;
        }
    }
    
    if (!searchButtonFound) {
        // Try clicking any button that might be the search button
        await providerPage.evaluate(() => {
            const searchButtons = document.querySelectorAll('input[type="button"], input[type="submit"], button');
            for (const btn of searchButtons) {
                if (btn.value?.toLowerCase().includes('search') || 
                    btn.textContent?.toLowerCase().includes('search') ||
                    btn.onclick?.toString().includes('search')) {
                    console.log('Found search button by content, clicking...');
                    btn.click();
                    return true;
                }
            }
            return false;
        });
    }
} catch (error) {
    console.log('🚨 Search button click failed:', error.message);
}

console.log('⏳ Waiting for provider search results...');
await new Promise(resolve => setTimeout(resolve, 3000));

// Check if there are no results
let providerHasNoResults = false;
try {
  providerHasNoResults = await providerPage.evaluate(() => {
    const tdElements = Array.from(document.querySelectorAll('td'));
    return tdElements.some(td => td.textContent.trim().toLowerCase() === 'no data results to display.');
  });

  if (providerHasNoResults) {
    console.warn('❌ No provider search results found, closing popup...');
    await providerPage.close();
  }
} catch (err) {
  console.warn('⚠️ Error checking provider result message:', err.message);
}

// If results exist, proceed to selection
if (!providerHasNoResults) {
  try {
    // Try various selectors that may appear
    const resultSelectors = [
      'a[onclick*="Select"]',
      'a[href*="Select"]',
      'input[value="Select"]',
      'button[onclick*="Select"]',
      '.grid tbody tr',
      'table tr td a'
    ];

    let found = false;
    for (const selector of resultSelectors) {
      try {
        await providerPage.waitForSelector(selector, { visible: true, timeout: 3000 });
        console.log(`✅ Found result using selector: ${selector}`);
        found = true;
        break;
      } catch (_) {}
    }

    if (found) {
      console.log('🖱️ Attempting to select the first provider...');
      const success = await providerPage.evaluate(() => {
        const selectLink = document.querySelector('a[onclick*="Select"], a[href*="Select"]');
        if (selectLink) {
          selectLink.scrollIntoView({ behavior: 'smooth', block: 'center' });
          selectLink.click();
          return true;
        }

        const selectButton = document.querySelector('input[value="Select"], button[onclick*="Select"]');
        if (selectButton) {
          selectButton.click();
          return true;
        }

        const tableRows = document.querySelectorAll('table tr');
        for (let i = 1; i < tableRows.length; i++) {
          const row = tableRows[i];
          const clickable = row.querySelector('a, button, input[type="button"]');
          if (clickable) {
            clickable.click();
            return true;
          }
        }

        return false;
      });

      if (success) {
        console.log('✅ Provider selected successfully.');
      } else {
        console.warn('⚠️ Could not find a clickable provider element.');
      }
    } else {
      console.warn('⚠️ Provider search result selectors not found.');
    }
  } catch (err) {
    console.error('🚨 Error during provider selection:', err.message);
  }
}


// Wait for the selection to process and popup to close
await new Promise(resolve => setTimeout(resolve, 5000));

console.log('🎉 Provider selection process completed.');

// ⏳ Wait for popup trigger button
console.log('⏳ Waiting for Facility Search button...');
await page.waitForSelector('#ctl00_phFolderContent_Button35', { visible: true, timeout: 6000 });

console.log('🖱️ Setting up facility popup listener...');

// Set up popup listener BEFORE clicking the button
const facilityPopupPromise = new Promise((resolve) => {
    const onPopup = async (target) => {
        console.log('✅ Facility popup detected!');
        browser.off('targetcreated', onPopup);
        const popup = await target.page();
        resolve(popup);
    };
    browser.on('targetcreated', onPopup);
});

// 🔘 Click facility search button
console.log('🖱️ Clicking facility search button...');
await page.click('#ctl00_phFolderContent_Button35');

// Wait for popup
console.log('⏳ Waiting for facility popup to load...');
let facilityPage;
try {
    facilityPage = await Promise.race([
        facilityPopupPromise,
        new Promise((_, reject) => setTimeout(() => reject(new Error('Facility popup timeout')), 8000))
    ]);
    console.log('✅ Facility popup loaded');
} catch (error) {
    const allPages = await browser.pages();
    if (allPages.length > 1) {
        facilityPage = allPages[allPages.length - 1];
        console.log('✅ Detected new tab as facility popup');
    } else {
        throw new Error('❌ No popup or new tab for facility found');
    }
}

// ⏳ Wait for search field
console.log('⏳ Waiting for facility search input...');
await facilityPage.waitForSelector('#ctl04_popupBase_txtSearch', { visible: true, timeout: 6000 });

// 💬 Enter facility name
console.log(`⌨️ Entering Facility Name: ${facilityName}`);
await facilityPage.evaluate((name) => {
    const input = document.querySelector('#ctl04_popupBase_txtSearch');
    if (input) {
        input.value = '';
        input.focus();
    }
}, facilityName);

await facilityPage.type('#ctl04_popupBase_txtSearch', facilityName, { delay: 100 });

// 🔍 Click Search button
console.log('🖱️ Clicking facility search button...');
await facilityPage.waitForSelector('#ctl04_popupBase_btnSearch', { visible: true, timeout: 4000 });
await facilityPage.click('#ctl04_popupBase_btnSearch');

console.log('⏳ Waiting for facility search results...');
await new Promise(resolve => setTimeout(resolve, 3000)); // Optional wait

let facilityHasNoResults = false;
let facilityNotFoundFlag = false; // ✅ NEW FLAG to track facility not found

try {
  facilityHasNoResults = await facilityPage.evaluate(() => {
    const cells = Array.from(document.querySelectorAll('td'));
    return cells.some(td => td.textContent.trim().toLowerCase() === 'no data results to display.');
  });

  if (facilityHasNoResults) {
    console.warn('❌ No facility search results found, closing popup...');
    facilityNotFoundFlag = true; // ✅ SET FLAG
    await facilityPage.close();
  }
} catch (err) {
  console.warn('⚠️ Error checking facility results:', err.message);
}

if (!facilityHasNoResults) {
  try {
    await facilityPage.waitForSelector('a[id$="_lnkSelect"]', { visible: true, timeout: 8000 });
    console.log('🖱️ Selecting the first facility...');
    await facilityPage.evaluate(() => {
      const link = document.querySelector('a[id$="_lnkSelect"]');
      if (link) {
        link.scrollIntoView({ behavior: 'smooth', block: 'center' });
        link.click();
      }
    });
    await new Promise(resolve => setTimeout(resolve, 3000));
    console.log('✅ Facility selected.');
  } catch (err) {
    console.error('🚨 Error selecting facility:', err.message);
  }
}

// ✅ Done
console.log('🎉 Facility selection completed!');
await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for popup to close

console.log('🔍 Processing complete. Please verify all entries on the form.');
if (['21', '61', '31'].includes(pos)) {
  console.log(`🩺 POS is ${pos}, entering Hospital From Date: ${admitDate}`);

  try {
    const [month, day, year] = admitDate.split('/');

    // Month field
    const monthSelector = '#ctl00_phFolderContent_HospitalFromDate_Month';
    await page.waitForSelector(monthSelector, { visible: true, timeout: 10000 });
    await page.focus(monthSelector);
    await page.type(monthSelector, month, { delay: 50 });
    await page.keyboard.press('Tab');

    // Day field
    const daySelector = '#ctl00_phFolderContent_HospitalFromDate_Day';
    await page.waitForSelector(daySelector, { visible: true, timeout: 3000 });
    await page.focus(daySelector);
    await page.type(daySelector, day, { delay: 50 });
    await page.keyboard.press('Tab');

    // Year field
    const yearSelector = '#ctl00_phFolderContent_HospitalFromDate_Year';
    await page.waitForSelector(yearSelector, { visible: true, timeout: 3000 });
    await page.focus(yearSelector);
    await page.type(yearSelector, year, { delay: 50 });

    // ✅ Do NOT press Tab or click again — just let it be
    console.log('✅ Hospital From Date filled successfully.');
  } catch (error) {
    console.log(`❌ Error filling Hospital From Date: ${error.message}`);
  }
} else {
  console.log(`ℹ️ POS is ${pos}. Skipping Hospital From Date entry.`);
}

await new Promise(resolve => setTimeout(resolve, 1500)); // Shorter wait


// Handle Prior Authorization Number
if (priornumber && priornumber.trim() !== '' && priornumber.trim() !== '-') {
    console.log(`📝 Entering Prior Authorization Number: ${priornumber}`);
    
    try {
        // Wait for the prior auth field with longer timeout
        await page.waitForSelector('#ctl00_phFolderContent_PriorAuthorizationNumber', { visible: true, timeout: 10000 });

        // Scroll to the field to ensure it's visible
        await page.evaluate(() => {
            const priorField = document.querySelector('#ctl00_phFolderContent_PriorAuthorizationNumber');
            if (priorField) {
                priorField.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
        });

        // Wait for scroll to complete
        await new Promise(resolve => setTimeout(resolve, 1000));

        // Click on the field to focus it
        await page.click('#ctl00_phFolderContent_PriorAuthorizationNumber');

        // Clear the field using keyboard shortcuts
        await page.keyboard.down('Control');
        await page.keyboard.press('KeyA');
        await page.keyboard.up('Control');
        await page.keyboard.press('Delete');

        // Type the prior authorization number
        await page.type('#ctl00_phFolderContent_PriorAuthorizationNumber', priornumber);

        // Verify the value was entered
        const enteredValue = await page.$eval('#ctl00_phFolderContent_PriorAuthorizationNumber', el => el.value);
        console.log(`✅ Prior Authorization Number entered. Value: ${enteredValue}`);

        // Trigger change event to ensure the form recognizes the input
        await page.evaluate(() => {
            const priorField = document.querySelector('#ctl00_phFolderContent_PriorAuthorizationNumber');
            if (priorField) {
                priorField.dispatchEvent(new Event('change', { bubbles: true }));
                priorField.dispatchEvent(new Event('input', { bubbles: true }));
            }
        });

    } catch (error) {
        console.log(`❌ Error entering Prior Authorization Number: ${error.message}`);
        
        // Try alternative approach if the main selector fails
        try {
            console.log('🔄 Trying alternative approach for Prior Auth field...');
            
            // Try to find the field by partial name match
            const alternativeSelector = 'input[name*="PriorAuthorizationNumber"]';
            await page.waitForSelector(alternativeSelector, { visible: true, timeout: 5000 });
            
            await page.click(alternativeSelector);
            await page.keyboard.down('Control');
            await page.keyboard.press('KeyA');
            await page.keyboard.up('Control');
            await page.keyboard.press('Delete');
            await page.type(alternativeSelector, priornumber);
            
            console.log('✅ Prior Authorization Number entered using alternative selector.');
            
        } catch (altError) {
            console.log(`❌ Alternative approach also failed: ${altError.message}`);
        }
    }
} else {
    console.log('ℹ️ Prior Authorization Number not provided. Skipping field.');
}

// Final wait before proceeding
await new Promise(resolve => setTimeout(resolve, 1000));

// ✅ Click Update Button
await page.waitForSelector('#ctl00_phFolderContent_btnUpdate', { visible: true });
await page.click('#ctl00_phFolderContent_btnUpdate');
console.log("🚀 Update button clicked.");
await new Promise(resolve => setTimeout(resolve, 1000));

// ✅ Define Selectors
const visitIdSelector = '.message-header span';
const alertOkButtonSelector = 'body > div.ui-dialog.ui-widget.ui-widget-content.ui-corner-all.ui-front.quickadd.oa-popup-gray.ui-dialog-buttons.ui-draggable > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button';
const cancelButtonSelector = '#ctl00_phFolderContent_btnCancel';

let visitId = null;
let success = false;

try {
  // ⏳ Wait for either Visit ID message OR Alert popup
  const result = await Promise.race([
    page.waitForSelector(visitIdSelector, { visible: true }).then(() => 'visit'),
    page.waitForSelector(alertOkButtonSelector, { visible: true }).then(() => 'alert')
  ]);

  if (result === 'visit') {
    // ✅ Visit ID appeared
    const messageText = await page.$eval(visitIdSelector, el => el.textContent.trim());
    const visitIdMatch = messageText.match(/Visit\s+(\d+)/);

    if (visitIdMatch) {
      visitId = visitIdMatch[1];
      success = true;
      console.log("✅ Visit ID found:", visitId);
    } else {
      console.log("❌ Visit ID pattern not matched.");
    }

  } else if (result === 'alert') {
    // ❌ Alert appeared
    console.log("⚠️ Alert popup detected. Handling as failure...");

    await page.click(alertOkButtonSelector);
    console.log("✅ OK button clicked on alert popup.");

    // 🛑 Click Cancel Button
    await page.waitForSelector(cancelButtonSelector, { visible: true });
    await page.click(cancelButtonSelector);
    console.log("🚪 Cancel button clicked.");
  }

} catch (err) {
  console.error("❌ Error during update result handling:", err.message);
}

// 🕒 Final delay
await new Promise(resolve => setTimeout(resolve, 1000));

// 🔚 Final return with success/failure
return { success, visitId, zeroBalanceCPTs, facilityNotFound: facilityNotFoundFlag, cptNotFoundList ,diagnosisNotFoundList };



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
        
        // Check if "Patient ID" exists in headers
        console.log(headers);
        if (!headers.includes("Patient ID")) {
            throw new Error("'Patient ID' Column is not defined.");
        }
        
        // Map headers to their index positions
        headers.forEach((header, index) => {
            columnMap[header] = index;
        });
        
        const requiredColumns = ["Patient ID", "Provider", "Facility", "DOS"];
        for (const column of requiredColumns) {
            if (!headers.includes(column)) {
                throw new Error(`Required column '${column}' is missing from the spreadsheet.`);
            }
        }
    
       // Ensure "Visit ID" column is added if missing (before Result)
const visitIdHeader = "Visit ID";
if (!headers.includes(visitIdHeader)) {
    // Find the position of Result column or add at the end
    const resultIndex = headers.indexOf("Result");
    const insertIndex = resultIndex !== -1 ? resultIndex : headers.length;

    // Insert Visit ID column
    headers.splice(insertIndex, 0, visitIdHeader);

    // Add empty Visit ID cells to all existing rows
    for (let i = 1; i < originalData.length; i++) {
        originalData[i] = originalData[i] || [];
        originalData[i].splice(insertIndex, 0, "");
    }
}

// Ensure "Billed Fee" column is added if missing (before Result)
const billedFeeHeader = "Billed Fee";
if (!headers.includes(billedFeeHeader)) {
    // Find the updated position of Result column (may have shifted after Visit ID insertion)
    const resultIndex = headers.indexOf("Result");
    const insertIndex = resultIndex !== -1 ? resultIndex : headers.length;

    // Insert Billed Fee column
    headers.splice(insertIndex, 0, billedFeeHeader);

    // Add empty Billed Fee cells to all existing rows
    for (let i = 1; i < originalData.length; i++) {
        originalData[i] = originalData[i] || [];
        originalData[i].splice(insertIndex, 0, "");
    }
}

// Ensure only "Result" column is added if missing
const resultHeader = "Result";
if (!headers.includes(resultHeader)) {
    columnMap[resultHeader] = headers.length;
    headers.push(resultHeader);
    for (let i = 1; i < originalData.length; i++) {
        originalData[i] = originalData[i] || [];
        originalData[i].push("");
    }
}


        // Update column mapping after adding new columns
        headers.forEach((header, idx) => columnMap[header] = idx);
        
        let newRowsToProcess = 0;
        const processedAccounts = new Set();
        
        let missingCallDiagnoses = 0;  // Track rows missing Diagnoses

        for (let i = 1; i < originalData.length; i++) {
            const row = originalData[i];
            if (!row) continue; // Skip undefined rows
            
            const accountNumberRaw = row[columnMap["Patient ID"]] || "";
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
        
        newPage = await continueWithLoggedInSession();

        
        const results = { successful: [], failed: [] };
        
        for (let i = 1; i < originalData.length; i++) {
            const row = originalData[i];
            if (!row) continue;
        
            const accountNumberRaw = row[columnMap["Patient ID"]] || "";
            const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
            if (!accountNumberRaw || status === "done") continue;
        
            const provider = row[columnMap["Provider"]] || "";
            const facilityName = row[columnMap["Facility"]] || "";
            const billingDate = row[columnMap["DOS"]] || "";
            const diagnosisText = row[columnMap["Diagnoses"]] || "";
            const chargesText = row[columnMap["CPT"]] || ""; // CPT codes
            const admitDate = row[columnMap["Hospital Date"]] || ""; // ✅ NEW LINE
            const pos = row[columnMap["POS"]] || ""; // ✅ NEW LINE
            const modifier = row[columnMap["Modifier"]] || ""; // ✅ NEW LINE
            const priornumber = row[columnMap["Prior Authorization Number"]] || ""; // ✅ NEW LINE
            const accountNumber = accountNumberRaw.split("-")[0].trim();
            console.log(`Processing account: ${accountNumber}`);
        
            // Validate data before processing
            if (!provider) console.log('Provider value is missing or undefined');
            if (!facilityName) console.log('Facility is missing or undefined');
            if (!billingDate) console.log('DOS is missing or undefined');
            if (!admitDate) console.log('Hospital Date is missing or undefined'); // ✅ Optional

          // Array to hold valid diagnosis codes
const diagnosisCodes = [];

// Regex:
// - ICD-10-CM: Starts with a letter, 2–4 digits, optional dot, followed by 1–4 alphanumerics
// - ICD-10-PCS: Exactly 7 alphanumeric characters (A-Z, 0-9)
const dxCodeRegex = /\b(?:[A-Z][0-9]{2,4}(?:\.[0-9A-Z]{1,4})?|[A-Z0-9]{7})\b/g;

// Match codes
const matches = diagnosisText.match(dxCodeRegex);

if (matches) {
  // Normalize: trim, uppercase, remove duplicates
  const uniqueCodes = [...new Set(matches.map(code => code.trim().toUpperCase()))];

  // Filter: Must contain at least one letter and one digit
  const validCodes = uniqueCodes.filter(code => /[A-Z]/i.test(code) && /\d/.test(code));

  diagnosisCodes.push(...validCodes);

  // Output
  console.log("✅ Final Unique, Valid Diagnosis Codes:");
  diagnosisCodes.forEach((code, index) => {
    console.log(`  ${index + 1}. ${code}`);
  });
  console.log("🔢 Total:", diagnosisCodes.length);
} else {
  console.log("⚠️ No valid diagnosis codes found.");
}
          
          
// ✅ Normalize and extract CPT codes and units
const cleanedText = chargesText.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();

const cptCodes = [];
const cptUnits = [];

const cptRegex = /\b(\d{4,5}[A-Z]?)(?:\s*\*\s*(\d+))?\b/gi;
let match;

while ((match = cptRegex.exec(cleanedText)) !== null) {
  const code = match[1]; // CPT code
  const units = match[2] ? match[2].trim() : null;

  cptCodes.push(code);
  cptUnits.push(units);
  console.log(`🧩 CPT Code: ${code}, Units: ${units ?? 'SKIP'}`);
}

const uniqueCptCodes = [...new Set(cptCodes)];
console.log('✅ Extracted Unique CPT Codes:', uniqueCptCodes);



const { success, visitId, zeroBalanceCPTs, facilityNotFound, cptNotFoundList ,diagnosisNotFoundList } = await processAccount(
    newPage,
    newPage.browser(), // Add browser parameter
    accountNumber,
    provider,
    facilityName,
    billingDate,
    diagnosisCodes,
    cptCodes, // ✅ only the array
    admitDate,
    pos,
    modifier,
    priornumber,
    cptUnits
);

if (success) {
    results.successful.push(accountNumber);
    row[columnMap["Result"]] = "Done";
    row[columnMap["Visit ID"]] = visitId ? visitId.toString() : "";
    
    let billedFeeMessages = [];
    
    if (facilityNotFound) {
        billedFeeMessages.push("Facility Not Found");
    }
    
    // CPTs with zero balance (0.00 or 0.01) - write "Billed Fee is Zero" for each CPT
    if (Array.isArray(zeroBalanceCPTs) && zeroBalanceCPTs.length > 0) {
        const zeroCPTs = zeroBalanceCPTs
            .filter(f => {
                const balance = parseFloat(f.Balance);
                return balance === 0.00 || balance === 0.01;
            })
            .map(f => `${f.CPTCode} Billed Fee is Zero`);
        
        if (zeroCPTs.length > 0) {
            billedFeeMessages.push(...zeroCPTs);
        }
    }
    
    // CPTs not found even if Visit ID is present
    if (Array.isArray(cptNotFoundList) && cptNotFoundList.length > 0) {
        billedFeeMessages.push(`${cptNotFoundList.join(', ')} CPT not found`);
    }
    
    if (diagnosisNotFoundList.length > 0) {
        billedFeeMessages.push(`Dx Code not found: ${diagnosisNotFoundList.join(', ')}`);
    }
    
    // Join with newlines for separate lines in Excel
    row[columnMap["Billed Fee"]] = billedFeeMessages.join('\n');
} else {
    results.failed.push(accountNumber);
    row[columnMap["Result"]] = "Failed";
    row[columnMap["Visit ID"]] = visitId ? visitId.toString() : "";
    
    let billedFeeMessages = [];
    
    if (facilityNotFound) {
        billedFeeMessages.push("Facility Not Found");
    }
    
    // CPTs with zero balance in failure case (0.00 or 0.01) - write "Billed Fee is Zero" for each CPT
    if (Array.isArray(zeroBalanceCPTs) && zeroBalanceCPTs.length > 0) {
        const zeroCPTs = zeroBalanceCPTs
            .filter(f => {
                const balance = parseFloat(f.Balance);
                return balance === 0.00 || balance === 0.01;
            })
            .map(f => `${f.CPTCode} Billed Fee is Zero`);
        
        if (zeroCPTs.length > 0) {
            billedFeeMessages.push(...zeroCPTs);
        }
    }
    
    // CPTs not found in failure
    if (Array.isArray(cptNotFoundList) && cptNotFoundList.length > 0) {
        billedFeeMessages.push(`${cptNotFoundList.join(', ')} CPT not found`);
    }
    
    if (diagnosisNotFoundList.length > 0) {
        billedFeeMessages.push(`Dx Code not found: ${diagnosisNotFoundList.join(', ')}`);
    }
    
    // Join with newlines for separate lines in Excel
    row[columnMap["Billed Fee"]] = billedFeeMessages.join('\n');
}
        }            
        // Convert data back to an ExcelJS workbook
       const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Sheet1");

// Write headers with gray background
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

// Define column widths (adjusted to include Visit ID column)
const columnWidths = [6, 11, 30, 15, 15, 23, 54, 10, 10, 11, 50, 20, 30, 15 ,10]; 
// Added one more width for Visit ID column (15)

sheet.columns = columnWidths.map((width, index) => ({
    width,
    style: (index === columnMap["Diagnoses"] || index === columnMap["CPT"])
        ? { alignment: { wrapText: true } }
        : {},
}));

originalData.slice(1).forEach((row) => {
    if (!row) return;

    const hasData = row.some(cell => cell && cell.toString().trim() !== "");
    if (!hasData) return;

    const rowData = headers.map(header => row[columnMap[header]] || "");
    const newRow = sheet.addRow(rowData);
    newRow.height = 100;

    headers.forEach((header, colIndex) => {
        const cell = newRow.getCell(colIndex + 1);

        // Apply border
        cell.border = {
            top: { style: "thin", color: { argb: "000000" } },
            left: { style: "thin", color: { argb: "000000" } },
            bottom: { style: "thin", color: { argb: "000000" } },
            right: { style: "thin", color: { argb: "000000" } }
        };

        cell.font = { size: 10 };

        // Format Patient ID
        if (header === "Patient ID" && cell.value) {
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFFFFF" }
            };
            cell.alignment = { horizontal: "center", vertical: "bottom" };
        }

        // Format Visit ID
        if (header === "Visit ID" && cell.value) {
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "E6F3FF" } // Light blue background
            };
            cell.alignment = { horizontal: "center", vertical: "middle" };
        }
        if (header === "Billed Fee" && cell.value) {
                    // Set the column width to 20
                    sheet.getColumn(columnMap["Billed Fee"] + 1).width = 20;
                
                    // Set text color to red
                    cell.font = {
                        color: { argb: "FF0000" }, // Red text
                    };
                    cell.font = { size: 10 }

                    // Align the text
                    cell.alignment = { horizontal: "center", vertical: "bottom",wrapText: true };
                }
        // Format Result
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
            await new Promise(resolve => setTimeout(resolve, 2000));
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

app.listen(port, () => console.log(`Server running on http://localhost:${port}`));   //complete whittier code 