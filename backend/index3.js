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
  // Complete code block to replace the problematic section
   try {
      // Wait and click Add New Visit
// Wait and click Add New Visit
await new Promise(resolve => setTimeout(resolve, 3000));
console.log('⏳ Waiting for Add New Visit button to appear...');
await page.waitForSelector('#addNewVisit', { visible: true, timeout: 16000 });

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
          newPage = await Promise.race([
              popupPromise,
              new Promise((_, reject) => setTimeout(() => reject(new Error('Popup timeout')), 7000))
          ]);
          console.log('✅ Popup loaded successfully');
      } catch (error) {
          console.log('⚠️ No popup detected, checking for new tabs...');
          
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

      // Wait for page to stabilize
      await new Promise(resolve => setTimeout(resolve, 2000));

      // Wait for Search button to appear
      console.log('⏳ Waiting for Search button to appear...');
      const searchButtonSelectors = [
          '#ctl00_phFolderContent_Button1',
          'input[value="Search"]',
          'button[onclick*="Search"]',
          '.button:contains("Search")',
          '#Button1'
      ];

      let searchButtonFound = false;
      let workingSelector = null;

      // Try each selector sequentially for better reliability
      for (const selector of searchButtonSelectors) {
          try {
              await newPage.waitForSelector(selector, { visible: true, timeout: 3000 });
              workingSelector = selector;
              searchButtonFound = true;
              console.log(`✅ Found search button with selector: ${selector}`);
              break;
          } catch (error) {
              continue;
          }
      }

      if (!searchButtonFound) {
          throw new Error('Search button not found with any selector');
      }

      // Click initial popup search button
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

      // Wait for dropdown to appear with better error handling
      console.log('⏳ Waiting for dropdown to appear...');
      const dropdownSelectors = [
          '#ctl04_popupBase_ddlSearch',
          'select[name*="ddlSearch"]',
          'select[id*="Search"]',
          'select:first-of-type'
      ];

      let dropdownSelector = null;
      let dropdownFound = false;

      // Try each dropdown selector sequentially
      for (const selector of dropdownSelectors) {
          try {
              await newPage.waitForSelector(selector, { visible: true, timeout: 4000 });
              dropdownSelector = selector;
              dropdownFound = true;
              console.log(`✅ Found dropdown with selector: ${selector}`);
              break;
          } catch (error) {
              continue;
          }
      }

      if (!dropdownFound) {
          console.log('❌ Dropdown not found with any selector');
          throw new Error('Dropdown not found');
      }

      // Wait for dropdown to be fully loaded
      await new Promise(resolve => setTimeout(resolve, 1000));

      // IMPROVED: Select "Patient ID" with better validation and error handling
console.log('🔽 Selecting "Patient ID" from dropdown...');

let patientIdSelected = false;
let attempts = 0;
const maxAttempts = 3;

while (!patientIdSelected && attempts < maxAttempts) {
    attempts++;
    console.log(`Attempt ${attempts} to select Patient ID...`);

    try {
        // Wait for dropdown to be fully loaded with options
        await newPage.waitForFunction((selector) => {
            const dropdown = document.querySelector(selector);
            return dropdown && dropdown.options && dropdown.options.length > 1;
        }, { timeout: 5000 }, dropdownSelector);

        const selectionResult = await newPage.evaluate((selector) => {
            const dropdown = document.querySelector(selector);
            
            if (!dropdown || !dropdown.options || dropdown.options.length === 0) {
                return { success: false, error: 'Dropdown not found or empty' };
            }

            // Log all available options for debugging
            const optionsInfo = Array.from(dropdown.options).map((option, index) => ({
                index,
                text: option.text.trim(),
                value: option.value.trim()
            }));
            
            console.log('Available dropdown options:', optionsInfo);

            // More specific patterns for Patient ID (ordered by specificity)
            const searchPatterns = [
                /^patient\s*id$/i,           // Exact match: "Patient ID"
                /^patient\s*identifier$/i,    // Alternative: "Patient Identifier"  
                /^patient\s*number$/i,        // Alternative: "Patient Number"
                /patient.*id/i,               // Contains both "patient" and "id"
                /^patient$/i                  // Just "Patient" as fallback
            ];

            let patientIdIndex = -1;
            let matchedPattern = null;
            let matchedText = '';

            // Try each pattern in order of specificity
            for (const pattern of searchPatterns) {
                patientIdIndex = Array.from(dropdown.options).findIndex(option => {
                    const text = option.text.trim();
                    const value = option.value.trim();
                    return pattern.test(text) || pattern.test(value);
                });
                
                if (patientIdIndex !== -1) {
                    matchedPattern = pattern;
                    matchedText = dropdown.options[patientIdIndex].text.trim();
                    console.log(`Found match with pattern ${pattern}: "${matchedText}" at index ${patientIdIndex}`);
                    break;
                }
            }

            // If no pattern match, don't use fallback - it's too risky
            if (patientIdIndex === -1) {
                return {
                    success: false,
                    error: 'Patient ID option not found with any pattern',
                    availableOptions: optionsInfo
                };
            }

            // Validate that we found a reasonable match
            const selectedOptionText = dropdown.options[patientIdIndex].text.trim().toLowerCase();
            if (!selectedOptionText.includes('patient')) {
                return {
                    success: false,
                    error: `Selected option "${matchedText}" doesn't contain 'patient' - likely incorrect match`,
                    availableOptions: optionsInfo
                };
            }

            // Select the option
            const previousSelection = dropdown.selectedIndex;
            dropdown.selectedIndex = patientIdIndex;
            
            // Trigger multiple events to ensure selection is registered
            dropdown.dispatchEvent(new Event('change', { bubbles: true }));
            dropdown.dispatchEvent(new Event('input', { bubbles: true }));
            dropdown.dispatchEvent(new Event('blur', { bubbles: true }));
            
            const selectedOption = dropdown.options[patientIdIndex];
            
            // Verify selection actually changed
            if (dropdown.selectedIndex !== patientIdIndex) {
                return {
                    success: false,
                    error: `Selection failed to stick. Previous: ${previousSelection}, Attempted: ${patientIdIndex}, Current: ${dropdown.selectedIndex}`
                };
            }
            
            return {
                success: true,
                selectedIndex: patientIdIndex,
                selectedText: selectedOption.text.trim(),
                selectedValue: selectedOption.value.trim(),
                matchedPattern: matchedPattern?.toString()
            };
        }, dropdownSelector);

        if (selectionResult.success) {
            console.log(`✅ Selected Patient ID: "${selectionResult.selectedText}" at index ${selectionResult.selectedIndex}`);
            
            // Double-check the selection persisted
            await new Promise(resolve => setTimeout(resolve, 500));
            
            const verificationResult = await newPage.evaluate((selector, expectedIndex) => {
                const dropdown = document.querySelector(selector);
                if (!dropdown) return { verified: false, error: 'Dropdown disappeared' };
                
                return {
                    verified: dropdown.selectedIndex === expectedIndex,
                    currentIndex: dropdown.selectedIndex,
                    currentText: dropdown.options[dropdown.selectedIndex]?.text?.trim() || 'N/A'
                };
            }, dropdownSelector, selectionResult.selectedIndex);
            
            if (verificationResult.verified) {
                console.log(`✅ Selection verified: "${verificationResult.currentText}"`);
                patientIdSelected = true;
            } else {
                console.log(`❌ Selection verification failed. Expected index ${selectionResult.selectedIndex}, got ${verificationResult.currentIndex}`);
            }
        } else {
            console.log(`❌ Failed to select Patient ID: ${selectionResult.error}`);
            if (selectionResult.availableOptions) {
                console.log('Available options:', selectionResult.availableOptions);
            }
        }
        
    } catch (error) {
        console.log(`❌ Error selecting Patient ID (attempt ${attempts}):`, error.message);
    }

    if (!patientIdSelected && attempts < maxAttempts) {
        console.log('⏳ Waiting before retry...');
        await new Promise(resolve => setTimeout(resolve, 1500));
    }
}

if (!patientIdSelected) {
    // Log final state for debugging
    const finalState = await newPage.evaluate((selector) => {
        const dropdown = document.querySelector(selector);
        if (!dropdown) return { error: 'Dropdown not found' };
        
        return {
            currentIndex: dropdown.selectedIndex,
            currentText: dropdown.options[dropdown.selectedIndex]?.text?.trim() || 'N/A',
            totalOptions: dropdown.options.length,
            allOptions: Array.from(dropdown.options).map((opt, idx) => ({
                index: idx,
                text: opt.text.trim(),
                value: opt.value.trim()
            }))
        };
    }, dropdownSelector);
    
    console.log('Final dropdown state:', finalState);
    throw new Error('Failed to select Patient ID after multiple attempts');
}

// Wait for selection to take effect
await new Promise(resolve => setTimeout(resolve, 1000));

  try {
      // Enhanced text field handling with validation
      console.log(`⌨️ Entering Patient ID: ${accountNumber}`);
      const textFieldSelectors = [
          '#ctl04_popupBase_txtSearch',
          'input[name*="txtSearch"]',
          'input[type="text"]:not([style*="display: none"])',
          'input[maxlength="50"]'
      ];

      let textFieldSelector = null;
      let textFieldFound = false;

      // Try each text field selector
      for (const selector of textFieldSelectors) {
          try {
              await newPage.waitForSelector(selector, { visible: true, timeout: 6000 });
              textFieldSelector = selector;
              textFieldFound = true;
              console.log(`✅ Found text field with selector: ${selector}`);
              break;
          } catch (error) {
              continue;
          }
      }

      if (!textFieldFound) {
          console.log('❌ Text field not found with any selector');
          throw new Error('Text field not found');
      }

      // Clear and enter text with validation
      await newPage.evaluate((selector) => {
          const textField = document.querySelector(selector);
          if (textField) {
              textField.value = '';
              textField.focus();
          }
      }, textFieldSelector);

      await newPage.type(textFieldSelector, accountNumber, { delay: 50 });

      // Validate text was entered correctly
      const textValidation = await newPage.$eval(textFieldSelector, el => el.value);
      if (textValidation !== accountNumber) {
          console.log(`⚠️ Text validation failed. Expected: "${accountNumber}", Got: "${textValidation}"`);
          // Retry typing
          await newPage.evaluate((selector) => {
              document.querySelector(selector).value = '';
          }, textFieldSelector);
          await newPage.type(textFieldSelector, accountNumber, { delay: 100 });
      }

      // Click the search button with improved reliability
      console.log('🖱️ Clicking search button...');
      const searchBtnSelectors = [
          '#ctl04_popupBase_btnSearch',
          'input[name*="btnSearch"]',
          'input[value="Search"]',
          'button[onclick*="Search"]'
      ];

      let searchBtnClicked = false;

      for (const selector of searchBtnSelectors) {
          try {
              await newPage.waitForSelector(selector, { visible: true, timeout: 6000 });
              await newPage.click(selector);
              searchBtnClicked = true;
              console.log(`✅ Clicked search button with selector: ${selector}`);
              break;
          } catch (error) {
              continue;
          }
      }

      if (!searchBtnClicked) {
          // Fallback click method
          const fallbackResult = await newPage.evaluate(() => {
              const searchBtn = document.querySelector('#ctl04_popupBase_btnSearch') ||
                              document.querySelector('input[name*="btnSearch"]') ||
                              document.querySelector('input[value="Search"]');
              if (searchBtn) {
                  searchBtn.click();
                  return true;
              }
              return false;
          });
          
          if (!fallbackResult) {
              throw new Error('Failed to click search button');
          }
      }

      // Wait for search results
      await new Promise(resolve => setTimeout(resolve, 2000));
      console.log('✅ Search completed.');

      // Wait for search results with improved selector handling
      console.log('⏳ Waiting for search result row selector to appear...');
      const resultSelectors = [
          '#ctl04_popupBase_grvPopup_ctl02_lnkSelect',
          'a[id*="lnkSelect"]',
          'a[href*="Select"]',
          'input[value="Select"]'
      ];

      let resultSelector = null;
      let resultFound = false;

      for (const selector of resultSelectors) {
          try {
              await newPage.waitForSelector(selector, { visible: true, timeout: 5000 });
              resultSelector = selector;
              resultFound = true;
              console.log(`✅ Found result selector: ${selector}`);
              break;
          } catch (error) {
              continue;
          }
      }

      if (resultFound && resultSelector) {
          // Click the first row's "Select" link
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

      await new Promise(resolve => setTimeout(resolve, 1500));

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
                // Trigger change event
                field.dispatchEvent(new Event('change', { bubbles: true }));
                field.dispatchEvent(new Event('input', { bubbles: true }));
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

// Always extract the first part as Last Name (ignoring MD, suffixes, commas, etc.)
let providerSearchTerm = "";

if (provider && provider.trim()) {
  // Split by space or comma
  let parts = provider.split(/[\s,]+/).filter(Boolean);

  // Take the very first word that is not "MD"
  if (parts.length > 0) {
    providerSearchTerm = parts[0].toUpperCase() === "MD" ? (parts[1] || "") : parts[0];
  }
}

console.log(`⌨️ Will search for Provider (Last Name): ${providerSearchTerm}`);



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
    // Split by newlines for multiple providers
    const providers = provider.split(/\n/);

    providers.forEach(providerEntry => {
        const trimmedProvider = providerEntry.trim();
        if (trimmedProvider) {
            let firstName = '';

            if (trimmedProvider.includes(',')) {
                // Always take part AFTER the last comma
                const afterComma = trimmedProvider.split(',').pop().trim();
                // First name is the first word after the last comma
                firstName = afterComma.split(' ')[0].trim();
            } else {
                // If no comma, just take the first word
                firstName = trimmedProvider.split(' ')[0].trim();
            }

            if (firstName) {
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
let popupMessages = [];

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
    const alertText = await page.$eval('.ui-dialog-content .alertClass', el => el.innerText.trim());
    
    // Remove the header text and clean up messages
    const cleanedText = alertText.replace(/Please correct the following\(s\):\s*/i, '');
    popupMessages = cleanedText.split('\n').map(line => line.trim()).filter(Boolean);
    console.log("📋 Popup Messages:", popupMessages);

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
return { success, visitId, zeroBalanceCPTs, facilityNotFound: facilityNotFoundFlag, cptNotFoundList ,diagnosisNotFoundList ,popupMessages };

} catch (error) {
    console.error('Error processing account:', error);
    return { success: false, error: error.message };
}
}
// Function to generate Excel file (moved outside main endpoint)
async function generateExcelFile(originalData, headers, columnMap) {
    try {
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
    const columnWidths = [11, 30, 15, 15, 23, 54, 10, 10, 11, 50, 20, 30, 15 ,10]; 
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
    
        return workbook;
    } catch (error) {
        console.error("Error generating Excel file:", error);
        throw error;
    }
}

// Updated main processing endpoint
app.post("/process", upload.single("file"), async (req, res) => {
    if (!req.file || !req.body.originalPath) {
        return res.status(400).json({ error: "No file or original path provided" });
    }
    
    let newPage = null;
    let originalData = [];
    let headers = [];
    let columnMap = {};
    let results = { successful: [], failed: [] };
    let processingError = null;
    
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
        originalData = xlsx.utils.sheet_to_json(sheetXLSX, {
            header: 1,
            raw: false,
            dateNF: 'mm/dd/yyyy'
        });

        if (!originalData.length) {
            throw new Error("Empty file or missing headers.");
        }
        
        headers = originalData[0].map(h => h?.toString().trim());
        
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
            // Mark this as an error but don't throw - we still want to generate Excel
            processingError = "Diagnoses not found in Excel.";
            
            // Mark all unprocessed rows as failed
            for (let i = 1; i < originalData.length; i++) {
                const row = originalData[i];
                if (!row) continue;
                
                const accountNumberRaw = row[columnMap["Patient ID"]] || "";
                const status = (row[columnMap["Result"]] || "").toString().trim().toLowerCase();
                
                if (accountNumberRaw && status !== "done" && status !== "failed") {
                    const accountNumber = accountNumberRaw.split("-")[0].trim();
                    results.failed.push(accountNumber);
                    row[columnMap["Result"]] = "Failed";
                    row[columnMap["Billed Fee"]] = "Diagnoses not found";
                }
            }
        } else if (newRowsToProcess === 0) {
            // All rows already processed - this is success, not error
            processingError = null;
        } else {
            // Process normally
            newPage = await continueWithLoggedInSession();
            
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
                const admitDate = row[columnMap["Hospital Date"]] || "";
                const pos = row[columnMap["POS"]] || "";
                const modifier = row[columnMap["Modifier"]] || "";
                const priornumber = row[columnMap["Prior Authorization Number"]] || "";
                const accountNumber = accountNumberRaw.split("-")[0].trim();
                console.log(`Processing account: ${accountNumber}`);
            
                // Validate data before processing
                if (!provider) console.log('Provider value is missing or undefined');
                if (!facilityName) console.log('Facility is missing or undefined');
                if (!billingDate) console.log('DOS is missing or undefined');
                if (!admitDate) console.log('Hospital Date is missing or undefined');

                // Array to hold valid diagnosis codes
                const diagnosisCodes = [];

                // Regex for ICD codes
                // ICD-like diagnosis codes, 2–8 chars, with or without dots, with or without suffix
                const dxCodeRegex = /\b[A-Z][A-Z0-9\.]{1,7}\b/gi;



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
                
                // Normalize and extract CPT codes and units
                const cleanedText = chargesText.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();

                const cptCodes = [];
                const cptUnits = [];
                const seenCodes = new Set(); // Track codes we've already seen

                const cptRegex = /\b(\d{4,5}[A-Z]?)(?:\s*\*\s*(\d+))?\b/gi;
                let match;

                while ((match = cptRegex.exec(cleanedText)) !== null) {
                  const code = match[1]; // CPT code
                  const units = match[2] ? match[2].trim() : null;

                  // Only add if we haven't seen this code before
                  if (!seenCodes.has(code)) {
                    seenCodes.add(code);
                    cptCodes.push(code);
                    cptUnits.push(units);
                    console.log(`🧩 CPT Code: ${code}, Units: ${units ?? 'SKIP'}`);
                  } else {
                    console.log(`⚠️ Duplicate CPT Code skipped: ${code}`);
                  }
                }

                const uniqueCptCodes = cptCodes;
                console.log('✅ Extracted Unique CPT Codes:', uniqueCptCodes);
                console.log('✅ Corresponding Units:', cptUnits);
                console.log(`📊 Total unique CPT codes found: ${uniqueCptCodes.length}`);

                const { success, visitId, zeroBalanceCPTs, facilityNotFound, cptNotFoundList, diagnosisNotFoundList,popupMessages } = await processAccount(
                    newPage,
                    newPage.browser(),
                    accountNumber,
                    provider,
                    facilityName,
                    billingDate,
                    diagnosisCodes,
                    cptCodes,
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

  if (Array.isArray(cptNotFoundList) && cptNotFoundList.length > 0) {
    billedFeeMessages.push(`${cptNotFoundList.join(', ')} CPT not found`);
  }

  if (diagnosisNotFoundList && diagnosisNotFoundList.length > 0) {
    billedFeeMessages.push(`Dx Code not found: ${diagnosisNotFoundList.join(', ')}`);
  }

  row[columnMap["Billed Fee"]] = billedFeeMessages.join('\n');

} else {
  results.failed.push(accountNumber);
  row[columnMap["Result"]] = "Failed";
  row[columnMap["Visit ID"]] = visitId ? visitId.toString() : "";

  let billedFeeMessages = [];

  // 🔍 Handle popup error messages
// 🔍 Handle popup error messages
if (Array.isArray(popupMessages)) {
  popupMessages.forEach(msg => {
    if (msg.includes("Patient ID is required")) {
      billedFeeMessages.push("Patient Not Found");
    } else if (msg.includes("Provider ID is required")) {
      billedFeeMessages.push("Provider Not Found");
    } else if (msg.toLowerCase().includes("diagnosis code pointer") || msg.toLowerCase().includes("line item")) {
      billedFeeMessages.push("ICD 10 Pointer is missing");
    } else {
      billedFeeMessages.push(msg); // generic fallback
    }
  });
}
  if (facilityNotFound) {
    billedFeeMessages.push("Facility Not Found");
  }

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

  if (Array.isArray(cptNotFoundList) && cptNotFoundList.length > 0) {
    billedFeeMessages.push(`${cptNotFoundList.join(', ')} CPT not found`);
  }

  if (diagnosisNotFoundList && diagnosisNotFoundList.length > 0) {
    billedFeeMessages.push(`Dx Code not found: ${diagnosisNotFoundList.join(', ')}`);
  }

  row[columnMap["Billed Fee"]] = billedFeeMessages.join('\n');
}

            }
        }

    } catch (error) {
        console.error("Error during processing:", error);
        processingError = error.message;
        
        // Ensure results arrays exist
        if (!results.successful) results.successful = [];
        if (!results.failed) results.failed = [];
    } finally {
        // ALWAYS generate Excel file if we have data, regardless of success/failure
        let fileId = null;
        let excelError = null;
        
        try {
            if (originalData.length > 0 && headers.length > 0) {
                const workbook = await generateExcelFile(originalData, headers, columnMap);
                
                // Generate a unique file ID
                fileId = Date.now().toString();

                // Create a buffer of the Excel file and store it
                const fileBuffer = await workbook.xlsx.writeBuffer();
                processedFiles.set(fileId, {
                    buffer: fileBuffer,
                    filename: "updated_file.xlsx"
                });
                
                console.log("✅ Excel file generated successfully");
            } else {
                console.log("⚠️ No data available for Excel generation");
            }
        } catch (error) {
            console.error("❌ Error generating Excel file:", error);
            excelError = error.message;
        }
        
        // Clean up browser
        if (newPage) {
            try {
                await new Promise(resolve => setTimeout(resolve, 2000));
                await newPage.close();
                console.log("Browser closed");
            } catch (closeError) {
                console.error("Error closing browser:", closeError);
            }
        }
        
        // Send response
       // Send response
const successfulCount = (results.successful || []).length;
const failedCount = (results.failed || []).length;

if (processingError === "Diagnoses not found in Excel.") {
    // Special case for missing diagnoses
    res.json({
        success: false,
        message: processingError,
        results: results,
        fileId: fileId
    });
} else if (processingError && processingError.includes("Column is not defined")) {
    // Handle column definition errors specifically
    res.status(400).json({
        success: false,
        error: processingError,
        results: results,
        fileId: fileId
    });
} else if (processingError && !fileId) {
    // Error occurred and no Excel file generated
    res.status(500).json({
        success: false,
        error: processingError,
        excelError: excelError,
        results: results
    });
} else if (successfulCount === 0 && failedCount === 0) {
    // All rows already processed
    res.json({
        success: true,
        message: "All rows were already updated.",
        results: results,
        fileId: fileId
    });
} else {
    // Normal case - processing completed (with or without some failures)
    res.json({
        success: !processingError,
        message: `${successfulCount} Charges Entered successfully. ${failedCount} rows failed.`,
        results: results,
        fileId: fileId,
        ...(processingError && { processingError })
    });
}
    }
});

// Keep your existing download endpoint unchanged
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
    
    // Remove the file from memory after some time
    setTimeout(() => {
        processedFiles.delete(fileId);
        console.log(`File ${fileId} removed from memory`);
    }, 3 * 60 * 60 * 1000); // Remove after 3 hours
});

app.listen(port, () => console.log(`Server running on http://localhost:${port}`));