const GEMINI_API_KEY = '####'; // Replace with your actual key
const FOLDER_ID = '####';  //Replace with your folder id
const GEMINI_MODEL = 'gemini-3.8-flash';    // The model we are using!

// 1. Build the Custom Menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Vintage Tools')
      .addItem('1. Process Photos (Manual)', 'runManual')
      .addItem('1. Process Photos (Auto-Restart)', 'runAuto')
      .addSeparator()
      .addItem('2. Generate eBay Drafts Export', 'createEbayExport')
      .addToUi();
}

function runManual() {
  PropertiesService.getScriptProperties().setProperty('AUTO_RESTART', 'false');
  processStampsBatches();
}

function runAuto() {
  PropertiesService.getScriptProperties().setProperty('AUTO_RESTART', 'true');
  SpreadsheetApp.getActiveSpreadsheet().toast("Auto-restart enabled! The script will quietly restart itself every 5.5 minutes until the folder is empty.", "Listings"); // replace Listings with the name of the tab you want it to fill
  processStampsBatches();
}

// Cleans up background triggers when the job is done
function clearTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processStampsBatches') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// 2. Main Processing Function
function processStampsBatches() {
  const startTime = Date.now();
  const MAX_RUN_TIME = 5.5 * 60 * 1000; // 5.5 minutes

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("11-10"); 
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  // Create or grab the "Processed" folder
  let processedFolder;
  const folders = folder.getFoldersByName("Processed");
  if (folders.hasNext()) {
    processedFolder = folders.next();
  } else {
    processedFolder = folder.createFolder("Processed");
  }

  const filesIterator = folder.getFiles();
  let filesArray = [];

  while (filesIterator.hasNext()) {
    filesArray.push(filesIterator.next());
  }
  filesArray.sort((a, b) => a.getName().localeCompare(b.getName()));

  // If the folder is empty, we are completely done! Clean up triggers.
  if (filesArray.length === 0) {
    clearTriggers();
    PropertiesService.getScriptProperties().deleteProperty('AUTO_RESTART');
    return;
  }

  let row = sheet.getLastRow() + 1;
  if (row < 2) row = 2;

  let currentItemImages = [];
  let currentItemFileNames = [];
  let currentItemFiles = []; 

  for (let i = 0; i < filesArray.length; i++) {
    
    // SAFETY CLOCK: Auto-Restart Logic
    if (Date.now() - startTime > MAX_RUN_TIME) {
      const autoRestart = PropertiesService.getScriptProperties().getProperty('AUTO_RESTART');
      if (autoRestart === 'true') {
        Logger.log("Time limit reached. Scheduling auto-restart.");
        ScriptApp.newTrigger('processStampsBatches').timeBased().after(1000).create();
        return; 
      } else {
        Logger.log("Time limit reached. Stopping manually.");
        return; 
      }
    }

    const file = filesArray[i];
    const base64Image = Utilities.base64Encode(file.getBlob().getBytes());
    const mimeType = file.getMimeType();

    const isScreenshot = checkIfScreenshot(base64Image, mimeType);

    if (isScreenshot) {
      if (currentItemImages.length > 0) {
        const success = generateListingForBatch(currentItemImages, currentItemFileNames, sheet, row);
        
        if (success) {
          row++;
          for (let f of currentItemFiles) {
            f.moveTo(processedFolder);
          }
          file.moveTo(processedFolder); 
        }
        
        currentItemImages = [];
        currentItemFileNames = [];
        currentItemFiles = [];
      } else {
         file.moveTo(processedFolder);
      }
    } else {
      currentItemImages.push({
        inlineData: { mimeType: mimeType, data: base64Image }
      });
      currentItemFileNames.push(file.getName());
      currentItemFiles.push(file);
    }
  }

  if (currentItemImages.length > 0) {
    const success = generateListingForBatch(currentItemImages, currentItemFileNames, sheet, row);
    if (success) {
      for (let f of currentItemFiles) {
        f.moveTo(processedFolder);
      }
    }
  }
  
  // Clean up triggers once loop finishes
  clearTriggers();
  PropertiesService.getScriptProperties().deleteProperty('AUTO_RESTART');
}

// 3. Helper: Screenshot Detection
function checkIfScreenshot(base64Image, mimeType) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  
  const payload = {
    contents: [{
      parts: [
        { text: "Is this image a screenshot of a smartphone home screen with app icons? Reply strictly with the word TRUE or FALSE." },
        { inlineData: { mimeType: mimeType, data: base64Image } }
      ]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.error) return false;
    const textResponse = json.candidates[0].content.parts[0].text.trim().toUpperCase();
    return textResponse.includes("TRUE");
  } catch (e) {
    return false;
  }
}

// 4. Helper: AI Listing Generator
function generateListingForBatch(imageParts, fileNames, sheet, row) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  
  const promptText = `
    You are "The Vintage Listing & Business Assistant" for VintageReveries, a brand specializing in high-quality vintage (1930s–1980s), stamps, postcards, magazines, books and ephemera, mid-century objects, and select modern items.
    I am providing you with multiple photos of the SAME single item (front, back, closeups). Analyze all images together to create a single comprehensive listing.
    
    CORE RESPONSIBILITIES:
    1. Analyze Visuals: Identify era, materials (giving confidence levels), flaws, and features. Provide historical context.
    2. Calculate Measurements: Name any measurements specified. Do the math to provide doubled measurements where appropriate.
    3. Tone & Style: Expert but accessible. Write the WooCommerce description in the first person. Quirky tone. Tell the story. Do not use emojis. Do not use the pound # symbol in hashtags.
    
    CRITICAL FORMATTING RULES: 
    1. Output plain text ONLY. Do NOT use HTML tags. Use standard keyboard line breaks (new lines) for paragraph spacing and dashes (-) for bullet points.
    2. STRICT TITLE LIMITS: Titles MUST be 80 characters or less. eBay will reject the listing if it is 81 characters. Count the characters carefully for the WooCommerce SEO Title and the Poshmark/eBay Title.
    3. Use single quotes (') instead of double quotes (") inside your descriptions to keep the JSON clean.
    
    Output the data based on this exact structure:
    - era_identification: Era, Material, Handling
    - condition_notes: Specific flaws or 'Excellent Vintage Condition'. (This must be clean text for the user to copy/paste directly).
    - pricing: Suggested List Price, Minimum Acceptable Offer
    - woocommerce: SEO Title (STRICTLY < 80 chars), Short Description, Long Description (At least 800-1600 words), Measurements.
    - poshmark_ebay: Title (STRICTLY < 80 chars), Description (<1500 chars), Tags, eBay Category
    - depop: Description (<1000 characters, casual and trendy tone for vintage lovers), Hashtags (no # symbols)
    - rationale: Your reasoning for the suggested pricing based on market/condition.
    - power_seller_tip: One quick tip for selling this specific item.
    - lot_grouping: Suggest a thematic lot to bundle this in. If high value, state 'Sell Individually'.
  `;

  const partsArray = [{ text: promptText }];
  
  const maxImages = 4;
  let imagesToAnalyze = [];

  if (imageParts.length <= maxImages) {
    imagesToAnalyze = imageParts;
  } else {
    const step = (imageParts.length - 1) / (maxImages - 1);
    for (let i = 0; i < maxImages; i++) {
      const index = Math.round(i * step);
      imagesToAnalyze.push(imageParts[index]);
    }
  }
  
  for (let i = 0; i < imagesToAnalyze.length; i++) {
    partsArray.push(imagesToAnalyze[i]);
  }

  const payload = {
    contents: [{ parts: partsArray }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: {
        type: "OBJECT",
        properties: {
          era_identification: { type: "STRING" },
          condition_notes: { type: "STRING" },
          pricing: { type: "STRING" },
          woocommerce: { type: "STRING" },
          poshmark_ebay: { type: "STRING" },
          depop: { type: "STRING" },
          rationale: { type: "STRING" },
          power_seller_tip: { type: "STRING" },
          lot_grouping: { type: "STRING" }
        },
        required: [
          "era_identification", "condition_notes", "pricing", "woocommerce", 
          "poshmark_ebay", "depop", "rationale", "power_seller_tip", "lot_grouping"
        ]
      }
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) return false; 
    
    const textResponse = json.candidates[0].content.parts[0].text;
    const geminiData = JSON.parse(textResponse);
    const joinedNames = fileNames.join(", ");
    
    sheet.getRange(row, 1, 1, 10).setValues([[
      joinedNames, geminiData.era_identification, geminiData.condition_notes,
      geminiData.pricing, geminiData.woocommerce, geminiData.poshmark_ebay,
      geminiData.depop, geminiData.rationale, geminiData.power_seller_tip, geminiData.lot_grouping
    ]]);
    return true; 
  } catch (e) {
    return false;
  }
}

// 5. eBay Draft Generator
function createEbayExport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet(); 
  const data = sourceSheet.getDataRange().getValues();
  
  let exportSheet = ss.getSheetByName("eBay Export");
  if (!exportSheet) {
    exportSheet = ss.insertSheet("eBay Export");
  } else {
    exportSheet.clear(); 
  }
  
  const infoRow1 = ["#INFO", "Version=0.0.2", "Template= eBay-draft-listings-template_US", "", "", "", "", "", "", "", ""];
  const infoRow2 = ["#INFO Action and Category ID are required fields. 1) Set Action to Draft 2) Please find the category ID for your listings here: https://pages.ebay.com/sellerinformation/news/categorychanges.html", "", "", "", "", "", "", "", "", "", ""];
  const infoRow3 = ["#INFO After you've successfully uploaded your draft from the Seller Hub Reports tab, complete your drafts to active listings here: https://www.ebay.com/sh/lst/drafts", "", "", "", "", "", "", "", "", "", ""];
  const infoRow4 = ["#INFO", "", "", "", "", "", "", "", "", "", ""];
  
  const ebayHeaders = [
    "Action(SiteID=US|Country=US|Currency=USD|Version=1193|CC=UTF-8)", 
    "Custom label (SKU)", "Category ID", "Title", "UPC", "Price", "Quantity", 
    "Item photo URL", "Condition ID", "Description", "Format"
  ];
  
  const exportData = [infoRow1, infoRow2, infoRow3, infoRow4, ebayHeaders];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || row[0] === "") continue; 
    
    const skuOrImages = row[0] || "";
    const conditionNotes = row[2] || "";
    const pricingCol = row[3] || "";
    const poshEbayCol = row[5] || "";
    const listedCol = row[10];
    const priceListedAt = row[11];
    
    if (listedCol && listedCol.toString().toLowerCase() === "yes") continue;
    
    let title = "";
    let description = "";
    
    const titleMatch = poshEbayCol.match(/Title:\s*(.*?)(?=\n+Description:|$)/i);
    const descMatch = poshEbayCol.match(/Description:\s*(.*)/is);
    
    if (titleMatch && descMatch) {
      title = titleMatch[1].trim();
      description = descMatch[1].trim();
    } else {
      const lines = poshEbayCol.split('\n');
      title = lines[0].replace(/Title:/i, '').trim();
      description = lines.slice(1).join('\n').replace(/Description:/i, '').trim();
    }
    
    description = description.replace(/(?:\n|^)*(?:Tags|eBay Category|Category):[\s\S]*$/i, '').trim();
    const finalDescription = `${description.replace(/\n/g, '<br>')}<br><br><b>Condition Notes:</b><br>${conditionNotes}`;
    
    let finalPrice = "";
    if (priceListedAt && priceListedAt !== "") {
      finalPrice = priceListedAt.toString().replace(/[^0-9.]/g, ''); 
    } else {
      const priceMatch = pricingCol.match(/([0-9]+\.[0-9]{2})/);
      if (priceMatch) finalPrice = priceMatch[1];
    }
    
    const ebayRow = [
      "Draft", skuOrImages.substring(0, 50), "365", title.substring(0, 80), "", 
      finalPrice, "1", "", "3000", finalDescription, "FixedPrice"
    ];
    
    exportData.push(ebayRow);
  }
  
  exportSheet.getRange(1, 1, exportData.length, ebayHeaders.length).setValues(exportData);
  SpreadsheetApp.getUi().alert("Success! Your eBay Drafts have been generated using the new Seller Hub format.");
}
