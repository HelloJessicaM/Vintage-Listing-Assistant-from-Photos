const GEMINI_API_KEY = '####'; // Replace with your actual key
const FOLDER_ID = '####';  //Replace with your folder id

function processStampsBatches() {
  const startTime = Date.now();
  const MAX_RUN_TIME = 4.5 * 60 * 1000; // 4.5 minutes (Stops safely before Google's 6 min limit)

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("listings"); // Replace if your tab is named differently
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  // Create a "Processed" folder if it doesn't exist yet
  let processedFolder;
  const folders = folder.getFoldersByName("Processed");
  if (folders.hasNext()) {
    processedFolder = folders.next();
  } else {
    processedFolder = folder.createFolder("Processed");
  }

  const filesIterator = folder.getFiles();
  let filesArray = [];

  // 1. Gather all files and sort them alphabetically
  while (filesIterator.hasNext()) {
    filesArray.push(filesIterator.next());
  }
  filesArray.sort((a, b) => a.getName().localeCompare(b.getName()));

  let row = sheet.getLastRow() + 1;
  if (row < 2) row = 2;

  let currentItemImages = [];
  let currentItemFileNames = [];
  let currentItemFiles = []; // We track the actual file objects so we can move them later

  // 2. Loop through the sorted files
  for (let i = 0; i < filesArray.length; i++) {
    
    // SAFETY CLOCK: Check if we are running out of time
    if (Date.now() - startTime > MAX_RUN_TIME) {
      Logger.log("Nearing 5 minute limit. Stopping gracefully.");
      SpreadsheetApp.getUi().alert("Time limit reached! The script processed as many as it could. Please click the button to run it again and it will pick up right where it left off.");
      return; 
    }

    const file = filesArray[i];
    const base64Image = Utilities.base64Encode(file.getBlob().getBytes());
    const mimeType = file.getMimeType();

    // 3. Ask Gemini if this specific image is your screenshot delimiter
    const isScreenshot = checkIfScreenshot(base64Image, mimeType);

    if (isScreenshot) {
      // It's a screenshot! Process the batch
      if (currentItemImages.length > 0) {
        Logger.log("Processing batch for: " + currentItemFileNames.join(", "));
        
        // Run the AI. We check if it was 'successful' before moving files
        const success = generateListingForBatch(currentItemImages, currentItemFileNames, sheet, row);
        
        if (success) {
          row++;
          
          // MOVE FILES: Get them out of the main folder so they aren't processed again
          for (let f of currentItemFiles) {
            f.moveTo(processedFolder);
          }
          file.moveTo(processedFolder); // Move the screenshot too
        }
        
        // Clear the arrays for the next item
        currentItemImages = [];
        currentItemFileNames = [];
        currentItemFiles = [];
      } else {
         // If there are back-to-back screenshots or empty batches, just move the screenshot out of the way
         file.moveTo(processedFolder);
      }
    } else {
      // It's a regular photo. Add it to the batch
      currentItemImages.push({
        inlineData: { mimeType: mimeType, data: base64Image }
      });
      currentItemFileNames.push(file.getName());
      currentItemFiles.push(file);
    }
  }

  // 4. Safety catch for the final batch (if you forgot a screenshot at the very end)
  if (currentItemImages.length > 0) {
    const success = generateListingForBatch(currentItemImages, currentItemFileNames, sheet, row);
    if (success) {
      for (let f of currentItemFiles) {
        f.moveTo(processedFolder);
      }
    }
  }
  
  SpreadsheetApp.getUi().alert("Hooray! All images in the folder have been processed.");
}

// Lightweight helper function to detect your phone screenshots
function checkIfScreenshot(base64Image, mimeType) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${GEMINI_API_KEY}`;
  
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
    const textResponse = json.candidates[0].content.parts[0].text.trim().toUpperCase();
    return textResponse.includes("TRUE");
  } catch (e) {
    Logger.log("Screenshot Check Error: " + e);
    return false;
  }
}

// Main listing generator modified to accept an array of images simultaneously
function generateListingForBatch(imageParts, fileNames, sheet, row) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${GEMINI_API_KEY}`;
  
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
  
  // SMART SAMPLING: If there are more than 4 photos, pick 4 spread evenly across the batch so Gemini sees the front AND back!
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
          "era_identification",
          "condition_notes",
          "pricing",
          "woocommerce", 
          "poshmark_ebay", 
          "depop",
          "rationale",
          "power_seller_tip",
          "lot_grouping"
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
    
    if (json.error) {
      Logger.log("API Error on batch: " + json.error.message);
      return false; // Tells the main script it failed so it DOESN'T move the files
    }
    
    const textResponse = json.candidates[0].content.parts[0].text;
    const geminiData = JSON.parse(textResponse);
    
    const joinedNames = fileNames.join(", ");
    
    sheet.getRange(row, 1, 1, 10).setValues([[
      joinedNames,
      geminiData.era_identification,
      geminiData.condition_notes,
      geminiData.pricing,
      geminiData.woocommerce,
      geminiData.poshmark_ebay,
      geminiData.depop,
      geminiData.rationale,
      geminiData.power_seller_tip,
      geminiData.lot_grouping
    ]]);
    
    return true; // Tells the main script it succeeded! Move the files!
    
  } catch (e) {
    Logger.log("Listing Generation Error: " + e);
    return false;
  }
}
