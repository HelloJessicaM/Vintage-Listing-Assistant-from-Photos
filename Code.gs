const GEMINI_API_KEY = 'XXXXXXX'; // Replace with your actual key
const FOLDER_ID = 'XXXXX';  //Replace with your folder id

function processStampsBatches() {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test-listings"); //Replace with the name of your tab
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const filesIterator = folder.getFiles();
  let filesArray = [];

  // 1. Gather all files and sort them alphabetically to maintain chronological order
  while (filesIterator.hasNext()) {
    filesArray.push(filesIterator.next());
  }
  filesArray.sort((a, b) => a.getName().localeCompare(b.getName()));

  let row = sheet.getLastRow() + 1; // Start adding at the bottom of your sheet
  if (row < 2) row = 2; // Assuming row 1 has headers

  let currentItemImages = [];
  let currentItemFileNames = [];

  // 2. Loop through the sorted files
  for (let i = 0; i < filesArray.length; i++) {
    const file = filesArray[i];
    const base64Image = Utilities.base64Encode(file.getBlob().getBytes());
    const mimeType = file.getMimeType();

    // 3. Ask Gemini if this specific image is your screenshot delimiter
    const isScreenshot = checkIfScreenshot(base64Image, mimeType);

    if (isScreenshot) {
      // It's a screenshot! Process the accumulated batch of item photos (if any exist)
      if (currentItemImages.length > 0) {
        Logger.log("Processing batch for: " + currentItemFileNames.join(", "));
        generateListingForBatch(currentItemImages, currentItemFileNames, sheet, row);
        row++;
        // Clear the arrays for the next item
        currentItemImages = [];
        currentItemFileNames = [];
      }
    } else {
      // It's a regular photo. Add it to the current item's batch
      currentItemImages.push({
        inlineData: { mimeType: mimeType, data: base64Image }
      });
      currentItemFileNames.push(file.getName());
    }
  }

  // 4. Safety catch: Process the final batch if the folder didn't end with a screenshot
  if (currentItemImages.length > 0) {
    generateListingForBatch(currentItemImages, currentItemFileNames, sheet, row);
  }
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
    
    CRITICAL FORMATTING RULE: 
    Output plain text ONLY. Do NOT use HTML tags. Use standard keyboard line breaks (new lines) for paragraph spacing and dashes (-) for bullet points. Use single quotes (') instead of double quotes (") inside your descriptions to keep the text clean.
    
    Output the data based on this exact structure:
    - era_identification: Era, Material, Handling
    - pricing_strategy: Suggested List Price, Minimum Acceptable Offer, Rationale
    - woocommerce: SEO Title, Short Description, Long Description (At least 800-1600 words), Measurements.
    - poshmark_ebay: Title, Description (<1500 chars), Tags, eBay Category
    - depop: Description (<1000 characters, casual and trendy tone for vintage lovers), Hashtags (no # symbols)
    - seller_insights: Condition Notes, Power Seller Tip
    - lot_grouping: Suggest a thematic lot to bundle this in. If high value, state 'Sell Individually'.
  `;

  // Combine the single text prompt with the array of images for this item
  const partsArray = [{ text: promptText }];
  for (let i = 0; i < imageParts.length; i++) {
    partsArray.push(imageParts[i]);
  }

  const payload = {
    contents: [{ parts: partsArray }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: {
        type: "OBJECT",
        properties: {
          era_identification: { type: "STRING" },
          pricing_strategy: { type: "STRING" },
          woocommerce: { type: "STRING" },
          poshmark_ebay: { type: "STRING" },
          depop: { type: "STRING" },
          seller_insights: { type: "STRING" },
          lot_grouping: { type: "STRING" }
        },
        required: [
          "era_identification", 
          "pricing_strategy", 
          "woocommerce", 
          "poshmark_ebay", 
          "depop",
          "seller_insights", 
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
      return;
    }
    
    const textResponse = json.candidates[0].content.parts[0].text;
    const geminiData = JSON.parse(textResponse);
    
    // Join the filenames with a comma so your Sheet lists all images associated with this single row
    const joinedNames = fileNames.join(", ");
    
    sheet.getRange(row, 1, 1, 7).setValues([[
      joinedNames,
      geminiData.era_identification,
      geminiData.pricing_strategy,
      geminiData.woocommerce,
      geminiData.poshmark_ebay,
      geminiData.seller_insights,
      geminiData.lot_grouping
    ]]);
    
  } catch (e) {
    Logger.log("Listing Generation Error: " + e);
  }
}
