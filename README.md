# Vintage-Listing-Assistant-from-Photos
A Google Apps Script pipeline for vintage resellers and curators. This script automatically processes folders of inventory photos, groups them into individual items, and uses the Gemini 1.5 Flash Vision API to write SEO-optimized, multi-platform listings (WooCommerce, Poshmark, eBay) directly into a Google Sheet.

**License:** Public Domain (The Unlicense) - Feel free to use, modify, and share!

## 🌟 How It Works (The "Screenshot Delimiter" Method)

When taking photos of inventory, resellers often take multiple photos of a single item (front, back, tag, flaws) before moving to the next. 

This script solves the problem of grouping those photos automatically. **Between photographing each physical item, you simply take a screenshot of your phone's home screen.** 
1. The script sorts all images in your Google Drive folder chronologically.
2. It looks at each image and asks Gemini a lightning-fast boolean question: *"Is this a smartphone screenshot?"*
3. Every time it hits a screenshot, it knows the previous batch of photos belongs to one single item. 
4. It sends that grouped batch of images to Gemini to write a comprehensive listing and logs it as a single row in your Google Sheet.

## 🚀 Setup Instructions

### 1. Google Drive & Sheets Setup
1. Create a specific folder in Google Drive for your inventory photos and grab the **Folder ID** from the URL (the long string of letters/numbers after `folders/`).
2. Create a Google Sheet. 
3. **Important:** Name the specific tab/sheet you want to use (e.g., `Listings`). 
4. Add the following **8 Headings** to Row 1 of that tab (A through H):
   * `Image Files`
   * `Era & Material`
   * `Pricing Strategy`
   * `WooCommerce Desc`
   * `Poshmark/eBay Desc`
   * `Seller Insights`
   * `Lot Grouping`

### 2. Google AI Studio Setup
1. Go to [Google AI Studio](https://aistudio.google.com/) and sign in.
2. Click **Get API key** and generate a free API key.

### 3. Apps Script Setup
1. Open your Google Sheet, click **Extensions > Apps Script**.
2. Delete any code there and paste the contents of `Code.gs` from this repository.
3. Update the Configuration Variables at the very top of the script (see below).
4. Click the Save (floppy disk) icon, select `processStampsBatches` from the dropdown at the top, and click **Run**. (You will need to grant Google permissions the first time you run it).

## 🛠️ Where to Customize the Code

At the very top of `Code.gs`, you MUST fill in your specific details:

javascript
const GEMINI_API_KEY = 'YOUR_API_KEY_HERE'; 
const FOLDER_ID = 'YOUR_FOLDER_ID_HERE'; 

## Targeting Your Specific Tab:
By default, the script targets a tab named "Listings" to prevent it from accidentally overwriting data if you have the wrong tab open. If you name your tab something else, update this line in the `processStampsBatches` function:


\`\`\`javascript
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listings"); 
\`\`\`


## Customizing the Prompt / Brand Voice:**
Scroll down to the `promptText` variable inside the `generateListingForBatch` function. You can rewrite the prompt to fit your specific brand voice, change the length requirements, or alter the JSON output structure. If you change the requested JSON keys, ensure you also update the `responseSchema` requirement block right below it!
