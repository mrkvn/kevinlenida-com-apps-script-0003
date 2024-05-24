function loadSidebar() {
    const html = HtmlService.createTemplateFromFile("index").evaluate();
    SpreadsheetApp.getUi().showSidebar(html);
}

function createMenu() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("SecretLangTo");
    menu.addItem("Shh", "loadSidebar");
    menu.addToUi();
}

function onOpen() {
    createMenu();
}

function getAllColumnHeaders() {
    // Get the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Get all sheets in the spreadsheet
    const sheets = spreadsheet.getSheets();

    // Initialize an array to hold the headers from all sheets
    let allHeaders = [];

    // Loop through each sheet
    for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        // Get the range of the first row (headers) in the sheet
        const headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
        // Get the values of the headers
        const headers = headersRange.getValues()[0];

        // Push the headers and sheet name to the allHeaders array
        allHeaders.push(...headers);
    }

    allHeaders = [...new Set(allHeaders)];
    return allHeaders;
}

function getPngFilesInFolder(folderId) {
    // Get the folder by ID
    const folder = DriveApp.getFolderById(folderId);

    // Get all files in the folder
    const files = folder.getFiles();
    let pngFiles = [];

    // Loop through the files and log their names and IDs
    while (files.hasNext()) {
        const pngFile = files.next();

        // Check if the file is a PNG
        if (pngFile.getMimeType() === MimeType.PNG) {
            pngFiles.push(pngFile);
        }
    }

    return pngFiles;
}

function getBase64String(file) {
    const blob = file.getBlob();

    // Convert the blob to a base64 string
    const base64String = Utilities.base64Encode(blob.getBytes());

    return base64String;
}

function fetchApi(apiKey, fields, base64ImageString) {
    const url =
        "https://generativelanguage.googleapis.com/v1/models/gemini-pro-vision:generateContent";

    // Create the request data
    const requestData = {
        contents: [
            {
                role: "user",
                parts: [
                    {
                        inlineData: {
                            mimeType: "image/png",
                            data: base64ImageString, // base64 encoded image
                        },
                    },
                    {
                        text: `
  Extract the following FIELDS from the invoice image:
  
  FIELDS:
  ${fields}
  
  - It is important that your response must be in json format.
  - If there aren't any available data for the mentioned column above, leave it blank in your json response.
  - If the invoice has multiple items (which would be, most of the time), put the item fields in an \`items\` array field.
  - Do not remove or add any other fields other than the fields mentioned (aside of course, from the \`items\` array field, when applicable)
  - Do not forget to respond in json format.
  `,
                    },
                ],
            },
        ],
        generationConfig: {
            temperature: 0,
            topK: 1,
        },
    };

    // Set the request options
    const options = {
        method: "post",
        contentType: "application/json",
        headers: {
            "x-goog-api-key": apiKey,
        },
        payload: JSON.stringify(requestData),
    };

    try {
        // Make the POST request
        const response = UrlFetchApp.fetch(url, options);

        // Parse the JSON response
        const data = JSON.parse(response.getContentText());

        // Log the response data
        // Logger.log(data);
        // Logger.log(data.candidates[0].content.parts[0].text);
        return data.candidates[0].content.parts[0].text;
    } catch (error) {
        Logger.log("Error: " + error.message);
    }
}

// Function to check if a string contains valid JSON
function isJsonString(str) {
    try {
        JSON.parse(str);
    } catch (e) {
        return false;
    }
    return true;
}

function appendData(sheetName, data) {
    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    // Get all data in the sheet
    const sheetData = sheet.getDataRange().getValues();

    // Get the header row
    const headers = sheetData[0];

    // Create a new row array with the same length as headers
    const newRow = new Array(headers.length).fill("");

    // Populate the new row with data from the object
    for (let key in data) {
        const columnIndex = headers.indexOf(key);
        if (columnIndex !== -1) {
            newRow[columnIndex] = data[key];
        }
    }

    // Append the new row to the sheet
    sheet.appendRow(newRow);
}

function main(apiKey, folderId) {
    const columnHeaders = getAllColumnHeaders();
    let fields = "";
    for (let i = 0; i < columnHeaders.length; i++) {
        if (i === columnHeaders.length - 1) {
            fields += columnHeaders[i];
            break;
        }
        fields += columnHeaders[i] + ", ";
    }

    const invoiceImages = getPngFilesInFolder(folderId);
    let base64Strings = [];
    if (invoiceImages.length > 0) {
        for (let i = 0; i < invoiceImages.length; i++) {
            base64Strings.push(getBase64String(invoiceImages[i]));
        }
    }
    // const result = []
    base64Strings.forEach((base64String) => {
        // process image
        const jsonString = fetchApi(apiKey, fields, base64String);

        let extractedObject;
        // process string result
        if (
            isJsonString(
                jsonString
                    .replace(/```json/g, "")
                    .replace(/```/g, "")
                    .trim()
            )
        ) {
            const extractedJson = jsonString
                .replace(/```json/g, "")
                .replace(/```/g, "")
                .trim();
            extractedObject = JSON.parse(extractedJson);
        } else {
            console.error("The provided string is not valid JSON.");
            return;
        }

        const { items, ...headerObj } = extractedObject;

        // append header to sheet
        appendData("header", headerObj);

        // Add invoice_number to each item
        const updatedItems = items.map((item) => {
            return {
                ...item,
                invoice_number: extractedObject.invoice_number,
                invoice_date: extractedObject.invoice_date,
            };
        });

        // append items to sheet
        updatedItems.forEach((item) => {
            appendData("items", item);
        });

        // process another image after number of seconds
        Utilities.sleep(3000);
    });
}
