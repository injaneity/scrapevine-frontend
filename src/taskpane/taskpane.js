/* global console, document, Excel, Office, fetch */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initializeDropdown();
    document.querySelector(".search-icon").onclick = sendData;
  }
});

async function sendData() {
  const keyword = document.getElementById('additionalInput').value;
  const website = document.getElementById('searchInput').value;
  const sheetOption = document.querySelector('input[name="sheet-option"]:checked').value;

  updateCodeBlock("Searching...", "green");

  const data = {
    siteUrl: website,
    tags: keyword
  };

  try {
    console.log('Sending data to proxy...');
    const response = await fetch('https://localhost:3000/proxy', { // Proxy endpoint
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(data)
    });

    if (!response.ok) {
      throw new Error('Network response was not ok.');
    }

    const result = await response.json();
    const responseId = result.responseId;
    updateCodeBlock('Data sent successfully. Response ID: ' + responseId, 'green');
    console.log('Data sent successfully:', result);

    // Start polling the backend with the responseId
    pollBackend(responseId, sheetOption);
  } catch (error) {
    updateCodeBlock('Failed to send data', 'red');
    console.error('Failed to send data:', error);
  }
}

function initializeDropdown() {
  const input = document.getElementById('searchInput');
  const dropdown = document.getElementById('dropdownList');

  input.addEventListener('input', function() {
    const searchValue = this.value.toLowerCase();
    const dropdownItems = dropdown.querySelectorAll('div');
    dropdownItems.forEach(function(item) {
      if (item.textContent.toLowerCase().includes(searchValue)) {
        item.style.display = "";
      } else {
        item.style.display = "none";
      }
    });
  });

  input.addEventListener('focus', function() {
    dropdown.style.display = 'block';
  });

  input.addEventListener('blur', function() {
    setTimeout(() => {
      dropdown.style.display = 'none';
    }, 200); // delay to allow click event to process
  });
}

function selectOption(value) {
  const input = document.getElementById('searchInput');
  input.value = value;
  const dropdownItems = document.querySelectorAll('#dropdownList div');
  dropdownItems.forEach(item => {
    item.style.display = "none";
  });
}

function pollBackend(responseId, sheetOption) {
  const proxyUrl = 'https://localhost:3000/reply?responseId=' + encodeURIComponent(responseId);

  const intervalId = setInterval(async () => {
    try {
      console.log('Polling backend for responseId: ' + responseId);
      const response = await fetch(proxyUrl, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error('Network response was not ok.');
      }

      const responseData = await response.json();

      if (responseData.status && responseData.status === "processing") {
        console.log('Data still processing. Will check again later.');
      } else {
        console.log('Data received, processing.', responseData);
        clearInterval(intervalId); // Stop polling
        updateCodeBlock('Processing complete. Data received.', 'green');
        // Handle the received data (e.g., update the UI or process the data)
        processReceivedData(responseData, sheetOption);
      }
    } catch (error) {
      console.error('Error polling backend:', error);
      updateCodeBlock('Error polling backend: ' + error.toString(), 'red');
      clearInterval(intervalId); // Stop polling on error
    }
  }, 10000); // Poll every 10 seconds
}

function processReceivedData(data, sheetOption) {
  // Extract price summaries and trend analysis from the first item in the JSON array
  const priceSummary = data[0];
  const trendAnalysis = priceSummary.Trend;

  // Update the rectangle container with average, highest, and lowest prices
  document.querySelector('.rectangle-container .rectangle-with-label:nth-child(1) .rectangle span').textContent = priceSummary['Lowest Price'];
  document.querySelector('.rectangle-container .rectangle-with-label:nth-child(2) .rectangle span').textContent = priceSummary['Average Price'];
  document.querySelector('.rectangle-container .rectangle-with-label:nth-child(3) .rectangle span').textContent = priceSummary['Highest Price'];

  // Update the analysis container with the trend analysis
  const analysisContainer = document.querySelector('.analysis-container p');
  analysisContainer.textContent = trendAnalysis;

  console.log('Received data:', data);

  // Extract headers and rows from the data
  const { headers, rows } = extractDataFromJson(data);

  // Write the data to the spreadsheet
  writeToSpreadsheet(headers, rows, sheetOption);
}

function extractDataFromJson(jsonData) {
  const headers = jsonData[1].headers; // Extract headers from the second object in the array
  const rows = jsonData.slice(2); // Extract rows starting from the third object in the array

  return { headers, rows };
}

async function writeToSpreadsheet(headers, rows, sheetOption) {
  await Excel.run(async (context) => {
    let sheet;
    if (sheetOption === 'new') {
      let sheetName = 'New Sheet';
      let sheetExists = true;
      let counter = 1;

      // Check if the sheet name already exists and create a unique name
      while (sheetExists) {
        try {
          sheet = context.workbook.worksheets.getItem(sheetName);
          await sheet.load('name');
          await context.sync();
          sheetName = `New Sheet ${counter++}`;
        } catch (error) {
          sheetExists = false;
        }
      }

      sheet = context.workbook.worksheets.add(sheetName);
    } else {
      sheet = context.workbook.worksheets.getActiveWorksheet();
    }

    // Write headers to row 1
    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;

    // Determine the starting row for new data
    const usedRange = sheet.getUsedRange();
    let startRow = 1;
    if (usedRange) {
      await usedRange.load('rowCount');
      await context.sync();
      startRow = usedRange.rowCount > 1 ? usedRange.rowCount + 1 : 2;
    }

    // Write rows starting from row 2
    const dataRange = sheet.getRangeByIndexes(startRow - 1, 0, rows.length, headers.length);
    const values = rows.map(row => headers.map(header => row[header] || ""));
    dataRange.values = values;

    // Ensure data is not bold
    dataRange.format.font.bold = false;

    await context.sync();
  }).catch((error) => {
    console.error(error);
  });
}

function updateCodeBlock(message, color = "black") {
  var codeBlock = document.getElementById("code-block");
  codeBlock.style.color = color;
  codeBlock.textContent = message;
}
