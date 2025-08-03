// Excel file data
let itemsData = [];
let selectedItems = [];

// Debounce helper for search
let searchDebounceTimeout;

// DOM elements
const excelFileInput = document.getElementById("excel-file");
const loadDefaultBtn = document.getElementById("load-default-btn");
const searchInput = document.getElementById("search-input");
const searchBtn = document.getElementById("search-btn");
const fromIdInput = document.getElementById("from-id");
const toIdInput = document.getElementById("to-id");
const printRangeBtn = document.getElementById("print-range-btn");
const itemGrid = document.getElementById("item-grid");
const selectedItemsContainer = document.getElementById(
  "selected-items-container"
);
const itemCountSpan = document.querySelector(".item-count");
const printBtn = document.getElementById("print-btn");
const resetBtn = document.getElementById("reset-btn");
const priceEditModal = document.getElementById("price-edit-modal");
const closePriceModal = document.getElementById("close-price-modal");
const cancelEditBtn = document.getElementById("cancel-edit-btn");
const confirmPrintBtn = document.getElementById("confirm-print-btn");
const priceEditContainer = document.getElementById("price-edit-container");
const printPreviewModal = document.getElementById("print-preview");
const closeModal = document.querySelector(".close");
const actualPrintBtn = document.getElementById("actual-print-btn");
const printContent = document.getElementById("print-content");

// Event listeners
excelFileInput.addEventListener("change", handleFileUpload);
// loadDefaultBtn.addEventListener("click", loadDefaultCSV);
searchInput.addEventListener("input", debounceSearchItems); // Real-time search
printRangeBtn.addEventListener("click", printRangeOfItems);
printBtn.addEventListener("click", showPriceEditModal);
resetBtn.addEventListener("click", resetSelectedItems);
closePriceModal.addEventListener("click", closePriceEditModalHandler);
cancelEditBtn.addEventListener("click", closePriceEditModalHandler);
confirmPrintBtn.addEventListener("click", confirmAndPrint);
closeModal.addEventListener(
  "click",
  () => (printPreviewModal.style.display = "none")
);
actualPrintBtn.addEventListener("click", () => window.print());

// Initialize the app
document.addEventListener("DOMContentLoaded", function () {
  // Load default CSV on page load
  loadDefaultCSV();
});

// Handle Excel file upload
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const fileExtension = file.name.split(".").pop().toLowerCase();

  if (fileExtension === "csv") {
    handleCSVFile(file);
  } else if (fileExtension === "xlsx" || fileExtension === "xls") {
    handleExcelFile(file);
  } else {
    alert("Please select a valid CSV or Excel file");
  }
}

// Handle CSV file
function handleCSVFile(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const csv = e.target.result;
    itemsData = parseCSV(csv);
    displayItems(itemsData);
  };
  reader.readAsText(file);
}

// Handle Excel file
function handleExcelFile(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    // Assuming first sheet is the one we want
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

    // Convert to JSON
    itemsData = XLSX.utils.sheet_to_json(firstSheet);

    // Display all items initially
    displayItems(itemsData);
  };
  reader.readAsArrayBuffer(file);
}

// Parse CSV data
function parseCSV(csv) {
  const lines = csv.split("\n");
  const headers = [];
  const data = [];

  if (lines.length === 0) return data;

  // Parse headers
  const headerLine = lines[0];
  const headerParts = headerLine.split(",");
  for (let part of headerParts) {
    headers.push(part.trim().replace(/"/g, ""));
  }

  // Parse data rows
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line === "") continue;

    const values = line.split(",");
    const row = {};

    headers.forEach((header, index) => {
      let value = values[index] || "";
      // Remove quotes and trim
      value = value.trim().replace(/"/g, "");
      row[header] = value;
    });

    // Only add rows that have at least an ITEM_ID
    if (row.ITEM_ID && row.ITEM_ID.trim() !== "") {
      data.push(row);
    }
  }

  return data;
}

// Load default CSV file
function loadDefaultCSV() {
  fetch("itemmaster.csv")
    .then((response) => {
      if (!response.ok) {
        throw new Error("Could not load itemmaster.csv file");
      }
      return response.text();
    })
    .then((csv) => {
      itemsData = parseCSV(csv);
      displayItems(itemsData);
      // Only show alert when button is clicked, not on page load
      if (event && event.type === "click") {
        alert("Default ItemMaster.csv loaded successfully!");
      }
    })
    .catch((error) => {
      console.error("Error loading default CSV:", error);
      // Only show alert when button is clicked, not on page load
      if (event && event.type === "click") {
        alert(
          "Error loading default ItemMaster.csv file. Please upload a file manually."
        );
      }
    });
}

// Display items in the grid
function displayItems(items) {
  itemGrid.innerHTML = "";

  items.forEach((item, index) => {
    const itemCard = document.createElement("div");
    itemCard.className = "item-card";

    // Calculate the actual index in the full dataset
    const actualIndex = itemsData.indexOf(item) + 1;

    itemCard.innerHTML = `
            <div class="item-index">#${actualIndex}</div>
            <strong>${item.SHORT_NAME || "N/A"}</strong><br>
            <small>ID: ${item.ITEM_ID || "N/A"}</small><br>
            <small>MRP: ₹${item.MRP || "0"}</small>
        `;

    itemCard.addEventListener("click", () => selectItem(item));
    itemGrid.appendChild(itemCard);
  });
}

function debounceSearchItems() {
  clearTimeout(searchDebounceTimeout);
  searchDebounceTimeout = setTimeout(searchItems, 200);
}

function searchItems() {
  const searchTerm = searchInput.value.toLowerCase();
  if (!searchTerm) {
    displayItems(itemsData);
    return;
  }

  const filteredItems = itemsData.filter((item) => {
    return (
      (item.SHORT_NAME && item.SHORT_NAME.toLowerCase().includes(searchTerm)) ||
      (item.ITEM_ID &&
        item.ITEM_ID.toString().toLowerCase().includes(searchTerm)) ||
      (item.MAIN_EANCODE &&
        item.MAIN_EANCODE.toString().toLowerCase().includes(searchTerm))
    );
  });

  displayItems(filteredItems);
}

// Print range of items by product index
function printRangeOfItems() {
  const fromIndex = parseInt(fromIdInput.value.trim());
  const toIndex = parseInt(toIdInput.value.trim());

  if (!fromIndex || !toIndex || isNaN(fromIndex) || isNaN(toIndex)) {
    alert("Please enter valid index numbers for both From Index and To Index");
    return;
  }

  if (fromIndex < 1 || toIndex < 1) {
    alert("Index numbers must be 1 or greater");
    return;
  }

  if (fromIndex > toIndex) {
    alert("From Index must be less than or equal to To Index");
    return;
  }

  if (itemsData.length === 0) {
    alert("No items data loaded. Please load a CSV/Excel file first.");
    return;
  }

  // Check if index numbers are within the available data range
  const maxIndex = itemsData.length;
  if (fromIndex > maxIndex || toIndex > maxIndex) {
    alert(
      `Index numbers must be between 1 and ${maxIndex} (total available items)`
    );
    return;
  }

  // Get items in the specified index range (convert to 0-based index)
  const rangeItems = itemsData.slice(fromIndex - 1, toIndex);

  if (rangeItems.length === 0) {
    alert(`No items found in the range from index ${fromIndex} to ${toIndex}`);
    return;
  }

  if (rangeItems.length > 50) {
    const confirmPrint = confirm(
      `This will print ${rangeItems.length} items (index ${fromIndex} to ${toIndex}). This might take a while. Continue?`
    );
    if (!confirmPrint) return;
  }

  // Print the range items directly
  printRangeDirectly(rangeItems);

  alert(
    `Range print completed for ${rangeItems.length} items (index ${fromIndex} to ${toIndex})`
  );
}

// Fallback print function without external template
function printWithInlineTemplate(items) {
  console.log("Using inline template for printing");

  let productBoxesHtml = "";

  items.forEach((item, itemIndex) => {
    const nameClass =
      item.SHORT_NAME && item.SHORT_NAME.length >= 30
        ? "product-name-sm"
        : "product-name";

    const discount =
      parseFloat(item.MRP || 0) - parseFloat(item.SALE_PRICE || 0);

    // Check if discount has decimal values to adjust font size
    const hasDecimal = discount % 1 !== 0;
    const discountFontSize = hasDecimal ? "6.5rem" : "8rem";

    productBoxesHtml += `
        <div class="product-box">
            <div class="product-discount" style="display: flex; flex-direction: column; align-items: center; justify-content: center;">
              <div style="display: flex; align-items: flex-end; justify-content: center;position:relative;">
                <span style="font-size:3rem; margin-bottom:1.5rem;">₹</span>
                <span style="font-size:${discountFontSize}; font-weight:bold; margin:0 10px;">${
      discount || "0"
    }</span>
                <span style="font-size:8rem; font-weight:bold; margin:0 10px;"></span>
                 <span style="font-size:1.8rem; margin-top:-1rem;position:absolute;bottom:1.5rem;right:-1.5rem;">OFF</span>
              </div>
            </div>
             <div class="${nameClass}">${item.SHORT_NAME || "N/A"}</div>
            <div style="width:100%;display:flex;align-items:center;justify-content:center;margin-top:5px;">
              <span class="mrp-price">MRP ₹${item.MRP || "0"}</span>
              <span style="border-left:2px solid #000;height:1.5em;margin:0 10px;"></span>
              <span class="sale-price">Mauli Mart Price ₹${
                item.SALE_PRICE || "0"
              }</span>
            </div>
        </div>
    `;
  });

  const fullHtml = `
    <!DOCTYPE html>
    <html>
    <head>
        <title>Product Labels</title>
        <style>
            @page {
                size: A4;
                margin: 0;
            }
            body {
                margin: 0;
                padding: 2mm;
                box-sizing: border-box;
                display: grid;
                grid-template-columns: 9.4cm 9.4cm;
                grid-template-rows: 5.6cm 5.6cm 5.6cm 5.6cm;
                gap: 2mm;
                justify-content: center;
                align-content: center;
                font-family: Arial, sans-serif;
            }
            .product-box {
                width: 9.2cm;
                height: 5.6cm;
                border: 5px solid #000;
                padding: 2mm;
                box-sizing: border-box;
                display: flex;
                flex-direction: column;
                justify-content: space-between;
                page-break-inside: avoid;
            }
            .product-name {
                font-size: 13pt;
                font-weight: bold;
                text-align: center;
                flex-grow: 1;
                display: flex;
                align-items: center;
                justify-content: center;
            }
            .product-name-sm {
                font-size: 10pt;
                font-weight: bold;
                text-align: center;
                flex-grow: 1;
                display: flex;
                align-items: center;
                justify-content: center;
            }
            .mrp-price {
                text-decoration: line-through;
                color: #000;
            }
            .sale-price {
                font-weight: bold;
                color: #000;
            }
        </style>
    </head>
    <body>
        ${productBoxesHtml}
    </body>
    </html>
  `;

  // Create a hidden iframe for printing
  const iframe = document.createElement("iframe");
  iframe.style.position = "absolute";
  iframe.style.left = "-9999px";
  document.body.appendChild(iframe);

  // Write the print content to the iframe
  const printDocument = iframe.contentWindow.document;
  printDocument.open();
  printDocument.write(fullHtml);
  printDocument.close();

  // Wait for content to load then print
  setTimeout(() => {
    console.log("Triggering print dialog with inline template");
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
    document.body.removeChild(iframe);
  }, 500);
}

// Print range directly without price editing modal
function printRangeDirectly(itemsToPrint = selectedItems) {
  console.log("Starting printRangeDirectly with", itemsToPrint.length, "items");

  // Load the print template
  fetch("print-template.html")
    .then((response) => {
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      return response.text();
    })
    .then((templateHtml) => {
      console.log("Template loaded successfully");

      // Generate product boxes HTML for all range items
      let productBoxesHtml = "";

      itemsToPrint.forEach((item, itemIndex) => {
        console.log(`Processing item ${itemIndex + 1}:`, item);

        // Check if SHORT_NAME is 30 or more characters
        const nameClass =
          item.SHORT_NAME && item.SHORT_NAME.length >= 30
            ? "product-name-sm"
            : "product-name";

        const discount =
          item.MRP && item.SALE_PRICE
            ? parseFloat(item.MRP) - parseFloat(item.SALE_PRICE)
            : 0;

        // Check if discount has decimal values to adjust font size
        const hasDecimal = discount % 1 !== 0;
        const discountFontSize = hasDecimal ? "6.5rem" : "8rem";

        productBoxesHtml += `
            <div class="product-box">
                <div class="product-discount" style="display: flex; flex-direction: column; align-items: center; justify-content: center;">
                  <div style="display: flex; align-items: flex-end; justify-content: center;position:relative;">
                    <span style="font-size:3rem; margin-bottom:1.5rem;">₹</span>
                    <span style="font-size:${discountFontSize}; font-weight:bold; margin:0 10px;">${
          discount || "0"
        }</span>
                    <span style="font-size:8rem; font-weight:bold; margin:0 10px;"></span>
                     <span style="font-size:1.8rem; margin-top:-1rem;position:absolute;bottom:1.5rem;right:-1.5rem;">OFF</span>
                  </div>
                </div>
                 <div class="${nameClass}">${item.SHORT_NAME || "N/A"}</div>
                <div style="width:100%;display:flex;align-items:center;justify-content:center;margin-top:5px;">
                  <span class="mrp-price">MRP ₹${item.MRP || "0"}</span>
                  <span style="border-left:2px solid #000;height:1.5em;margin:0 10px;"></span>
                  <span class="sale-price">Mauli Mart Price ₹${
                    item.SALE_PRICE || "0"
                  }</span>
                </div>
            </div>
        `;
      });

      console.log("Generated HTML length:", productBoxesHtml.length);

      // Insert product boxes into template
      const finalHtml = templateHtml.replace(
        "<!-- Product boxes will be dynamically inserted here -->",
        productBoxesHtml
      );

      console.log("Final HTML prepared, creating iframe");

      // Create a hidden iframe for printing
      const iframe = document.createElement("iframe");
      iframe.style.position = "absolute";
      iframe.style.left = "-9999px";
      document.body.appendChild(iframe);

      // Write the print content to the iframe
      const printDocument = iframe.contentWindow.document;
      printDocument.open();
      printDocument.write(finalHtml);
      printDocument.close();

      // Wait for content to load then print
      setTimeout(() => {
        console.log("Triggering print dialog");
        iframe.contentWindow.focus();
        iframe.contentWindow.print();
        document.body.removeChild(iframe);
      }, 500); // Increased timeout to 500ms
    })
    .catch((error) => {
      console.error("Error in printRangeDirectly:", error);
      console.log("Falling back to inline template");
      printWithInlineTemplate(itemsToPrint);
    });
}

// Select an item
function selectItem(item) {
  // Check if already selected
  if (selectedItems.some((selected) => selected.ITEM_ID === item.ITEM_ID)) {
    alert("This item is already selected");
    return;
  }

  // Limit to 8 items
  if (selectedItems.length >= 10) {
    alert("You can select maximum 10 items");
    return;
  }

  selectedItems.push(item);
  updateSelectedItemsTable();

  // Enable print button if we have at least one item
  if (selectedItems.length > 0) {
    printBtn.disabled = false;
  }
}

// Update the selected items display
function updateSelectedItemsTable() {
  selectedItemsContainer.innerHTML = "";

  // Update item count
  itemCountSpan.textContent = `${selectedItems.length}/10`;

  selectedItems.forEach((item, index) => {
    const discount = (((item.MRP - item.SALE_PRICE) / item.MRP) * 100).toFixed(
      2
    );

    const itemDiv = document.createElement("div");
    itemDiv.className = "compact-item";
    itemDiv.innerHTML = `
      <div class="item-info">
        <div class="item-name">${item.SHORT_NAME || "N/A"}</div>
        <div class="item-details">
          ID: ${item.ITEM_ID || "N/A"} | MRP: ₹${item.MRP || "0"} | Sale: ₹${
      item.SALE_PRICE || "0"
    } | ${discount}% off
        </div>
      </div>
      <button class="remove-btn" data-index="${index}">×</button>
    `;

    selectedItemsContainer.appendChild(itemDiv);
  });

  // Add event listeners to remove buttons
  document.querySelectorAll(".remove-btn").forEach((btn) => {
    btn.addEventListener("click", function () {
      const index = parseInt(this.getAttribute("data-index"));
      selectedItems.splice(index, 1);
      updateSelectedItemsTable();

      // Disable print button if no items left
      if (selectedItems.length === 0) {
        printBtn.disabled = true;
      }
    });
  });
}

// Reset selected items
function resetSelectedItems() {
  if (selectedItems.length === 0) {
    alert("No items selected to reset");
    return;
  }

  const confirmReset = confirm(
    `Are you sure you want to remove all ${selectedItems.length} selected items?`
  );

  if (confirmReset) {
    selectedItems = [];
    updateSelectedItemsTable();

    // Disable print button
    printBtn.disabled = true;

    alert("All selected items have been cleared");
  }
}

// Show price edit modal before printing
function showPriceEditModal() {
  if (selectedItems.length === 0) {
    alert("Please select items first");
    return;
  }

  priceEditContainer.innerHTML = "";

  selectedItems.forEach((item, index) => {
    const itemDiv = document.createElement("div");
    itemDiv.className = "price-edit-item";

    const mrp = parseFloat(item.MRP) || 0;
    const salePrice = parseFloat(item.SALE_PRICE) || 0;
    const discount = mrp > 0 ? (((mrp - salePrice) / mrp) * 100).toFixed(1) : 0;

    itemDiv.innerHTML = `
      <div class="item-info">
        <div class="item-name">${item.SHORT_NAME || "N/A"}</div>
        <div class="item-id">ID: ${item.ITEM_ID || "N/A"}</div>
      </div>
      <div class="price-inputs">
        <div class="price-group">
          <label>MRP (₹)</label>
          <input type="number" 
                 class="mrp-input" 
                 data-index="${index}" 
                 value="${mrp}" 
                 min="0" 
                 step="0.01">
        </div>
        <div class="price-group">
          <label>Sale Price (₹)</label>
          <input type="number" 
                 class="sale-input" 
                 data-index="${index}" 
                 value="${salePrice}" 
                 min="0" 
                 step="0.01">
        </div>
        <div class="discount-display" data-index="${index}">
          ${discount}% OFF
        </div>
      </div>
    `;

    priceEditContainer.appendChild(itemDiv);
  });

  // Add event listeners for price inputs to update discount
  document.querySelectorAll(".mrp-input, .sale-input").forEach((input) => {
    input.addEventListener("input", updateDiscount);
  });

  priceEditModal.style.display = "block";
}

// Update discount display when prices change
function updateDiscount(e) {
  const index = e.target.getAttribute("data-index");
  const mrpInput = document.querySelector(`.mrp-input[data-index="${index}"]`);
  const saleInput = document.querySelector(
    `.sale-input[data-index="${index}"]`
  );
  const discountDisplay = document.querySelector(
    `.discount-display[data-index="${index}"]`
  );

  const mrp = parseFloat(mrpInput.value) || 0;
  const salePrice = parseFloat(saleInput.value) || 0;

  if (mrp > 0 && salePrice <= mrp) {
    const discount = (((mrp - salePrice) / mrp) * 100).toFixed(1);
    discountDisplay.textContent = `${discount}% OFF`;
    discountDisplay.style.backgroundColor = "#28a745";
  } else {
    discountDisplay.textContent = "Invalid";
    discountDisplay.style.backgroundColor = "#dc3545";
  }
}

// Close price edit modal
function closePriceEditModalHandler() {
  priceEditModal.style.display = "none";
}

// Confirm prices and proceed to print
function confirmAndPrint() {
  // Update selected items with new prices
  document.querySelectorAll(".price-edit-item").forEach((itemDiv, index) => {
    const mrpInput = itemDiv.querySelector(".mrp-input");
    const saleInput = itemDiv.querySelector(".sale-input");

    const newMrp = parseFloat(mrpInput.value) || 0;
    const newSalePrice = parseFloat(saleInput.value) || 0;

    // Validate prices
    if (newMrp <= 0 || newSalePrice < 0 || newSalePrice > newMrp) {
      alert(`Invalid prices for item: ${selectedItems[index].SHORT_NAME}`);
      return;
    }

    // Update the selected item with new prices
    selectedItems[index].MRP = newMrp;
    selectedItems[index].SALE_PRICE = newSalePrice;
  });

  // Update the sidebar display with new prices
  updateSelectedItemsTable();

  // Close modal and proceed to print
  closePriceEditModalHandler();
  showPrintPreview();
}

// Show print preview with specific template
function showPrintPreview() {
  console.log("Starting showPrintPreview with", selectedItems.length, "items");

  // Load the print template
  fetch("print-template.html")
    .then((response) => {
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      return response.text();
    })
    .then((templateHtml) => {
      console.log("Template loaded successfully");

      // Generate product boxes HTML
      let productBoxesHtml = "";

      selectedItems.forEach((item, itemIndex) => {
        console.log(`Processing item ${itemIndex + 1}:`, item.SHORT_NAME);

        // Check if SHORT_NAME is 30 or more characters
        const nameClass =
          item.SHORT_NAME && item.SHORT_NAME.length >= 30
            ? "product-name-sm"
            : "product-name";

        const discount =
          parseFloat(item.MRP || 0) - parseFloat(item.SALE_PRICE || 0);

        // Check if discount has decimal values to adjust font size
        const hasDecimal = discount % 1 !== 0;
        const discountFontSize = hasDecimal ? "6.5rem" : "8rem";

        productBoxesHtml += `
            <div class="product-box">
                <div class="product-discount" style="display: flex; flex-direction: column; align-items: center; justify-content: center;">
                  <div style="display: flex; align-items: flex-end; justify-content: center;position:relative;">
                    <span style="font-size:3rem; margin-bottom:1.5rem;">₹</span>
                    <span style="font-size:${discountFontSize}; font-weight:bold; margin:0 10px;">${
          discount || "0"
        }</span>
                    <span style="font-size:8rem; font-weight:bold; margin:0 10px;"></span>
                     <span style="font-size:1.8rem; margin-top:-1rem;position:absolute;bottom:1.5rem;right:-1.5rem;">OFF</span>
                  </div>
                </div>
                 <div class="${nameClass}">${item.SHORT_NAME || "N/A"}</div>
                <div style="width:100%;display:flex;align-items:center;justify-content:center;margin-top:2px;margin-bottom:3px;">
                  <span class="mrp-price">MRP ₹${item.MRP || "0"}</span>
                  <span style="border-left:2px solid #000;height:1.5em;margin:0 10px;"></span>
                  <span class="sale-price">Mauli Mart Price ₹${
                    item.SALE_PRICE || "0"
                  }</span>
                </div>
            </div>
        `;
      });

      console.log("Generated HTML length:", productBoxesHtml.length);

      // Insert product boxes into template
      const finalHtml = templateHtml.replace(
        "<!-- Product boxes will be dynamically inserted here -->",
        productBoxesHtml
      );

      console.log("Final HTML prepared, creating iframe");

      // Create a hidden iframe for printing
      const iframe = document.createElement("iframe");
      iframe.style.position = "absolute";
      iframe.style.left = "-9999px";
      document.body.appendChild(iframe);

      // Write the print content to the iframe
      const printDocument = iframe.contentWindow.document;
      printDocument.open();
      printDocument.write(finalHtml);
      printDocument.close();

      // Wait for content to load then print
      setTimeout(() => {
        console.log("Triggering print dialog");
        iframe.contentWindow.focus();
        iframe.contentWindow.print();
        document.body.removeChild(iframe);
      }, 500); // Increased timeout to 500ms
    })
    .catch((error) => {
      console.error("Error in showPrintPreview:", error);
      console.log("Falling back to inline template");
      printWithInlineTemplate(selectedItems);
    });
}
