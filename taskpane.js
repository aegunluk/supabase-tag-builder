/**
 * Supabase Tag Builder Word Taskpane Add-in
 * Wilson & Associates Template System
 */

// SUPABASE CONFIGURATION (Pre-filled from project settings)
const SUPABASE_URL = "https://bgytxhaboxdghbfwqivh.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_2fsFLIM_EVlNDYi_n9C_Yg_6ZwprUQH";

// Global cache for loaded OpenAPI definitions
let apiSchema = null;
let allTables = [];
let currentColumns = [];

// Initialize Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "complete" || document.readyState === "interactive") {
      initApp();
    } else {
      document.addEventListener("DOMContentLoaded", () => {
        initApp();
      });
    }
  }
});

/**
 * Main application initializer
 */
async function initApp() {
  const statusBadge = document.getElementById("connectionStatus");
  const statusText = document.getElementById("statusText");
  const tableSelect = document.getElementById("tableSelect");
  const columnSelect = document.getElementById("columnSelect");
  const insertBtn = document.getElementById("insertBtn");
  const searchFilter = document.getElementById("searchFilter");
  const searchClear = document.getElementById("searchClear");

  // Setup Event Listeners
  tableSelect.addEventListener("change", handleTableChange);
  columnSelect.addEventListener("change", handleColumnChange);
  insertBtn.addEventListener("click", insertTagIntoDocument);
  searchFilter.addEventListener("input", handleSearchInput);
  searchClear.addEventListener("click", clearSearch);

  try {
    // Set Connecting status
    updateStatus("connecting", "Connecting to Supabase...");

    // Fetch the PostgREST OpenAPI spec directly using the anon key.
    // This gives us the complete metadata schema of all exposed tables and columns.
    const response = await fetch(`${SUPABASE_URL}/rest/v1/`, {
      headers: {
        "apikey": SUPABASE_ANON_KEY,
        "Authorization": `Bearer ${SUPABASE_ANON_KEY}`
      }
    });

    if (!response.ok) {
      throw new Error(`Connection failed: ${response.status} ${response.statusText}`);
    }

    const openApiData = await response.json();
    apiSchema = openApiData.definitions || {};
    
    // Sort tables alphabetically
    allTables = Object.keys(apiSchema).sort((a, b) => a.localeCompare(b));

    if (allTables.length === 0) {
      throw new Error("No public tables exposed in the public schema.");
    }

    // Enable inputs
    searchFilter.disabled = false;
    tableSelect.disabled = false;

    // Populate tables
    populateTables(allTables);
    updateStatus("connected", "Connected to Supabase");

  } catch (error) {
    console.error("Database connection error:", error);
    updateStatus("error", "Connection Error");
    tableSelect.innerHTML = `<option value="">Error: ${error.message}</option>`;
  }
}

/**
 * Update the UI status badge
 */
function updateStatus(state, text) {
  const badge = document.getElementById("connectionStatus");
  const statusText = document.getElementById("statusText");
  
  badge.className = `status-badge ${state}`;
  statusText.textContent = text;
}

/**
 * Populate table select dropdown
 */
function populateTables(tables) {
  const tableSelect = document.getElementById("tableSelect");
  let optionsHtml = '<option value="">-- Select a Table / View --</option>';
  
  tables.forEach(table => {
    optionsHtml += `<option value="${table}">${table}</option>`;
  });
  
  tableSelect.innerHTML = optionsHtml;
}

/**
 * Handle table dropdown change
 */
function handleTableChange() {
  const tableSelect = document.getElementById("tableSelect");
  const columnSelect = document.getElementById("columnSelect");
  const insertBtn = document.getElementById("insertBtn");
  const selectedTable = tableSelect.value;

  // Reset fields
  columnSelect.innerHTML = '<option value="">Choose a table first...</option>';
  columnSelect.disabled = true;
  insertBtn.disabled = true;
  updatePreview(null, null);

  if (!selectedTable || !apiSchema || !apiSchema[selectedTable]) {
    currentColumns = [];
    return;
  }

  // Retrieve columns for this table
  const properties = apiSchema[selectedTable].properties || {};
  currentColumns = Object.keys(properties).sort((a, b) => a.localeCompare(b));

  populateColumns(currentColumns);
  columnSelect.disabled = false;
}

/**
 * Populate column dropdown list
 */
function populateColumns(columns) {
  const columnSelect = document.getElementById("columnSelect");
  let optionsHtml = '<option value="">-- Select a Column (Field) --</option>';
  
  columns.forEach(col => {
    optionsHtml += `<option value="${col}">${col}</option>`;
  });
  
  columnSelect.innerHTML = optionsHtml;
}

/**
 * Handle column selection change
 */
function handleColumnChange() {
  const tableSelect = document.getElementById("tableSelect");
  const columnSelect = document.getElementById("columnSelect");
  const insertBtn = document.getElementById("insertBtn");

  const table = tableSelect.value;
  const col = columnSelect.value;

  if (table && col) {
    updatePreview(table, col);
    insertBtn.disabled = false;
  } else {
    updatePreview(null, null);
    insertBtn.disabled = true;
  }
}

/**
 * Update the tag preview box
 */
function updatePreview(table, col) {
  const previewBox = document.getElementById("tagPreview");
  
  if (table && col) {
    previewBox.innerHTML = `{{${table}.${col}}}`;
    previewBox.classList.remove("preview-placeholder");
  } else {
    previewBox.innerHTML = '<span class="preview-placeholder">No field selected</span>';
  }
}

/**
 * Perform client-side quick filtering on dropdown options
 */
function handleSearchInput() {
  const filterVal = document.getElementById("searchFilter").value.toLowerCase();
  const searchClear = document.getElementById("searchClear");
  const tableSelect = document.getElementById("tableSelect");
  const columnSelect = document.getElementById("columnSelect");

  if (filterVal.length > 0) {
    searchClear.style.display = "block";
  } else {
    searchClear.style.display = "none";
  }

  // If we have tables loaded
  if (allTables.length > 0) {
    // 1. Filter tables list
    const filteredTables = allTables.filter(t => t.toLowerCase().includes(filterVal));
    
    // Remember current selected table value
    const currentSelectedTable = tableSelect.value;
    populateTables(filteredTables);
    
    // Attempt to restore selection if it's still in the filtered list
    if (filteredTables.includes(currentSelectedTable)) {
      tableSelect.value = currentSelectedTable;
    }

    // 2. If a table is active, also filter columns
    if (tableSelect.value && currentColumns.length > 0) {
      const filteredCols = currentColumns.filter(c => c.toLowerCase().includes(filterVal));
      const currentSelectedCol = columnSelect.value;
      populateColumns(filteredCols);
      
      if (filteredCols.includes(currentSelectedCol)) {
        columnSelect.value = currentSelectedCol;
      }
    }
  }
}

/**
 * Clear the filter input
 */
function clearSearch() {
  const searchFilter = document.getElementById("searchFilter");
  searchFilter.value = "";
  handleSearchInput();
}

/**
 * Run the Word.run API to insert the placeholder tag into the document
 */
async function insertTagIntoDocument() {
  const tableSelect = document.getElementById("tableSelect");
  const columnSelect = document.getElementById("columnSelect");
  
  const table = tableSelect.value;
  const col = columnSelect.value;
  
  if (!table || !col) return;
  
  const tagText = `{{${table}.${col}}}`;

  try {
    // Use OfficeJS Word API to run a batch against the active document context
    await Word.run(async (context) => {
      // Get the current cursor selection range in the active document
      const selection = context.document.getSelection();
      
      // Insert the text tag directly at the selection
      // Word.InsertLocation.replace replaces any currently highlighted text, 
      // or collapses to cursor position and inserts if no text is selected.
      selection.insertText(tagText, Word.InsertLocation.replace);
      
      // Synchronize the document state
      await context.sync();
      
      // Show success toast
      showToast(`Inserted: ${tagText}`);
    });
  } catch (error) {
    console.error("Office.js insertion error:", error);
    alert(`Could not insert tag: ${error.message}`);
  }
}

/**
 * Visual success toast
 */
function showToast(message) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.add("show");
  
  setTimeout(() => {
    toast.classList.remove("show");
  }, 2500);
}
