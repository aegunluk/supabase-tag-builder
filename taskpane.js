/**
 * Supabase Tag Builder Word Taskpane Add-in
 * Wilson & Associates Template System
 */

// SUPABASE CONFIGURATION (Pre-filled from project settings)
const SUPABASE_URL = "https://bgytxhaboxdghbfwqivh.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_2fsFLIM_EVlNDYi_n9C_Yg_6ZwprUQH";

// Embedded database schema metadata (statically mirrored from src/types/database.ts)
const apiSchema = {
  "AppUser": {
    "properties": {
      "id": {}, "email": {}, "full_name": {}, "role": {}, "phone": {}, "avatar_url": {}, "is_active": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Applicant": {
    "properties": {
      "id": {}, "user_id": {}, "given_names": {}, "family_name": {}, "date_of_birth": {}, "gender": {}, "nationality": {}, "country_of_birth": {}, "passport_number": {}, "passport_expiry": {}, "passport_country": {}, "email": {}, "phone": {}, "address_line1": {}, "address_line2": {}, "suburb": {}, "state": {}, "postcode": {}, "country": {}, "current_visa_subclass": {}, "current_visa_expiry": {}, "english_proficiency_level": {}, "skills_assessment_body": {}, "skills_assessment_number": {}, "skills_assessment_expiry": {}, "eoi_submitted": {}, "eoi_score": {}, "tax_file_number": {}, "anzsco_code": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Dependent": {
    "properties": {
      "id": {}, "applicant_id": {}, "relationship": {}, "given_names": {}, "family_name": {}, "date_of_birth": {}, "gender": {}, "nationality": {}, "passport_number": {}, "passport_expiry": {}, "is_included_in_application": {}, "visa_grant_number": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Employer": {
    "properties": {
      "id": {}, "user_id": {}, "business_name": {}, "trading_name": {}, "abn": {}, "acn": {}, "business_type": {}, "industry_anzsic_code": {}, "contact_person_name": {}, "contact_email": {}, "contact_phone": {}, "address_line1": {}, "suburb": {}, "state": {}, "postcode": {}, "website": {}, "standard_business_sponsorship_number": {}, "sbs_approval_date": {}, "sbs_expiry_date": {}, "labour_agreement_type": {}, "labour_agreement_number": {}, "turnover_aud": {}, "number_of_employees": {}, "is_active": {}, "notes": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Case": {
    "properties": {
      "id": {}, "case_number": {}, "case_type": {}, "applicant_id": {}, "employer_id": {}, "assigned_agent_id": {}, "description": {}, "status": {}, "priority": {}, "visa_subclass_target": {}, "stream": {}, "lodge_date": {}, "target_lodge_date": {}, "decision_date": {}, "visa_grant_date": {}, "visa_expiry_date": {}, "visa_grant_number": {}, "immi_account_id": {}, "internal_notes": {}, "client_visible_notes": {}, "is_archived": {}, "created_at": {}, "updated_at": {}
    }
  },
  "VisaApplication": {
    "properties": {
      "id": {}, "case_id": {}, "visa_subclass": {}, "application_reference_number": {}, "tran_id": {}, "lodgement_date": {}, "bridging_visa_subclass": {}, "bridging_visa_expiry": {}, "health_examination_required": {}, "health_exam_date": {}, "health_ref_number": {}, "character_clearance_required": {}, "police_check_countries": {}, "application_fee_aud": {}, "second_instalment_fee_aud": {}, "outcome": {}, "refusal_reason": {}, "aar_lodged": {}, "aar_number": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Document": {
    "properties": {
      "id": {}, "name": {}, "document_type": {}, "related_to_type": {}, "related_to_id": {}, "uploaded_by_user_id": {}, "file_url": {}, "file_size_bytes": {}, "file_mime_type": {}, "is_sensitive": {}, "is_client_visible": {}, "expiry_date": {}, "notes": {}, "verified_by_agent": {}, "verified_at": {}, "esign_status": {}, "esign_envelope_id": {}, "esign_signer_email": {}, "esign_url": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Template": {
    "properties": {
      "id": {}, "name": {}, "description": {}, "template_type": {}, "category": {}, "file_url": {}, "html_body": {}, "subject": {}, "field_mappings": {}, "available_data_sources": {}, "is_active": {}, "created_by": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Task": {
    "properties": {
      "id": {}, "case_id": {}, "assigned_to": {}, "title": {}, "description": {}, "due_date": {}, "status": {}, "priority": {}, "is_client_action_required": {}, "completed_at": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Invoice": {
    "properties": {
      "id": {}, "case_id": {}, "invoice_number": {}, "issued_to_type": {}, "issued_to_id": {}, "subtotal_aud": {}, "gst_aud": {}, "total_aud": {}, "status": {}, "due_date": {}, "paid_date": {}, "payment_method": {}, "notes": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Lead": {
    "properties": {
      "id": {}, "name": {}, "email": {}, "phone": {}, "message": {}, "lead_type": {}, "status": {}, "invited_user_id": {}, "invited_at": {}, "internal_notes": {}, "created_at": {}, "updated_at": {}
    }
  },
  "Signature": {
    "properties": {
      "signature:Applicant": {},
      "signature:Sponsor": {},
      "signature:Agent": {},
      "date:Applicant": {},
      "date:Sponsor": {},
      "date:Agent": {}
    }
  }
};

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
    // Sort tables alphabetically
    allTables = Object.keys(apiSchema).sort((a, b) => a.localeCompare(b));

    if (allTables.length === 0) {
      throw new Error("No tables found in schema metadata.");
    }

    // Enable inputs
    searchFilter.disabled = false;
    tableSelect.disabled = false;

    // Populate tables
    populateTables(allTables);
    updateStatus("connected", "Connected to Supabase schema");

  } catch (error) {
    console.error("Database connection error:", error);
    updateStatus("error", "Error loading schema");
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
    let label = col;
    if (col === "signature:Applicant") label = "Applicant Signature";
    else if (col === "signature:Sponsor") label = "Sponsor Signature";
    else if (col === "signature:Agent") label = "Agent Signature";
    else if (col === "date:Applicant") label = "Applicant Date Signed";
    else if (col === "date:Sponsor") label = "Sponsor Date Signed";
    else if (col === "date:Agent") label = "Agent Date Signed";

    optionsHtml += `<option value="${col}">${label}</option>`;
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
    const tagText = table === "Signature" ? `{{${col}}}` : `{{${table}.${col}}}`;
    previewBox.innerHTML = tagText;
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
  
  const tagText = table === "Signature" ? `{{${col}}}` : `{{${table}.${col}}}`;

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
