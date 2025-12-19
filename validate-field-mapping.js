import https from "https";
import axios from "axios";
import YAML from "yaml";
import fs from "fs";
import path from "path";


// ============================================================
// LOAD CONFIGURATIONS
// ============================================================

const cfg = YAML.parse(fs.readFileSync("config.yaml", "utf-8"));
const fieldMapping = YAML.parse(fs.readFileSync("field-mapping.yaml", "utf-8"));

// Redmine Configuration
const REDMINE_BASE = cfg.redmine.base_url.replace(/\/+$/, "");
const REDMINE_HEADERS = { "X-Redmine-API-Key": cfg.redmine.api_key };

const httpsAgent = new https.Agent({ rejectUnauthorized: false });

const redmineHttp = axios.create({ 
  baseURL: REDMINE_BASE, 
  headers: REDMINE_HEADERS, 
  timeout: 60_000,
  httpsAgent
});

// Azure DevOps Configuration
const ADO_ORG = cfg.azure_devops?.organization;
const ADO_PROJECT = cfg.azure_devops?.project;
const ADO_PAT = cfg.azure_devops?.pat;

if (!ADO_ORG || !ADO_PROJECT || !ADO_PAT) {
  console.error("Error: Missing Azure DevOps configuration in config.yaml");
  process.exit(1);
}

const ADO_BASE_URL = `https://dev.azure.com/${ADO_ORG}`;
const ADO_AUTH = Buffer.from(`:${ADO_PAT}`).toString('base64');

const adoHttp = axios.create({
  baseURL: ADO_BASE_URL,
  headers: {
    'Authorization': `Basic ${ADO_AUTH}`,
    'Content-Type': 'application/json'
  },
  timeout: 60_000
});


// ============================================================
// REDMINE API FUNCTIONS
// ============================================================

async function getRedmineCustomFields() {
  try {
    const { data } = await redmineHttp.get("/custom_fields.json");
    return data.custom_fields || [];
  } catch (err) {
    if (err.response?.status === 403) {
      console.warn("⚠ Warning: Cannot fetch Redmine custom fields (admin rights required)");
      return [];
    }
    throw err;
  }
}

async function getRedmineTrackers() {
  try {
    const { data } = await redmineHttp.get("/trackers.json");
    return data.trackers || [];
  } catch (err) {
    console.error("Error fetching Redmine trackers:", err.message);
    return [];
  }
}

async function getRedmineStatuses() {
  try {
    const { data } = await redmineHttp.get("/issue_statuses.json");
    return data.issue_statuses || [];
  } catch (err) {
    console.error("Error fetching Redmine statuses:", err.message);
    return [];
  }
}

async function getRedminePriorities() {
  try {
    const { data } = await redmineHttp.get("/enumerations/issue_priorities.json");
    return data.issue_priorities || [];
  } catch (err) {
    console.error("Error fetching Redmine priorities:", err.message);
    return [];
  }
}


// ============================================================
// AZURE DEVOPS API FUNCTIONS
// ============================================================

async function getADOFields() {
  try {
    const { data } = await adoHttp.get(`/_apis/wit/fields`, {
      params: { 'api-version': '7.1-preview.2' }
    });
    return data.value || [];
  } catch (err) {
    console.error("Error fetching ADO fields:", err.message);
    if (err.response) {
      console.error("Status:", err.response.status);
      console.error("Data:", err.response.data);
    }
    throw err;
  }
}

async function getADOWorkItemTypes() {
  try {
    const { data } = await adoHttp.get(`/${ADO_PROJECT}/_apis/wit/workitemtypes`, {
      params: { 'api-version': '7.1-preview.2' }
    });
    return data.value || [];
  } catch (err) {
    console.error("Error fetching ADO work item types:", err.message);
    return [];
  }
}


// ============================================================
// VALIDATION FUNCTIONS
// ============================================================

function validateWorkItemTypeMappings(redmineTrackers, adoWorkItemTypes) {
  console.log("\n=== WORK ITEM TYPE MAPPINGS ===\n");
  
  const mappings = fieldMapping.work_item_types || {};
  const adoTypeNames = adoWorkItemTypes.map(t => t.name);
  
  let valid = 0;
  let invalid = 0;
  
  for (const [redmineTracker, adoType] of Object.entries(mappings)) {
    const redmineExists = redmineTrackers.some(t => t.name === redmineTracker);
    const adoExists = adoTypeNames.includes(adoType);
    
    if (redmineExists && adoExists) {
      console.log(`✓ ${redmineTracker} → ${adoType}`);
      valid++;
    } else if (!redmineExists && adoExists) {
      console.log(`⚠ ${redmineTracker} → ${adoType} (Redmine tracker not found)`);
      invalid++;
    } else if (redmineExists && !adoExists) {
      console.log(`✗ ${redmineTracker} → ${adoType} (ADO work item type not found)`);
      invalid++;
    } else {
      console.log(`✗ ${redmineTracker} → ${adoType} (Both not found)`);
      invalid++;
    }
  }
  
  console.log(`\nValid: ${valid}, Issues: ${invalid}`);
  return { valid, invalid };
}

function validateFieldMappings(redmineCustomFields, adoFields) {
  console.log("\n=== STANDARD FIELD MAPPINGS ===\n");
  
  const mappings = fieldMapping.field_mappings || {};
  const adoFieldRefs = adoFields.map(f => f.referenceName);
  
  // Redmine standard fields (always available)
  const redmineStandardFields = [
    'id', 'subject', 'description', 'status', 'priority', 'assigned_to', 
    'author', 'category', 'version', 'start_date', 'due_date', 
    'created_on', 'updated_on', 'closed_on', 'done_ratio', 'estimated_hours',
    'project', 'tracker', 'parent'
  ];
  
  let valid = 0;
  let invalid = 0;
  
  for (const [redmineField, adoField] of Object.entries(mappings)) {
    const redmineExists = redmineStandardFields.includes(redmineField);
    const adoExists = adoFieldRefs.includes(adoField) || adoField.includes('Custom.');
    
    if (redmineExists && adoExists) {
      console.log(`✓ ${redmineField} → ${adoField}`);
      valid++;
    } else if (!redmineExists) {
      console.log(`⚠ ${redmineField} → ${adoField} (May be custom field in Redmine)`);
      valid++;
    } else if (!adoExists) {
      console.log(`✗ ${redmineField} → ${adoField} (ADO field not found - may need to create)`);
      invalid++;
    }
  }
  
  console.log(`\nValid: ${valid}, Needs Review: ${invalid}`);
  return { valid, invalid };
}

function validateCustomFieldMappings(redmineCustomFields, adoFields) {
  console.log("\n=== CUSTOM FIELD MAPPINGS ===\n");
  
  const mappings = fieldMapping.custom_field_mappings || {};
  const adoFieldRefs = adoFields.map(f => f.referenceName);
  const redmineCustomFieldNames = redmineCustomFields.map(f => f.name);
  
  let valid = 0;
  let invalid = 0;
  
  if (Object.keys(mappings).length === 0) {
    console.log("No custom field mappings configured.");
    return { valid: 0, invalid: 0 };
  }
  
  for (const [redmineField, adoField] of Object.entries(mappings)) {
    const redmineExists = redmineCustomFieldNames.includes(redmineField);
    const adoExists = adoFieldRefs.includes(adoField) || adoField.startsWith('Custom.');
    
    if (redmineExists && adoExists) {
      console.log(`✓ ${redmineField} → ${adoField}`);
      valid++;
    } else if (!redmineExists) {
      console.log(`✗ ${redmineField} → ${adoField} (Redmine custom field not found)`);
      invalid++;
    } else if (!adoExists) {
      console.log(`⚠ ${redmineField} → ${adoField} (ADO custom field not found - may need to create)`);
      invalid++;
    }
  }
  
  console.log(`\nValid: ${valid}, Needs Review: ${invalid}`);
  return { valid, invalid };
}

function validateStatusMappings(redmineStatuses, adoWorkItemTypes) {
  console.log("\n=== STATUS MAPPINGS ===\n");
  
  const statusMappings = fieldMapping.status_mappings || {};
  const redmineStatusNames = redmineStatuses.map(s => s.name);
  
  let valid = 0;
  let invalid = 0;
  
  for (const [workItemType, statusMap] of Object.entries(statusMappings)) {
    if (workItemType === 'default') continue;
    
    console.log(`\n${workItemType}:`);
    for (const [redmineStatus, adoStatus] of Object.entries(statusMap)) {
      const redmineExists = redmineStatusNames.includes(redmineStatus);
      
      if (redmineExists) {
        console.log(`  ✓ ${redmineStatus} → ${adoStatus}`);
        valid++;
      } else {
        console.log(`  ⚠ ${redmineStatus} → ${adoStatus} (Redmine status not found)`);
        invalid++;
      }
    }
  }
  
  console.log(`\nValid: ${valid}, Needs Review: ${invalid}`);
  return { valid, invalid };
}

function validatePriorityMappings(redminePriorities) {
  console.log("\n=== PRIORITY MAPPINGS ===\n");
  
  const priorityMappings = fieldMapping.priority_mappings || {};
  const redminePriorityNames = redminePriorities.map(p => p.name);
  
  let valid = 0;
  let invalid = 0;
  
  for (const [redminePriority, adoPriority] of Object.entries(priorityMappings)) {
    if (redminePriority === 'default') continue;
    
    const redmineExists = redminePriorityNames.includes(redminePriority);
    const adoValid = [1, 2, 3, 4].includes(adoPriority);
    
    if (redmineExists && adoValid) {
      console.log(`✓ ${redminePriority} → ${adoPriority}`);
      valid++;
    } else if (!redmineExists) {
      console.log(`⚠ ${redminePriority} → ${adoPriority} (Redmine priority not found)`);
      invalid++;
    } else if (!adoValid) {
      console.log(`✗ ${redminePriority} → ${adoPriority} (Invalid ADO priority - must be 1-4)`);
      invalid++;
    }
  }
  
  console.log(`\nValid: ${valid}, Needs Review: ${invalid}`);
  return { valid, invalid };
}


// ============================================================
// MAIN
// ============================================================

async function main() {
  console.log("=== FIELD MAPPING VALIDATION ===");
  console.log(`\nRedmine: ${REDMINE_BASE}`);
  console.log(`Azure DevOps: ${ADO_ORG}/${ADO_PROJECT}\n`);

  try {
    console.log("Fetching Redmine data...");
    const [redmineCustomFields, redmineTrackers, redmineStatuses, redminePriorities] = await Promise.all([
      getRedmineCustomFields(),
      getRedmineTrackers(),
      getRedmineStatuses(),
      getRedminePriorities()
    ]);
    
    console.log(`  - Custom Fields: ${redmineCustomFields.length}`);
    console.log(`  - Trackers: ${redmineTrackers.length}`);
    console.log(`  - Statuses: ${redmineStatuses.length}`);
    console.log(`  - Priorities: ${redminePriorities.length}`);

    console.log("\nFetching Azure DevOps data...");
    const [adoFields, adoWorkItemTypes] = await Promise.all([
      getADOFields(),
      getADOWorkItemTypes()
    ]);
    
    console.log(`  - Fields: ${adoFields.length}`);
    console.log(`  - Work Item Types: ${adoWorkItemTypes.length}`);

    // Validate mappings
    const results = {
      workItemTypes: validateWorkItemTypeMappings(redmineTrackers, adoWorkItemTypes),
      fields: validateFieldMappings(redmineCustomFields, adoFields),
      customFields: validateCustomFieldMappings(redmineCustomFields, adoFields),
      statuses: validateStatusMappings(redmineStatuses, adoWorkItemTypes),
      priorities: validatePriorityMappings(redminePriorities)
    };

    // Summary
    console.log("\n" + "=".repeat(60));
    console.log("VALIDATION SUMMARY");
    console.log("=".repeat(60));
    
    const totalValid = Object.values(results).reduce((sum, r) => sum + r.valid, 0);
    const totalInvalid = Object.values(results).reduce((sum, r) => sum + r.invalid, 0);
    
    console.log(`\nTotal Valid Mappings: ${totalValid}`);
    console.log(`Total Issues Found: ${totalInvalid}`);
    
    if (totalInvalid === 0) {
      console.log("\n✓ All mappings validated successfully!");
    } else {
      console.log(`\n⚠ Found ${totalInvalid} mapping(s) that need review.`);
      console.log("Please check the output above and update field-mapping.yaml as needed.");
    }

    // List available ADO custom fields
    console.log("\n" + "=".repeat(60));
    console.log("AVAILABLE ADO CUSTOM FIELDS");
    console.log("=".repeat(60));
    const customFields = adoFields.filter(f => f.referenceName.startsWith('Custom.'));
    if (customFields.length > 0) {
      customFields.forEach(f => {
        console.log(`  - ${f.name} (${f.referenceName}) - Type: ${f.type}`);
      });
    } else {
      console.log("  No custom fields found. You may need to create custom fields in ADO.");
    }

    // List Redmine custom fields
    console.log("\n" + "=".repeat(60));
    console.log("AVAILABLE REDMINE CUSTOM FIELDS");
    console.log("=".repeat(60));
    if (redmineCustomFields.length > 0) {
      redmineCustomFields.forEach(f => {
        console.log(`  - ${f.name} (ID: ${f.id}) - Format: ${f.field_format}`);
      });
    } else {
      console.log("  Unable to fetch (requires admin rights) or no custom fields exist.");
    }

  } catch (err) {
    console.error("\nError:", err.message);
    if (err.response) {
      console.error("Status:", err.response.status);
    }
    process.exit(1);
  }
}

main();
